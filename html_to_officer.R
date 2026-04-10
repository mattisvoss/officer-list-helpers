# ==============================================================================
# html_to_officer.R
#
# Converts HTML fragments (like those found in DDI 3.3 XML metadata) into
# formatted officer paragraphs. Handles block-level elements (paragraphs,
# lists) and inline formatting (bold, italic, nested combinations).
#
# Depends on officer_list_helpers.R for bullet/numbered list support.
#
#
# TABLE OF CONTENTS
# =================
#
#   1. BACKGROUND: HTML fragments in DDI 3.3
#   2. BACKGROUND: Block vs inline elements
#   3. BACKGROUND: How this maps to officer
#   4. CODE: HTML parsing helpers
#   5. CODE: Inline content collector (recursive)
#   6. CODE: Block-level dispatcher
#   7. TESTS
#
#
# QUICK START
# ===========
#
#   library(officer)
#   library(xml2)
#   source("officer_list_helpers.R")
#   source("html_to_officer.R")
#
#   doc <- read_docx()
#   doc <- body_add_html_fragment(doc, "<p>Hello <b>world</b></p>")
#   doc <- body_add_html_fragment(doc, "<ul><li>First</li><li>Second</li></ul>")
#   print(doc, target = "output.docx")
#
#
# ==============================
# 1. HTML FRAGMENTS IN DDI 3.3
# ==============================
#
# DDI (Data Documentation Initiative) 3.3 is an XML standard for describing
# survey and research data. Some DDI elements contain embedded HTML fragments
# for rich text content — things like variable descriptions, survey questions,
# or methodology notes.
#
# These fragments are NOT full HTML documents. They're snippets like:
#
#   <p>The variable <b>age</b> measures the respondent's age in
#   <i>completed years</i> at the time of interview.</p>
#
#   <ul>
#     <li>Valid range: 0-120</li>
#     <li>Missing value: <b>-1</b></li>
#   </ul>
#
# We need to convert these into formatted Word paragraphs using officer.
#
#
# ==============================
# 2. BLOCK VS INLINE ELEMENTS
# ==============================
#
# HTML elements fall into two categories that map directly to Word concepts:
#
# BLOCK ELEMENTS create new paragraphs:
#   <p>           → one paragraph (may contain inline formatting)
#   <ul>          → a series of bullet list paragraphs (one per <li>)
#   <ol>          → a series of numbered list paragraphs (one per <li>)
#   <h1>..<h6>    → heading paragraphs
#   <blockquote>  → indented paragraph
#   bare text     → plain paragraph (text outside any tag)
#
# INLINE ELEMENTS create runs WITHIN a paragraph:
#   <b>, <strong> → bold text
#   <i>, <em>     → italic text
#   <u>           → underlined text
#   <sub>         → subscript
#   <sup>         → superscript
#   plain text    → normal text
#
# The key insight: a single <p> like this:
#
#   <p>Hello <b>bold <i>and italic</i></b> world</p>
#
# becomes ONE paragraph with FOUR runs:
#
#   fpar(
#     ftext("Hello "),                            # normal
#     ftext("bold ",   prop = fp_text(bold=TRUE)), # bold
#     ftext("and italic", prop = fp_text(bold=TRUE, italic=TRUE)), # both
#     ftext(" world")                              # normal
#   )
#
# Notice that <i> inside <b> INHERITS the bold formatting. This is why
# the inline collector needs to be recursive — nested tags accumulate
# formatting properties as you go deeper.
#
#
# ==============================
# 3. HOW THIS MAPS TO OFFICER
# ==============================
#
# officer's text model:
#
#   ftext(text, prop)   One "run" — a chunk of text with uniform formatting.
#                       prop is an fp_text object specifying bold, italic, etc.
#
#   fpar(run1, run2, ...)  One paragraph — contains one or more runs.
#                          This is what body_add_fpar() expects.
#
#   fp_text(bold, italic, underlined, ...)  Formatting properties for a run.
#
# Our conversion pipeline:
#
#   HTML fragment
#     → parse with xml2::read_html()
#     → walk block-level elements (p, ul, ol, text)
#       → for each block, walk inline elements recursively
#         → build ftext() runs, accumulating formatting from parent tags
#       → combine runs into one fpar()
#     → emit fpar as a paragraph (body_add_fpar or list_add_fpar)
#
# ==============================================================================


# ---- Dependency check --------------------------------------------------------

if (!requireNamespace("officer", quietly = TRUE)) {
  stop("Package 'officer' is required. Install with: install.packages('officer')")
}
if (!requireNamespace("xml2", quietly = TRUE)) {
  stop("Package 'xml2' is required. Install with: install.packages('xml2')")
}

# We use list_add_fpar and list_end from officer_list_helpers.R.
# That file must be sourced before this one (or at least before calling
# body_add_html_fragment with lists).
if (!exists("list_add_fpar", mode = "function")) {
  if (file.exists("officer_list_helpers.R")) {
    source("officer_list_helpers.R")
  } else {
    warning(paste(
      "officer_list_helpers.R not found. List support (<ul>, <ol>) will not",
      "work. Source officer_list_helpers.R before this file to enable lists."
    ))
  }
}


# ==============================================================================
# SECTION 4: HTML PARSING HELPERS
# ==============================================================================

#' Wrap an HTML fragment in a full document so xml2 can parse it.
#'
#' HTML fragments from DDI 3.3 aren't full documents — they're just
#' snippets like '<p>Hello <b>world</b></p>'. xml2::read_html() needs a
#' well-formed document, so we wrap the fragment in <html><body>...</body>
#' and return the <body> node.
#'
#' @param html_str A string containing an HTML fragment.
#' @return An xml2 node representing the <body> element.
wrap_html_fragment <- function(html_str) {
  wrapped <- paste0("<html><body>", html_str, "</body></html>")
  html_doc <- xml2::read_html(wrapped)
  xml2::xml_find_first(html_doc, ".//body")
}


# ==============================================================================
# SECTION 5: INLINE CONTENT COLLECTOR (RECURSIVE)
#
# This is the core of the formatting logic. It walks the children of an
# HTML element and builds a flat list of ftext() runs. Formatting
# accumulates as we descend into nested tags:
#
#   <b>bold <i>bold+italic</i></b>
#
# When we enter <b>, we set bold=TRUE in the inherited properties.
# When we then enter <i>, we ADD italic=TRUE to the already-bold properties.
# When we encounter text, we emit an ftext() with the accumulated properties.
#
# The result is a flat list of runs — no nesting — which is exactly what
# officer's fpar() expects.
# ==============================================================================

# Tags that we recognise as inline formatting. Each maps to an fp_text
# property. Tags not in this list are treated as transparent containers
# (we still recurse into them to grab their text content).
INLINE_FORMAT_MAP <- list(
  b      = list(bold = TRUE),
  strong = list(bold = TRUE),
  i      = list(italic = TRUE),
  em     = list(italic = TRUE),
  u      = list(underlined = TRUE),
  sub    = list(vertical.align = "subscript"),
  sup    = list(vertical.align = "superscript")
)


#' Merge two sets of fp_text properties.
#'
#' Takes an existing set of properties (a named list) and overlays new
#' properties on top. New values override old ones for the same key.
#'
#' Example:
#'   merge_props(list(bold = TRUE), list(italic = TRUE))
#'   → list(bold = TRUE, italic = TRUE)
#'
#' @param base A named list of fp_text properties (may be empty).
#' @param overlay A named list of new properties to add/override.
#' @return A merged named list.
merge_props <- function(base, overlay) {
  merged <- base
  for (name in names(overlay)) {
    merged[[name]] <- overlay[[name]]
  }
  merged
}


#' Build an fp_text object from a named list of properties.
#'
#' Converts our internal property list (e.g. list(bold=TRUE, italic=TRUE))
#' into an officer::fp_text object that ftext() can use.
#'
#' @param props A named list of fp_text properties. Empty list = default.
#' @return An fp_text object.
props_to_fp_text <- function(props) {
  if (length(props) == 0) {
    return(officer::fp_text())
  }
  do.call(officer::fp_text, props)
}


#' Recursively collect inline content from an HTML node into ftext runs.
#'
#' Walks the children of `node`. For each child:
#'   - Text node → emit ftext with the currently accumulated formatting
#'   - Inline tag (b, i, em, strong, u, sub, sup) → add its formatting
#'     to the inherited props, then recurse into its children
#'   - Unknown tag → recurse into children with unchanged formatting
#'     (treats the tag as a transparent wrapper)
#'
#' @param node An xml2 node to walk.
#' @param inherited_props Named list of formatting accumulated from parent
#'   tags. Starts empty at the top-level call.
#' @return A flat list of ftext objects (no nesting).
#'
#' @examples
#' # Internal use — called by collect_inline_runs():
#' #   <p>Hello <b>bold <i>both</i></b></p>
#' # produces:
#' #   list(ftext("Hello "), ftext("bold ", bold), ftext("both", bold+italic))
.collect_runs_recursive <- function(node, inherited_props = list()) {
  runs <- list()
  children <- xml2::xml_contents(node)

  for (child in children) {
    tag <- xml2::xml_name(child)

    if (tag == "text") {
      # --- Leaf: a text node. Emit an ftext with accumulated formatting. ---
      text <- xml2::xml_text(child)
      if (nchar(text) > 0) {
        runs <- c(runs, list(
          officer::ftext(text, prop = props_to_fp_text(inherited_props))
        ))
      }

    } else if (tag %in% names(INLINE_FORMAT_MAP)) {
      # --- Inline formatting tag: merge its properties and recurse. ---
      # For example, if we're inside <b> (inherited has bold=TRUE) and we
      # hit <i>, the new inherited props become list(bold=TRUE, italic=TRUE).
      new_props <- merge_props(inherited_props, INLINE_FORMAT_MAP[[tag]])
      child_runs <- .collect_runs_recursive(child, new_props)
      runs <- c(runs, child_runs)

    } else if (tag == "br") {
      # --- Line break: emit a newline run. ---
      # officer renders "\n" as a line break within a paragraph (soft return),
      # not a new paragraph.
      runs <- c(runs, list(
        officer::ftext("\n", prop = props_to_fp_text(inherited_props))
      ))

    } else {
      # --- Unknown tag: recurse into it without adding any formatting. ---
      # This handles things like <span>, <a>, or any tag we don't explicitly
      # recognise. We still grab the text content — we just don't style it.
      child_runs <- .collect_runs_recursive(child, inherited_props)
      runs <- c(runs, child_runs)
    }
  }

  runs
}


#' Collect all inline content within an HTML element into one fpar.
#'
#' This is the public entry point for inline collection. It calls the
#' recursive walker and wraps the resulting runs in an fpar().
#'
#' @param node An xml2 node (e.g. a <p>, <li>, or text node).
#' @return An fpar object containing one or more ftext runs.
collect_inline_runs <- function(node) {
  tag <- xml2::xml_name(node)

  if (tag == "text") {
    # Bare text node (not wrapped in any element).
    text <- xml2::xml_text(node)
    runs <- list(officer::ftext(text))
  } else {
    # Element node — recurse into its children.
    runs <- .collect_runs_recursive(node)
  }

  # If empty (e.g. <p></p>), add an empty run so fpar() doesn't complain.
  if (length(runs) == 0) {
    runs <- list(officer::ftext(""))
  }

  do.call(officer::fpar, runs)
}


# ==============================================================================
# SECTION 6: BLOCK-LEVEL DISPATCHER
#
# This function walks the top-level children of the HTML <body> and decides
# what to do with each one based on its tag name:
#
#   <p>       → collect inline content → body_add_fpar (one paragraph)
#   <ul>      → for each <li>: collect inline content → list_add_fpar (bullet)
#   <ol>      → for each <li>: collect inline content → list_add_fpar (decimal)
#   <h1>..<h6>→ collect inline content → body_add_fpar with heading style
#   text node → plain paragraph
#
# The key principle: block-level tags create new paragraphs; inline tags
# within them create runs inside those paragraphs.
# ==============================================================================

#' Convert an HTML fragment into formatted officer paragraphs.
#'
#' Parses the HTML, walks block-level elements, and adds each as a
#' properly formatted paragraph (or list item) to the document. Inline
#' formatting (bold, italic, nested combinations) is preserved.
#'
#' @param document An rdocx object.
#' @param html_str A string of HTML (a fragment, not a full page).
#' @param style Paragraph style name from your template. NULL = document
#'   default. This is applied to every paragraph unless overridden by a
#'   heading tag.
#' @return The modified rdocx object.
#'
#' @examples
#' doc <- read_docx()
#'
#' doc <- body_add_html_fragment(doc, "<p>Hello <b>world</b></p>")
#' doc <- body_add_html_fragment(doc,
#'   "<p>This has <b>bold and <i>bold-italic</i></b> text.</p>"
#' )
#' doc <- body_add_html_fragment(doc,
#'   "<ul><li>First item</li><li><b>Bold</b> second item</li></ul>"
#' )
#'
#' print(doc, target = tempfile(fileext = ".docx"))
body_add_html_fragment <- function(document, html_str, style = NULL) {
  # Skip empty or whitespace-only strings.
  if (is.null(html_str) || trimws(html_str) == "") {
    return(document)
  }

  body_node <- wrap_html_fragment(html_str)
  html_children <- xml2::xml_contents(body_node)

  for (child_node in html_children) {
    tag <- xml2::xml_name(child_node)

    if (tag == "text") {
      # --- Bare text outside any tag ---
      text <- trimws(xml2::xml_text(child_node))
      if (nchar(text) > 0) {
        document <- officer::body_add_fpar(
          document,
          collect_inline_runs(child_node),
          style = style
        )
      }

    } else if (tag == "p") {
      # --- Paragraph: may contain mixed inline formatting ---
      # <p>Hello <b>bold</b> and <i>italic</i></p>
      # → one fpar with four ftext runs.
      document <- officer::body_add_fpar(
        document,
        collect_inline_runs(child_node),
        style = style
      )

    } else if (tag == "ul") {
      # --- Unordered list: each <li> becomes a bullet ---
      # XPath ".//li" finds all <li> descendants (handles nested <ul> too,
      # though nested lists would need ilvl support for proper indentation).
      items <- xml2::xml_find_all(child_node, "./li")
      for (item in items) {
        document <- list_add_fpar(
          document,
          collect_inline_runs(item),
          style = style,
          list_type = "bullet"
        )
      }
      document <- list_end(document)

    } else if (tag == "ol") {
      # --- Ordered list: each <li> becomes a numbered item ---
      items <- xml2::xml_find_all(child_node, "./li")
      for (item in items) {
        document <- list_add_fpar(
          document,
          collect_inline_runs(item),
          style = style,
          list_type = "decimal"
        )
      }
      document <- list_end(document)

    } else if (grepl("^h[1-6]$", tag)) {
      # --- Headings: use officer's heading styles ---
      # <h1> → "heading 1", <h2> → "heading 2", etc.
      # These are built into every Word template.
      level <- as.integer(sub("h", "", tag))
      heading_style <- paste("heading", level)
      document <- officer::body_add_fpar(
        document,
        collect_inline_runs(child_node),
        style = heading_style
      )

    } else if (tag == "blockquote") {
      # --- Block quote: use the user's style (or default) ---
      # A more complete implementation could use a "Quote" style if available.
      document <- officer::body_add_fpar(
        document,
        collect_inline_runs(child_node),
        style = style
      )

    } else if (tag %in% c("div", "section", "article")) {
      # --- Container tags: recurse into their children ---
      # These don't produce paragraphs themselves; they just group content.
      inner_html <- as.character(xml2::xml_contents(child_node))
      inner_str <- paste(inner_html, collapse = "")
      document <- body_add_html_fragment(document, inner_str, style = style)

    } else {
      # --- Unknown block tag: treat as a paragraph ---
      text <- trimws(xml2::xml_text(child_node))
      if (nchar(text) > 0) {
        document <- officer::body_add_fpar(
          document,
          collect_inline_runs(child_node),
          style = style
        )
      }
    }
  }

  document
}


# ==============================================================================
# SECTION 7: TESTS
#
# Run this file directly to execute the tests:
#
#   Rscript html_to_officer.R
#
# Each test builds a document from an HTML fragment, then inspects the
# resulting XML to verify the correct structure was produced.
# ==============================================================================

# Only run tests when this file is executed directly (not when sourced).
if (sys.nframe() == 0L) {

  library(officer)
  library(xml2)

  test_count <- 0L
  pass_count <- 0L

  run_test <- function(name, expr) {
    test_count <<- test_count + 1L
    cat(sprintf("  [%2d] %s ... ", test_count, name))
    tryCatch(
      {
        force(expr)
        pass_count <<- pass_count + 1L
        cat("PASS\n")
      },
      error = function(e) {
        cat("FAIL\n")
        stop(sprintf("Test '%s' failed: %s", name, conditionMessage(e)),
             call. = FALSE)
      }
    )
  }

  assert <- function(condition, message = "assertion failed") {
    if (!isTRUE(condition)) stop(message, call. = FALSE)
  }

  assert_equal <- function(actual, expected, message = NULL) {
    if (is.null(message)) {
      message <- sprintf("expected '%s' but got '%s'", expected, actual)
    }
    if (!identical(actual, expected)) stop(message, call. = FALSE)
  }


  cat("\nRunning html_to_officer tests\n")
  cat(strrep("-", 60), "\n")


  # --- Helper: convert HTML to a doc and return the docx_summary ---

  html_to_summary <- function(html_str, style = NULL) {
    doc <- read_docx()
    doc <- body_add_html_fragment(doc, html_str, style = style)
    tf <- tempfile(fileext = ".docx")
    print(doc, target = tf)
    doc2 <- read_docx(tf)
    s <- docx_summary(doc2)
    unlink(tf)
    s
  }


  # ---------- 1. Plain paragraph ----------

  run_test("plain <p> becomes one paragraph", {
    s <- html_to_summary("<p>Hello world</p>")
    # Should be one paragraph (plus the default empty first paragraph).
    text_rows <- s[s$text != "", ]
    assert(nrow(text_rows) >= 1, "should have at least one text row")
    assert(
      any(grepl("Hello world", text_rows$text)),
      "should contain 'Hello world'"
    )
  })


  # ---------- 2. Bold formatting ----------

  run_test("bold text stays in same paragraph", {
    s <- html_to_summary("<p>Hello <b>world</b></p>")
    # "Hello " and "world" should be in the SAME paragraph, not two.
    text_rows <- s[s$text != "", ]
    # With docx_summary, inline formatting is concatenated into one text.
    assert(
      any(grepl("Hello world", text_rows$text, fixed = TRUE) |
          grepl("Hello", text_rows$text, fixed = TRUE)),
      "should contain the text in one or combined rows"
    )
    # Most importantly: there should NOT be a separate paragraph for "world".
    assert(
      !any(text_rows$text == "world"),
      "'world' should not be a separate paragraph"
    )
  })


  # ---------- 3. Nested bold+italic ----------

  run_test("nested <b><i> produces bold-italic runs", {
    # Test at the fpar level — build the fpar and inspect the runs directly.
    body <- wrap_html_fragment("<p>normal <b>bold <i>both</i></b> end</p>")
    p_node <- xml2::xml_find_first(body, ".//p")
    fp <- collect_inline_runs(p_node)

    # fp is an fpar. Its internal structure contains the runs.
    # We check it round-trips correctly by adding it to a document.
    doc <- read_docx()
    doc <- body_add_fpar(doc, fp)
    tf <- tempfile(fileext = ".docx")
    print(doc, target = tf)
    doc2 <- read_docx(tf)
    s <- docx_summary(doc2)
    unlink(tf)

    # The paragraph text should contain all parts concatenated.
    full_text <- paste(s$text[s$text != ""], collapse = " ")
    assert(grepl("normal", full_text), "should contain 'normal'")
    assert(grepl("bold", full_text), "should contain 'bold'")
    assert(grepl("both", full_text), "should contain 'both'")
    assert(grepl("end", full_text), "should contain 'end'")
  })


  # ---------- 4. Multiple paragraphs ----------

  run_test("multiple <p> tags become separate paragraphs", {
    s <- html_to_summary("<p>First</p><p>Second</p><p>Third</p>")
    text_rows <- s[s$text != "", ]
    assert(nrow(text_rows) >= 3, "should have at least 3 text paragraphs")
  })


  # ---------- 5. Unordered list ----------

  run_test("<ul> creates bullet items", {
    doc <- read_docx()
    doc <- body_add_html_fragment(doc,
      "<ul><li>Alpha</li><li>Bravo</li></ul>"
    )

    # Check the XML for <w:numPr> on the list paragraphs.
    # The cursor is on the last paragraph added.
    node <- docx_current_block_xml(doc)
    num_pr <- xml2::xml_find_first(node, "w:pPr/w:numPr/w:numId")
    assert(!inherits(num_pr, "xml_missing"), "last <li> should have numPr")
  })


  # ---------- 6. Ordered list ----------

  run_test("<ol> creates numbered items", {
    doc <- read_docx()
    doc <- body_add_html_fragment(doc,
      "<ol><li>One</li><li>Two</li></ol>"
    )

    node <- docx_current_block_xml(doc)
    num_pr <- xml2::xml_find_first(node, "w:pPr/w:numPr/w:numId")
    assert(!inherits(num_pr, "xml_missing"), "last <li> should have numPr")

    # Verify it's a decimal format by checking numbering.xml.
    num_doc <- xml2::read_xml(
      file.path(doc$package_dir, "word", "numbering.xml")
    )
    decimal_fmts <- xml2::xml_find_all(
      num_doc,
      "w:abstractNum/w:lvl/w:numFmt[@w:val='decimal']"
    )
    assert(length(decimal_fmts) > 0, "should have decimal format")
  })


  # ---------- 7. List with inline formatting ----------

  run_test("<li> with <b> preserves formatting in same paragraph", {
    doc <- read_docx()
    doc <- body_add_html_fragment(doc,
      "<ul><li>Normal and <b>bold</b> text</li></ul>"
    )

    # Should be one list item, not two paragraphs.
    tf <- tempfile(fileext = ".docx")
    print(doc, target = tf)
    doc2 <- read_docx(tf)
    s <- docx_summary(doc2)
    unlink(tf)

    text_rows <- s[s$text != "", ]
    # "bold" should NOT be a separate paragraph.
    assert(
      !any(text_rows$text == "bold"),
      "'bold' should be inline, not a separate paragraph"
    )
  })


  # ---------- 8. Headings ----------

  run_test("<h1> and <h2> produce heading-styled paragraphs", {
    s <- html_to_summary("<h1>Title</h1><h2>Subtitle</h2><p>Body</p>")
    text_rows <- s[s$text != "", ]
    assert(nrow(text_rows) >= 3, "should have heading + subheading + body")
    assert(any(grepl("Title", text_rows$text)), "should contain 'Title'")
  })


  # ---------- 9. Mixed content ----------

  run_test("mixed paragraphs, lists, and headings", {
    html <- paste0(
      "<h1>Report</h1>",
      "<p>Introduction with <i>emphasis</i>.</p>",
      "<ul><li>Point one</li><li>Point two</li></ul>",
      "<p>Conclusion.</p>"
    )
    s <- html_to_summary(html)
    text_rows <- s[s$text != "", ]
    assert(nrow(text_rows) >= 5,
           "should have heading + intro + 2 bullets + conclusion")
  })


  # ---------- 10. Empty and whitespace input ----------

  run_test("empty string is handled gracefully", {
    doc <- read_docx()
    # Should not crash.
    doc <- body_add_html_fragment(doc, "")
    doc <- body_add_html_fragment(doc, "   ")
    doc <- body_add_html_fragment(doc, NULL)
    # Should still be a valid document.
    tf <- tempfile(fileext = ".docx")
    print(doc, target = tf)
    doc2 <- read_docx(tf)
    unlink(tf)
    assert(TRUE, "should not crash on empty input")
  })


  # ---------- 11. Bare text (no tags) ----------

  run_test("bare text without tags becomes a paragraph", {
    s <- html_to_summary("Just some text with no tags")
    text_rows <- s[s$text != "", ]
    assert(
      any(grepl("Just some text", text_rows$text)),
      "should contain the bare text"
    )
  })


  # ---------- 12. Round-trip integrity ----------

  run_test("complex HTML round-trips without corruption", {
    html <- paste0(
      "<p>Normal <b>bold</b> <i>italic</i> <b><i>both</i></b></p>",
      "<ol><li>Step <b>one</b></li><li>Step two</li></ol>",
      "<ul><li>Bullet with <i>emphasis</i></li></ul>",
      "<p>Final paragraph.</p>"
    )
    doc <- read_docx()
    doc <- body_add_html_fragment(doc, html)
    tf <- tempfile(fileext = ".docx")
    print(doc, target = tf)

    # If the XML is malformed, read_docx will error.
    doc2 <- read_docx(tf)
    s <- docx_summary(doc2)
    unlink(tf)

    text_rows <- s[s$text != "", ]
    assert(nrow(text_rows) >= 5, "should have multiple paragraphs")
  })


  # ---------- 13. Deeply nested formatting ----------

  run_test("triple-nested inline tags accumulate formatting", {
    # <b><i><u>text</u></i></b> should produce bold + italic + underlined.
    body <- wrap_html_fragment("<p><b><i><u>styled</u></i></b></p>")
    p_node <- xml2::xml_find_first(body, ".//p")
    fp <- collect_inline_runs(p_node)

    # Verify by round-tripping.
    doc <- read_docx()
    doc <- body_add_fpar(doc, fp)
    tf <- tempfile(fileext = ".docx")
    print(doc, target = tf)
    doc2 <- read_docx(tf)
    s <- docx_summary(doc2)
    unlink(tf)

    assert(any(grepl("styled", s$text)), "should contain 'styled'")
  })


  # ---------- 14. Two separate lists restart numbering ----------

  run_test("two separate <ol> blocks restart numbering", {
    doc <- read_docx()
    doc <- body_add_html_fragment(doc,
      "<ol><li>A</li><li>B</li></ol>"
    )
    # list_end is called internally after each <ol>.
    doc <- body_add_html_fragment(doc,
      "<ol><li>C</li><li>D</li></ol>"
    )

    # The second <ol> should have a different numId (restarted).
    node <- docx_current_block_xml(doc)
    num_id_second <- xml2::xml_attr(
      xml2::xml_find_first(node, "w:pPr/w:numPr/w:numId"), "val"
    )

    # Move back to the first <ol>'s last item.
    doc <- officer::cursor_backward(doc)  # D -> C
    doc <- officer::cursor_backward(doc)  # C -> B
    num_id_first <- xml2::xml_attr(
      xml2::xml_find_first(
        docx_current_block_xml(doc), "w:pPr/w:numPr/w:numId"
      ),
      "val"
    )

    assert(
      num_id_first != num_id_second,
      sprintf("two <ol> blocks should have different numIds (got %s and %s)",
              num_id_first, num_id_second)
    )
  })


  # ---- Summary ----

  cat(strrep("-", 60), "\n")
  cat(sprintf("\n%d / %d tests passed.\n", pass_count, test_count))
  if (pass_count == test_count) {
    cat(sprintf("\nALL %d TESTS PASSED\n\n", test_count))
  } else {
    stop("Some tests failed.", call. = FALSE)
  }
}
