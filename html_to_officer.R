# ==============================================================================
# html_to_officer.R
#
# Converts HTML fragments (like those in DDI 3.3 XML metadata) into formatted
# officer paragraphs. Depends on officer_list_helpers.R for list support.
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
#   doc <- body_add_html_fragment(doc, "Line one\nLine two")
#   print(doc, target = "output.docx")
#
#
# WHAT THIS DOES
# ==============
#
# HTML has two kinds of elements:
#
#   BLOCK elements become separate paragraphs:
#     <p>, <ul>, <ol>, <h1>-<h6>, <blockquote>
#
#   INLINE elements become runs WITHIN a paragraph:
#     <b>/<strong> (bold), <i>/<em> (italic), <u> (underline),
#     <sub>, <sup>, <br>, <span>, <a>, plain text
#
# This maps directly to officer's model:
#   - One paragraph = one fpar() containing one or more ftext() runs
#   - Bold/italic = fp_text properties on individual runs
#   - Line breaks = run_linebreak() objects between runs
#
# The tricky cases this handles:
#
#   1. NESTED FORMATTING: <b>bold <i>bold-italic</i></b>
#      The inline collector is recursive — formatting accumulates as we
#      descend. <i> inside <b> inherits bold and adds italic.
#
#   2. BARE INLINE TAGS: "Hello <b>world</b>" (no <p> wrapper)
#      The dispatcher groups consecutive inline nodes and emits them as
#      one paragraph. A block tag (or end of input) flushes the group.
#
#   3. NEWLINES: "Line one\nLine two"
#      xml2::read_html() collapses whitespace per the HTML spec, eating
#      newlines. We convert \n to <br/> BEFORE parsing, so they survive
#      as elements and become run_linebreak() in the output.
#
# ==============================================================================


# ---- Dependencies ------------------------------------------------------------

if (!requireNamespace("officer", quietly = TRUE))
  stop("Package 'officer' is required.")
if (!requireNamespace("xml2", quietly = TRUE))
  stop("Package 'xml2' is required.")

if (!exists("list_add_fpar", mode = "function")) {
  if (file.exists("officer_list_helpers.R")) {
    source("officer_list_helpers.R")
  } else {
    warning("officer_list_helpers.R not found. <ul>/<ol> support disabled.")
  }
}


# ---- Constants ---------------------------------------------------------------

# Inline formatting tags mapped to fp_text properties.
INLINE_FORMAT_MAP <- list(
  b      = list(bold = TRUE),
  strong = list(bold = TRUE),
  i      = list(italic = TRUE),
  em     = list(italic = TRUE),
  u      = list(underlined = TRUE),
  sub    = list(vertical.align = "subscript"),
  sup    = list(vertical.align = "superscript")
)

# Block-level tags. Everything else is inline.
BLOCK_TAGS <- c("p", "ul", "ol", "h1", "h2", "h3", "h4", "h5", "h6",
                "blockquote", "pre", "div", "section", "article", "table", "hr")


# ---- HTML parsing ------------------------------------------------------------

#' Parse an HTML fragment into an xml2 body node.
#' Converts \n to <br/> before parsing so newlines survive.
wrap_html_fragment <- function(html_str) {
  html_str <- gsub("\r?\n", "<br/>", html_str)
  xml2::xml_find_first(
    xml2::read_html(paste0("<html><body>", html_str, "</body></html>")),
    ".//body"
  )
}

.is_inline <- function(node) {
  tag <- xml2::xml_name(node)
  tag == "text" || !(tag %in% BLOCK_TAGS)
}


# ---- Inline content collector ------------------------------------------------
#
# Walks an element's children recursively. Formatting accumulates:
#   <b>bold <i>both</i></b>  →  ftext("bold ", bold) + ftext("both", bold+italic)

#' Recursively collect ftext runs from an element's children.
.collect_runs <- function(node, props = list()) {
  runs <- list()
  for (child in xml2::xml_contents(node)) {
    tag <- xml2::xml_name(child)

    if (tag == "text") {
      text <- xml2::xml_text(child)
      if (nchar(text) > 0) {
        fp <- if (length(props) == 0) officer::fp_text()
              else do.call(officer::fp_text, props)
        runs <- c(runs, list(officer::ftext(text, prop = fp)))
      }
    } else if (tag == "br") {
      runs <- c(runs, list(officer::run_linebreak()))
    } else if (tag %in% names(INLINE_FORMAT_MAP)) {
      new_props <- props
      for (nm in names(INLINE_FORMAT_MAP[[tag]]))
        new_props[[nm]] <- INLINE_FORMAT_MAP[[tag]][[nm]]
      runs <- c(runs, .collect_runs(child, new_props))
    } else {
      runs <- c(runs, .collect_runs(child, props))
    }
  }
  runs
}

#' Collect inline content from a single element into one fpar.
collect_inline_runs <- function(node) {
  runs <- if (xml2::xml_name(node) == "text") {
    list(officer::ftext(xml2::xml_text(node)))
  } else {
    .collect_runs(node)
  }
  if (length(runs) == 0) runs <- list(officer::ftext(""))
  do.call(officer::fpar, runs)
}

#' Collect inline content from a list of sibling nodes into one fpar.
#' Used when bare inline tags appear at the top level without a <p>.
collect_inline_runs_from_nodes <- function(nodes) {
  runs <- list()
  for (node in nodes) {
    tag <- xml2::xml_name(node)
    if (tag == "text") {
      text <- xml2::xml_text(node)
      if (nchar(text) > 0) runs <- c(runs, list(officer::ftext(text)))
    } else if (tag == "br") {
      runs <- c(runs, list(officer::run_linebreak()))
    } else if (tag %in% names(INLINE_FORMAT_MAP)) {
      runs <- c(runs, .collect_runs(node, INLINE_FORMAT_MAP[[tag]]))
    } else {
      runs <- c(runs, .collect_runs(node))
    }
  }
  if (length(runs) == 0) runs <- list(officer::ftext(""))
  do.call(officer::fpar, runs)
}


# ---- Block-level dispatcher --------------------------------------------------
#
# Walks top-level children. Groups consecutive inline nodes and flushes them
# as one paragraph when a block element (or end of input) is reached.

#' Convert an HTML fragment into formatted officer paragraphs.
#'
#' @param document An rdocx object.
#' @param html_str An HTML fragment string.
#' @param style    Paragraph style name from your template. NULL = default.
#' @return The modified rdocx object.
body_add_html_fragment <- function(document, html_str, style = NULL) {
  if (is.null(html_str) || trimws(html_str) == "") return(document)

  children   <- xml2::xml_contents(wrap_html_fragment(html_str))
  inline_acc <- list()

  flush_inline <- function() {
    if (length(inline_acc) == 0) return()
    combined <- trimws(paste0(vapply(inline_acc, xml2::xml_text, character(1)),
                              collapse = ""))
    has_br <- any(vapply(inline_acc, function(n) xml2::xml_name(n) == "br",
                         logical(1)))
    if (nchar(combined) > 0 || has_br) {
      document <<- officer::body_add_fpar(
        document, collect_inline_runs_from_nodes(inline_acc), style = style)
    }
    inline_acc <<- list()
  }

  for (child in children) {
    tag <- xml2::xml_name(child)

    if (.is_inline(child)) {
      inline_acc <- c(inline_acc, list(child))
      next
    }

    flush_inline()

    if (tag == "p") {
      document <- officer::body_add_fpar(
        document, collect_inline_runs(child), style = style)

    } else if (tag == "ul") {
      for (li in xml2::xml_find_all(child, "./li"))
        document <- list_add_fpar(
          document, collect_inline_runs(li), style = style, list_type = "bullet")
      document <- list_end(document)

    } else if (tag == "ol") {
      for (li in xml2::xml_find_all(child, "./li"))
        document <- list_add_fpar(
          document, collect_inline_runs(li), style = style, list_type = "decimal")
      document <- list_end(document)

    } else if (grepl("^h[1-6]$", tag)) {
      document <- officer::body_add_fpar(
        document, collect_inline_runs(child),
        style = paste("heading", as.integer(sub("h", "", tag))))

    } else if (tag %in% c("div", "section", "article")) {
      inner <- paste(as.character(xml2::xml_contents(child)), collapse = "")
      document <- body_add_html_fragment(document, inner, style = style)

    } else if (nchar(trimws(xml2::xml_text(child))) > 0) {
      document <- officer::body_add_fpar(
        document, collect_inline_runs(child), style = style)
    }
  }

  flush_inline()
  document
}
