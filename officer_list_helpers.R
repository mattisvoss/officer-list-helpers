# ==============================================================================
# officer_list_helpers.R
#
# Helper functions that add bullet and numbered list support to the R 'officer'
# package. Use these alongside officer's standard API — they work with any
# template and any custom paragraph style.
#
#
# TABLE OF CONTENTS
# =================
#
#   1. BACKGROUND: How .docx files work
#   2. BACKGROUND: How lists work in .docx
#   3. BACKGROUND: The tools we use (xml2 and XPath)
#   4. BACKGROUND: How officer fits in
#   5. BACKGROUND: Indentation and the tab_pos approach
#   6. CODE: Constants
#   7. CODE: XML string builders (pure functions, no side effects)
#   8. CODE: Numbering.xml file management (reads/writes the XML file)
#   9. CODE: Public API (the functions you actually call)
#
#
# QUICK START
# ===========
#
#   library(officer)
#   library(xml2)
#   source("officer_list_helpers.R")
#
#   doc <- read_docx()
#   doc <- list_add_par(doc, "First bullet",  list_type = "bullet")
#   doc <- list_add_par(doc, "Second bullet",  list_type = "bullet")
#   doc <- list_end(doc)
#   doc <- list_add_par(doc, "Step one", list_type = "decimal")
#   doc <- list_add_par(doc, "Step two", list_type = "decimal")
#   print(doc, target = "output.docx")
#
#   # For styles with borders/fills, use flush-left mode:
#   doc <- read_docx()
#   doc <- list_setup(doc, tab_pos = 360L)
#   doc <- list_add_par(doc, "Bullet inside a bordered style",
#                       style = "My Bordered Style", list_type = "bullet")
#
#
# ==========================
# 1. HOW .DOCX FILES WORK
# ==========================
#
# A .docx file is just a .zip archive. Rename "report.docx" to "report.zip",
# unzip it, and you'll find a folder structure like:
#
#   report/
#     [Content_Types].xml        <- registry of what's in the package
#     word/
#       document.xml             <- the actual document content
#       numbering.xml            <- list/bullet definitions (our focus)
#       styles.xml               <- paragraph and character styles
#       _rels/
#         document.xml.rels      <- links between files
#
# Everything is XML (a structured text format with nested tags, like HTML).
# A paragraph that says "Hello world" in bold looks like this in document.xml:
#
#   <w:p>                        <- paragraph (a block of text)
#     <w:pPr>                    <- paragraph properties (style, alignment, etc.)
#       <w:pStyle w:val="Normal"/>
#     </w:pPr>
#     <w:r>                      <- run (a chunk of text with uniform formatting)
#       <w:rPr>                  <- run properties (bold, italic, font, etc.)
#         <w:b/>                 <- bold
#       </w:rPr>
#       <w:t>Hello world</w:t>   <- the actual text content
#     </w:r>
#   </w:p>
#
# Key vocabulary:
#
#   PARAGRAPH (<w:p>)
#     One block of text — like a line or paragraph in a document. Contains
#     one or more runs.
#
#   RUN (<w:r>)
#     A chunk of text within a paragraph that shares the same character
#     formatting. The sentence "Hello **world**" has two runs: "Hello "
#     (normal) and "world" (bold). Runs are how Word handles mixed formatting
#     within a single paragraph.
#
#   PARAGRAPH PROPERTIES (<w:pPr>)
#     Settings that apply to the whole paragraph: style, alignment,
#     indentation, and — crucially for us — list numbering.
#
# The "w:" prefix on every tag is an XML NAMESPACE. All Word XML elements
# belong to the namespace:
#
#   http://schemas.openxmlformats.org/wordprocessingml/2006/main
#
# The "w:" is just a shorthand alias for that long URL. Think of it like a
# module prefix — it says "this element is defined by the Word spec, not by
# some other XML format."
#
#
# ==========================
# 2. HOW LISTS WORK IN .DOCX
# ==========================
#
# A paragraph becomes a list item when its <w:pPr> (paragraph properties)
# contains a <w:numPr> element. Here's what a bulleted paragraph looks like:
#
#   <w:p>
#     <w:pPr>
#       <w:pStyle w:val="Normal"/>
#       <w:numPr>                          <- THIS makes it a list item
#         <w:ilvl w:val="0"/>              <- indent level (0=top, 1=sub, ...)
#         <w:numId w:val="3"/>             <- which list it belongs to
#       </w:numPr>
#     </w:pPr>
#     <w:r><w:t>Buy groceries</w:t></w:r>
#   </w:p>
#
# The numId points into word/numbering.xml, which has a two-layer system:
#
#
# LAYER 1 — <w:abstractNum>: A FORMAT TEMPLATE
#
#   Defines how each indent level of a list looks. Contains up to 9 <w:lvl>
#   elements (levels 0 through 8), each specifying:
#
#     <w:numFmt>   What kind of marker: "bullet" or "decimal" (or upperRoman,
#                  lowerLetter, etc.)
#     <w:lvlText>  What to display: "•" for bullets, "%1." for decimals.
#                  The %1 token means "insert the counter for this level."
#     <w:ind>      How far to indent, measured in "twips" (1/20 of a point).
#                  720 twips = 0.5 inch.
#
#   Example — a two-level bullet format:
#
#     <w:abstractNum w:abstractNumId="0">
#       <w:lvl w:ilvl="0">                 <- top level: solid bullet
#         <w:numFmt w:val="bullet"/>
#         <w:lvlText w:val="•"/>
#         <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
#       </w:lvl>
#       <w:lvl w:ilvl="1">                 <- sub-item: open circle
#         <w:numFmt w:val="bullet"/>
#         <w:lvlText w:val="◦"/>
#         <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
#       </w:lvl>
#     </w:abstractNum>
#
#
# LAYER 2 — <w:num>: AN INSTANCE (a specific list in the document)
#
#   Points to an abstractNum for its formatting. Multiple <w:num> elements
#   can share the same abstractNum but maintain independent counters. This
#   is how you get two separate "1. 2. 3." lists in the same document —
#   they look identical but count independently.
#
#   Example — two list instances sharing the same format:
#
#     <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>  <- list A
#     <w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num>  <- list B
#
#
# RESTARTING NUMBERING:
#
#   By default, a new <w:num> pointing to the same abstractNum may continue
#   counting from where the previous one left off (behaviour varies by
#   renderer). To FORCE it to restart at 1, you add a <w:lvlOverride>:
#
#     <w:num w:numId="3">
#       <w:abstractNumId w:val="0"/>
#       <w:lvlOverride w:ilvl="0">
#         <w:startOverride w:val="1"/>    <- forces restart at 1
#       </w:lvlOverride>
#     </w:num>
#
#   This is what list_end() triggers under the hood.
#
#
# ==========================
# 3. THE TOOLS WE USE: xml2 AND XPATH
# ==========================
#
# We use the R package 'xml2' to read, modify, and write XML files.
# Here is every xml2 function used in this file:
#
#   read_xml(path)               Read an XML file from disk into memory.
#   write_xml(doc, path)         Write an in-memory XML document back to disk.
#   as_xml_document(string)      Parse an XML string into a document object.
#   xml_find_all(doc, xpath)     Find ALL nodes matching an XPath query.
#   xml_find_first(doc, xpath)   Find the FIRST node matching an XPath query.
#   xml_attr(node, name)         Get the value of an attribute. Returns NA if
#                                the attribute doesn't exist.
#   xml_add_child(parent, child) Insert a child node INSIDE a parent element.
#   xml_add_sibling(node, sib)   Insert a node NEXT TO an existing node.
#
# XPath is a query language for finding things in XML. Quick primer:
#
#   "w:abstractNum"                   Find all <w:abstractNum> children
#   "w:abstractNum/w:lvl"             Find <w:lvl> inside <w:abstractNum>
#   "w:num[@w:numId='3']"             Find <w:num> where numId = 3
#   "w:pPr/w:numPr/w:numId"          Navigate: pPr -> numPr -> numId
#
#
# ==========================
# 4. HOW OFFICER FITS IN
# ==========================
#
# officer::read_docx() unzips the .docx into a temp directory and gives you
# an R object (class "rdocx"). x$package_dir is the path to that temp folder.
#
# Our strategy: use officer to add paragraphs normally, then immediately
# modify the XML of the paragraph officer just created to add <w:numPr>.
# Your paragraph style is always preserved — we layer list formatting on top.
#
#
# ==========================
# 5. INDENTATION AND THE TAB_POS APPROACH
# ==========================
#
# By default, list items are indented: the bullet/number hangs to the left,
# and the text is pushed rightward. This is the standard Word behaviour:
#
#   STANDARD (default):
#     <w:ind w:left="720" w:hanging="360"/>
#
#     |          • Bullet text starts here
#     |            ↑ left=720 (text indent from margin)
#     |          ↑ hanging=360 (bullet hangs back from text)
#     |        ↑ bullet is at 720-360 = 360 twips from margin
#
# This works fine for normal paragraphs, but if your paragraph style has
# a BORDER or FILL (background colour), the indent pushes the content area
# rightward, creating an ugly gap between the border/fill edge and the text.
#
# The fix: FLUSH-LEFT mode. Set the paragraph indent to zero and use a
# TAB STOP to create the gap between the bullet/number and the text:
#
#   FLUSH-LEFT (tab_pos mode):
#     <w:ind w:left="0" w:firstLine="0"/>
#     <w:tabs><w:tab w:val="left" w:pos="360"/></w:tabs>
#     <w:suff w:val="tab"/>
#
#     |• → Bullet text starts here
#     |↑    ↑ tab stop at pos=360 pushes text inward
#     |↑ bullet at left margin (position 0)
#     |↑ border/fill edge — no gap!
#
# The <w:suff w:val="tab"/> element tells Word: "after the bullet/number,
# insert a tab character (not a space)." The tab stop then controls where
# the text begins. The paragraph indent is zero, so borders and fills
# extend all the way to the margin.
#
# To enable flush-left mode, call list_setup(doc, tab_pos = 360L) before
# adding any list items. The tab_pos value (in twips) controls the gap.
# 360 twips = 0.25 inch is a good default.
#
# Sub-levels (ilvl > 0) in flush-left mode use multiples of tab_pos for
# both the indent and the tab stop, giving a clean nested appearance
# without breaking the border/fill.
#
# ==============================================================================


# ---- Dependency check --------------------------------------------------------

if (!requireNamespace("officer", quietly = TRUE)) {
  stop("Package 'officer' is required. Install it with: install.packages('officer')")
}
if (!requireNamespace("xml2", quietly = TRUE)) {
  stop("Package 'xml2' is required. Install it with: install.packages('xml2')")
}


# ==============================================================================
# SECTION 6: CONSTANTS
# ==============================================================================

OOXML_NS <- "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
OOXML_NS_DECL <- sprintf('xmlns:w="%s"', OOXML_NS)
BULLET_GLYPHS <- c("\u2022", "\u25e6", "\u25aa")


# ==============================================================================
# SECTION 7: XML STRING BUILDERS
# ==============================================================================

#' Build the XML for one indent level of a list format definition.
#'
#' @param ilvl Integer 0-8. Which indent level this defines.
#' @param list_type "bullet" or "decimal".
#' @param tab_pos Integer or NULL. If NULL (default), uses standard hanging
#'   indent. If an integer (twips), uses flush-left mode with a tab stop.
#' @return An XML string: '<w:lvl w:ilvl="0">...</w:lvl>'.
.build_lvl_xml <- function(ilvl, list_type, tab_pos = NULL) {
  if (list_type == "bullet") {
    fmt <- "bullet"
    text <- BULLET_GLYPHS[ilvl %% 3L + 1L]
  } else {
    fmt <- "decimal"
    text <- sprintf("%%%d.", ilvl + 1L)
  }

  if (is.null(tab_pos)) {
    # STANDARD MODE: hanging indent.
    # Each level indents 0.5 inch (720 twips) further.
    # The bullet/number "hangs" 360 twips to the left of the text.
    left_indent <- 720L * (ilvl + 1L)

    sprintf(
      paste0(
        '<w:lvl w:ilvl="%d">',
          '<w:start w:val="1"/>',
          '<w:numFmt w:val="%s"/>',
          '<w:lvlText w:val="%s"/>',
          '<w:lvlJc w:val="left"/>',
          '<w:pPr>',
            '<w:ind w:left="%d" w:hanging="360"/>',
          '</w:pPr>',
        '</w:lvl>'
      ),
      ilvl, fmt, text, left_indent
    )
  } else {
    # FLUSH-LEFT MODE: zero indent, tab stop for bullet-to-text gap.
    #
    # For ilvl 0: indent = 0, tab at tab_pos
    # For ilvl 1: indent = tab_pos, tab at tab_pos * 2
    # For ilvl 2: indent = tab_pos * 2, tab at tab_pos * 3
    #
    # This gives nested levels a staircase effect while keeping ilvl 0
    # flush with the left margin (preserving borders/fills).
    level_indent <- as.integer(tab_pos) * ilvl
    level_tab <- as.integer(tab_pos) * (ilvl + 1L)

    sprintf(
      paste0(
        '<w:lvl w:ilvl="%d">',
          '<w:start w:val="1"/>',
          '<w:numFmt w:val="%s"/>',
          '<w:lvlText w:val="%s"/>',
          '<w:suff w:val="tab"/>',
          '<w:lvlJc w:val="left"/>',
          '<w:pPr>',
            '<w:ind w:left="%d" w:firstLine="0"/>',
            '<w:tabs><w:tab w:val="left" w:pos="%d"/></w:tabs>',
          '</w:pPr>',
        '</w:lvl>'
      ),
      ilvl, fmt, text, level_indent, level_tab
    )
  }
}


#' Build a complete <w:abstractNum> — a list format template.
#'
#' @param abstract_num_id Integer. Unique ID.
#' @param list_type "bullet" or "decimal".
#' @param tab_pos Integer or NULL. Passed through to .build_lvl_xml.
#' @return A complete XML string for injection into numbering.xml.
.build_abstract_num_xml <- function(abstract_num_id, list_type, tab_pos = NULL) {
  lvls <- vapply(
    0:8,
    .build_lvl_xml,
    character(1),
    list_type = list_type,
    tab_pos = tab_pos
  )

  sprintf(
    paste0(
      '<w:abstractNum %s w:abstractNumId="%d">',
        '<w:multiLevelType w:val="multilevel"/>',
        '%s',
      '</w:abstractNum>'
    ),
    OOXML_NS_DECL, abstract_num_id, paste0(lvls, collapse = "")
  )
}


#' Build a <w:num> — an instance of a list.
#'
#' @param num_id Integer. Unique ID for this list instance.
#' @param abstract_num_id Integer. Which format template to use.
#' @param restart Logical. If TRUE, adds <w:startOverride> to reset counter.
#' @return A complete XML string.
.build_num_xml <- function(num_id, abstract_num_id, restart = FALSE) {
  override <- ""
  if (restart) {
    override <- paste0(
      '<w:lvlOverride w:ilvl="0">',
        '<w:startOverride w:val="1"/>',
      '</w:lvlOverride>'
    )
  }

  sprintf(
    '<w:num %s w:numId="%d"><w:abstractNumId w:val="%d"/>%s</w:num>',
    OOXML_NS_DECL, num_id, abstract_num_id, override
  )
}


#' Build a <w:numPr> node — the element that makes a paragraph a list item.
#'
#' @param num_id Integer. Which list instance.
#' @param ilvl Integer. Indent level (0 = top, 1 = sub-item, ...).
#' @return An xml2 node object ready for xml_add_child().
.build_num_pr_node <- function(num_id, ilvl) {
  xml2::read_xml(sprintf(
    '<w:numPr xmlns:w="%s"><w:ilvl w:val="%d"/><w:numId w:val="%d"/></w:numPr>',
    OOXML_NS, as.integer(ilvl), as.integer(num_id)
  ))
}


# ==============================================================================
# SECTION 8: NUMBERING.XML FILE MANAGEMENT
# ==============================================================================

#' Read numbering.xml from the document's temp directory.
.read_numbering_xml <- function(x) {
  xml2::read_xml(file.path(x$package_dir, "word", "numbering.xml"))
}

#' Write numbering.xml back to the document's temp directory.
.write_numbering_xml <- function(x, doc) {
  xml2::write_xml(doc, file = file.path(x$package_dir, "word", "numbering.xml"))
}

#' Find the next available IDs in numbering.xml.
.next_available_ids <- function(doc) {
  abs_ids <- as.integer(
    xml2::xml_attr(xml2::xml_find_all(doc, "w:abstractNum"), "abstractNumId")
  )
  num_ids <- as.integer(
    xml2::xml_attr(xml2::xml_find_all(doc, "w:num"), "numId")
  )
  list(
    abstract_num_id = max(abs_ids) + 1L,
    num_id = max(num_ids) + 1L
  )
}


#' Set up the list numbering system for a document.
#'
#' Called automatically on first list_add_* call, or explicitly via
#' list_setup(). Creates abstractNum and num definitions in numbering.xml.
#'
#' @param x An rdocx object.
#' @param tab_pos Integer or NULL. If set, uses flush-left mode.
#' @return The rdocx object with x$.list_state initialized.
.init_list_state <- function(x, tab_pos = NULL) {
  doc <- .read_numbering_xml(x)
  ids <- .next_available_ids(doc)

  bullet_abs_id  <- ids$abstract_num_id
  decimal_abs_id <- ids$abstract_num_id + 1L
  bullet_num_id  <- ids$num_id
  decimal_num_id <- ids$num_id + 1L

  # Inject the two format templates, passing tab_pos through.
  abs_nodes <- xml2::xml_find_all(doc, "w:abstractNum")
  last_abs <- abs_nodes[[length(abs_nodes)]]
  xml2::xml_add_sibling(last_abs, xml2::as_xml_document(
    .build_abstract_num_xml(bullet_abs_id, "bullet", tab_pos = tab_pos)
  ))
  xml2::xml_add_sibling(last_abs, xml2::as_xml_document(
    .build_abstract_num_xml(decimal_abs_id, "decimal", tab_pos = tab_pos)
  ))

  # Inject the two initial list instances.
  num_nodes <- xml2::xml_find_all(doc, "w:num")
  last_num <- num_nodes[[length(num_nodes)]]
  xml2::xml_add_sibling(last_num, xml2::as_xml_document(
    .build_num_xml(bullet_num_id, bullet_abs_id)
  ))
  xml2::xml_add_sibling(last_num, xml2::as_xml_document(
    .build_num_xml(decimal_num_id, decimal_abs_id)
  ))

  .write_numbering_xml(x, doc)

  x$.list_state <- list(
    abstract_ids    = list(bullet = bullet_abs_id, decimal = decimal_abs_id),
    current_num_ids = list(bullet = bullet_num_id, decimal = decimal_num_id),
    active_type     = NULL,
    next_num_id     = decimal_num_id + 1L,
    tab_pos         = tab_pos
  )

  x
}


#' Ensure the list system is initialized. No-op if already done.
.ensure_init <- function(x) {
  if (is.null(x$.list_state)) {
    x <- .init_list_state(x)
  }
  x
}


#' Create a new list instance (<w:num>) with a restart override.
.restart_num <- function(x, list_type) {
  doc <- .read_numbering_xml(x)
  num_nodes <- xml2::xml_find_all(doc, "w:num")

  new_id <- x$.list_state$next_num_id
  abs_id <- x$.list_state$abstract_ids[[list_type]]

  xml2::xml_add_sibling(
    num_nodes[[length(num_nodes)]],
    xml2::as_xml_document(.build_num_xml(new_id, abs_id, restart = TRUE))
  )
  .write_numbering_xml(x, doc)

  x$.list_state$current_num_ids[[list_type]] <- new_id
  x$.list_state$next_num_id <- new_id + 1L
  x
}


#' Inject <w:numPr> into the paragraph at the current cursor position.
.inject_num_pr <- function(x, num_id, ilvl) {
  node <- officer::docx_current_block_xml(x)
  ppr <- xml2::xml_find_first(node, "w:pPr")
  xml2::xml_add_child(ppr, .build_num_pr_node(num_id, ilvl))
  x
}


#' Shared setup logic for all list_add_* functions.
.prepare_list_item <- function(x, list_type) {
  list_type <- match.arg(list_type, c("bullet", "decimal"))
  x <- .ensure_init(x)

  needs_restart <- is.null(x$.list_state$active_type) ||
                   x$.list_state$active_type != list_type

  if (needs_restart) {
    x <- .restart_num(x, list_type)
  }

  x$.list_state$active_type <- list_type
  num_id <- x$.list_state$current_num_ids[[list_type]]

  list(x = x, num_id = num_id)
}


# ==============================================================================
# SECTION 9: PUBLIC API
# ==============================================================================

#' Configure list formatting before adding list items.
#'
#' Call this ONCE before your first list_add_* call to control how list
#' items are indented. If you don't call this, standard hanging-indent
#' mode is used (fine for most cases).
#'
#' The main reason to call this: if your paragraph style has BORDERS or
#' FILLS, the standard indent pushes the content rightward, creating a gap
#' between the border and the bullet. Flush-left mode fixes this by keeping
#' the paragraph indent at zero and using a tab stop for the bullet-to-text
#' gap.
#'
#' @param x An rdocx object.
#' @param tab_pos Integer (twips) or NULL.
#'   NULL (default) = standard hanging indent (720 twips per level).
#'   Integer = flush-left mode. The value controls the gap between the
#'   bullet/number and the text. 360 = 0.25 inch, a good default.
#' @return The rdocx object with list formatting configured.
#'
#' @examples
#' doc <- read_docx()
#'
#' # Standard mode (default) — no need to call list_setup
#' doc <- list_add_par(doc, "Normal bullet", list_type = "bullet")
#'
#' # Flush-left mode — call list_setup first
#' doc2 <- read_docx()
#' doc2 <- list_setup(doc2, tab_pos = 360L)
#' doc2 <- list_add_par(doc2, "Flush bullet", list_type = "bullet")
list_setup <- function(x, tab_pos = NULL) {
  if (!is.null(x$.list_state)) {
    warning("list_setup() must be called before any list_add_* calls. Ignoring.")
    return(x)
  }
  x <- .init_list_state(x, tab_pos = tab_pos)
  x
}


#' Add a plain text paragraph as a list item.
#'
#' @param x An rdocx object.
#' @param value Character string. The paragraph text.
#' @param style Paragraph style name. NULL = document default.
#' @param list_type "bullet" or "decimal".
#' @param ilvl Indent level: 0 = top, 1 = sub-item, 2 = sub-sub.
#' @param pos "after" (default), "before", or "on".
#' @return The modified rdocx object.
#'
#' @examples
#' doc <- read_docx()
#' doc <- list_add_par(doc, "Buy groceries", list_type = "bullet")
#' doc <- list_add_par(doc, "Milk",          list_type = "bullet", ilvl = 1L)
#' print(doc, target = tempfile(fileext = ".docx"))
list_add_par <- function(x, value, style = NULL, list_type = "bullet",
                         ilvl = 0L, pos = "after") {
  prep <- .prepare_list_item(x, list_type)
  x <- prep$x
  x <- officer::body_add_par(x, value, style = style, pos = pos)
  x <- .inject_num_pr(x, prep$num_id, ilvl)
  x
}


#' Add a formatted paragraph (fpar) as a list item.
#'
#' @param x An rdocx object.
#' @param value An fpar object.
#' @param style Paragraph style name. NULL = document default.
#' @param list_type "bullet" or "decimal".
#' @param ilvl Indent level (default 0).
#' @param pos "after" (default), "before", or "on".
#' @return The modified rdocx object.
list_add_fpar <- function(x, value, style = NULL, list_type = "bullet",
                          ilvl = 0L, pos = "after") {
  prep <- .prepare_list_item(x, list_type)
  x <- prep$x
  x <- officer::body_add_fpar(x, value, style = style, pos = pos)
  x <- .inject_num_pr(x, prep$num_id, ilvl)
  x
}


#' Add multiple formatted paragraphs as list items.
#'
#' @param x An rdocx object.
#' @param blocks A block_list object.
#' @param list_type "bullet" or "decimal".
#' @param ilvl Indent level applied to ALL items (default 0).
#' @param pos "after" (default), "before", or "on".
#' @return The modified rdocx object.
list_add_blocks <- function(x, blocks, list_type = "bullet",
                            ilvl = 0L, pos = "after") {
  prep <- .prepare_list_item(x, list_type)
  x <- prep$x
  x <- officer::body_add_blocks(x, blocks, pos = pos)

  n <- length(blocks)
  for (i in seq_len(n)) {
    x <- .inject_num_pr(x, prep$num_id, ilvl)
    if (i < n) {
      x <- officer::cursor_backward(x)
    }
  }
  for (i in seq_len(n - 1L)) {
    x <- officer::cursor_forward(x)
  }

  x
}


#' End the current list.
#'
#' Call between two same-type lists to force restart. Switching types
#' restarts automatically. Safe to call anywhere, any number of times.
#'
#' @param x An rdocx object.
#' @return The rdocx object.
list_end <- function(x) {
  if (!is.null(x$.list_state)) {
    x$.list_state$active_type <- NULL
  }
  x
}


#' Inspect numbering definitions in a document.
#'
#' Prints a readable summary of every abstractNum and num in the document.
#'
#' @param x An rdocx object.
#' @return NULL (called for side effect of printing).
list_inspect <- function(x) {
  doc <- .read_numbering_xml(x)

  cat("=== Format templates (abstractNum) ===\n")
  cat("These define HOW a list looks at each indent level.\n\n")

  abs_nodes <- xml2::xml_find_all(doc, "w:abstractNum")
  for (node in abs_nodes) {
    abs_id <- xml2::xml_attr(node, "abstractNumId")
    cat(sprintf("  abstractNumId = %s\n", abs_id))
    lvls <- xml2::xml_find_all(node, "w:lvl")
    for (lvl in lvls) {
      fmt  <- xml2::xml_attr(xml2::xml_find_first(lvl, "w:numFmt"), "val")
      text <- xml2::xml_attr(xml2::xml_find_first(lvl, "w:lvlText"), "val")

      # Show indent info.
      ind <- xml2::xml_find_first(lvl, "w:pPr/w:ind")
      left <- xml2::xml_attr(ind, "left")
      hang <- xml2::xml_attr(ind, "hanging")
      first <- xml2::xml_attr(ind, "firstLine")

      # Show tab stop if present.
      tab <- xml2::xml_find_first(lvl, "w:pPr/w:tabs/w:tab")
      tab_info <- ""
      if (!inherits(tab, "xml_missing")) {
        tab_info <- sprintf("  tab=%s", xml2::xml_attr(tab, "pos"))
      }

      indent_info <- ""
      if (!is.na(hang)) {
        indent_info <- sprintf("left=%s hang=%s", left, hang)
      } else {
        indent_info <- sprintf("left=%s firstLine=%s", left, first)
      }

      cat(sprintf("    level %s: format=%-10s display=%-4s %s%s\n",
                  xml2::xml_attr(lvl, "ilvl"), fmt, text, indent_info, tab_info))
    }
    cat("\n")
  }

  cat("=== List instances (num) ===\n")
  cat("Each num is an active list. Paragraphs point to these via numId.\n\n")

  num_nodes <- xml2::xml_find_all(doc, "w:num")
  for (node in num_nodes) {
    override <- xml2::xml_find_first(node, "w:lvlOverride/w:startOverride")
    restart_note <- ""
    if (!inherits(override, "xml_missing")) {
      restart_note <- sprintf("  [restarts at %s]",
                              xml2::xml_attr(override, "val"))
    }
    cat(sprintf("  numId = %s  ->  abstractNumId = %s%s\n",
                xml2::xml_attr(node, "numId"),
                xml2::xml_attr(xml2::xml_find_first(node, "w:abstractNumId"), "val"),
                restart_note))
  }
  cat("\n")

  invisible(NULL)
}
