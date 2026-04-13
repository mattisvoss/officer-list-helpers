# ==============================================================================
# officer_list_helpers.R
#
# Bullet and numbered list support for the R 'officer' package.
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
#
# HOW .DOCX LISTS WORK
# ====================
#
# A .docx file is a zip of XML files. A paragraph becomes a list item when
# its <w:pPr> (paragraph properties) contains a <w:numPr> element:
#
#   <w:p>
#     <w:pPr>
#       <w:numPr>
#         <w:ilvl w:val="0"/>       <- indent level (0 = top, 1 = sub, ...)
#         <w:numId w:val="3"/>      <- which list this belongs to
#       </w:numPr>
#     </w:pPr>
#     <w:r><w:t>Buy groceries</w:t></w:r>
#   </w:p>
#
# The numId points into word/numbering.xml, which has two layers:
#
#   <w:abstractNum> — a FORMAT TEMPLATE defining bullet glyphs, number
#     formats, and indentation for each of 9 possible indent levels.
#
#   <w:num> — an INSTANCE referencing an abstractNum. Multiple <w:num>
#     elements can share the same abstractNum but count independently.
#     This is how you get two separate "1. 2. 3." lists.
#
# To restart numbering, a <w:num> includes <w:lvlOverride>/<w:startOverride>:
#
#   <w:num w:numId="3">
#     <w:abstractNumId w:val="0"/>
#     <w:lvlOverride w:ilvl="0">
#       <w:startOverride w:val="1"/>
#     </w:lvlOverride>
#   </w:num>
#
#
# INDENTATION: THE FLUSH-LEFT PROBLEM
# ====================================
#
# Standard Word list indentation uses a "hanging indent":
#
#   <w:ind w:left="720" w:hanging="360"/>
#
#   |        •  Text starts here
#   |           and wrapped lines align here.
#
# This looks fine normally, but if your paragraph style has a BORDER or
# FILL, the indent in the numbering definition overrides the style and
# pushes the entire content area rightward — breaking the border.
#
# Our fix uses TWO layers that work together:
#
# 1. The numbering definition (<w:lvl>) has ZERO indent:
#
#      <w:ind w:left="0" w:hanging="0"/>
#      <w:tabs><w:tab w:val="left" w:pos="360"/></w:tabs>
#      <w:suff w:val="tab"/>
#
#    This keeps the numbering definition "clean" — it won't interfere
#    with any paragraph style's border or fill.
#
# 2. Each paragraph gets a <w:ind> injected directly into its own <w:pPr>:
#
#      <w:ind w:left="360" w:hanging="360"/>
#
#    Paragraph-level indent moves TEXT within the paragraph box but does
#    NOT move the box itself. Borders and fills stay at the margin.
#
#    Since left == hanging, the bullet sits at 360-360 = 0 (flush left),
#    and both text and wrapped lines align at 360 (the tab stop position).
#
# Result:
#
#   |•  Text starts here after the tab stop         |
#   |   and wrapped lines align here too.            |
#   ↑   ↑                                            |
#   |   text at 360 (w:left, also the tab stop)      |
#   bullet at 0 (flush with margin/border)           |
#
# Enable this with: doc <- list_setup(doc, tab_pos = 360L)
# Default (no list_setup call) uses standard Word hanging indents.
#
#
# HOW OFFICER FITS IN
# ===================
#
# officer::read_docx() unzips the .docx into a temp directory (x$package_dir).
# We use officer to add paragraphs normally, then immediately modify the XML
# of the paragraph officer just created to add <w:numPr> (and optionally
# <w:ind>). Your paragraph style is always preserved.
#
#
# XML2 AND XPATH QUICK REFERENCE
# ==============================
#
#   read_xml(path)             — read XML file into memory
#   write_xml(doc, path)       — write XML back to disk
#   as_xml_document(string)    — parse an XML string into a node
#   xml_find_all(doc, xpath)   — find all nodes matching an XPath query
#   xml_find_first(doc, xpath) — find the first matching node
#   xml_attr(node, name)       — get an attribute value (NA if missing)
#   xml_add_child(parent, kid) — add a child node inside parent
#   xml_add_sibling(node, sib) — add a sibling node next to node
#
# XPath patterns used in this file:
#
#   "w:abstractNum"                 — all abstractNum elements
#   "w:num"                         — all num elements
#   "w:pPr/w:numPr/w:numId"        — navigate: pPr → numPr → numId
#   "w:num[@w:numId='3']"           — num where numId attribute = 3
#
# ==============================================================================


# ---- Dependencies ------------------------------------------------------------

if (!requireNamespace("officer", quietly = TRUE))
  stop("Package 'officer' is required.")
if (!requireNamespace("xml2", quietly = TRUE))
  stop("Package 'xml2' is required.")


# ---- Constants ---------------------------------------------------------------

OOXML_NS      <- "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
OOXML_NS_DECL <- sprintf('xmlns:w="%s"', OOXML_NS)
BULLET_GLYPHS <- c("\u2022", "\u25e6", "\u25aa")  # •, ◦, ▪ — cycles per level


# ---- Pure functions: build XML strings ---------------------------------------
# These take parameters and return strings. No side effects, no file I/O.

#' Build XML for one indent level of a list definition.
#'
#' @param ilvl  Integer 0-8.
#' @param list_type "bullet" or "decimal".
#' @param tab_pos NULL for standard indent, or integer (twips) for flush-left.
.build_lvl_xml <- function(ilvl, list_type, tab_pos = NULL) {
  if (list_type == "bullet") {
    fmt  <- "bullet"
    text <- BULLET_GLYPHS[ilvl %% 3L + 1L]
  } else {
    fmt  <- "decimal"
    text <- sprintf("%%%d.", ilvl + 1L)
  }

  if (is.null(tab_pos)) {
    # STANDARD: each level indents 720 twips (0.5") further, bullet hangs 360.
    sprintf(
      paste0(
        '<w:lvl w:ilvl="%d">',
          '<w:start w:val="1"/>',
          '<w:numFmt w:val="%s"/>',
          '<w:lvlText w:val="%s"/>',
          '<w:lvlJc w:val="left"/>',
          '<w:pPr><w:ind w:left="%d" w:hanging="360"/></w:pPr>',
        '</w:lvl>'
      ),
      ilvl, fmt, text, 720L * (ilvl + 1L)
    )
  } else {
    # FLUSH-LEFT: zero indent on the numbering definition.
    # The tab stop positions text; paragraph-level <w:ind> (injected later
    # by .inject_num_pr) handles wrapped-line alignment.
    sprintf(
      paste0(
        '<w:lvl w:ilvl="%d">',
          '<w:start w:val="1"/>',
          '<w:numFmt w:val="%s"/>',
          '<w:lvlText w:val="%s"/>',
          '<w:suff w:val="tab"/>',
          '<w:lvlJc w:val="left"/>',
          '<w:pPr>',
            '<w:ind w:left="0" w:hanging="0"/>',
            '<w:tabs><w:tab w:val="left" w:pos="%d"/></w:tabs>',
          '</w:pPr>',
        '</w:lvl>'
      ),
      ilvl, fmt, text, as.integer(tab_pos) * (ilvl + 1L)
    )
  }
}

#' Build a complete <w:abstractNum>.
.build_abstract_num_xml <- function(abstract_num_id, list_type, tab_pos = NULL) {
  lvls <- vapply(0:8, .build_lvl_xml, character(1),
                 list_type = list_type, tab_pos = tab_pos)
  sprintf(
    '<w:abstractNum %s w:abstractNumId="%d"><w:multiLevelType w:val="multilevel"/>%s</w:abstractNum>',
    OOXML_NS_DECL, abstract_num_id, paste0(lvls, collapse = "")
  )
}

#' Build a <w:num> instance. Set restart = TRUE to force counter back to 1.
.build_num_xml <- function(num_id, abstract_num_id, restart = FALSE) {
  override <- if (restart) {
    '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>'
  } else ""
  sprintf(
    '<w:num %s w:numId="%d"><w:abstractNumId w:val="%d"/>%s</w:num>',
    OOXML_NS_DECL, num_id, abstract_num_id, override
  )
}

#' Build a <w:numPr> node (makes a paragraph a list item).
.build_num_pr_node <- function(num_id, ilvl) {
  xml2::read_xml(sprintf(
    '<w:numPr xmlns:w="%s"><w:ilvl w:val="%d"/><w:numId w:val="%d"/></w:numPr>',
    OOXML_NS, as.integer(ilvl), as.integer(num_id)
  ))
}

#' Build a <w:ind> node for paragraph-level indent (flush-left mode only).
.build_ind_node <- function(tab_pos, ilvl) {
  indent <- as.integer(tab_pos) * (as.integer(ilvl) + 1L)
  xml2::read_xml(sprintf(
    '<w:ind xmlns:w="%s" w:left="%d" w:hanging="%d"/>',
    OOXML_NS, indent, indent
  ))
}


# ---- File I/O: numbering.xml management -------------------------------------

.read_numbering_xml <- function(x) {
  xml2::read_xml(file.path(x$package_dir, "word", "numbering.xml"))
}

.write_numbering_xml <- function(x, doc) {
  xml2::write_xml(doc, file.path(x$package_dir, "word", "numbering.xml"))
}

.next_available_ids <- function(doc) {
  list(
    abstract_num_id = max(as.integer(xml2::xml_attr(
      xml2::xml_find_all(doc, "w:abstractNum"), "abstractNumId"))) + 1L,
    num_id = max(as.integer(xml2::xml_attr(
      xml2::xml_find_all(doc, "w:num"), "numId"))) + 1L
  )
}


# ---- State management --------------------------------------------------------
#
# We store tracking state on x$.list_state:
#
#   abstract_ids    — list(bullet = <id>, decimal = <id>)  [set once, never changes]
#   current_num_ids — list(bullet = <id>, decimal = <id>)  [changes on restart]
#   active_type     — "bullet", "decimal", or NULL          [tracks current list]
#   next_num_id     — integer counter for new <w:num> IDs
#   tab_pos         — NULL or integer (flush-left mode)

.init_list_state <- function(x, tab_pos = NULL) {
  doc <- .read_numbering_xml(x)
  ids <- .next_available_ids(doc)

  bullet_abs_id  <- ids$abstract_num_id
  decimal_abs_id <- ids$abstract_num_id + 1L
  bullet_num_id  <- ids$num_id
  decimal_num_id <- ids$num_id + 1L

  # Inject abstractNums.
  abs_nodes <- xml2::xml_find_all(doc, "w:abstractNum")
  last_abs <- abs_nodes[[length(abs_nodes)]]
  xml2::xml_add_sibling(last_abs, xml2::as_xml_document(
    .build_abstract_num_xml(bullet_abs_id, "bullet", tab_pos)))
  xml2::xml_add_sibling(last_abs, xml2::as_xml_document(
    .build_abstract_num_xml(decimal_abs_id, "decimal", tab_pos)))

  # Inject initial nums.
  num_nodes <- xml2::xml_find_all(doc, "w:num")
  last_num <- num_nodes[[length(num_nodes)]]
  xml2::xml_add_sibling(last_num, xml2::as_xml_document(
    .build_num_xml(bullet_num_id, bullet_abs_id)))
  xml2::xml_add_sibling(last_num, xml2::as_xml_document(
    .build_num_xml(decimal_num_id, decimal_abs_id)))

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

.ensure_init <- function(x) {
  if (is.null(x$.list_state)) x <- .init_list_state(x)
  x
}

.restart_num <- function(x, list_type) {
  doc <- .read_numbering_xml(x)
  num_nodes <- xml2::xml_find_all(doc, "w:num")
  new_id <- x$.list_state$next_num_id
  abs_id <- x$.list_state$abstract_ids[[list_type]]
  xml2::xml_add_sibling(
    num_nodes[[length(num_nodes)]],
    xml2::as_xml_document(.build_num_xml(new_id, abs_id, restart = TRUE)))
  .write_numbering_xml(x, doc)
  x$.list_state$current_num_ids[[list_type]] <- new_id
  x$.list_state$next_num_id <- new_id + 1L
  x
}


# ---- Paragraph XML injection ------------------------------------------------

#' Inject list numbering (and optional indent) into the current paragraph.
.inject_num_pr <- function(x, num_id, ilvl) {
  node <- officer::docx_current_block_xml(x)
  ppr  <- xml2::xml_find_first(node, "w:pPr")

  # Add <w:numPr> — this makes the paragraph a list item.
  xml2::xml_add_child(ppr, .build_num_pr_node(num_id, ilvl))

  # In flush-left mode, also add paragraph-level <w:ind>.
  # This indents the TEXT (and wrapped lines) without moving the paragraph
  # box — so borders and fills stay at the margin.
  tab_pos <- x$.list_state$tab_pos
  if (!is.null(tab_pos)) {
    xml2::xml_add_child(ppr, .build_ind_node(tab_pos, ilvl))
  }

  x
}

#' Decide whether to restart, then return the current num_id.
.prepare_list_item <- function(x, list_type) {
  list_type <- match.arg(list_type, c("bullet", "decimal"))
  x <- .ensure_init(x)

  if (is.null(x$.list_state$active_type) ||
      x$.list_state$active_type != list_type) {
    x <- .restart_num(x, list_type)
  }

  x$.list_state$active_type <- list_type
  list(x = x, num_id = x$.list_state$current_num_ids[[list_type]])
}


# ---- Public API --------------------------------------------------------------

#' Configure list formatting. Call ONCE before any list_add_* call.
#'
#' @param x An rdocx object.
#' @param tab_pos NULL = standard hanging indent (default).
#'   Integer (twips) = flush-left mode. The bullet sits at the left margin
#'   and text starts at tab_pos. 360 = 0.25 inch, a good default.
#'   Use this when your paragraph style has borders or fills.
list_setup <- function(x, tab_pos = NULL) {
  if (!is.null(x$.list_state)) {
    warning("list_setup() must be called before any list_add_* calls.")
    return(x)
  }
  .init_list_state(x, tab_pos = tab_pos)
}

#' Add a plain text paragraph as a list item.
list_add_par <- function(x, value, style = NULL, list_type = "bullet",
                         ilvl = 0L, pos = "after") {
  prep <- .prepare_list_item(x, list_type)
  x <- prep$x
  x <- officer::body_add_par(x, value, style = style, pos = pos)
  .inject_num_pr(x, prep$num_id, ilvl)
}

#' Add a formatted paragraph (fpar) as a list item.
list_add_fpar <- function(x, value, style = NULL, list_type = "bullet",
                          ilvl = 0L, pos = "after") {
  prep <- .prepare_list_item(x, list_type)
  x <- prep$x
  x <- officer::body_add_fpar(x, value, style = style, pos = pos)
  .inject_num_pr(x, prep$num_id, ilvl)
}

#' Add a block_list as list items (each paragraph gets numbered).
list_add_blocks <- function(x, blocks, list_type = "bullet",
                            ilvl = 0L, pos = "after") {
  prep <- .prepare_list_item(x, list_type)
  x <- prep$x
  x <- officer::body_add_blocks(x, blocks, pos = pos)

  # Tag each paragraph. Cursor is on the last one; walk backwards.
  n <- length(blocks)
  for (i in seq_len(n)) {
    x <- .inject_num_pr(x, prep$num_id, ilvl)
    if (i < n) x <- officer::cursor_backward(x)
  }
  for (i in seq_len(n - 1L)) x <- officer::cursor_forward(x)
  x
}

#' End the current list. Next list_add_* of the same type restarts at 1.
list_end <- function(x) {
  if (!is.null(x$.list_state)) x$.list_state$active_type <- NULL
  x
}

#' Print a summary of all numbering definitions in the document.
list_inspect <- function(x) {
  doc <- .read_numbering_xml(x)

  cat("=== abstractNum (format templates) ===\n\n")
  for (node in xml2::xml_find_all(doc, "w:abstractNum")) {
    cat(sprintf("  abstractNumId = %s\n", xml2::xml_attr(node, "abstractNumId")))
    for (lvl in xml2::xml_find_all(node, "w:lvl")) {
      fmt  <- xml2::xml_attr(xml2::xml_find_first(lvl, "w:numFmt"), "val")
      text <- xml2::xml_attr(xml2::xml_find_first(lvl, "w:lvlText"), "val")
      ind  <- xml2::xml_find_first(lvl, "w:pPr/w:ind")
      tab  <- xml2::xml_find_first(lvl, "w:pPr/w:tabs/w:tab")
      tab_info <- if (!inherits(tab, "xml_missing"))
        sprintf("  tab=%s", xml2::xml_attr(tab, "pos")) else ""
      cat(sprintf("    lvl %s: %s %s  left=%s hang=%s%s\n",
                  xml2::xml_attr(lvl, "ilvl"), fmt, text,
                  xml2::xml_attr(ind, "left"),
                  xml2::xml_attr(ind, "hanging"), tab_info))
    }
    cat("\n")
  }

  cat("=== num (list instances) ===\n\n")
  for (node in xml2::xml_find_all(doc, "w:num")) {
    ov <- xml2::xml_find_first(node, "w:lvlOverride/w:startOverride")
    restart <- if (!inherits(ov, "xml_missing"))
      sprintf("  [restarts at %s]", xml2::xml_attr(ov, "val")) else ""
    cat(sprintf("  numId=%s -> abstractNumId=%s%s\n",
                xml2::xml_attr(node, "numId"),
                xml2::xml_attr(xml2::xml_find_first(node, "w:abstractNumId"), "val"),
                restart))
  }
  cat("\n")
  invisible(NULL)
}
