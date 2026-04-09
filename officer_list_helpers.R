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
#   5. CODE: Constants
#   6. CODE: XML string builders (pure functions, no side effects)
#   7. CODE: Numbering.xml file management (reads/writes the XML file)
#   8. CODE: Public API (the functions you actually call)
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
#     Settings that apply to the whole paragraph: which style to use,
#     text alignment, indentation, and — crucially for us — list numbering.
#     This is where we inject <w:numPr> to make a paragraph a list item.
#
# The "w:" prefix on every tag is an XML NAMESPACE. All Word XML elements
# belong to the namespace:
#
#   http://schemas.openxmlformats.org/wordprocessingml/2006/main
#
# The "w:" is just a shorthand alias for that long URL. Think of it like a
# module prefix — it says "this element is defined by the Word spec, not by
# some other XML format." Every tool that reads or writes these files needs
# to know about this namespace, which is why you'll see it referenced
# throughout this code.
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
#                                Returns a special "xml_missing" object if
#                                nothing matches (check with inherits()).
#   xml_attr(node, name)         Get the value of an attribute. Returns NA if
#                                the attribute doesn't exist.
#   xml_add_child(parent, child) Insert a child node INSIDE a parent element.
#   xml_add_sibling(node, sib)   Insert a node NEXT TO an existing node
#                                (as a sibling, not inside it).
#
# XPATH is a query language for finding things in XML. It works like file
# paths, but for XML elements instead of folders. Here are the patterns
# used in this file:
#
#   "w:abstractNum"
#       Find all <w:abstractNum> elements (direct children of the root).
#
#   "w:abstractNum/w:lvl"
#       Find <w:lvl> elements that are children of <w:abstractNum>.
#       The "/" means "child of" — like a folder separator.
#
#   "w:num[@w:numId='3']"
#       Find the <w:num> element whose numId attribute equals 3.
#       The [@...] syntax filters by attribute value.
#
#   "w:pPr/w:numPr/w:numId"
#       Navigate a chain: pPr -> numPr -> numId.
#       Each "/" goes one level deeper into the tree.
#
#   "w:abstractNum/w:lvl/w:numFmt[@w:val='bullet']"
#       Find numFmt elements with val="bullet" that are inside lvl elements
#       that are inside abstractNum elements. Combines path navigation
#       with attribute filtering.
#
#   "w:num[@w:numId='3']/w:lvlOverride/w:startOverride[@w:val='1']"
#       Find startOverride elements (with val=1) inside lvlOverride, inside
#       a specific num (with numId=3). This is how we verify that a restart
#       override was correctly created.
#
# That's all the XPath you need to understand this code.
#
#
# ==========================
# 4. HOW OFFICER FITS IN
# ==========================
#
# The 'officer' R package provides a high-level API for building .docx files.
# When you call read_docx(), officer unzips the template .docx into a
# temporary directory and gives you an R object (class "rdocx") that
# represents the document in memory.
#
# Key officer functions and concepts used in this file:
#
#   read_docx(path)
#       Open a .docx template. Returns an rdocx object. The template is
#       unzipped into a temp folder — x$package_dir is the path to that
#       folder. This is where we find word/numbering.xml.
#
#   body_add_par(x, text, style, pos)
#       Add a plain text paragraph to the document. The cursor moves to
#       the new paragraph.
#
#   body_add_fpar(x, fpar_obj, style, pos)
#       Add a formatted paragraph. fpar() and ftext() let you mix bold,
#       italic, and other formatting within a single paragraph.
#
#   body_add_blocks(x, block_list, pos)
#       Add multiple formatted paragraphs at once.
#
#   docx_current_block_xml(x)
#       Get the XML node at the current cursor position. The cursor always
#       points at the most recently added paragraph. This is how we grab
#       the paragraph XML to inject <w:numPr>.
#
#   cursor_backward(x) / cursor_forward(x)
#       Move the cursor to the previous/next paragraph. We use these in
#       list_add_blocks() to tag each paragraph individually.
#
#   print(x, target = path)
#       Save the document. Officer re-zips the temp directory (including
#       our modified numbering.xml) into a .docx file.
#
# OUR STRATEGY:
#
#   1. Let officer add the paragraph normally (preserving your custom style).
#   2. Immediately grab the XML of that paragraph via docx_current_block_xml().
#   3. Find its <w:pPr> and add <w:numPr> inside it.
#
#   This means your paragraph style is always preserved — we just layer
#   list formatting on top. officer doesn't know or care about our changes;
#   it just includes them when it saves the file.
#
# ==============================================================================


# ---- Dependency check --------------------------------------------------------
# These packages must be loaded before sourcing this file.
# We check rather than call library() to avoid side effects.

if (!requireNamespace("officer", quietly = TRUE)) {
  stop("Package 'officer' is required. Install it with: install.packages('officer')")
}
if (!requireNamespace("xml2", quietly = TRUE)) {
  stop("Package 'xml2' is required. Install it with: install.packages('xml2')")
}


# ==============================================================================
# SECTION 5: CONSTANTS
# ==============================================================================

# The XML namespace for all Word document elements.
# This is the long URL that the "w:" prefix stands for in every .docx XML file.
OOXML_NS <- "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# When building XML strings to be parsed by xml2, we need to declare the
# namespace inline. This string gets pasted into our XML fragments so that
# xml2's parser knows what "w:" means.
OOXML_NS_DECL <- sprintf('xmlns:w="%s"', OOXML_NS)

# Bullet characters that cycle across indent levels: •, ◦, ▪, •, ◦, ▪, ...
# These match Word's built-in bullet defaults.
BULLET_GLYPHS <- c("\u2022", "\u25e6", "\u25aa")


# ==============================================================================
# SECTION 6: XML STRING BUILDERS
#
# These functions build XML strings. They are PURE FUNCTIONS — they don't
# read or write any files, they don't modify any state, and they don't depend
# on any external context. They just take parameters and return strings.
#
# This makes them easy to test, easy to understand, and easy to modify if
# you want to change the list formatting (e.g. different bullet glyphs or
# different indentation).
# ==============================================================================

#' Build the XML for one indent level of a list format definition.
#'
#' Each list format (abstractNum) has up to 9 indent levels (0 through 8).
#' This function builds the XML for one of those levels.
#'
#' @param ilvl Integer 0-8. Which indent level this defines.
#' @param list_type "bullet" or "decimal".
#' @return An XML string: '<w:lvl w:ilvl="0">...</w:lvl>'.
.build_lvl_xml <- function(ilvl, list_type) {
  # Each level indents 0.5 inch (720 twips) further than the last.
  # Level 0 = 720 twips from left margin, level 1 = 1440 twips, etc.
  # The "hanging" indent (360 twips) is how far the bullet/number hangs
  # to the left of the text — this is what creates the visual alignment
  # where the bullet sticks out and the text is neatly indented.
  left_indent <- 720L * (ilvl + 1L)

  if (list_type == "bullet") {
    fmt <- "bullet"
    # Cycle through bullet glyphs: level 0 = •, level 1 = ◦, level 2 = ▪,
    # then it repeats: level 3 = •, level 4 = ◦, etc.
    text <- BULLET_GLYPHS[ilvl %% 3L + 1L]
  } else {
    fmt <- "decimal"
    # The %N token means "insert the counter for level N".
    # %1. at level 0 produces "1.", "2.", "3."
    # %2. at level 1 produces "1.", "2.", "3." (independently)
    text <- sprintf("%%%d.", ilvl + 1L)
  }

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
}


#' Build a complete <w:abstractNum> — a list format template.
#'
#' An abstractNum defines HOW a list looks (bullet glyphs, decimal numbers,
#' indentation) across all 9 possible indent levels. It does NOT create an
#' actual list in the document — for that, you need a <w:num> instance that
#' references this abstractNum.
#'
#' @param abstract_num_id Integer. Unique ID (must not collide with existing).
#' @param list_type "bullet" or "decimal".
#' @return A complete XML string ready to be injected into numbering.xml.
.build_abstract_num_xml <- function(abstract_num_id, list_type) {
  # Build all 9 levels (0 through 8).
  lvls <- vapply(0:8, .build_lvl_xml, character(1), list_type = list_type)

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
#' A <w:num> is a specific, concrete list in the document. Paragraphs point
#' to it via their <w:numPr> element. It references an <w:abstractNum> for
#' its visual formatting.
#'
#' You can have multiple <w:num> instances pointing to the same abstractNum.
#' Each one maintains its own counter. This is how you create two separate
#' "1. 2. 3." lists that look identical but number independently.
#'
#' @param num_id Integer. Unique ID for this list instance.
#' @param abstract_num_id Integer. Which format template to use.
#' @param restart Logical. If TRUE, adds <w:lvlOverride>/<w:startOverride>
#'   to force the counter back to 1. Use this when creating a new list that
#'   should start fresh despite sharing a format with a previous list.
#' @return A complete XML string.
.build_num_xml <- function(num_id, abstract_num_id, restart = FALSE) {
  override <- ""
  if (restart) {
    # This element tells Word: "ignore the running counter from any previous
    # lists that use this same abstractNum — start fresh at 1."
    # Without this, some renderers (LibreOffice, WPS) continue counting.
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
#' This gets added inside a paragraph's <w:pPr> (paragraph properties).
#' It says: "this paragraph belongs to list instance [num_id] at indent
#' level [ilvl]."
#'
#' @param num_id Integer. Which <w:num> instance this paragraph belongs to.
#' @param ilvl Integer. Indent level (0 = top, 1 = sub-item, 2 = sub-sub).
#' @return An xml2 node object (not a string), ready for xml_add_child().
.build_num_pr_node <- function(num_id, ilvl) {
  # We return a parsed xml2 node (not a string) because xml_add_child()
  # expects a node object when adding children to an existing XML tree.
  xml2::read_xml(sprintf(
    '<w:numPr xmlns:w="%s"><w:ilvl w:val="%d"/><w:numId w:val="%d"/></w:numPr>',
    OOXML_NS, as.integer(ilvl), as.integer(num_id)
  ))
}


# ==============================================================================
# SECTION 7: NUMBERING.XML FILE MANAGEMENT
#
# These functions read and write word/numbering.xml inside the temporary
# directory where officer unpacked the .docx template. They also manage the
# tracking state stored on x$.list_state.
#
# The state tracks:
#   - Which abstractNum IDs we created (so we can make new num instances)
#   - Which num instance is currently active for each list type
#   - Which list type was used most recently (for restart detection)
#   - A counter for allocating new unique num IDs
# ==============================================================================

#' Read numbering.xml from the document's temp directory.
#'
#' officer::read_docx() unzips the .docx template into a temp folder.
#' x$package_dir is the path to that folder. numbering.xml lives at
#' word/numbering.xml inside it.
#'
#' @param x An rdocx object.
#' @return An xml2 document object representing numbering.xml.
.read_numbering_xml <- function(x) {
  xml2::read_xml(file.path(x$package_dir, "word", "numbering.xml"))
}


#' Write numbering.xml back to the document's temp directory.
#'
#' After modifying the XML in memory, we write it back so that officer
#' includes our changes when it saves the final .docx.
#'
#' @param x An rdocx object.
#' @param doc An xml2 document object (the modified numbering.xml).
.write_numbering_xml <- function(x, doc) {
  xml2::write_xml(doc, file = file.path(x$package_dir, "word", "numbering.xml"))
}


#' Find the next available IDs in numbering.xml.
#'
#' Both abstractNum and num elements need unique integer IDs. This reads
#' all existing IDs in the file and returns values that won't collide.
#'
#' @param doc An xml2 document (numbering.xml).
#' @return A list: $abstract_num_id and $num_id (both integers).
.next_available_ids <- function(doc) {
  # XPath "w:abstractNum" finds all <w:abstractNum> elements.
  # xml_attr(..., "abstractNumId") extracts the ID from each one.
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
#' Called automatically the first time you use any list_add_* function.
#' Creates one bullet and one decimal <w:abstractNum> (format template),
#' plus one <w:num> instance for each, and stores tracking state on the
#' document object at x$.list_state.
#'
#' THE STATE OBJECT (x$.list_state) contains:
#'
#'   abstract_ids    — list(bullet = 10, decimal = 11)
#'       Which abstractNum ID to use for each list type.
#'       Set once at init; never changes.
#'
#'   current_num_ids — list(bullet = 12, decimal = 13)
#'       The currently active <w:num> instance for each type.
#'       Changes every time we restart a list (new instance created).
#'
#'   active_type     — "bullet", "decimal", or NULL
#'       Which list type was used most recently.
#'       NULL means no list is active (list_end() was called, or we
#'       haven't started yet). This is how we detect when to restart.
#'
#'   next_num_id     — Integer
#'       Counter for allocating new unique <w:num> IDs.
#'
#' @param x An rdocx object.
#' @return The rdocx object with x$.list_state initialized.
.init_list_state <- function(x) {
  doc <- .read_numbering_xml(x)
  ids <- .next_available_ids(doc)

  # Allocate IDs for our two format templates and two initial instances.
  bullet_abs_id  <- ids$abstract_num_id
  decimal_abs_id <- ids$abstract_num_id + 1L
  bullet_num_id  <- ids$num_id
  decimal_num_id <- ids$num_id + 1L

  # Inject the two format templates into numbering.xml.
  # xml_add_sibling() inserts the new node right after the reference node.
  abs_nodes <- xml2::xml_find_all(doc, "w:abstractNum")
  last_abs <- abs_nodes[[length(abs_nodes)]]
  xml2::xml_add_sibling(last_abs, xml2::as_xml_document(
    .build_abstract_num_xml(bullet_abs_id, "bullet")
  ))
  xml2::xml_add_sibling(last_abs, xml2::as_xml_document(
    .build_abstract_num_xml(decimal_abs_id, "decimal")
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

  # Write the modified numbering.xml back to disk.
  .write_numbering_xml(x, doc)

  # Store our tracking state on the document object.
  x$.list_state <- list(
    abstract_ids    = list(bullet = bullet_abs_id, decimal = decimal_abs_id),
    current_num_ids = list(bullet = bullet_num_id, decimal = decimal_num_id),
    active_type     = NULL,
    next_num_id     = decimal_num_id + 1L
  )

  x
}


#' Ensure the list system is initialized. No-op if already done.
#'
#' @param x An rdocx object.
#' @return The rdocx object (with .list_state set if it wasn't already).
.ensure_init <- function(x) {
  if (is.null(x$.list_state)) {
    x <- .init_list_state(x)
  }
  x
}


#' Create a new list instance (<w:num>) with a restart override.
#'
#' Called when we need numbering to restart at 1. The new <w:num> uses the
#' same format template (abstractNum) as before, but includes
#' <w:startOverride> to reset the counter.
#'
#' @param x An rdocx object with .list_state initialized.
#' @param list_type "bullet" or "decimal".
#' @return The rdocx object with updated current_num_ids and next_num_id.
.restart_num <- function(x, list_type) {
  doc <- .read_numbering_xml(x)

  # Find all <w:num> nodes so we can add our new one after the last.
  num_nodes <- xml2::xml_find_all(doc, "w:num")

  new_id <- x$.list_state$next_num_id
  abs_id <- x$.list_state$abstract_ids[[list_type]]

  # Build a <w:num> with restart = TRUE (includes <w:startOverride>).
  xml2::xml_add_sibling(
    num_nodes[[length(num_nodes)]],
    xml2::as_xml_document(.build_num_xml(new_id, abs_id, restart = TRUE))
  )
  .write_numbering_xml(x, doc)

  # Update our tracking state.
  x$.list_state$current_num_ids[[list_type]] <- new_id
  x$.list_state$next_num_id <- new_id + 1L
  x
}


#' Inject <w:numPr> into the paragraph at the current cursor position.
#'
#' officer's cursor always points at the most recently added paragraph.
#' docx_current_block_xml(x) returns the XML node for that paragraph.
#' We find its <w:pPr> child and add our <w:numPr> inside it.
#'
#' @param x An rdocx object (cursor must be on the target paragraph).
#' @param num_id Integer. Which list instance to reference.
#' @param ilvl Integer. Indent level (0 = top, 1 = sub-item, etc.)
#' @return The rdocx object (XML modified in place).
.inject_num_pr <- function(x, num_id, ilvl) {
  # Get the XML node for the paragraph officer just added.
  node <- officer::docx_current_block_xml(x)

  # XPath "w:pPr" finds the <w:pPr> child of this paragraph.
  # Every paragraph has one — officer always creates it.
  ppr <- xml2::xml_find_first(node, "w:pPr")

  # xml_add_child() inserts our <w:numPr> node inside <w:pPr>.
  xml2::xml_add_child(ppr, .build_num_pr_node(num_id, ilvl))

  x
}


#' Shared setup logic for all list_add_* functions.
#'
#' Handles three things:
#'   1. Lazy initialization (first call creates the numbering definitions)
#'   2. Restart detection (type changed, or list_end() was called)
#'   3. State tracking (records which type is currently active)
#'
#' RESTART LOGIC:
#'   - active_type is NULL → no active list → create new num (restart)
#'   - active_type != list_type → type changed → create new num (restart)
#'   - active_type == list_type → same list continues → reuse current num
#'
#' @param x An rdocx object.
#' @param list_type "bullet" or "decimal".
#' @return A list: $x (updated rdocx) and $num_id (integer to use).
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
# SECTION 8: PUBLIC API
#
# These are the functions you call in your code. Each one follows the same
# pattern:
#
#   1. Call .prepare_list_item() to handle init/restart/state
#   2. Use an officer function to add the paragraph (preserving your style)
#   3. Call .inject_num_pr() to add list formatting to that paragraph's XML
#
# Your paragraph style is ALWAYS preserved — list formatting is layered on top.
# ==============================================================================

#' Add a plain text paragraph as a list item.
#'
#' Works exactly like officer::body_add_par(), but the paragraph also gets
#' bullet or numbered list formatting.
#'
#' @param x An rdocx object (from officer::read_docx()).
#' @param value Character string. The paragraph text.
#' @param style Paragraph style name from your template (e.g. "Normal",
#'   "My Custom Style"). NULL uses the document's default paragraph style.
#' @param list_type "bullet" for bullet points, "decimal" for numbered lists.
#' @param ilvl Indent level: 0 = top level, 1 = sub-item, 2 = sub-sub-item.
#' @param pos Where to insert: "after" (default), "before", or "on".
#' @return The modified rdocx object.
#'
#' @examples
#' doc <- read_docx()
#' doc <- list_add_par(doc, "Buy groceries", list_type = "bullet")
#' doc <- list_add_par(doc, "Milk",          list_type = "bullet", ilvl = 1L)
#' doc <- list_add_par(doc, "Eggs",          list_type = "bullet", ilvl = 1L)
#' doc <- list_end(doc)
#' doc <- list_add_par(doc, "First step",    list_type = "decimal")
#' doc <- list_add_par(doc, "Second step",   list_type = "decimal")
#' print(doc, target = tempfile(fileext = ".docx"))
list_add_par <- function(x, value, style = NULL, list_type = "bullet",
                         ilvl = 0L, pos = "after") {
  prep <- .prepare_list_item(x, list_type)
  x <- prep$x
  x <- officer::body_add_par(x, value, style = style, pos = pos)
  x <- .inject_num_pr(x, prep$num_id, ilvl)
  x
}


#' Add a formatted paragraph as a list item.
#'
#' Use this when you need mixed formatting within one paragraph — bold words,
#' italic phrases, hyperlinks, etc. Build your content with officer::fpar()
#' and officer::ftext().
#'
#' @param x An rdocx object.
#' @param value An fpar object (from officer::fpar()).
#' @param style Paragraph style name. NULL uses the document default.
#' @param list_type "bullet" or "decimal".
#' @param ilvl Indent level (default 0).
#' @param pos "after" (default), "before", or "on".
#' @return The modified rdocx object.
#'
#' @examples
#' doc <- read_docx()
#' formatted <- fpar(
#'   ftext("Important: ", prop = fp_text(bold = TRUE)),
#'   ftext("remember to buy milk")
#' )
#' doc <- list_add_fpar(doc, formatted, list_type = "bullet")
#' print(doc, target = tempfile(fileext = ".docx"))
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
#' Each paragraph in the block_list becomes a separate list item. All items
#' share the same list sequence (continuous numbering, no restart between).
#'
#' @param x An rdocx object.
#' @param blocks A block_list object (from officer::block_list()).
#' @param list_type "bullet" or "decimal".
#' @param ilvl Indent level applied to ALL items (default 0).
#' @param pos "after" (default), "before", or "on".
#' @return The modified rdocx object.
#'
#' @examples
#' doc <- read_docx()
#' items <- block_list(
#'   fpar(ftext("First point")),
#'   fpar(ftext("Second point")),
#'   fpar(ftext("Third point"))
#' )
#' doc <- list_add_blocks(doc, items, list_type = "decimal")
#' print(doc, target = tempfile(fileext = ".docx"))
list_add_blocks <- function(x, blocks, list_type = "bullet",
                            ilvl = 0L, pos = "after") {
  prep <- .prepare_list_item(x, list_type)
  x <- prep$x
  x <- officer::body_add_blocks(x, blocks, pos = pos)

  # body_add_blocks adds multiple paragraphs. The cursor ends up on the last
  # one. We need to tag EACH paragraph, so we walk backwards, inject numPr
  # into each, then walk forwards to restore the cursor position.
  n <- length(blocks)
  for (i in seq_len(n)) {
    x <- .inject_num_pr(x, prep$num_id, ilvl)
    if (i < n) {
      x <- officer::cursor_backward(x)
    }
  }
  # Restore the cursor to the last paragraph.
  for (i in seq_len(n - 1L)) {
    x <- officer::cursor_forward(x)
  }

  x
}


#' End the current list.
#'
#' Call this between two list sequences of the SAME type to force the next
#' one to restart numbering at 1.
#'
#' You do NOT need this when switching types. Switching from bullet to
#' decimal (or vice versa) restarts automatically.
#'
#' Safe to call multiple times, or when no list is active.
#'
#' @param x An rdocx object.
#' @return The rdocx object.
#'
#' @examples
#' doc <- read_docx()
#' doc <- list_add_par(doc, "A", list_type = "decimal")  # renders as 1.
#' doc <- list_add_par(doc, "B", list_type = "decimal")  # renders as 2.
#' doc <- list_end(doc)
#' doc <- list_add_par(doc, "C", list_type = "decimal")  # renders as 1.
#' doc <- list_add_par(doc, "D", list_type = "decimal")  # renders as 2.
#' print(doc, target = tempfile(fileext = ".docx"))
list_end <- function(x) {
  if (!is.null(x$.list_state)) {
    x$.list_state$active_type <- NULL
  }
  x
}


#' Inspect the numbering definitions in a document.
#'
#' Prints a readable summary of every list format template (abstractNum)
#' and list instance (num) in the document. Useful for:
#'   - Understanding what your template provides out of the box
#'   - Debugging: seeing what this library has created
#'   - Learning: seeing the OOXML structure in plain English
#'
#' @param x An rdocx object.
#' @return NULL (called for the side effect of printing).
#'
#' @examples
#' doc <- read_docx()
#' doc <- list_add_par(doc, "test", list_type = "bullet")
#' list_inspect(doc)
list_inspect <- function(x) {
  doc <- .read_numbering_xml(x)

  # --- Format templates ---
  cat("=== Format templates (abstractNum) ===\n")
  cat("These define HOW a list looks at each indent level.\n\n")

  # XPath "w:abstractNum" finds all <w:abstractNum> elements in the file.
  abs_nodes <- xml2::xml_find_all(doc, "w:abstractNum")
  for (node in abs_nodes) {
    abs_id <- xml2::xml_attr(node, "abstractNumId")
    cat(sprintf("  abstractNumId = %s\n", abs_id))

    # XPath "w:lvl" finds all <w:lvl> children of this abstractNum.
    lvls <- xml2::xml_find_all(node, "w:lvl")
    for (lvl in lvls) {
      # XPath "w:numFmt" finds the format element; xml_attr gets its "val".
      fmt  <- xml2::xml_attr(xml2::xml_find_first(lvl, "w:numFmt"), "val")
      text <- xml2::xml_attr(xml2::xml_find_first(lvl, "w:lvlText"), "val")
      cat(sprintf("    level %s: format = %-10s  display = %s\n",
                  xml2::xml_attr(lvl, "ilvl"), fmt, text))
    }
    cat("\n")
  }

  # --- List instances ---
  cat("=== List instances (num) ===\n")
  cat("Each num is an active list. Paragraphs point to these via numId.\n\n")

  # XPath "w:num" finds all <w:num> elements.
  num_nodes <- xml2::xml_find_all(doc, "w:num")
  for (node in num_nodes) {
    # Check if this num has a restart override.
    # XPath "w:lvlOverride/w:startOverride" navigates two levels deep.
    # xml_find_first returns an "xml_missing" object if nothing matches.
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
