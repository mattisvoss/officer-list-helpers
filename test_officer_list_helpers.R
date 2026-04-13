# ==============================================================================
# test_officer_list_helpers.R
#
# Run:  Rscript test_officer_list_helpers.R
#
# Each test demonstrates a usage pattern AND verifies it works correctly.
# Read them top-to-bottom as a guide to the API.
# ==============================================================================

library(officer)
library(xml2)
source("officer_list_helpers.R")

test_count <- 0L
pass_count <- 0L

run_test <- function(name, expr) {
  test_count <<- test_count + 1L
  cat(sprintf("  [%2d] %s ... ", test_count, name))
  tryCatch({
    force(expr)
    pass_count <<- pass_count + 1L
    cat("PASS\n")
  }, error = function(e) {
    cat("FAIL\n")
    stop(sprintf("'%s' failed: %s", name, conditionMessage(e)), call. = FALSE)
  })
}

assert <- function(cond, msg = "assertion failed") {
  if (!isTRUE(cond)) stop(msg, call. = FALSE)
}

# Helper: get numId and ilvl from the paragraph at the cursor.
get_num_pr <- function(x) {
  node <- docx_current_block_xml(x)
  list(
    num_id = xml_attr(xml_find_first(node, "w:pPr/w:numPr/w:numId"), "val"),
    ilvl   = xml_attr(xml_find_first(node, "w:pPr/w:numPr/w:ilvl"), "val")
  )
}

# Helper: get paragraph-level indent from the paragraph at the cursor.
get_ind <- function(x) {
  node <- docx_current_block_xml(x)
  ind  <- xml_find_first(node, "w:pPr/w:ind")
  list(
    left    = xml_attr(ind, "left"),
    hanging = xml_attr(ind, "hanging")
  )
}


cat("\nRunning officer_list_helpers tests\n")
cat(strrep("-", 50), "\n")


# --- 1. Basic bullet list ---
# The simplest usage: add bullet items to a document.

run_test("bullet list", {
  doc <- read_docx()
  doc <- list_add_par(doc, "Item one", list_type = "bullet")
  doc <- list_add_par(doc, "Item two", list_type = "bullet")

  np <- get_num_pr(doc)
  assert(!is.na(np$num_id), "should have numId")
  assert(identical(np$ilvl, "0"), "should be indent level 0")
})


# --- 2. Numbered list ---
# Same API, just change list_type.

run_test("numbered list", {
  doc <- read_docx()
  doc <- list_add_par(doc, "Step one", list_type = "decimal")
  doc <- list_add_par(doc, "Step two", list_type = "decimal")

  np <- get_num_pr(doc)
  assert(!is.na(np$num_id))
})


# --- 3. Numbered list restart ---
# list_end() between two same-type lists forces the second to start at 1.

run_test("list_end restarts numbering", {
  doc <- read_docx()
  doc <- list_add_par(doc, "A", list_type = "decimal")
  id_before <- get_num_pr(doc)$num_id

  doc <- list_end(doc)
  doc <- list_add_par(doc, "B", list_type = "decimal")
  id_after <- get_num_pr(doc)$num_id

  # Different numId = different list = numbering restarts.
  assert(id_before != id_after, "should get a new numId after list_end")
})


# --- 4. Switching types auto-restarts ---
# Going from bullet to decimal (or vice versa) automatically starts a new list.

run_test("switching types auto-restarts", {
  doc <- read_docx()
  doc <- list_add_par(doc, "Bullet", list_type = "bullet")
  id1 <- get_num_pr(doc)$num_id

  doc <- list_add_par(doc, "Number", list_type = "decimal")
  id2 <- get_num_pr(doc)$num_id

  assert(id1 != id2, "bullet and decimal should have different numIds")
})


# --- 5. Continuation ---
# Consecutive same-type calls share the same numId (one continuous list).

run_test("consecutive calls continue the same list", {
  doc <- read_docx()
  doc <- list_add_par(doc, "One", list_type = "decimal")
  id1 <- get_num_pr(doc)$num_id

  doc <- list_add_par(doc, "Two", list_type = "decimal")
  id2 <- get_num_pr(doc)$num_id

  assert(identical(id1, id2), "should share the same numId")
})


# --- 6. Preserves custom paragraph style ---
# Your template style is kept; list formatting is layered on top.

run_test("custom style preserved", {
  doc <- read_docx()
  doc <- list_add_par(doc, "Styled bullet", style = "heading 2",
                      list_type = "bullet")

  node <- docx_current_block_xml(doc)
  assert(!is.na(xml_attr(xml_find_first(node, "w:pPr/w:pStyle"), "val")),
         "should keep the paragraph style")
  assert(!is.na(get_num_pr(doc)$num_id),
         "should also have numPr")
})


# --- 7. Formatted paragraphs (fpar) ---
# Use list_add_fpar for mixed bold/italic within a list item.

run_test("list_add_fpar with formatted text", {
  doc <- read_docx()
  fp <- fpar(ftext("Bold", prop = fp_text(bold = TRUE)), ftext(" normal"))
  doc <- list_add_fpar(doc, fp, list_type = "bullet")

  assert(!is.na(get_num_pr(doc)$num_id))
})


# --- 8. Flush-left mode: numbering definition ---
# list_setup(tab_pos=360) creates numbering definitions with zero indent.

run_test("flush-left: numbering definition has zero indent", {
  doc <- read_docx()
  doc <- list_setup(doc, tab_pos = 360L)
  doc <- list_add_par(doc, "Flush", list_type = "bullet")

  # Find our bullet abstractNum — it should have left=0, hanging=0, tab=360.
  num_doc <- .read_numbering_xml(doc)
  for (node in xml_find_all(num_doc, "w:abstractNum")) {
    lvl0 <- xml_find_first(node, "w:lvl[@w:ilvl='0']")
    if (inherits(lvl0, "xml_missing")) next
    if (xml_attr(xml_find_first(lvl0, "w:numFmt"), "val") != "bullet") next

    ind <- xml_find_first(lvl0, "w:pPr/w:ind")
    tab <- xml_find_first(lvl0, "w:pPr/w:tabs/w:tab")
    if (inherits(tab, "xml_missing")) next

    assert(identical(xml_attr(ind, "left"), "0"), "left should be 0")
    assert(identical(xml_attr(ind, "hanging"), "0"), "hanging should be 0")
    assert(identical(xml_attr(tab, "pos"), "360"), "tab should be 360")
    return(invisible(NULL))  # found and verified
  }
  stop("did not find flush-left bullet abstractNum")
})


# --- 9. Flush-left mode: paragraph-level indent ---
# Each list paragraph gets <w:ind> with left == hanging == tab_pos.
# This keeps the bullet at 0 and block-indents the text.

run_test("flush-left: paragraph gets w:ind for block indent", {
  doc <- read_docx()
  doc <- list_setup(doc, tab_pos = 360L)
  doc <- list_add_par(doc, "Flush", list_type = "bullet")

  ind <- get_ind(doc)
  assert(identical(ind$left, "360"), "paragraph left should be 360")
  assert(identical(ind$hanging, "360"), "paragraph hanging should be 360")
})


# --- 10. Round-trip: save and reload ---
# Verifies our XML doesn't corrupt the document.

run_test("document round-trips without corruption", {
  doc <- read_docx()
  doc <- list_setup(doc, tab_pos = 360L)
  doc <- list_add_par(doc, "Bullet", list_type = "bullet")
  doc <- list_end(doc)
  doc <- list_add_par(doc, "Number", list_type = "decimal")

  tf <- tempfile(fileext = ".docx")
  print(doc, target = tf)
  doc2 <- read_docx(tf)  # would error if XML is malformed
  assert(nrow(docx_summary(doc2)) >= 2)
  unlink(tf)
})


# ---- Summary -----------------------------------------------------------------

cat(strrep("-", 50), "\n")
cat(sprintf("\n%d / %d passed.", pass_count, test_count))
if (pass_count == test_count) {
  cat(sprintf("  ALL %d TESTS PASSED\n\n", test_count))
} else {
  stop("Some tests failed.", call. = FALSE)
}
