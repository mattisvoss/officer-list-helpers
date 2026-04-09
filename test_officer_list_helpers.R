# ==============================================================================
# test_officer_list_helpers.R
#
# Tests for officer_list_helpers.R.
#
# HOW TO RUN
# ==========
#
#   Put this file in the same directory as officer_list_helpers.R, then:
#
#     Rscript test_officer_list_helpers.R
#
#   Each test creates a temporary document in memory — nothing is saved to
#   disk. If all tests pass you'll see "ALL 12 TESTS PASSED". If any test
#   fails, the script stops immediately with an error message explaining
#   what went wrong.
#
#
# WHAT EACH TEST VERIFIES
# =======================
#
#    1. Bullet items get <w:numPr> and numbering.xml has a bullet format
#    2. Decimal items get <w:numPr> and numbering.xml has a decimal format
#    3. The ilvl parameter produces the correct indent level in the XML
#    4. list_end() causes a new <w:num> with <w:startOverride> (restart)
#    5. Switching list types (bullet -> decimal) auto-restarts
#    6. Custom paragraph styles coexist with list numbering
#    7. list_add_fpar() works with formatted text (bold, italic, etc.)
#    8. list_add_blocks() tags every paragraph in the block
#    9. Saved .docx can be reopened without corruption (round-trip)
#   10. list_inspect() runs without error (smoke test)
#   11. Consecutive same-type calls reuse the same numId (continuation)
#   12. list_end() is safe to call anywhere, any number of times
#
# ==============================================================================

library(officer)
library(xml2)
source("officer_list_helpers.R")


# ==============================================================================
# TEST UTILITIES
#
# Small helpers that make the test code more readable.
# ==============================================================================

test_count <- 0L
pass_count <- 0L

#' Run one named test. Stops with a clear message on failure.
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

#' Assert a condition is TRUE, or stop with the given message.
assert <- function(condition, message = "assertion failed") {
  if (!isTRUE(condition)) stop(message, call. = FALSE)
}

#' Assert two values are identical.
assert_equal <- function(actual, expected, message = NULL) {
  if (is.null(message)) {
    message <- sprintf("expected '%s' but got '%s'", expected, actual)
  }
  assert(identical(actual, expected), message)
}

#' Get the numId and ilvl from the paragraph at the current cursor position.
#'
#' Returns a list with $num_id and $ilvl as character strings.
#' Both are NA if the paragraph is not a list item.
#'
#' XPath used:
#'   "w:pPr/w:numPr/w:numId" — navigate: paragraph props -> numPr -> numId
#'   "w:pPr/w:numPr/w:ilvl"  — navigate: paragraph props -> numPr -> ilvl
get_cursor_num_pr <- function(x) {
  node <- docx_current_block_xml(x)
  list(
    num_id = xml_attr(xml_find_first(node, "w:pPr/w:numPr/w:numId"), "val"),
    ilvl   = xml_attr(xml_find_first(node, "w:pPr/w:numPr/w:ilvl"), "val")
  )
}


# ==============================================================================
# TESTS
# ==============================================================================

cat("\nRunning officer_list_helpers tests\n")
cat(strrep("-", 60), "\n")


# ---------- 1. Basic bullets ----------

run_test("list_add_par creates bullet items", {
  doc <- read_docx()
  doc <- list_add_par(doc, "Item one", list_type = "bullet")
  doc <- list_add_par(doc, "Item two", list_type = "bullet")

  # The cursor is on "Item two". It should have a <w:numPr>.
  np <- get_cursor_num_pr(doc)
  assert(!is.na(np$num_id), "paragraph should have a numId")
  assert_equal(np$ilvl, "0")

  # numbering.xml should contain a bullet format definition.
  #
  # XPath breakdown:
  #   w:abstractNum                    — find abstractNum elements
  #   /w:lvl                           — their lvl children
  #   /w:numFmt[@w:val='bullet']       — that have numFmt with val="bullet"
  num_doc <- .read_numbering_xml(doc)
  bullet_fmts <- xml_find_all(
    num_doc,
    "w:abstractNum/w:lvl/w:numFmt[@w:val='bullet']"
  )
  assert(length(bullet_fmts) > 0, "numbering.xml should have a bullet format")
})


# ---------- 2. Basic numbered list ----------

run_test("list_add_par creates decimal items", {
  doc <- read_docx()
  doc <- list_add_par(doc, "Step one", list_type = "decimal")
  doc <- list_add_par(doc, "Step two", list_type = "decimal")

  np <- get_cursor_num_pr(doc)
  assert(!is.na(np$num_id), "paragraph should have a numId")
  assert_equal(np$ilvl, "0")

  # Check for decimal format in numbering.xml.
  num_doc <- .read_numbering_xml(doc)
  decimal_fmts <- xml_find_all(
    num_doc,
    "w:abstractNum/w:lvl/w:numFmt[@w:val='decimal']"
  )
  assert(length(decimal_fmts) > 0, "numbering.xml should have a decimal format")
})


# ---------- 3. Indent levels ----------

run_test("ilvl parameter sets correct indent level", {
  doc <- read_docx()
  doc <- list_add_par(doc, "Level 0", list_type = "bullet", ilvl = 0L)
  doc <- list_add_par(doc, "Level 1", list_type = "bullet", ilvl = 1L)
  doc <- list_add_par(doc, "Level 2", list_type = "bullet", ilvl = 2L)

  # Cursor is on "Level 2".
  assert_equal(get_cursor_num_pr(doc)$ilvl, "2")

  # Move back and check each level.
  doc <- cursor_backward(doc)
  assert_equal(get_cursor_num_pr(doc)$ilvl, "1")

  doc <- cursor_backward(doc)
  assert_equal(get_cursor_num_pr(doc)$ilvl, "0")
})


# ---------- 4. Restart with list_end() ----------

run_test("list_end causes numbering restart", {
  doc <- read_docx()
  doc <- list_add_par(doc, "A", list_type = "decimal")
  doc <- list_add_par(doc, "B", list_type = "decimal")
  num_id_before <- get_cursor_num_pr(doc)$num_id

  # End this list.
  doc <- list_end(doc)

  # Start a new decimal list — should get a DIFFERENT numId.
  doc <- list_add_par(doc, "C", list_type = "decimal")
  num_id_after <- get_cursor_num_pr(doc)$num_id

  assert(
    num_id_before != num_id_after,
    sprintf("numIds should differ after restart (both were %s)", num_id_before)
  )

  # The new <w:num> should have a <w:startOverride> element.
  #
  # XPath breakdown:
  #   w:num[@w:numId='X']                — find the num with our new ID
  #   /w:lvlOverride                     — its lvlOverride child
  #   /w:startOverride[@w:val='1']       — the startOverride with val=1
  num_doc <- .read_numbering_xml(doc)
  xpath <- sprintf(
    "w:num[@w:numId='%s']/w:lvlOverride/w:startOverride[@w:val='1']",
    num_id_after
  )
  overrides <- xml_find_all(num_doc, xpath)
  assert(length(overrides) > 0, "restarted <w:num> should have <w:startOverride>")
})


# ---------- 5. Auto-restart on type switch ----------

run_test("switching list types auto-restarts", {
  doc <- read_docx()
  doc <- list_add_par(doc, "Bullet", list_type = "bullet")
  bullet_id <- get_cursor_num_pr(doc)$num_id

  # Switching to decimal should use a different numId automatically.
  doc <- list_add_par(doc, "Number", list_type = "decimal")
  decimal_id <- get_cursor_num_pr(doc)$num_id

  assert(
    bullet_id != decimal_id,
    "bullet and decimal should have different numIds"
  )
})


# ---------- 6. Custom style preserved ----------

run_test("custom paragraph style coexists with list formatting", {
  doc <- read_docx()

  # "heading 2" exists in officer's default template.
  doc <- list_add_par(doc, "Styled bullet", style = "heading 2",
                      list_type = "bullet")

  node <- docx_current_block_xml(doc)

  # The paragraph should have BOTH a style reference AND a numPr.
  # XPath "w:pPr/w:pStyle" finds the style element.
  pstyle <- xml_attr(xml_find_first(node, "w:pPr/w:pStyle"), "val")
  assert(!is.na(pstyle), "paragraph should have a pStyle")

  np <- get_cursor_num_pr(doc)
  assert(!is.na(np$num_id), "paragraph should also have numPr")
})


# ---------- 7. Formatted paragraph (fpar) ----------

run_test("list_add_fpar works with formatted text", {
  doc <- read_docx()

  # fpar() creates a paragraph with mixed formatting.
  # ftext() creates one run (text chunk) with specific properties.
  formatted <- fpar(
    ftext("Bold part", prop = fp_text(bold = TRUE)),
    ftext(" and normal part")
  )
  doc <- list_add_fpar(doc, formatted, list_type = "bullet")

  np <- get_cursor_num_pr(doc)
  assert(!is.na(np$num_id), "fpar paragraph should have numPr")
  assert_equal(np$ilvl, "0")
})


# ---------- 8. Block list ----------

run_test("list_add_blocks numbers every paragraph", {
  doc <- read_docx()

  items <- block_list(
    fpar(ftext("First")),
    fpar(ftext("Second")),
    fpar(ftext("Third"))
  )
  doc <- list_add_blocks(doc, items, list_type = "decimal")

  # Cursor is on "Third". All three should share the same numId.
  np3 <- get_cursor_num_pr(doc)
  assert(!is.na(np3$num_id), "third paragraph should have numPr")

  doc <- cursor_backward(doc)
  np2 <- get_cursor_num_pr(doc)
  assert(!is.na(np2$num_id), "second paragraph should have numPr")

  doc <- cursor_backward(doc)
  np1 <- get_cursor_num_pr(doc)
  assert(!is.na(np1$num_id), "first paragraph should have numPr")

  # All three share the same numId — one continuous list.
  assert_equal(np1$num_id, np2$num_id)
  assert_equal(np2$num_id, np3$num_id)
})


# ---------- 9. Round-trip (save and reload) ----------

run_test("document round-trips through save and reload", {
  doc <- read_docx()
  doc <- list_add_par(doc, "Bullet", list_type = "bullet")
  doc <- list_add_par(doc, "Number", list_type = "decimal")

  # Save to a temp file.
  tf <- tempfile(fileext = ".docx")
  print(doc, target = tf)

  # Reload. If our XML is malformed, read_docx() will error.
  doc2 <- read_docx(tf)
  summary <- docx_summary(doc2)
  assert(nrow(summary) >= 2, "reloaded doc should have at least 2 paragraphs")

  unlink(tf)
})


# ---------- 10. list_inspect smoke test ----------

run_test("list_inspect runs without error", {
  doc <- read_docx()
  doc <- list_add_par(doc, "Item", list_type = "bullet")

  # Capture output to keep test output clean.
  output <- capture.output(list_inspect(doc))
  assert(length(output) > 0, "should produce output")
  assert(any(grepl("bullet", output)), "should mention 'bullet'")
})


# ---------- 11. Same-type continuation ----------

run_test("consecutive same-type calls share numId (no restart)", {
  doc <- read_docx()
  doc <- list_add_par(doc, "One",   list_type = "decimal")
  id1 <- get_cursor_num_pr(doc)$num_id

  doc <- list_add_par(doc, "Two",   list_type = "decimal")
  id2 <- get_cursor_num_pr(doc)$num_id

  doc <- list_add_par(doc, "Three", list_type = "decimal")
  id3 <- get_cursor_num_pr(doc)$num_id

  # All three use the same numId — they're one continuous list.
  assert_equal(id1, id2)
  assert_equal(id2, id3)
})


# ---------- 12. list_end is safe to call anywhere ----------

run_test("list_end is idempotent and safe", {
  doc <- read_docx()

  # Before any list — should not crash.
  doc <- list_end(doc)

  doc <- list_add_par(doc, "Item", list_type = "bullet")

  # Multiple times — should not crash.
  doc <- list_end(doc)
  doc <- list_end(doc)

  # Should still work.
  doc <- list_add_par(doc, "New item", list_type = "bullet")
  np <- get_cursor_num_pr(doc)
  assert(!is.na(np$num_id), "should still work after multiple list_end calls")
})


# ==============================================================================
# SUMMARY
# ==============================================================================

cat(strrep("-", 60), "\n")
cat(sprintf("\n%d / %d tests passed.\n", pass_count, test_count))
if (pass_count == test_count) {
  cat(sprintf("\nALL %d TESTS PASSED\n\n", test_count))
} else {
  stop("Some tests failed.", call. = FALSE)
}
