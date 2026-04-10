# ==============================================================================
# test_officer_list_helpers.R
#
# Tests for officer_list_helpers.R. Run with:
#
#   Rscript test_officer_list_helpers.R
#
# ==============================================================================

library(officer)
library(xml2)
source("officer_list_helpers.R")


# ---- Test helpers ------------------------------------------------------------

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
  assert(identical(actual, expected), message)
}

get_cursor_num_pr <- function(x) {
  node <- docx_current_block_xml(x)
  list(
    num_id = xml_attr(xml_find_first(node, "w:pPr/w:numPr/w:numId"), "val"),
    ilvl   = xml_attr(xml_find_first(node, "w:pPr/w:numPr/w:ilvl"), "val")
  )
}


# ---- Tests -------------------------------------------------------------------

cat("\nRunning officer_list_helpers tests\n")
cat(strrep("-", 60), "\n")


# ---------- 1. Basic bullets ----------

run_test("list_add_par creates bullet items", {
  doc <- read_docx()
  doc <- list_add_par(doc, "Item one", list_type = "bullet")
  doc <- list_add_par(doc, "Item two", list_type = "bullet")

  np <- get_cursor_num_pr(doc)
  assert(!is.na(np$num_id), "paragraph should have a numId")
  assert_equal(np$ilvl, "0")

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

  assert_equal(get_cursor_num_pr(doc)$ilvl, "2")

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

  doc <- list_end(doc)
  doc <- list_add_par(doc, "C", list_type = "decimal")
  num_id_after <- get_cursor_num_pr(doc)$num_id

  assert(num_id_before != num_id_after,
         sprintf("numIds should differ (both were %s)", num_id_before))

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

  doc <- list_add_par(doc, "Number", list_type = "decimal")
  decimal_id <- get_cursor_num_pr(doc)$num_id

  assert(bullet_id != decimal_id,
         "bullet and decimal should have different numIds")
})


# ---------- 6. Custom style preserved ----------

run_test("custom paragraph style coexists with list formatting", {
  doc <- read_docx()
  doc <- list_add_par(doc, "Styled bullet", style = "heading 2",
                      list_type = "bullet")

  node <- docx_current_block_xml(doc)
  pstyle <- xml_attr(xml_find_first(node, "w:pPr/w:pStyle"), "val")
  assert(!is.na(pstyle), "paragraph should have a style")

  np <- get_cursor_num_pr(doc)
  assert(!is.na(np$num_id), "paragraph should also have numPr")
})


# ---------- 7. Formatted paragraph (fpar) ----------

run_test("list_add_fpar works with formatted text", {
  doc <- read_docx()
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

  np3 <- get_cursor_num_pr(doc)
  assert(!is.na(np3$num_id), "third paragraph should have numPr")

  doc <- cursor_backward(doc)
  np2 <- get_cursor_num_pr(doc)
  assert(!is.na(np2$num_id), "second paragraph should have numPr")

  doc <- cursor_backward(doc)
  np1 <- get_cursor_num_pr(doc)
  assert(!is.na(np1$num_id), "first paragraph should have numPr")

  assert_equal(np1$num_id, np2$num_id)
  assert_equal(np2$num_id, np3$num_id)
})


# ---------- 9. Round-trip ----------

run_test("document round-trips through save and reload", {
  doc <- read_docx()
  doc <- list_add_par(doc, "Bullet", list_type = "bullet")
  doc <- list_add_par(doc, "Number", list_type = "decimal")

  tf <- tempfile(fileext = ".docx")
  print(doc, target = tf)
  doc2 <- read_docx(tf)
  summary <- docx_summary(doc2)
  assert(nrow(summary) >= 2, "reloaded doc should have at least 2 paragraphs")
  unlink(tf)
})


# ---------- 10. list_inspect ----------

run_test("list_inspect runs without error", {
  doc <- read_docx()
  doc <- list_add_par(doc, "Item", list_type = "bullet")
  output <- capture.output(list_inspect(doc))
  assert(length(output) > 0, "should produce output")
  assert(any(grepl("bullet", output)), "should mention 'bullet'")
})


# ---------- 11. Same-type continuation ----------

run_test("consecutive same-type calls share numId", {
  doc <- read_docx()
  doc <- list_add_par(doc, "One",   list_type = "decimal")
  id1 <- get_cursor_num_pr(doc)$num_id
  doc <- list_add_par(doc, "Two",   list_type = "decimal")
  id2 <- get_cursor_num_pr(doc)$num_id
  doc <- list_add_par(doc, "Three", list_type = "decimal")
  id3 <- get_cursor_num_pr(doc)$num_id

  assert_equal(id1, id2)
  assert_equal(id2, id3)
})


# ---------- 12. list_end is idempotent ----------

run_test("list_end is idempotent and safe", {
  doc <- read_docx()
  doc <- list_end(doc)
  doc <- list_add_par(doc, "Item", list_type = "bullet")
  doc <- list_end(doc)
  doc <- list_end(doc)
  doc <- list_add_par(doc, "New item", list_type = "bullet")
  np <- get_cursor_num_pr(doc)
  assert(!is.na(np$num_id), "should still work after multiple list_end calls")
})


# ---------- 13. list_setup with tab_pos ----------

run_test("list_setup with tab_pos produces flush-left abstractNum", {
  doc <- read_docx()
  doc <- list_setup(doc, tab_pos = 360L)
  doc <- list_add_par(doc, "Flush bullet", list_type = "bullet")

  # Check that the abstractNum has w:ind with left="0" and a tab stop.
  num_doc <- .read_numbering_xml(doc)
  abs_nodes <- xml_find_all(num_doc, "w:abstractNum")

  # Our bullet abstractNum is one of the last ones added.
  # Find the one with bullet format that has left="0".
  found_flush <- FALSE
  for (node in abs_nodes) {
    lvl0 <- xml_find_first(node, "w:lvl[@w:ilvl='0']")
    if (inherits(lvl0, "xml_missing")) next

    fmt <- xml_attr(xml_find_first(lvl0, "w:numFmt"), "val")
    if (fmt != "bullet") next

    ind <- xml_find_first(lvl0, "w:pPr/w:ind")
    left_val <- xml_attr(ind, "left")
    first_val <- xml_attr(ind, "firstLine")

    tab <- xml_find_first(lvl0, "w:pPr/w:tabs/w:tab")
    if (inherits(tab, "xml_missing")) next

    tab_val <- xml_attr(tab, "pos")

    if (left_val == "0" && first_val == "0" && tab_val == "360") {
      found_flush <- TRUE
      break
    }
  }
  assert(found_flush,
         "should have bullet abstractNum with left=0, firstLine=0, tab=360")
})


# ---------- 14. tab_pos sub-levels have correct indents ----------

run_test("tab_pos sub-levels use multiples of tab_pos", {
  doc <- read_docx()
  doc <- list_setup(doc, tab_pos = 360L)
  doc <- list_add_par(doc, "Level 0", list_type = "bullet", ilvl = 0L)
  doc <- list_add_par(doc, "Level 1", list_type = "bullet", ilvl = 1L)

  # Check level 1 in the abstractNum: left should be 360, tab should be 720.
  num_doc <- .read_numbering_xml(doc)
  abs_nodes <- xml_find_all(num_doc, "w:abstractNum")

  found <- FALSE
  for (node in abs_nodes) {
    lvl1 <- xml_find_first(node, "w:lvl[@w:ilvl='1']")
    if (inherits(lvl1, "xml_missing")) next

    fmt <- xml_attr(xml_find_first(lvl1, "w:numFmt"), "val")
    if (fmt != "bullet") next

    ind <- xml_find_first(lvl1, "w:pPr/w:ind")
    left_val <- xml_attr(ind, "left")

    tab <- xml_find_first(lvl1, "w:pPr/w:tabs/w:tab")
    if (inherits(tab, "xml_missing")) next
    tab_val <- xml_attr(tab, "pos")

    if (left_val == "360" && tab_val == "720") {
      found <- TRUE
      break
    }
  }
  assert(found, "level 1 should have left=360, tab=720")
})


# ---------- 15. Default mode has standard hanging indent ----------

run_test("default mode (no list_setup) uses hanging indent", {
  doc <- read_docx()
  doc <- list_add_par(doc, "Standard bullet", list_type = "bullet")

  num_doc <- .read_numbering_xml(doc)
  abs_nodes <- xml_find_all(num_doc, "w:abstractNum")

  found_hanging <- FALSE
  for (node in abs_nodes) {
    lvl0 <- xml_find_first(node, "w:lvl[@w:ilvl='0']")
    if (inherits(lvl0, "xml_missing")) next

    fmt <- xml_attr(xml_find_first(lvl0, "w:numFmt"), "val")
    if (fmt != "bullet") next

    ind <- xml_find_first(lvl0, "w:pPr/w:ind")
    left_val <- xml_attr(ind, "left")
    hang_val <- xml_attr(ind, "hanging")

    if (!is.na(left_val) && left_val == "720" &&
        !is.na(hang_val) && hang_val == "360") {
      found_hanging <- TRUE
      break
    }
  }
  assert(found_hanging,
         "default mode should have left=720, hanging=360")
})


# ---------- 16. list_setup warns if called after list_add ----------

run_test("list_setup warns if called after list activity", {
  doc <- read_docx()
  doc <- list_add_par(doc, "Item", list_type = "bullet")

  # Calling list_setup after list_add should produce a warning.
  warned <- FALSE
  tryCatch(
    {
      doc <- withCallingHandlers(
        list_setup(doc, tab_pos = 360L),
        warning = function(w) {
          warned <<- TRUE
          invokeRestart("muffleWarning")
        }
      )
    }
  )
  assert(warned, "list_setup after list_add should warn")
})


# ---------- 17. tab_pos round-trip ----------

run_test("tab_pos document round-trips without corruption", {
  doc <- read_docx()
  doc <- list_setup(doc, tab_pos = 360L)
  doc <- list_add_par(doc, "Flush bullet", list_type = "bullet")
  doc <- list_add_par(doc, "Flush number", list_type = "decimal")

  tf <- tempfile(fileext = ".docx")
  print(doc, target = tf)
  doc2 <- read_docx(tf)
  s <- docx_summary(doc2)
  assert(nrow(s) >= 2, "reloaded doc should have at least 2 paragraphs")
  unlink(tf)
})


# ---------- 18. list_inspect shows tab info in flush mode ----------

run_test("list_inspect shows tab info for flush-left mode", {
  doc <- read_docx()
  doc <- list_setup(doc, tab_pos = 360L)
  doc <- list_add_par(doc, "Item", list_type = "bullet")

  output <- capture.output(list_inspect(doc))
  assert(any(grepl("tab=", output)), "inspect output should show tab info")
  assert(any(grepl("firstLine=", output)), "inspect output should show firstLine")
})


# ---- Summary -----------------------------------------------------------------

cat(strrep("-", 60), "\n")
cat(sprintf("\n%d / %d tests passed.\n", pass_count, test_count))
if (pass_count == test_count) {
  cat(sprintf("\nALL %d TESTS PASSED\n\n", test_count))
} else {
  stop("Some tests failed.", call. = FALSE)
}
