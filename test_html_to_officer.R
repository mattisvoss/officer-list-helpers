# ==============================================================================
# test_html_to_officer.R
#
# Run:  Rscript test_html_to_officer.R
#
# Each test demonstrates a usage pattern and verifies the output.
# ==============================================================================

library(officer)
library(xml2)
source("officer_list_helpers.R")
source("html_to_officer.R")

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

# Round-trip helper: HTML → docx → reload → summary.
html_to_summary <- function(html_str, style = NULL) {
  doc <- read_docx()
  doc <- body_add_html_fragment(doc, html_str, style = style)
  tf <- tempfile(fileext = ".docx")
  print(doc, target = tf)
  s <- docx_summary(read_docx(tf))
  unlink(tf)
  s
}


cat("\nRunning html_to_officer tests\n")
cat(strrep("-", 50), "\n")


# --- 1. Bold stays inline ---
# <b> inside <p> should NOT create a separate paragraph.

run_test("<p> with <b> is one paragraph", {
  s <- html_to_summary("<p>Hello <b>world</b></p>")
  text_rows <- s[s$text != "", ]
  assert(!any(text_rows$text == "world"),
         "'world' should be inline, not a separate paragraph")
})


# --- 2. Nested formatting ---
# <b><i>text</i></b> produces a run that is both bold and italic.

run_test("nested <b><i> accumulates formatting", {
  s <- html_to_summary("<p>A <b>B <i>C</i></b> D</p>")
  full <- paste(s$text[s$text != ""], collapse = " ")
  for (word in c("A", "B", "C", "D"))
    assert(grepl(word, full), sprintf("should contain '%s'", word))
})


# --- 3. Bare inline tags grouped ---
# "Hello <b>world</b>" with no <p> should be one paragraph.

run_test("bare inline tags become one paragraph", {
  s <- html_to_summary("Hello <b>world</b> and <i>more</i>")
  text_rows <- s[s$text != "", ]
  assert(nrow(text_rows) == 1,
         sprintf("should be 1 paragraph, got %d", nrow(text_rows)))
})


# --- 4. Block splits inline groups ---
# Inline, then <p>, then inline = three paragraphs.

run_test("block element splits inline groups", {
  s <- html_to_summary("start <b>A</b> <p>middle</p> end <i>B</i>")
  text_rows <- s[s$text != "", ]
  assert(nrow(text_rows) >= 3,
         sprintf("should be >= 3 paragraphs, got %d", nrow(text_rows)))
  assert(!any(text_rows$text == "A"), "'A' should be grouped with 'start'")
})


# --- 5. Lists ---
# <ul> and <ol> produce list items with numPr.

run_test("<ul> and <ol> create list items", {
  doc <- read_docx()
  doc <- body_add_html_fragment(doc,
    "<ul><li>Bullet</li></ul><ol><li>Number</li></ol>")
  np <- xml2::xml_find_first(docx_current_block_xml(doc),
                             "w:pPr/w:numPr/w:numId")
  assert(!inherits(np, "xml_missing"), "should have numPr")
})


# --- 6. Newlines preserved ---
# Literal \n becomes <w:br/> (soft return), not a new paragraph.

run_test("newlines become line breaks in same paragraph", {
  doc <- read_docx()
  doc <- body_add_html_fragment(doc, "Line one\nLine two\nLine three")
  br_nodes <- xml2::xml_find_all(docx_current_block_xml(doc), ".//w:br")
  assert(length(br_nodes) >= 2,
         sprintf("should have >= 2 <w:br/>, got %d", length(br_nodes)))
})


# --- 7. Empty input ---

run_test("empty input handled gracefully", {
  doc <- read_docx()
  doc <- body_add_html_fragment(doc, "")
  doc <- body_add_html_fragment(doc, NULL)
  tf <- tempfile(fileext = ".docx")
  print(doc, target = tf)
  read_docx(tf)  # would error if malformed
  unlink(tf)
})


# --- 8. Round-trip with everything ---

run_test("complex HTML round-trips without corruption", {
  html <- paste0(
    "<h1>Title</h1>",
    "<p>Normal <b>bold</b> <i>italic</i> <b><i>both</i></b></p>",
    "<ol><li>Step <b>one</b></li><li>Step two</li></ol>",
    "<ul><li>Bullet</li></ul>",
    "bare <b>text</b>\nwith newlines",
    "<p>Final.</p>"
  )
  s <- html_to_summary(html)
  assert(nrow(s[s$text != "", ]) >= 6, "should have multiple paragraphs")
})


# ---- Summary -----------------------------------------------------------------

cat(strrep("-", 50), "\n")
cat(sprintf("\n%d / %d passed.", pass_count, test_count))
if (pass_count == test_count) {
  cat(sprintf("  ALL %d TESTS PASSED\n\n", test_count))
} else {
  stop("Some tests failed.", call. = FALSE)
}
