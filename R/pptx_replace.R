
#' @export
#' @title Replace text anywhere in a pptx document, or only on the current slide
#' @description Replace all occurrences of old_value with new_value. This method
#' uses \code{\link{grepl}}/\code{\link{gsub}} for pattern matching; you may
#' supply arguments as required (and therefore use \code{\link{regex}} features)
#' using the optional \code{...} argument.
#'
#' Note that by default, grepl/gsub will use \code{fixed=FALSE}, which means
#' that \code{old_value} and \code{new_value} will be interepreted as regular
#' expressions.
#'
#' \strong{Chunking of text}
#'
#' Note that the behind-the-scenes representation of text in a document is
#' frequently not what you might expect! Sometimes a paragraph of text is broken
#' up (or "chunked") into several "runs," as a result of style changes, pauses
#' in text entry, later revisions and edits, etc. If you have not styled the
#' text, and have entered it in an "all-at-once" fashion, e.g. by pasting it or
#' by outputing it programmatically into your document, then this will
#' likely not be a problem. If you are working with a manually-edited document,
#' however, this can lead to unexpected failures to find text.
#'
#' @seealso \code{\link{grep}}, \code{\link{regex}}
#' @author Sushila Shenoy, \email{sushila.alyssa.shenoy@gmail.com}
#' @param x a pptx device
#' @param old_value the value to replace
#' @param new_value the value to replace it with
#' @param slide_index if \code{NULL} (default), search-and-replace on the current
#' slide; otherwise the index of the slide to search-and-replace.
#' @param warn warn if \code{old_value} could not be found.
#' @param ... optional arguments to grepl/gsub (e.g. \code{fixed=TRUE})
#' @examples
#' doc <- read_pptx()
#' bl <- block_list(
#'   fpar(ftext("hello PERSON. ", shortcuts$fp_bold(color = "pink"))),
#'   fpar(
#'     ftext("hello ", shortcuts$fp_bold()),
#'     ftext("person. ", shortcuts$fp_italic(color="red"))
#'   ),
#'   fpar(ftext("No need to panic. ", fp_text())))
#' doc <- add_slide(doc)
#' doc <- ph_with(doc, "Hello, PERSON ", location = ph_location_type(type="title"))
#' doc <- ph_with(x = doc, value = bl,
#'                location = ph_location_type(type="body") )
#'
#' # Show slide contents before modification
#' slide_summary(doc)[, c(1, 8)]
#'
#' # Simple search-and-replace, with regex turned off
#' doc <- replace_text_on_slide(doc, old_value = "PERSON", new_value = "Alice", fixed = TRUE)
#'
#' # Show slide contents after modification
#' slide_summary(doc)[, c(1, 8)]
#'
#' # Do the same, but in the entire document and ignoring case
#' doc <- replace_text_on_slide(doc, old_value = "PERSON", new_value = "Bob", ignore.case = TRUE)
#' slide_summary(doc)[, c(1, 8)]
#'
#' # Show slide contents after modification
#' slide_summary(doc)[, c(1, 8)]
#'
#' # Use regex: replace all words starting with lowercase "n" with the word "example"
#' doc <- replace_text_on_slide(doc, "\\bn.*?\\b", "example")
#' slide_summary(doc)[, c(1, 8)]
#'
#' # Show slide contents after modification
#' slide_summary(doc)[, c(1, 8)]

replace_text_on_slide <- function( x, old_value, new_value,
                                   slide_index = NULL,
                                   warn = TRUE, ...){
  stopifnot(is_scalar_character(old_value),
            is_scalar_character(new_value),
            is_scalar_logical(warn))

  if (is.null(slide_index)) slide_index <- doc$cursor
  doc$slide$get_slide(slide_index)$replace_all_text(old_value, new_value, warn, ...)

  x
}
