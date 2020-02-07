#' Create quality assurance report from Power BI list
#
#' @param report A Power BI report list read in using read_pbi or read_layout
#' @param output_file Output file path.
#' @return xlsx
#' @export
#' @keywords report, qa
#' @examples
#' \dontrun{
#' create_qa_report(report)
#' }

create_qa_report <- function(report, output_file = "pbi_qa_report.xlsx"){

  # Create datasets based off report
  report_summary <- create_report_summary(report)
  measures_summary <- create_measures_summary(report)
  sections_summary <- create_sections_summary(report)
  visuals_summary <- create_visuals_summary(report)

  # Create empty workbook
  wb <- openxlsx::createWorkbook()

  # Create styles
  tite_style <- openxlsx::createStyle(textDecoration = "bold")
  header_style <- openxlsx::createStyle(fgFill = "#4F81BD", halign = "left", textDecoration = "Bold", border = "TopBottomLeftRight", fontColour = "white", wrapText = TRUE)
  unlocked <- openxlsx::createStyle(locked = FALSE, border = "TopBottomLeftRight")
  grey <- openxlsx::createStyle(fgFill = "#f9f9f9", border = "TopBottomLeftRight")

  # Validation Tab
  openxlsx::addWorksheet(wb, "validation")
  openxlsx::writeData(wb, "validation", tibble::data_frame(response = c("Yes", "No", "N/A")))

  # Report Summary
  openxlsx::addWorksheet(wb, "report_summary")
  openxlsx::writeData(wb, "report_summary", "Report Summary")
  openxlsx::addStyle(wb, "report_summary", rows = 1, cols = 1, style = tite_style)

  ncol_report_summary_init <- ncol(report_summary)

  report_summary <- report_summary %>%
    dplyr::mutate(
      report_comment_for_qa = NA, report_filters_correct = NA,
      report_qa_comment = NA
    )

  validation_cells_report_summary_final <- which(grepl("correct", names(report_summary)))

  openxlsx::writeData(wb, "report_summary", report_summary, startRow = 4, headerStyle = header_style,
            borders = "all", borderStyle = "thin")

  lapply(validation_cells_report_summary_final, function(x){openxlsx::dataValidation(wb, "report_summary", col = x, rows = 5:(nrow(report_summary) + 4), type = "list", value = "'validation'!$A$2:$A$4")})
  openxlsx::setColWidths(wb, "report_summary", cols = 1:ncol(report_summary), widths = "auto")
  openxlsx::protectWorksheet(wb, "report_summary", protect = TRUE, lockFormattingCells = FALSE, lockFormattingColumns = FALSE, lockInsertingColumns = TRUE, lockDeletingColumns = TRUE)
  openxlsx::addStyle(wb, "report_summary", style = grey, col = 1:ncol_report_summary_init, rows = 5:(nrow(report_summary) + 4), gridExpand = TRUE)
  openxlsx::addStyle(wb, "report_summary", style = unlocked, col = (ncol_report_summary_init + 1):ncol(report_summary), rows = 5:(nrow(report_summary) + 4), gridExpand = TRUE)
  openxlsx::mergeCells(wb, "report_summary", cols = 1:2, rows = 3)
  openxlsx::mergeCells(wb, "report_summary", cols = 3, rows = 3)
  openxlsx::mergeCells(wb, "report_summary", cols = 4:ncol(report_summary), rows = 3)
  openxlsx::writeData(wb, "report_summary", "report_data", startCol = 1, startRow = 3)
  openxlsx::writeData(wb, "report_summary", "report_owner_checks", startCol = 3, startRow = 3)
  openxlsx::writeData(wb, "report_summary", "quality_assurer_checks", startCol = 4, startRow = 3)
  openxlsx::addStyle(wb, "report_summary", style = header_style, col = 1:ncol(report_summary), row = 3, gridExpand = TRUE)

  # Measures Summary
  openxlsx::addWorksheet(wb, "measures_summary")
  openxlsx::writeData(wb, "measures_summary", "Measures Summary")
  openxlsx::addStyle(wb, "measures_summary", rows = 1, cols = 1, style = tite_style)

  ncol_measures_summary_init <- ncol(measures_summary)

  measures_summary <- measures_summary %>%
    dplyr::mutate(
      measure_comment_for_qa = NA, measure_naming_convention_correct = NA,
      measure_linked_to_correct_table = NA,measure_code_correct = NA,
      measure_qa_comment = NA
    )

  validation_cells_measures_summary_final <- which(grepl("correct", names(measures_summary)))

  openxlsx::writeData(wb, "measures_summary", x = measures_summary, startRow = 4, headerStyle = header_style,
            borders = "all", borderStyle = "thin")

 lapply(validation_cells_measures_summary_final, function(x){openxlsx::dataValidation(wb, "measures_summary", col = x, rows = 5:(nrow(measures_summary) + 4), type = "list", value = "'validation'!$A$2:$A$4")})
  openxlsx::setColWidths(wb, "measures_summary", cols = 1:ncol(measures_summary), widths = "auto")
  openxlsx::protectWorksheet(wb, "measures_summary", protect = TRUE, lockFormattingCells = FALSE, lockFormattingColumns = FALSE, lockInsertingColumns = TRUE, lockDeletingColumns = TRUE)
  openxlsx::addStyle(wb, "measures_summary", style = grey, col = 1:ncol_measures_summary_init, rows = 5:(nrow(measures_summary) + 4), gridExpand = TRUE)
  openxlsx::addStyle(wb, "measures_summary", style = unlocked, col = (ncol_measures_summary_init + 1):ncol(measures_summary), rows = 5:(nrow(measures_summary) + 4), gridExpand = TRUE)
  openxlsx::mergeCells(wb, "measures_summary", cols = 1:3, rows = 3)
  openxlsx::mergeCells(wb, "measures_summary", cols = 4, rows = 3)
  openxlsx::mergeCells(wb, "measures_summary", cols = 5:ncol(measures_summary), rows = 3)
  openxlsx::writeData(wb, "measures_summary", "report_data", startCol = 1, startRow = 3)
  openxlsx::writeData(wb, "measures_summary", "report_owner_checks", startCol = 4, startRow = 3)
  openxlsx::writeData(wb, "measures_summary", "quality_assurer_checks", startCol = 5, startRow = 3)
  openxlsx::addStyle(wb, "measures_summary", style = header_style, col = 1:ncol(measures_summary), row = 3, gridExpand = TRUE)

  # Section Summary
  openxlsx::addWorksheet(wb, "sections_summary")

  openxlsx::writeData(wb, "sections_summary", "Sections Summary")
  openxlsx::addStyle(wb, "sections_summary", rows = 1, cols = 1, style = tite_style)

  ncol_sections_summary_init <- ncol(sections_summary)

  sections_summary <- sections_summary %>%
    dplyr::mutate(
      section_comment_for_qa = NA, section_filters_correct = NA,
      section_makes_sense = NA, section_qa_comment = NA
    )

  validation_cells_sections_summary_final <- which(grepl("correct|sense", names(sections_summary)))

  openxlsx::writeData(wb, "sections_summary", sections_summary, startRow = 4, headerStyle = header_style,
            borders = "all", borderStyle = "thin")

  lapply(validation_cells_sections_summary_final, function(x){openxlsx::dataValidation(wb, "sections_summary", col = x, rows = 5:(nrow(sections_summary) + 4), type = "list", value = "'validation'!$A$2:$A$4")})
  openxlsx::setColWidths(wb, "sections_summary", cols = 1:ncol(sections_summary), widths = "auto")
  openxlsx::protectWorksheet(wb, "sections_summary", protect = TRUE, lockFormattingCells = FALSE, lockFormattingColumns = FALSE, lockInsertingColumns = TRUE, lockDeletingColumns = TRUE)
  openxlsx::addStyle(wb, "sections_summary", style = grey, col = 1:ncol_sections_summary_init, rows = 5:(nrow(sections_summary) + 4), gridExpand = TRUE)
  openxlsx::addStyle(wb, "sections_summary", style = unlocked, col = (ncol_sections_summary_init + 1):ncol(sections_summary), rows = 5:(nrow(sections_summary) + 4), gridExpand = TRUE)
  openxlsx::mergeCells(wb, "sections_summary", cols = 1:2, rows = 3)
  openxlsx::mergeCells(wb, "sections_summary", cols = 3, rows = 3)
  openxlsx::mergeCells(wb, "sections_summary", cols = 4:ncol(sections_summary), rows = 3)
  openxlsx::writeData(wb, "sections_summary", "report_data", startCol = 1, startRow = 3)
  openxlsx::writeData(wb, "sections_summary", "report_owner_checks", startCol = 3, startRow = 3)
  openxlsx::writeData(wb, "sections_summary", "quality_assurer_checks", startCol = 4, startRow = 3)
  openxlsx::addStyle(wb, "sections_summary", style = header_style, col = 1:ncol(sections_summary), row = 3, gridExpand = TRUE)

  # Visuals Summary
  openxlsx::addWorksheet(wb, "visuals_summary")

  openxlsx::writeData(wb, "visuals_summary", "Visuals Summary")
  openxlsx::addStyle(wb, "visuals_summary", rows = 1, cols = 1, style = tite_style)

  ncol_visuals_summary_init <- ncol(visuals_summary)

  visuals_summary <- visuals_summary %>%
    dplyr::filter(!(visual_type %in% c("textbox", "image", "slicer", "basicShape", "actionButton"))) %>%
    dplyr::group_by(visual_title) %>%
    dplyr::mutate(
      visual_title_duplicate = ifelse(sum(!is.na(visual_title), na.rm = TRUE) > 1, 1, 0)
    ) %>%
    dplyr::mutate(visual_updates_published = NA, visual_published_url = NA,
           visual_position_published_url = NA, visual_changes_narrative_from_published = NA,
           visual_differences_from_published_explained = NA, visual_other_comment_for_qa = NA,
           visual_title_correct = NA, visual_filters_correct = NA, visual_values_correct = NA,
           visual_make_sense_and_clear = NA, visual_qa_comment = NA)

  validation_cells_visuals_summary_final <- which(grepl("correct|updates|changes|clear|differences", names(visuals_summary)))

  openxlsx::writeData(wb, "visuals_summary", visuals_summary, startRow = 4, headerStyle = header_style,
            borders = "all", borderStyle = "thin")

  lapply(validation_cells_visuals_summary_final, function(x){openxlsx::dataValidation(wb, "visuals_summary", col = x, rows = 5:(nrow(visuals_summary) + 4), type = "list", value = "'validation'!$A$2:$A$4")})
  openxlsx::setColWidths(wb, "visuals_summary", cols = 1:ncol(visuals_summary), widths = "auto")
  openxlsx::protectWorksheet(wb, "visuals_summary", protect = TRUE, lockFormattingCells = FALSE, lockFormattingColumns = FALSE, lockInsertingColumns = TRUE, lockDeletingColumns = TRUE)
  openxlsx::addStyle(wb, "visuals_summary", style = grey, col = 1:ncol_visuals_summary_init, rows = 5:(nrow(visuals_summary) + 4), gridExpand = TRUE)
  openxlsx::addStyle(wb, "visuals_summary", style = unlocked, col = (ncol_visuals_summary_init + 1):ncol(visuals_summary), rows = 5:(nrow(visuals_summary) + 4), gridExpand = TRUE)
  openxlsx::mergeCells(wb, "visuals_summary", cols = 1:6, rows = 3)
  openxlsx::mergeCells(wb, "visuals_summary", cols = 7:12, rows = 3)
  openxlsx::mergeCells(wb, "visuals_summary", cols = 13:ncol(visuals_summary), rows = 3)
  openxlsx::writeData(wb, "visuals_summary", "report_data", startCol = 1, startRow = 3)
  openxlsx::writeData(wb, "visuals_summary", "report_owner_checks", startCol = 7, startRow = 3)
  openxlsx::writeData(wb, "visuals_summary", "quality_assurer_checks", startCol = 13, startRow = 3)
  openxlsx::addStyle(wb, "visuals_summary", style = header_style, col = 1:ncol(visuals_summary), row = 3, gridExpand = TRUE)

  # Clean up
  openxlsx::sheetVisibility(wb)[1] <- FALSE

  # Save
  openxlsx::saveWorkbook(wb, output_file, overwrite = TRUE)

}
