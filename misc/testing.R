report <- pbiAssure::read_layout("Layout")

pbiAssure::create_sections_summary(report)

pbiAssure::create_report_summary(report)

pbiAssure::create_measures_summary(report)

pbiAssure::create_visuals_summary(report) %>%
  View

pbiAssure::create_qa_report(report)
