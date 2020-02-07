
# Measures ----------------------------------------------------------------

#' Extract details of an individual power bi measure
#
#' @param measure
#' @return tibble
#' @keywords measure, extract
#' @examples
#' \dontrun{
#' pbiAssure:::extract_measure(RJSONIO::fromJSON(report$config)$modelExtensions[[1]]$entities[[1]]$measures[[1]])de
#' }
extract_measure <- function(measure){

  tibble::tibble(measure$name, measure$expression)

}

#' Extract an individual measure from an individual entity (table) within a power bi report list.
#
#' @param entity
#' @return tibble
#' @keywords measures, extract
#' @examples
#' \dontrun{
#' pbiAssure:::extract_entity_measures(RJSONIO::fromJSON(report$config)$modelExtensions[[1]]$entities[[1]])
#' }
extract_entity_measures <- function(entity){

  base_table <- entity$name

  lapply(entity$measures, extract_measure) %>%
    dplyr::bind_rows() %>%
    dplyr::transmute(
      measure_base_table = base_table,
      measure_name = `measure$name`,
      measure_code = trimws(`measure$expression`)
    )


}

#' Extract all measures associated with all model extensions.
#
#' @param modelExtensions
#' @return tibble
#' @keywords measures, extract
#' @examples
#' \dontrun{
#' pbiAssure:::extract_extensions_measures(RJSONIO::fromJSON(report$config)$modelExtensions[[1]])
#' }
extract_extensions_measures <- function(modelExtensions){

  lapply(modelExtensions$entities, extract_entity_measures) %>%
    dplyr::bind_rows()

}

#' Creates a measures summary from a Power BI report list
#
#' @param report A Power BI report list read in using read_pbi or read_layout
#' @return tibble
#' @keywords measures, extract
#' @export
#' @examples
#' \dontrun{
#' create_measures_summary(report)
#' }
create_measures_summary <- function(report){

  config_json <-  RJSONIO::fromJSON(report$config)

  lapply(config_json$modelExtensions, extract_extensions_measures) %>%
    dplyr::bind_rows()

}

# Filters -----------------------------------------------------------------

#' Extract value from where in filter in Power BI list
#
#' @param where_element A where in element in Power BI list
#' @return tibble
#' @keywords filters, extract
#' @examples
#' \dontrun{
#' pbiAssure:::extract_filter_where_in_value(RJSONIO::fromJSON(report$sections[[6]]$visualContainers[[4]]$filters)[[2]]$filter$Where[[1]])
#' }
extract_filter_where_in_value <- function(where){

  lapply(where$Condition$In$Values, function(x) {as.character(x[[1]]$`Literal`["Value"])}) %>%
    paste(collapse = ",") %>%
    paste0("c(", ., ")")

}

#' Extract value from where comparisson filter in Power BI list
#
#' @param where_element A where in element in Power BI list
#' @return tibble
#' @keywords filters, extract
#' @examples
#' \dontrun{
#' pbiAssure:::extract_filter_where_comparison_value(RJSONIO::fromJSON(report$sections[[6]]$visualContainers[[4]]$filters)[[1]]$filter$Where[[1]])
#' }
extract_filter_where_comparison_value <- function(where){

  comparison_kind <- where$Condition$Comparison$ComparisonKind

  left <- where$Condition$Comparison$Left$Literal[["Value"]]

  right <- where$Condition$Comparison$Right$Literal[["Value"]]

  dplyr::case_when(
    comparison_kind == 1 ~ paste(">", left),
    comparison_kind == 2 ~ paste(">=", left),
    comparison_kind == 3 ~ paste("<", right),
    comparison_kind == 4 ~ paste("<=", right),
    TRUE ~ "NA"
  )

}

#' Extract invividual filter from filters element of Power BI list
#
#' @param filters_element A filters element in Power BI list
#' @return tibble
#' @keywords filters, extract
#' @examples
#' \dontrun{
#' pbiAssure:::extract_filter(RJSONIO::fromJSON(report$sections[[6]]$visualContainers[[4]]$filters)[[1]])
#' }
extract_filter <- function(filters_element){

  if (length(names(filters_element$filter$Where[[1]]$Condition)) > 0) {

    condition_type <- names(filters_element$filter$Where[[1]]$Condition)

    if (condition_type == "In") {

      tibble::data_frame(
        filter_table = if_null_na(filters_element$filter$From[[1]]["Entity"]),
        filter_column = if_null_na(filters_element$filter$Where[[1]]$Condition$In$Expressions[[1]]$Column$Property),
        filter_value =  if_null_na(extract_filter_where_in_value(filters_element$filter$Where[[1]]))
      ) %>%
        dplyr::filter(filter_value != "c()") %>%
        dplyr::mutate(
          conc = paste0("[", filter_table, "].[", filter_column, "] in ", filter_value)
        ) %>%
        dplyr::filter(!is.na(filter_table))

    } else if (condition_type == "Comparison") {

      tibble::data_frame(
        filter_table = if_null_na(filters_element$filter$From[[1]]["Entity"]),
        filter_column = if_null_na(filters_element$filter$Where[[1]]$Condition$Comparison$Left$Column$Property),
        filter_value =  if_null_na(extract_filter_where_comparison_value(filters_element$filter$Where[[1]]))
      ) %>%
        dplyr::filter(filter_value != "c()") %>%
        dplyr::mutate(
          conc = paste0("[", filter_table, "].[", filter_column, "] ", filter_value)
        ) %>%
        dplyr::filter(!is.na(filter_table))

    }

  }

}

#' Extract all filters from an object
#
#' @param object A Power BI object that has attached filters
#' @return tibble
#' @keywords filters, extract
#' @examples
#' \dontrun{
#' pbiAssure:::extract_filters(report$sections[[6]]$visualContainers[[4]])
#' }
extract_filters <- function(object){

  if (object$filters == "[]") {
    NA
  } else {

    filters_list <- RJSONIO::fromJSON(object$filters)

    df <- lapply(filters_list, extract_filter) %>%
      dplyr::bind_rows()

    ifelse(nrow(df) == 0, NA, paste(df["conc"]))

  }

}

# Report Summary -------------------------------------------------------------

#' Provides a high level summary of a report
#
#' @param report A Power BI report list read in using read_pbi or read_layout
#' @return tibble
#' @export
#' @keywords report, filters, extract
#' @examples
#' \dontrun{
#' create_report_summary(report)
#' }
create_report_summary <- function(report){

  tibble::data_frame(
    report_name = "report",
    report_filters = extract_filters(report)
  )

}

# Section Summary -----------------------------------------------------------

#' Provides a high level summary of a section
#
#' @param section A section element in a Power BI list
#' @return tibble
#' @keywords section, filters, extract
#' @examples
#' \dontrun{
#' pbiAssure:::create_section_summary(report$sections[[6]])
#' }
create_section_summary <- function(section){

  tibble::data_frame(
    slide_name = section$displayName,
    slide_filters = extract_filters(section)
    )

}

#' Provides a high level summary of all report sections
#
#' @param report A Power BI report list read in using read_pbi or read_layout
#' @return tibble
#' @export
#' @keywords section, filters, extract
#' @examples
#' \dontrun{
#' create_sections_summary(report)
#' }
create_sections_summary <- function(report){

  lapply(report$sections, create_section_summary) %>%
    dplyr::bind_rows()

}

# Visuals -----------------------------------------------------------------

#' Create a summmary of a single Power BI visual
#
#' @param single_visual A visual container element in a Power BI list
#' @return tibble
#' @keywords visual, filters, extract
#' @examples
#' \dontrun{
#' pbiAssure:::extract_single_visual(report$sections[[6]]$visualContainers[[2]])
#' }
extract_single_visual <- function(single_visual){

  config_json <- RJSONIO::fromJSON(single_visual$config)

  tibble::tibble(
    visual_id = config_json$name,
    visual_type = config_json$singleVisual$`visualType`,
    visual_title = if_null_na(config_json$singleVisual$vcObjects$title[[1]]$properties$text$expr$Literal["Value"]),
    visual_filters = extract_filters(single_visual)
    )

}

#' Create a summmary of all Power BI visuals in a section
#
#' @param section A section element in a Power BI list
#' @return tibble
#' @keywords visual, filters, extract
#' @examples
#' \dontrun{
#' pbiAssure:::extract_section_visuals(report$sections[[6]])
#' }
extract_section_visuals <- function(section){

  lapply(section, extract_single_visual) %>%
    dplyr::bind_rows()

}

#' Provides a high level summary of all report visuals
#
#' @param report A Power BI report list read in using read_pbi or read_layout
#' @return tibble
#' @export
#' @keywords section, filters, extract
#' @examples
#' \dontrun{
#' create_visuals_summary(report)
#' }
create_visuals_summary <- function(report){

  lapply(report$sections, function(x){extract_section_visuals(x$visualContainers) %>%
    dplyr::transmute(slide_name = x$displayName, visual_id, visual_type, visual_title, visual_filters)}) %>%
    dplyr::bind_rows()

}



