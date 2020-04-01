# pbiAssure [![Travis build status](https://travis-ci.org/adamrobinson361/pbiAssure.svg?branch=master)](https://travis-ci.org/adamrobinson361/pbiAssure)

pbiAssure is an R package that lets you read a Power BI file into R, create summaries of its core elements, and generate a bespoke quality assurance excel with the summaries and approval fields.

The summaries currently extractable and visible in the QA excel are as follows:

- Custom Measures - All custom measures in a report with DAX code.

- Report Summary - All filters applied at a report level.

- Sections Summary - All filters applied in each section (tab).

- Visuals Summary - All visuals (plots, text boxes, tables etc.) with associated titles, selections and filters.

**Note:** the package is primarily developed for SSAS direct query Power BI reports. As such you may find that some of the functionality doesn't work for different report types, particularly contained Power BI reports.

## Install

pbiAssure is **not on CRAN**. The only way to install the package currently is directly from GitHub. There are a number of ways to do this such as:

`
remotes::install_github("adamrobinson361/pbiAssure")
`

If you are behind a corporate network you'll have to ensure that your proxy is set up. You will get a timeout error if this is not set up correctly.

## Use

Current functionality can be broken down as follows: 

- Read a Power BI report into R.

- Create summaries of different parts of the report.

- Create an Excel quality assurance template.

### Reading a Power BI report into R

To read the Power BI report into R you use the `read_pbi` function. A report can be loaded in as follows:

`
report <- read_pbi("report_name.pbix")
`

**Note:** if your report has any special characters copied from Word the reading function will fail. Please convert any characters not native to Power BI to characters that can be created in Power BI and try again. e.g. Word bullet to -.

### Create summaries of different parts of the report

There are 4 summary functions. These are as follows:

- `create_measures_summary`

- `create_report_summary`

- `create_sections_summary`

- `create_visuals_summary`

All functions take one parameter which is a report read into R using `read_pbi`.

For example, to create a measures summary of the report we previously read in we would run the following:

`
create_measures_summary(report)
`

This should output a data frame with all of the custom measures in the report.

**Note:** The create summary functions are not essential if you are just interested in generating a qa report. They can be useful if you want to operate in the summaries in R or use them in your own reports.

### Create an Excel quality assurance template

The `create_qa_report` function utilises all of the create summary functions to create a formatted excel quality assurance template that is bespoke to your Power BI. The output is simply an excel with a tab for each summary extended with validation fields.

You use the function as follows:

`
create_qa_report(report)
`

This will create a file called `pbi_qa_report.xlsx` in your working directory. You can change this if you would like using the `output_file` parameter.

## Further Developments

Going forward functionality will be developed to compare Power BI reports and to read back in old QA excels. This will be used to allow iterative quality assurance where changes between versions are flagged for QA and previous QA for elements that haven't changed is brought forward.
