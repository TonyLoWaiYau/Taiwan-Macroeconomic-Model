library(shiny)
source("model_excel_workflow.R")

default_outdir <- file.path(Sys.getenv("USERPROFILE"), "Documents", "Taiwan Macroeconomic Model_Output")

ui <- fluidPage(
  titlePanel("Taiwan Macroeconomic Model v3.5 (Tony Lo, Feb 2026)"),
  sidebarLayout(
    sidebarPanel(
      fileInput("datafile", "Upload data.xlsx", accept = ".xlsx"),
      textInput("sheet", "Sheet name", value = "clean"),
      textInput("outdir", "Output folder (existing)", value = default_outdir),
      
      tags$hr(),
      actionButton("btn_part1", "Run Part 1 (Create result_raw.xlsx)"),
      actionButton("btn_open_raw", "Open result_raw.xlsx in Excel"),
      
      tags$hr(),
      helpText("Edit ONLY the 'forecast_condition' sheet in result_raw.xlsx, save, and CLOSE Excel (file must not be locked)."),
      actionButton("btn_part2", "Run Part 2 (Conditional Forecast + result_summary.xlsx)"),
      actionButton("btn_open_summary", "Open result_summary.xlsx"),
      
      tags$hr(),
      verbatimTextOutput("status")
    ),
    mainPanel(
      h4("Workflow"),
      tags$ol(
        tags$li("Upload data.xlsx and run Part 1."),
        tags$li("Excel opens: edit forecast_condition, save, close Excel."),
        tags$li("Run Part 2. Repeat editing + Part 2 until satisfied.")
      )
    )
  )
)

server <- function(input, output, session) {
  
  state <- reactiveValues(part1 = NULL, part2 = NULL, msg = "")
  
  output$status <- renderPrint({
    list(
      msg = state$msg,
      result_raw = if (!is.null(state$part1)) state$part1$out_raw else NA,
      result_summary = if (!is.null(state$part2)) state$part2$out_summary else NA
    )
  })
  
  observeEvent(input$btn_part1, {
    req(input$datafile$datapath, input$outdir)
    
    tryCatch({
      dir.create(input$outdir, showWarnings = FALSE, recursive = TRUE)
      
      state$msg <- "Running Part 1..."
      state$part1 <- run_part1(
        data_xlsx = input$datafile$datapath,
        sheet = input$sheet,
        out_dir = input$outdir
      )
      
      if (!file.exists(state$part1$out_raw)) {
        stop("Part 1 finished but result_raw.xlsx was not created. Output dir: ", input$outdir)
      }
      
      state$msg <- paste("Part 1 done. Created:", state$part1$out_raw)
      showNotification("Created result_raw.xlsx", type = "message", duration = 5)
      
    }, error = function(e) {
      state$msg <- paste("Part 1 ERROR:", conditionMessage(e))
      showNotification(conditionMessage(e), type = "error", duration = 10)
    })
  })
  
  observeEvent(input$btn_open_raw, {
    req(state$part1$out_raw)
    tryCatch(
      open_file(state$part1$out_raw),
      error = function(e) showNotification(conditionMessage(e), type = "error", duration = 8)
    )
  })
  
  observeEvent(input$btn_part2, {
    req(state$part1$out_raw)
    
    tryCatch({
      state$msg <- "Running Part 2... (make sure Excel is CLOSED)"
      state$part2 <- run_part2(state$part1, edited_result_raw_xlsx = state$part1$out_raw)
      
      state$msg <- paste("Part 2 done. Updated:", state$part2$out_raw,
                         "and created:", state$part2$out_summary)
      showNotification("Part 2 completed successfully.", type = "message", duration = 5)
      
    }, error = function(e) {
      # Do NOT crash; just inform user and let them fix Excel and retry
      state$msg <- paste("Part 2 ERROR:", conditionMessage(e))
      showNotification(
        paste0("Part 2 failed:\n", conditionMessage(e),
               "\n\nFix 'forecast_condition' in result_raw.xlsx, save, CLOSE Excel, and click 'Run Part 2' again."),
        type = "error",
        duration = NULL   # keep message visible until user closes it
      )
    })
  })
  
  observeEvent(input$btn_open_summary, {
    req(state$part2$out_summary)
    tryCatch(
      open_file(state$part2$out_summary),
      error = function(e) showNotification(conditionMessage(e), type = "error", duration = 8)
    )
  })}

shinyApp(ui, server)
