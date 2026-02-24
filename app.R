# =============================================================================
# CTG Assay IC50 Analyser - R Shiny App - Created by Ashley Ng, February 2026
# Compatible with BMG CLARIOstar / standard 96-well plate luminescence output
#
# PLATE LAYOUT CONVENTION:
#   Row     A    = vehicle/DMSO control (100% viability reference — no drug)
#   Rows    B-H  = 7-point drug dilution series, ASCENDING in concentration
#                  (Row B = lowest dose; Row H = highest / starting concentration)
#   Columns 1-12 = each column is one drug × one sample (replicates share same Drug+Sample name)
#
# Requires: shiny, shinydashboard, DT, drc, ggplot2, dplyr, tidyr, readxl,
#           openxlsx, shinyjs, scales
#
# Install dependencies (run once):
# install.packages(c("shiny","shinydashboard","DT","drc","ggplot2","dplyr",
#                    "tidyr","readxl","openxlsx","shinyjs","scales"))
# =============================================================================

library(shiny)
library(shinydashboard)
library(DT)
library(drc)
library(ggplot2)
library(dplyr)
library(tidyr)
library(readxl)
library(openxlsx)
library(shinyjs)
library(scales)

# ---- Utility: Find the 8x12 data block in a sheet -------------------------
find_plate_block <- function(df_raw) {
  df_raw    <- as.data.frame(df_raw, stringsAsFactors = FALSE)
  row_labels <- c("A","B","C","D","E","F","G","H")
  max_row   <- nrow(df_raw) - 7
  if (max_row < 1) return(NULL)
  for (i in seq_len(max_row)) {
    for (col_i in seq_len(ncol(df_raw))) {
      vals <- as.character(df_raw[i:(i+7), col_i])
      if (all(!is.na(vals)) && all(trimws(toupper(vals)) == row_labels)) {
        data_rows <- i:(i + 7)
        num_cols  <- which(apply(df_raw[data_rows, ], 2, function(x)
          sum(!is.na(suppressWarnings(as.numeric(as.character(x))))) >= 6))
        if (length(num_cols) >= 12) {
          return(list(data_rows = data_rows, data_cols = num_cols[1:12]))
        }
      }
    }
  }
  return(NULL)
}

extract_plate <- function(path, sheet_name) {
  tmp <- tempfile(fileext = ".xlsx")
  file.copy(path, tmp)
  on.exit(unlink(tmp))
  df_raw <- read_excel(tmp, sheet = sheet_name, col_names = FALSE)
  block  <- find_plate_block(df_raw)
  if (is.null(block)) stop("Could not locate the 8\u00d712 plate block. Check your file.")
  mat <- as.data.frame(df_raw)[block$data_rows, block$data_cols]
  mat <- apply(mat, 2, function(x) as.numeric(as.character(x)))
  rownames(mat) <- c("A","B","C","D","E","F","G","H")
  colnames(mat) <- as.character(1:12)
  return(mat)
}

# ---- Helpers ---------------------------------------------------------------
# Generates n concentration steps for the drug rows (always 7 when one row is ctrl).
# Concentrations are returned in ASCENDING order: lowest dose first, highest last.
# The caller is responsible for mapping these to the correct plate rows (nearest
# to ctrl_row = lowest, furthest from ctrl_row = highest / starting conc).
# e.g. starting=10, factor=3, n=7 → 0.0137, 0.041, 0.123, 0.370, 1.11, 3.33, 10
calc_concs <- function(starting, factor, n = 7) {
  starting / (factor ^ ((n-1):0))
}

# Returns the 7 drug row letters ordered from LOWEST to HIGHEST concentration,
# given which row is the vehicle control.  The rule is: the row immediately
# adjacent to the control gets the lowest dose; the row furthest away gets the
# highest dose (starting concentration).
#
# ctrl_row = "A"  →  drug rows nearest→furthest: B,C,D,E,F,G,H  (ascending ri)
# ctrl_row = "H"  →  drug rows nearest→furthest: G,F,E,D,C,B,A  (descending ri)
# ctrl_row = "D"  →  rows nearest: C or E first, then outward both ways, then wrap
#
drug_rows_ordered <- function(ctrl_row) {
  ctrl_ri   <- which(LETTERS[1:8] == ctrl_row)
  other_ri  <- setdiff(1:8, ctrl_ri)
  # Sort by distance from ctrl_ri (ascending distance = lowest dose first)
  # Ties (equidistant) are broken by taking the row with lower ri first.
  other_ri[order(abs(other_ri - ctrl_ri), other_ri)]
}

# ---- UI -------------------------------------------------------------------
ui <- dashboardPage(
  skin = "blue",
  dashboardHeader(title = tags$span(
    "CTG Assay IC50 Analyser",
    tags$br(),
    tags$small(style = "font-size:11px; font-weight:normal; color:#cce0ff;",
               "Created by Ashley Ng")
  )),

  dashboardSidebar(
    sidebarMenu(
      menuItem("1. Upload & Preview",      tabName = "upload",    icon = icon("upload")),
      menuItem("2. Plate Layout",          tabName = "layout",    icon = icon("th")),
      menuItem("3. Concentrations",        tabName = "concs",     icon = icon("flask")),
      menuItem("4. Results & Curves",      tabName = "results",   icon = icon("chart-line")),
      menuItem("5. IC Unit Conversion",    tabName = "ic_conv",   icon = icon("exchange-alt")),
      menuItem("6. Export",                tabName = "export",    icon = icon("download"))
    )
  ),

  dashboardBody(
    useShinyjs(),
    tags$head(tags$style(HTML("
      .content-wrapper { background: #f4f6f9; }
      .tbl-note { font-size:11px; color:#777; margin-top:4px; }
      .tbl-note b { color:#555; }

      /* ── Fix invisible white text in DT inline cell editors ── */
      table.dataTable input[type='text'],
      table.dataTable input[type='number'],
      table.dataTable textarea {
        color: #222 !important;
        background-color: #fffde7 !important;
        border: 1px solid #f0c040 !important;
        border-radius: 3px;
        padding: 2px 4px;
      }
    "))),

    tabItems(

      # ── Tab 1: Upload ────────────────────────────────────────────────────
      tabItem(tabName = "upload",
        fluidRow(
          box(title = "Upload CTG File", width = 5, status = "primary",
              fileInput("file1", "Choose Excel file (.xlsx)",
                        accept = ".xlsx", buttonLabel = "Browse..."),
              uiOutput("sheet_selector"),
              actionButton("load_btn", "Load Plate", icon = icon("table"),
                           class = "btn-primary btn-lg"),
              br(), br(),
              p(class = "tbl-note",
                HTML("Plate rows A\u2013H: Row A = vehicle/DMSO control (100% viability reference);
                     Rows B\u2013H = 7-point drug dilution series ascending in concentration
                     (Row H = highest / starting concentration)."),
                br(), "Plate columns 1\u201312 = drug \u00d7 sample groups.")
          ),
          box(title = "Raw Plate (luminescence counts)", width = 7, status = "info",
              p(class = "tbl-note",
                "Rows A\u2013H are dilution steps. Columns are drug\u00d7sample groups."),
              DT::dataTableOutput("raw_plate")
          )
        ),
        fluidRow(
          box(title = "Upload Data Model (optional)", width = 12, status = "success",
              p(class = "tbl-note",
                HTML("<b>Optional:</b> Upload a previously saved Data Model (.xlsx) to
                     pre-populate the Plate Layout and Starting Concentrations.
                     Download a blank template from the <b>Export</b> tab first.")),
              fluidRow(
                column(6,
                  fileInput("datamodel_file", "Choose Data Model file (.xlsx)",
                            accept = ".xlsx", buttonLabel = "Browse...")
                ),
                column(3,
                  br(),
                  actionButton("load_datamodel_btn", "Apply Data Model",
                               icon  = icon("file-import"),
                               class = "btn-success",
                               title = "Populates Layout and Concentrations tabs from the uploaded template")
                ),
                column(3,
                  p(class = "tbl-note", br(),
                    "Load the CTG plate file first, then apply the Data Model.")
                )
              )
          )
        )
      ),

      # ── Tab 2: Plate Layout ──────────────────────────────────────────────
      tabItem(tabName = "layout",
        fluidRow(
          box(title = "Column Annotations", width = 12, status = "warning",
              p(class = "tbl-note",
                HTML("<b>Instructions:</b> For each plate column set
                     <b>Sample</b> (cell line / genotype), <b>Drug</b> (compound),
                     <b>Replicate</b> (1, 2, 3\u2026 within that drug\u00d7sample group),
                     <b>Unit</b>, and tick <b>Blank</b> for media-only columns.
                     Columns sharing the same Sample <em>and</em> Drug name are averaged
                     as replicates before curve fitting.")),
              br(),
              fluidRow(
                column(3,
                  numericInput("n_replicates",
                               "Number of replicates per group:",
                               value = 2, min = 1, max = 12, step = 1)
                ),
                column(3,
                  br(),
                  actionButton("auto_rep_btn",
                               "Auto-populate replicates",
                               icon  = icon("clone"),
                               class = "btn-warning",
                               title = "Fills columns after each rep-1 anchor with the same Sample & Drug and increments Replicate number. Overwrites existing values.")
                ),
                column(6,
                  p(class = "tbl-note",
                    br(),
                    HTML("Set <b>Replicate = 1</b> for the first column of each group,
                         then click <b>Auto-populate</b> to fill subsequent replicate columns.
                         You can still edit any cell afterwards."))
                )
              ),
              br(),
              DT::dataTableOutput("col_annot_tbl")
          )
        ),
        fluidRow(
          box(title = "Control Row (Normalisation)", width = 12, status = "success",
              p(class = "tbl-note",
                HTML("Select the plate row that contains <b>vehicle/DMSO-only wells</b>
                     (no drug — these define 100% viability).<br>
                     <b>Convention used here:</b> the selected control row is excluded from
                     curve fitting. The remaining rows carry the drug dilution series,
                     ascending in concentration from the row <em>after</em> the control
                     down to the last row — i.e. if control = Row A, then
                     Row B = lowest dose and Row H = highest dose (starting concentration).")),
              fluidRow(
                column(4,
                  selectInput("ctrl_row", "Control row:",
                              choices = LETTERS[1:8], selected = "A")
                ),
                column(8,
                  p(class = "tbl-note", br(),
                    HTML("The <b>Starting Concentration</b> in the Concentrations tab refers
                         to the <b>highest drug dose on the plate</b> (the row furthest from
                         the control row). The dilution series is calculated ascending toward
                         that row."))
                )
              )
          )
        )
      ),

      # ── Tab 3: Concentrations ────────────────────────────────────────────
      tabItem(tabName = "concs",
        fluidRow(
          box(title = "Global Dilution Settings", width = 12, status = "primary",
              fluidRow(
                column(4,
                  numericInput("dilution_factor",
                               "Dilution factor (e.g. 3 for 1:3 serial dilution):",
                               value = 3, min = 1.01, max = 1000, step = 0.5)
                ),
                column(4,
                  br(),
                  actionButton("recalc_all_btn",
                               "Recalculate all drugs from starting conc",
                               icon  = icon("sync"),
                               class = "btn-info")
                ),
                column(4,
                  p(class = "tbl-note", br(),
                    HTML("Changing the dilution factor and clicking
                         <b>Recalculate</b> resets all concentrations from each
                         drug's starting value. Manual edits in the table below
                         are preserved until you recalculate."))
                )
              )
          )
        ),
        fluidRow(
          box(title = "Starting Concentrations per Drug", width = 12, status = "warning",
              p(class = "tbl-note",
                HTML("Edit the <b>Starting_Conc</b> — the <b>highest drug concentration</b>
                     on the plate (assigned to the row furthest from the control row).
                     The full dilution series (7 steps, ascending away from the control) is
                     recalculated automatically. You can fine-tune individual steps below.")),
              br(),
              DT::dataTableOutput("drug_start_tbl")
          )
        ),
        fluidRow(
          box(title = "Full Dilution Series (editable per step)", width = 12, status = "info",
              p(class = "tbl-note",
                HTML("Drug, Row, and Step columns are locked.
                     Edit <b>Concentration</b> to override individual steps.<br>
                     <b>Step 1 = lowest dose (row nearest the control row);
                     Step 7 = highest dose / starting concentration (row furthest from control).</b>
                     The control row itself carries no drug and is excluded from the fit.")),
              DT::dataTableOutput("drug_conc_tbl")
          )
        )
      ),

      # ── Tab 4: Results ───────────────────────────────────────────────────
      tabItem(tabName = "results",
        fluidRow(
          box(title = "Run Analysis", width = 12, status = "primary",
              actionButton("run_analysis", "Run Dose-Response Analysis",
                           icon = icon("play"), class = "btn-success btn-lg"),
              br(), br(),
              p("Fits a 4-parameter log-logistic (Hill) curve per Drug \u00d7 Sample."),
              p("y = d + (c \u2212 d) / (1 + (x/e)^b)   |   b=slope, c=lower, d=upper, e=IC50")
          )
        ),
        fluidRow(
          box(title = "IC Values Summary (original concentration units, linear scale)",
              width = 12, status = "success",
              DT::dataTableOutput("ic_table"),
              p(class = "tbl-note",
                HTML("IC values are in the <b>original linear concentration units</b>
                     entered in the Concentrations tab (e.g. µM, nM, UN) — <b>not log scale</b>.
                     The dose-response x-axis is log\u2081\u2080 for visualisation only;
                     the IC estimates themselves are back-transformed to the real concentration scale.
                     95% delta CI shown for IC50. See the <b>IC Unit Conversion</b> tab for nM values."))
          )
        ),
        fluidRow(
          box(title = "Dose-Response Curves", width = 12, status = "info",
              fluidRow(
                column(5, selectInput("plot_drug",   "Drug:",   choices = NULL)),
                column(5, selectInput("plot_sample", "Sample:", choices = NULL)),
                column(2, br(),
                       downloadButton("dl_plot", "PDF", class = "btn-info btn-sm"))
              ),
              plotOutput("dr_plot", height = "500px")
          )
        )
      ),

      # ── Tab 5: IC Unit Conversion ────────────────────────────────────────
      tabItem(tabName = "ic_conv",
        fluidRow(
          box(title = "IC Values — Unit Converted", width = 12, status = "success",
              p(class = "tbl-note",
                HTML("IC values are in the <b>original linear concentration units</b>
                     entered in the Concentrations tab — <b>NOT log-transformed</b>.<br><br>
                     <b>Column layout:</b><br>
                     &bull; <b>IC10 … IC90_&lt;unit&gt;</b> — original units
                     (e.g. IC50_\u00b5M for micromolar drugs).<br>
                     &bull; <b>IC10 … IC90_nM</b> — nanomolar conversion
                     (original \u00d7 1\u202f000, <em>only for \u00b5M drugs</em>;
                     NA for drugs entered in other units such as U, nM, UN).<br><br>
                     All columns are also included in the <b>Download IC Table (.xlsx)</b>
                     on the Export tab.")),
              br(),
              DT::dataTableOutput("ic_conv_table")
          )
        )
      ),

      # ── Tab 6: Export ────────────────────────────────────────────────────
      tabItem(tabName = "export",
        fluidRow(
          box(title = "Data Model Template", width = 12, status = "warning",
              p(class = "tbl-note",
                HTML("<b>Step 1:</b> Download the blank template, fill it in with your
                     sample names, drugs, replicates, units, blank flags, starting
                     concentrations and dilution factor, then upload it on the
                     <b>Upload</b> tab to auto-populate the Layout and Concentrations tabs.")),
              fluidRow(
                column(4,
                  downloadButton("dl_datamodel_template", "Download Blank Template (.xlsx)",
                                 class = "btn-warning")
                ),
                column(4,
                  downloadButton("dl_datamodel_current", "Download Current Settings (.xlsx)",
                                 class = "btn-info",
                                 title = "Saves your current Layout + Concentration settings as a reusable Data Model")
                ),
                column(4,
                  p(class = "tbl-note", br(),
                    HTML("Grey placeholder text in the template shows example values.
                         Replace all grey cells with your own data before uploading."))
                )
              )
          )
        ),
        fluidRow(
          box(title = "Export Results", width = 5, status = "primary",
              downloadButton("dl_results_xlsx", "Download IC Table (.xlsx)",
                             class = "btn-success"),
              br(), br(),
              downloadButton("dl_norm_xlsx", "Download Normalised Data (.xlsx)",
                             class = "btn-info"),
              br(), br(),
              downloadButton("dl_all_plots", "Download All Curves (PDF)",
                             class = "btn-warning")
          ),
          box(title = "Normalised Data Preview", width = 7, status = "info",
              DT::dataTableOutput("norm_preview")
          )
        )
      )
    )
  )
)

# ---- SERVER ---------------------------------------------------------------
server <- function(input, output, session) {

  rv <- reactiveValues(
    plate_mat        = NULL,  # 8×12 numeric matrix (rows=dilution steps A-H, cols=plate cols 1-12)
    col_annot        = NULL,  # data.frame: Column, Sample, Drug, Replicate, Unit, Blank
    drug_start_concs = NULL,  # data.frame: Drug, Starting_Conc  (one row per unique drug)
    drug_concs       = NULL,  # data.frame: Drug, Row, Step, Concentration  (8 rows per drug)
    norm_data        = NULL,  # long-form normalised data
    fit_results      = NULL,  # named list: "Drug | Sample" -> fit info
    ic_summary       = NULL,  # data.frame: IC values in original units (linear, NOT log)
    ic_converted     = NULL,  # data.frame: IC values with unit conversions (e.g. µM → nM)
    ctg_filename     = NULL,  # original name of the uploaded CTG plate xlsx
    datamodel_filename = NULL # original name of the uploaded Data Model xlsx (if any)
  )

  # ── Helper: detect whether a unit string means micromolar ──────────────────
  # toupper() corrupts U+00B5 (micro sign) → Greek capital mu, breaking %in% checks.
  # Use a regex instead: match µ (U+00B5), μ (U+03BC Greek mu), or plain 'u', followed by 'M'.
  is_micromolar <- function(unit_str) {
    grepl("^[\u00b5\u03bcuU][Mm]$", trimws(unit_str))
  }

  # ── Helper: build unit-converted IC table from ic_summary ──────────────────
  # Produces a fixed-column-order table:
  #   Drug | Sample | Unit | Hill_Slope |
  #   IC10_<unit> … IC90_<unit> IC50_CI_Lo_<unit> IC50_CI_Hi_<unit> |   ← original units
  #   IC10_nM     … IC90_nM     IC50_CI_Lo_nM     IC50_CI_Hi_nM          ← nM (µM rows only)
  #
  # For non-µM drugs (e.g. U, nM, UN) the _nM columns are NA.
  # This gives a consistent, readable layout regardless of how many unit types are present.
  make_ic_converted <- function(ic_df) {
    if (is.null(ic_df) || nrow(ic_df) == 0) return(data.frame())

    ic_cols <- c("IC10","IC30","IC50","IC60","IC90","IC50_CI_Lo","IC50_CI_Hi")

    # Determine the dominant original-unit label for column naming.
    # Each row contributes its own unit suffix; we'll use per-row units below.
    result <- lapply(seq_len(nrow(ic_df)), function(i) {
      row    <- ic_df[i, ]
      unit   <- trimws(row$Unit)
      is_uM  <- is_micromolar(unit)

      out <- data.frame(
        Drug       = row$Drug,
        Sample     = row$Sample,
        Unit       = unit,
        Hill_Slope = row$Hill_Slope,
        stringsAsFactors = FALSE
      )

      # Original-unit IC columns (suffixed with the drug's own unit)
      for (col in ic_cols) {
        out[[paste0(col, "_", unit)]] <- row[[col]]
      }

      # nM conversion: only for µM drugs (×1000); NA for all others
      for (col in ic_cols) {
        out[[paste0(col, "_nM")]] <- if (is_uM) round(row[[col]] * 1000, 4) else NA_real_
      }

      out
    })

    # Merge: fill any missing original-unit columns with NA so rbind works cleanly
    all_cols <- unique(unlist(lapply(result, names)))
    result   <- lapply(result, function(r) {
      missing <- setdiff(all_cols, names(r))
      r[missing] <- NA_real_
      r[all_cols]
    })
    do.call(rbind, result)
  }

  # ── Sheet selector ─────────────────────────────────────────────────────────
  output$sheet_selector <- renderUI({
    req(input$file1)
    tmp <- tempfile(fileext = ".xlsx")
    file.copy(input$file1$datapath, tmp)
    on.exit(unlink(tmp))
    sheets <- excel_sheets(tmp)
    selectInput("sheet_name", "Select sheet:", choices = sheets, selected = sheets[1])
  })

  # ── Load plate ─────────────────────────────────────────────────────────────
  observeEvent(input$load_btn, {
    req(input$file1, input$sheet_name)
    tryCatch({
      mat <- extract_plate(input$file1$datapath, input$sheet_name)
      rv$plate_mat <- mat

      # Default column annotations (all 12 columns)
      rv$col_annot <- data.frame(
        Column    = 1:12,
        Sample    = paste0("Sample_", 1:12),
        Drug      = paste0("Drug_", 1:12),
        Replicate = 1L,
        Unit      = "\u00b5M",
        Blank     = FALSE,
        stringsAsFactors = FALSE
      )

      # Store original filename for provenance tracking in downloads
      rv$ctg_filename <- input$file1$name

      # Initialise starting concs and dilution series (default 10 µM, 1:3)
      init_drug_concs(rv$col_annot,
                      factor   = isolate(input$dilution_factor),
                      ctrl_row = isolate(input$ctrl_row))

      showNotification(
        "Plate loaded! Go to the Layout tab to annotate columns.",
        type = "message", duration = 6
      )
    }, error = function(e) {
      showNotification(paste("Error loading plate:", e$message), type = "error")
    })
  })

  # ── Helper: (re)build drug_start_concs and drug_concs from col_annot ───────
  # ctrl_row: the plate row used as the vehicle/DMSO control (default "A").
  # Drug rows are ordered nearest→furthest from ctrl_row = lowest→highest dose.
  init_drug_concs <- function(ca, factor = 3, ctrl_row = "A") {
    # Ensure Blank is always treated as logical (DT edits return character "TRUE"/"FALSE")
    ca$Blank <- as.logical(ca$Blank)
    ca$Blank[is.na(ca$Blank)] <- FALSE

    drugs <- unique(ca$Drug[!ca$Blank & nzchar(trimws(ca$Drug))])
    if (length(drugs) == 0) return()

    old_starts <- rv$drug_start_concs
    old_concs  <- rv$drug_concs

    # Row letters in ascending-concentration order for this ctrl_row setting
    d_rows <- LETTERS[drug_rows_ordered(ctrl_row)]  # length 7

    starts <- lapply(drugs, function(d) {
      # Always take exactly one scalar starting concentration
      s <- if (!is.null(old_starts) && d %in% old_starts$Drug)
        as.numeric(old_starts$Starting_Conc[old_starts$Drug == d][1])
      else 10
      if (length(s) != 1 || is.na(s)) s <- 10
      data.frame(Drug = d, Starting_Conc = s, stringsAsFactors = FALSE)
    })
    rv$drug_start_concs <- do.call(rbind, starts)

    concs <- lapply(drugs, function(d) {
      # Preserve manually-edited concs only if the row set still matches
      if (!is.null(old_concs) && d %in% old_concs$Drug) {
        existing <- old_concs[old_concs$Drug == d, ]
        if (nrow(existing) == 7 && identical(existing$Row, d_rows)) return(existing)
      }
      # Look up starting conc — guaranteed scalar by the starts block above
      s <- rv$drug_start_concs$Starting_Conc[rv$drug_start_concs$Drug == d][1]
      if (length(s) != 1 || is.na(s)) s <- 10
      # Steps 1-7: lowest dose (nearest ctrl_row) → highest dose (furthest)
      data.frame(
        Drug          = d,
        Row           = d_rows,
        Step          = 1:7,
        Concentration = calc_concs(s, factor),
        stringsAsFactors = FALSE
      )
    })
    rv$drug_concs <- do.call(rbind, concs)
  }

  # ── Helper: recalc one drug's series from its starting conc ────────────────
  recalc_drug <- function(drug_name, factor = isolate(input$dilution_factor)) {
    req(rv$drug_start_concs, rv$drug_concs)
    s  <- rv$drug_start_concs$Starting_Conc[rv$drug_start_concs$Drug == drug_name]
    dc <- rv$drug_concs
    idx <- which(dc$Drug == drug_name)
    if (length(idx) == 7) {
      dc$Concentration[idx] <- calc_concs(s, factor)
      rv$drug_concs <- dc
    }
  }

  # ── Raw plate table ────────────────────────────────────────────────────────
  output$raw_plate <- DT::renderDataTable({
    req(rv$plate_mat)
    df <- cbind(Row = rownames(rv$plate_mat), as.data.frame(rv$plate_mat))
    DT::datatable(df, rownames = FALSE,
                  options = list(dom = "t", pageLength = 8, scrollX = TRUE),
                  class = "compact")
  })

  # ── Column annotation table ────────────────────────────────────────────────
  output$col_annot_tbl <- DT::renderDataTable({
    req(rv$col_annot)
    DT::datatable(
      rv$col_annot,
      editable = list(target = "cell", disable = list(columns = 0L)),
      rownames = FALSE,
      options  = list(dom = "t", pageLength = 12, scrollX = TRUE),
      class    = "compact"
    )
  })

  observeEvent(input$col_annot_tbl_cell_edit, {
    info     <- input$col_annot_tbl_cell_edit
    col_name <- names(rv$col_annot)[info$col + 1]
    val      <- info$value

    # Coerce types — DT always returns character from cell edits
    if (col_name == "Replicate") val <- as.integer(val)
    if (col_name == "Blank")     val <- toupper(trimws(val)) %in% c("TRUE","1","YES")

    rv$col_annot[info$row, col_name] <- val

    # Rebuild concentration tables when Drug or Sample name changes
    if (col_name %in% c("Drug", "Sample", "Blank")) {
      init_drug_concs(rv$col_annot,
                      factor   = isolate(input$dilution_factor),
                      ctrl_row = isolate(input$ctrl_row))
    }
  })

  # ── Auto-populate replicate columns ────────────────────────────────────────
  observeEvent(input$auto_rep_btn, {
    req(rv$col_annot)
    n_rep <- as.integer(input$n_replicates)
    if (is.na(n_rep) || n_rep < 1) return()

    ca    <- rv$col_annot
    n_col <- nrow(ca)

    # Walk through columns in blocks of n_rep, using the first col as anchor
    for (start in seq(1, n_col, by = n_rep)) {
      end    <- min(start + n_rep - 1, n_col)
      anchor <- ca[start, ]
      for (j in start:end) {
        ca$Sample[j]    <- anchor$Sample
        ca$Drug[j]      <- anchor$Drug
        ca$Unit[j]      <- anchor$Unit
        ca$Blank[j]     <- anchor$Blank
        ca$Replicate[j] <- as.integer(j - start + 1L)
      }
    }

    rv$col_annot <- ca
    init_drug_concs(ca,
                    factor   = isolate(input$dilution_factor),
                    ctrl_row = isolate(input$ctrl_row))
    showNotification(
      paste0("Auto-populated: ", n_rep, " replicates per group across all 12 columns."),
      type = "message"
    )
  })

  # ── Starting concentration table ───────────────────────────────────────────
  output$drug_start_tbl <- DT::renderDataTable({
    req(rv$drug_start_concs)
    DT::datatable(
      rv$drug_start_concs,
      editable = list(target = "cell", disable = list(columns = 0L)),  # Drug locked
      rownames = FALSE,
      options  = list(dom = "t", pageLength = 24, scrollX = TRUE),
      class    = "compact"
    )
  })

  observeEvent(input$drug_start_tbl_cell_edit, {
    info <- input$drug_start_tbl_cell_edit
    col_name <- names(rv$drug_start_concs)[info$col + 1]
    if (col_name == "Starting_Conc") {
      rv$drug_start_concs[info$row, "Starting_Conc"] <- as.numeric(info$value)
      # Auto-recalculate that drug's full dilution series
      drug_name <- rv$drug_start_concs$Drug[info$row]
      recalc_drug(drug_name)
    }
  })

  # ── Full dilution series table ─────────────────────────────────────────────
  output$drug_conc_tbl <- DT::renderDataTable({
    req(rv$drug_concs)
    DT::datatable(
      rv$drug_concs,
      editable = list(target = "cell", disable = list(columns = 0:2)),  # Drug/Row/Step locked
      rownames = FALSE,
      options  = list(dom = "t", pageLength = 96, scrollX = TRUE),
      class    = "compact"
    )
  })

  observeEvent(input$drug_conc_tbl_cell_edit, {
    info <- input$drug_conc_tbl_cell_edit
    col_name <- names(rv$drug_concs)[info$col + 1]
    if (col_name == "Concentration") {
      rv$drug_concs[info$row, "Concentration"] <- as.numeric(info$value)
    }
  })

  # ── Recalculate ALL drugs from starting concs + current dilution factor ────
  observeEvent(input$recalc_all_btn, {
    req(rv$drug_concs, rv$drug_start_concs)
    factor <- input$dilution_factor
    dc <- rv$drug_concs
    for (d in unique(dc$Drug)) {
      s   <- rv$drug_start_concs$Starting_Conc[rv$drug_start_concs$Drug == d]
      idx <- which(dc$Drug == d)
      dc$Concentration[idx] <- calc_concs(s, factor, n = length(idx))
    }
    rv$drug_concs <- dc
    showNotification("All concentrations recalculated from starting values.", type = "message")
  })

  # ── CORE ANALYSIS ──────────────────────────────────────────────────────────
  observeEvent(input$run_analysis, {
    req(rv$plate_mat, rv$col_annot, rv$drug_concs)

    mat      <- rv$plate_mat
    ca       <- rv$col_annot
    dc       <- rv$drug_concs
    ctrl_row <- input$ctrl_row

    # Ensure Blank is logical (guards against character "TRUE"/"FALSE" from DT)
    ca$Blank <- as.logical(ca$Blank)
    ca$Blank[is.na(ca$Blank)] <- FALSE

    # Active (non-blank, non-empty-drug) columns
    active <- ca[!ca$Blank & nzchar(trimws(ca$Drug)), ]
    if (nrow(active) == 0) {
      showNotification("No active columns found. Check Layout tab.", type = "error")
      return()
    }

    # --- 1. Build long-form data ---------------------------------------------
    # Drug row indices ordered nearest→furthest from ctrl_row = Step 1 (lowest) → Step 7 (highest).
    # drug_rows_ordered() handles any ctrl_row choice correctly:
    #   ctrl_row="A" → drug_rows = [2,3,4,5,6,7,8]  (B lowest, H highest)
    #   ctrl_row="H" → drug_rows = [7,6,5,4,3,2,1]  (G lowest, A highest)
    drug_rows <- drug_rows_ordered(ctrl_row)   # integer row indices, length 7

    long_list <- vector("list", nrow(active) * 8)
    idx <- 1L
    for (ci in seq_len(nrow(active))) {
      col_num   <- active$Column[ci]
      sample    <- active$Sample[ci]
      drug      <- active$Drug[ci]
      replicate <- active$Replicate[ci]
      unit      <- active$Unit[ci]

      drug_dc <- dc[dc$Drug == drug, ]
      drug_dc <- drug_dc[order(drug_dc$Step), ]

      for (ri in 1:8) {
        plate_row <- LETTERS[ri]
        # The selected ctrl_row is the vehicle/DMSO control — no drug concentration.
        # The remaining 7 rows carry ascending drug concentrations mapped to drug_dc Steps 1-7.
        conc <- if (plate_row == ctrl_row) {
          NA_real_
        } else {
          # Position of this row within the drug rows (1 = lowest dose, 7 = highest)
          step_idx <- which(drug_rows == ri)
          if (length(step_idx) == 1 && step_idx <= nrow(drug_dc))
            drug_dc$Concentration[step_idx]
          else
            NA_real_
        }
        raw  <- mat[plate_row, as.character(col_num)]

        long_list[[idx]] <- data.frame(
          Plate_Col = col_num,
          Sample    = sample,
          Drug      = drug,
          Replicate = replicate,
          Unit      = unit,
          Row       = plate_row,
          Step      = ri,
          Conc      = conc,
          Raw       = as.numeric(raw),
          stringsAsFactors = FALSE
        )
        idx <- idx + 1L
      }
    }
    long_df <- do.call(rbind, long_list)

    # --- 2. Normalise WITHIN each plate column to its own control well --------
    # Each column is normalised independently to its own ctrl_row Raw value.
    # This corrects for technical variation in cell seeding or drug loading
    # per column before replicates are averaged.
    ctrl_per_col <- long_df[long_df$Row == ctrl_row,
                            c("Plate_Col", "Raw")]
    names(ctrl_per_col)[names(ctrl_per_col) == "Raw"] <- "Ctrl_Raw"

    long_df <- long_df %>%
      left_join(ctrl_per_col, by = "Plate_Col") %>%
      mutate(
        Viability  = (Raw / Ctrl_Raw) * 100,
        Inhibition = 100 - Viability
      )

    rv$norm_data <- long_df

    # --- 3. Average column-normalised replicates then fit 4PL ----------------
    # Now that each replicate column is independently normalised, averaging
    # across replicates (same Drug+Sample+Step+Conc) is unconfounded.
    avg_df <- long_df %>%
      filter(Row != ctrl_row) %>%
      group_by(Drug, Sample, Unit, Step, Conc) %>%
      summarise(
        Viability     = mean(Viability,   na.rm = TRUE),
        Inhibition    = mean(Inhibition,  na.rm = TRUE),
        SD_Viability  = sd(Viability,     na.rm = TRUE),
        SD_Inhibition = sd(Inhibition,    na.rm = TRUE),
        N_Replicates  = n(),
        .groups = "drop"
      ) %>%
      filter(!is.na(Conc), Conc > 0, !is.na(Inhibition))

    fit_list <- list()
    ic_rows  <- list()
    combos   <- avg_df %>% distinct(Drug, Sample) %>% arrange(Drug, Sample)

    for (ci in seq_len(nrow(combos))) {
      d <- combos$Drug[ci]
      s <- combos$Sample[ci]
      u <- unique(avg_df$Unit[avg_df$Drug == d & avg_df$Sample == s])[1]

      sub <- avg_df %>% filter(Drug == d, Sample == s) %>% arrange(Conc)
      key <- paste0(d, " | ", s)

      if (nrow(sub) < 4) {
        fit_list[[key]] <- list(fit = NULL, data = sub, drug = d, sample = s, unit = u)
        next
      }

      # Fit 4-parameter log-logistic on INHIBITION (increasing curve: 0 → 100%).
      # LL.4 parameters: b=Hill slope, c=lower asymptote, d=upper asymptote, e=EC50 (IC50).
      # Constraints: lower (c) >= 0; upper (d) <= 100 so the fit stays biologically sensible.
      fit <- tryCatch(
        drc::drm(Inhibition ~ Conc, data = sub,
                 fct     = drc::LL.4(names = c("b","c","d","e")),
                 lowerl  = c(NA,   0, NA, NA),
                 upperl  = c(NA, NA, 100, NA),
                 control = drmc(noMessage = TRUE)),
        error = function(e) NULL
      )

      fit_list[[key]] <- list(fit = fit, data = sub, drug = d, sample = s, unit = u)

      if (!is.null(fit)) {
        b_par <- coef(fit)["b:(Intercept)"]
        # ED() with type="relative" (default) on an increasing inhibition curve:
        # ED(fit, 50) → concentration where Inhibition = 50% of the fitted range → IC50.
        # ED(fit, 10) → IC10 (lower conc), ED(fit, 90) → IC90 (higher conc).
        # This correctly gives IC10 < IC30 < IC50 < IC60 < IC90.
        ed_vals <- tryCatch(
          drc::ED(fit, c(10, 30, 50, 60, 90),
                  interval = "delta",
                  type     = "relative",
                  display  = FALSE),
          error = function(e) NULL
        )
        if (!is.null(ed_vals)) {
          ic_rows[[length(ic_rows)+1]] <- data.frame(
            Drug       = d,
            Sample     = s,
            Unit       = u,
            Hill_Slope = round(abs(b_par), 3),
            IC10       = round(ed_vals["e:1:10","Estimate"], 4),
            IC30       = round(ed_vals["e:1:30","Estimate"], 4),
            IC50       = round(ed_vals["e:1:50","Estimate"], 4),
            IC60       = round(ed_vals["e:1:60","Estimate"], 4),
            IC90       = round(ed_vals["e:1:90","Estimate"], 4),
            IC50_CI_Lo = round(ed_vals["e:1:50","Lower"],    4),
            IC50_CI_Hi = round(ed_vals["e:1:50","Upper"],    4),
            stringsAsFactors = FALSE
          )
        }
      }
    }

    rv$fit_results <- fit_list
    rv$ic_summary  <- if (length(ic_rows) > 0) do.call(rbind, ic_rows) else data.frame()
    rv$ic_converted <- make_ic_converted(rv$ic_summary)

    all_drugs   <- unique(combos$Drug)
    all_samples <- unique(combos$Sample)
    updateSelectInput(session, "plot_drug",   choices = all_drugs,   selected = all_drugs[1])
    updateSelectInput(session, "plot_sample", choices = all_samples, selected = all_samples[1])

    showNotification(
      paste0("Analysis complete: ", length(fit_list), " Drug\u00d7Sample curves fitted."),
      type = "message"
    )
  })

  # ── Plot sample selector follows drug selection ────────────────────────────
  observeEvent(input$plot_drug, {
    req(rv$fit_results)
    keys    <- names(rv$fit_results)
    pattern <- paste0("^", gsub("([.|()\\^{}+$*?]|\\[|\\])", "\\\\\\1",
                                input$plot_drug), " \\| ")
    matched <- keys[grepl(pattern, keys)]
    samples <- sub("^.*? \\| ", "", matched)
    updateSelectInput(session, "plot_sample", choices = samples, selected = samples[1])
  })

  # ── IC table (original units, clearly labelled) ────────────────────────────
  output$ic_table <- DT::renderDataTable({
    req(rv$ic_summary)
    if (nrow(rv$ic_summary) == 0)
      return(DT::datatable(data.frame(Message = "No curves fitted yet.")))

    df <- rv$ic_summary
    # Rename IC columns to include unit suffix so it is unambiguous
    unit_label <- if (length(unique(df$Unit)) == 1) unique(df$Unit) else "orig.unit"
    ic_cols    <- c("IC10","IC30","IC50","IC60","IC90","IC50_CI_Lo","IC50_CI_Hi")
    new_names  <- paste0(ic_cols, " (", unit_label, ")")
    names(df)[match(ic_cols, names(df))] <- new_names

    DT::datatable(df, rownames = FALSE,
                  options = list(scrollX = TRUE, pageLength = 20),
                  class = "compact stripe") %>%
      DT::formatRound(columns = c("Hill_Slope", new_names), digits = 4)
  })

  # ── IC unit-converted table ────────────────────────────────────────────────
  output$ic_conv_table <- DT::renderDataTable({
    req(rv$ic_converted)
    if (is.null(rv$ic_converted) || nrow(rv$ic_converted) == 0)
      return(DT::datatable(data.frame(Message = "Run analysis first.")))

    df <- rv$ic_converted

    # Separate numeric columns into original-unit and nM groups for formatting
    nm_cols   <- names(df)[grepl("_nM$",  names(df)) & sapply(df, is.numeric)]
    orig_cols <- names(df)[!grepl("_nM$", names(df)) & sapply(df, is.numeric) &
                             !names(df) %in% c("Hill_Slope")]

    DT::datatable(
      df,
      rownames = FALSE,
      caption  = htmltools::tags$caption(
        style = "color:#555; font-size:11px; text-align:left; padding:4px 0;",
        htmltools::HTML(
          paste0("Columns suffixed <b>_", unique(df$Unit[!sapply(df$Unit, is_micromolar)]),
                 "</b> are in original units. ",
                 "Columns suffixed <b>_nM</b> are \u00d7\u202f1\u202f000 (µM drugs only). ",
                 "NA = not applicable for that drug's unit.")
        )
      ),
      options  = list(scrollX = TRUE, pageLength = 20),
      class    = "compact stripe"
    ) %>%
      DT::formatRound(columns = orig_cols, digits = 4) %>%
      DT::formatRound(columns = "Hill_Slope", digits = 3) %>%
      { if (length(nm_cols) > 0) DT::formatRound(., columns = nm_cols, digits = 1) else . }
  })

  # ── Dose-response plot ─────────────────────────────────────────────────────
  # ctg_file: optional filename string appended to the PDF caption (for downloads).
  # Leave NULL for the interactive plot (caption is omitted).
  make_dr_plot <- function(drug, sample, ctg_file = NULL) {
    key  <- paste0(drug, " | ", sample)
    req(rv$fit_results, key %in% names(rv$fit_results))
    item <- rv$fit_results[[key]]
    fit  <- item$fit
    sub  <- item$data
    unit <- item$unit

    ic_line_cols <- c("#e74c3c","#e67e22","#27ae60","#8e44ad","#2c3e50")

    has_reps <- "N_Replicates" %in% names(sub) && max(sub$N_Replicates, na.rm = TRUE) > 1
    has_sd   <- has_reps && "SD_Inhibition" %in% names(sub) &&
                  any(!is.na(sub$SD_Inhibition))

    # Replicate/SD description on its own line to prevent subtitle truncation
    rep_line <- if (has_reps)
      paste0(sub$N_Replicates[1], " replicates \u2014 mean \u00b1 SD shown.")
    else
      "Single replicate."

    # PDF caption: creator credit + timestamp + source file (only when saving to file)
    pdf_caption <- if (!is.null(ctg_file)) {
      paste0("Created with CTG Assay IC50 Analyser by Ashley Ng  |  ",
             format(Sys.time(), "%Y-%m-%d %H:%M"),
             "  |  Source: ", ctg_file)
    } else {
      NULL
    }

    # Plot on INHIBITION axis (increasing curve: low conc → 0%, high conc → 100%).
    # IC reference lines at 10/30/50/60/90% inhibition are drawn horizontally.
    p <- ggplot(sub, aes(x = Conc, y = Inhibition)) +
      scale_x_log10(labels = scales::label_number()) +
      coord_cartesian(ylim = c(-15, 110)) +
      labs(
        title    = paste0(drug, "  \u2014  ", sample),
        subtitle = paste0(
          "% Inhibition = 100 \u2212 (Raw / DMSO control) \u00d7 100, normalised per column.\n",
          "Curve fit: 4-parameter log-logistic on inhibition (increasing).  ",
          rep_line
        ),
        x       = paste0("Concentration (", unit, ")  [log\u2081\u2080 scale]"),
        y       = "Mean % Inhibition",
        caption = pdf_caption
      ) +
      geom_hline(yintercept = c(10, 30, 50, 60, 90),
                 linetype = "dashed", colour = ic_line_cols, alpha = 0.6) +
      annotate("text",
               x      = min(sub$Conc, na.rm = TRUE),
               y      = c(13, 33, 53, 63, 93),
               label  = c("IC10","IC30","IC50","IC60","IC90"),
               colour = ic_line_cols,
               hjust  = 0, size = 3.2) +
      geom_hline(yintercept = 0,  linetype = "solid", colour = "grey70", linewidth = 0.4) +
      geom_hline(yintercept = 100, linetype = "solid", colour = "grey70", linewidth = 0.4) +
      theme_bw(base_size = 13) +
      theme(plot.title    = element_text(face = "bold"),
            plot.subtitle = element_text(colour = "grey50", size = 9),
            plot.caption  = element_text(colour = "grey65", size = 7,
                                         hjust = 0, margin = margin(t = 6)))

    # SD error bars on inhibition (drawn before points so points sit on top)
    if (has_sd) {
      p <- p +
        geom_errorbar(aes(ymin = Inhibition - SD_Inhibition,
                          ymax = Inhibition + SD_Inhibition),
                      width = 0.06, colour = "#2980b9", alpha = 0.7, linewidth = 0.6)
    }
    p <- p + geom_point(size = 3, colour = "#2980b9")

    if (!is.null(fit)) {
      conc_range <- 10^seq(
        log10(max(min(sub$Conc) * 0.5, 1e-9)),
        log10(max(sub$Conc) * 1.5),
        length.out = 300
      )
      pred_df <- data.frame(
        Conc       = conc_range,
        Inhibition = predict(fit, newdata = data.frame(Conc = conc_range))
      )
      p <- p + geom_line(data = pred_df, colour = "#e74c3c", linewidth = 1.2)

      # IC50 vertical line + label
      if (!is.null(rv$ic_summary) && nrow(rv$ic_summary) > 0) {
        ic_row <- rv$ic_summary[rv$ic_summary$Drug == drug &
                                  rv$ic_summary$Sample == sample, ]
        if (nrow(ic_row) > 0 && !is.na(ic_row$IC50[1])) {
          p <- p +
            geom_vline(xintercept = ic_row$IC50[1], linetype = "dotted",
                       colour = "#27ae60", linewidth = 1) +
            annotate("text", x = ic_row$IC50[1], y = 47,
                     label    = paste0("IC50 = ", round(ic_row$IC50[1], 3), " ", unit),
                     colour   = "#27ae60", hjust = -0.05,
                     fontface = "bold", size = 3.5)
        }
      }
    }
    return(p)
  }

  output$dr_plot <- renderPlot({
    req(input$plot_drug, input$plot_sample, rv$fit_results)
    make_dr_plot(input$plot_drug, input$plot_sample)
  })

  # ── Normalised data preview ────────────────────────────────────────────────
  output$norm_preview <- DT::renderDataTable({
    req(rv$norm_data)
    want_cols <- c("Plate_Col","Sample","Drug","Replicate",
                   "Row","Step","Conc","Unit","Raw","Ctrl_Raw","Viability","Inhibition")
    keep      <- intersect(want_cols, names(rv$norm_data))
    df        <- rv$norm_data[, keep, drop = FALSE]
    round_cols <- setdiff(names(df)[sapply(df, is.numeric)],
                          c("Plate_Col","Step","Replicate"))
    df[round_cols] <- lapply(df[round_cols], round, digits = 2)
    DT::datatable(df, rownames = FALSE,
                  options = list(scrollX = TRUE, pageLength = 16),
                  class = "compact")
  })

  # ── Data Model helpers ─────────────────────────────────────────────────────

  # Build a styled workbook from either blank examples or current rv state
  build_datamodel_wb <- function(use_current = FALSE) {
    wb <- createWorkbook()

    # ── Sheet 1: Plate_Layout ──────────────────────────────────────────────
    addWorksheet(wb, "Plate_Layout")

    if (use_current && !is.null(rv$col_annot)) {
      layout_df <- rv$col_annot
    } else {
      layout_df <- data.frame(
        Column    = 1:12,
        Sample    = paste0("Sample_", rep(1:6, each = 2)),
        Drug      = paste0("Drug_",   rep(1:6, each = 2)),
        Replicate = rep(1:2, times = 6),
        Unit      = "\u00b5M",
        Blank     = FALSE,
        stringsAsFactors = FALSE
      )
    }
    writeData(wb, "Plate_Layout", layout_df, startRow = 1)

    # Style header row
    hdr_style <- createStyle(fontColour = "#FFFFFF", fgFill = "#2c3e50",
                             textDecoration = "bold", halign = "center",
                             border = "Bottom", borderColour = "#888888")
    addStyle(wb, "Plate_Layout", hdr_style,
             rows = 1, cols = 1:ncol(layout_df), gridExpand = TRUE)

    # Style example data rows in grey if blank template
    if (!use_current) {
      eg_style <- createStyle(fontColour = "#999999", fgFill = "#f5f5f5",
                              border = "TopBottomLeftRight", borderColour = "#cccccc")
      addStyle(wb, "Plate_Layout", eg_style,
               rows = 2:(nrow(layout_df)+1), cols = 1:ncol(layout_df),
               gridExpand = TRUE)
    }
    setColWidths(wb, "Plate_Layout", cols = 1:ncol(layout_df),
                 widths = c(8, 18, 18, 11, 8, 8))

    # ── Sheet 2: Concentrations ────────────────────────────────────────────
    addWorksheet(wb, "Concentrations")

    if (use_current && !is.null(rv$drug_start_concs)) {
      conc_df <- data.frame(
        Drug           = rv$drug_start_concs$Drug,
        Starting_Conc  = rv$drug_start_concs$Starting_Conc,
        Dilution_Factor = isolate(input$dilution_factor),
        stringsAsFactors = FALSE
      )
    } else {
      conc_df <- data.frame(
        Drug            = paste0("Drug_", 1:6),
        Starting_Conc   = 10,
        Dilution_Factor = 3,
        stringsAsFactors = FALSE
      )
    }
    writeData(wb, "Concentrations", conc_df, startRow = 1)

    addStyle(wb, "Concentrations", hdr_style,
             rows = 1, cols = 1:3, gridExpand = TRUE)

    if (!use_current) {
      eg_style2 <- createStyle(fontColour = "#999999", fgFill = "#f5f5f5",
                               border = "TopBottomLeftRight", borderColour = "#cccccc")
      addStyle(wb, "Concentrations", eg_style2,
               rows = 2:(nrow(conc_df)+1), cols = 1:3, gridExpand = TRUE)
    }
    setColWidths(wb, "Concentrations", cols = 1:3, widths = c(18, 16, 16))

    # ── Sheet 3: Instructions ──────────────────────────────────────────────
    addWorksheet(wb, "Instructions")
    instructions <- data.frame(
      Instructions = c(
        "CTG Assay IC50 Analyser — Data Model Template",
        "",
        "PLATE LAYOUT sheet:",
        "  Column      = plate column number (1-12, do not change)",
        "  Sample      = cell line or genotype name",
        "  Drug        = compound / drug name",
        "  Replicate   = replicate number within that Drug+Sample group (1, 2, 3...)",
        "  Unit        = concentration unit (e.g. uM, nM, UN)",
        "  Blank       = TRUE if media-only / background column; FALSE otherwise",
        "",
        "CONCENTRATIONS sheet:",
        "  Drug            = must match Drug names in Plate_Layout sheet exactly",
        "  Starting_Conc   = concentration in row A (highest concentration)",
        "  Dilution_Factor = serial dilution factor (e.g. 3 for 1:3 dilutions)",
        "",
        "HOW TO USE:",
        "  1. Fill in Plate_Layout and Concentrations sheets (replace grey example values)",
        "  2. Save the file",
        "  3. On the Upload tab: load your CTG .xlsx plate file first",
        "  4. Then upload this Data Model file and click 'Apply Data Model'",
        "  5. The Layout and Concentrations tabs will be pre-populated automatically"
      ),
      stringsAsFactors = FALSE
    )
    writeData(wb, "Instructions", instructions, colNames = FALSE)
    title_style <- createStyle(fontColour = "#2c3e50", textDecoration = "bold", fontSize = 12)
    addStyle(wb, "Instructions", title_style, rows = 1, cols = 1)
    setColWidths(wb, "Instructions", cols = 1, widths = 70)

    return(wb)
  }

  # ── Download blank template ────────────────────────────────────────────────
  output$dl_datamodel_template <- downloadHandler(
    filename = function() paste0("CTG_DataModel_Template_", Sys.Date(), ".xlsx"),
    content  = function(file) {
      wb <- build_datamodel_wb(use_current = FALSE)
      add_about_sheet(wb,
                      ctg_file       = rv$ctg_filename,
                      datamodel_file = rv$datamodel_filename)
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )

  # ── Download current settings as data model ────────────────────────────────
  output$dl_datamodel_current <- downloadHandler(
    filename = function() paste0("CTG_DataModel_", Sys.Date(), ".xlsx"),
    content  = function(file) {
      wb <- build_datamodel_wb(use_current = TRUE)
      add_about_sheet(wb,
                      ctg_file       = rv$ctg_filename,
                      datamodel_file = rv$datamodel_filename)
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )

  # ── Apply uploaded data model ──────────────────────────────────────────────
  observeEvent(input$load_datamodel_btn, {
    req(input$datamodel_file)

    tryCatch({
      tmp <- tempfile(fileext = ".xlsx")
      file.copy(input$datamodel_file$datapath, tmp)
      on.exit(unlink(tmp))

      sheets <- excel_sheets(tmp)

      # ── Read Plate_Layout sheet ──────────────────────────────────────────
      if (!"Plate_Layout" %in% sheets)
        stop("Sheet 'Plate_Layout' not found in the uploaded file.")

      layout_raw <- as.data.frame(
        read_excel(tmp, sheet = "Plate_Layout", col_types = "text"),
        stringsAsFactors = FALSE
      )

      required_cols <- c("Column","Sample","Drug","Replicate","Unit","Blank")
      missing <- setdiff(required_cols, names(layout_raw))
      if (length(missing) > 0)
        stop(paste("Plate_Layout sheet missing columns:", paste(missing, collapse = ", ")))

      layout_df <- data.frame(
        Column    = as.integer(layout_raw$Column),
        Sample    = trimws(layout_raw$Sample),
        Drug      = trimws(layout_raw$Drug),
        Replicate = as.integer(layout_raw$Replicate),
        Unit      = trimws(layout_raw$Unit),
        Blank     = toupper(trimws(layout_raw$Blank)) %in% c("TRUE","1","YES"),
        stringsAsFactors = FALSE
      )

      # ── Read Concentrations sheet ────────────────────────────────────────
      if (!"Concentrations" %in% sheets)
        stop("Sheet 'Concentrations' not found in the uploaded file.")

      conc_raw <- as.data.frame(
        read_excel(tmp, sheet = "Concentrations", col_types = "text"),
        stringsAsFactors = FALSE
      )

      required_conc_cols <- c("Drug","Starting_Conc","Dilution_Factor")
      missing_conc <- setdiff(required_conc_cols, names(conc_raw))
      if (length(missing_conc) > 0)
        stop(paste("Concentrations sheet missing columns:", paste(missing_conc, collapse = ", ")))

      conc_df <- data.frame(
        Drug            = trimws(conc_raw$Drug),
        Starting_Conc   = as.numeric(conc_raw$Starting_Conc),
        Dilution_Factor = as.numeric(conc_raw$Dilution_Factor),
        stringsAsFactors = FALSE
      )

      # ── Apply to rv ──────────────────────────────────────────────────────
      rv$col_annot <- layout_df

      # Set dilution factor from first row (assumed global)
      dil_factor <- conc_df$Dilution_Factor[1]
      if (!is.na(dil_factor)) {
        updateNumericInput(session, "dilution_factor", value = dil_factor)
      }

      # Build drug_start_concs
      rv$drug_start_concs <- data.frame(
        Drug          = conc_df$Drug,
        Starting_Conc = conc_df$Starting_Conc,
        stringsAsFactors = FALSE
      )

      # Build 7-step concentration series per drug using ctrl_row-aware row ordering
      factor_val  <- if (!is.na(dil_factor)) dil_factor else isolate(input$dilution_factor)
      cr          <- isolate(input$ctrl_row)
      d_rows      <- LETTERS[drug_rows_ordered(cr)]   # 7 letters, lowest→highest dose
      conc_rows <- lapply(seq_len(nrow(conc_df)), function(i) {
        d <- conc_df$Drug[i]
        s <- conc_df$Starting_Conc[i]
        if (is.na(s)) s <- 10
        data.frame(
          Drug          = d,
          Row           = d_rows,
          Step          = 1:7,
          Concentration = calc_concs(s, factor_val),
          stringsAsFactors = FALSE
        )
      })
      rv$drug_concs <- do.call(rbind, conc_rows)

      # Store data model filename for provenance tracking in downloads
      rv$datamodel_filename <- input$datamodel_file$name

      showNotification(
        paste0("Data Model applied: ", nrow(layout_df), " columns and ",
               nrow(conc_df), " drugs loaded. Check Layout and Concentrations tabs."),
        type = "message", duration = 8
      )

    }, error = function(e) {
      showNotification(paste("Error reading Data Model:", e$message), type = "error")
    })
  })

  # ── Helper: append a provenance 'About' sheet to any openxlsx workbook ────
  # wb            : an openxlsx workbook object
  # ctg_file      : name of the uploaded CTG plate xlsx (or NULL)
  # datamodel_file: name of the uploaded Data Model xlsx (or NULL)
  add_about_sheet <- function(wb, ctg_file = NULL, datamodel_file = NULL) {
    addWorksheet(wb, "About")

    about_df <- data.frame(
      Field = c(
        "Application",
        "Created by",
        "Export date/time",
        "CTG plate file",
        "Data Model file"
      ),
      Value = c(
        "CTG Assay IC50 Analyser",
        "Ashley Ng",
        format(Sys.time(), "%Y-%m-%d %H:%M:%S"),
        if (!is.null(ctg_file) && nzchar(ctg_file)) ctg_file else "(not recorded)",
        if (!is.null(datamodel_file) && nzchar(datamodel_file)) datamodel_file else "(not used)"
      ),
      stringsAsFactors = FALSE
    )

    writeData(wb, "About", about_df, startRow = 1)

    # Style the Field column bold, Value column normal
    field_style <- createStyle(textDecoration = "bold", fgFill = "#f0f4f8",
                               border = "TopBottomLeftRight", borderColour = "#cccccc")
    value_style <- createStyle(border = "TopBottomLeftRight", borderColour = "#cccccc")
    addStyle(wb, "About", field_style,
             rows = 2:(nrow(about_df) + 1), cols = 1, gridExpand = TRUE)
    addStyle(wb, "About", value_style,
             rows = 2:(nrow(about_df) + 1), cols = 2, gridExpand = TRUE)

    # Header row
    hdr_style <- createStyle(fontColour = "#FFFFFF", fgFill = "#2c3e50",
                             textDecoration = "bold")
    addStyle(wb, "About", hdr_style, rows = 1, cols = 1:2, gridExpand = TRUE)
    setColWidths(wb, "About", cols = 1:2, widths = c(22, 55))
  }

  # ── Downloads ──────────────────────────────────────────────────────────────
  output$dl_plot <- downloadHandler(
    filename = function()
      paste0("DR_", input$plot_drug, "_", input$plot_sample, ".pdf"),
    content = function(file)
      ggsave(file,
             plot   = make_dr_plot(input$plot_drug, input$plot_sample,
                                   ctg_file = rv$ctg_filename),
             device = "pdf", width = 9, height = 6)
  )

  output$dl_results_xlsx <- downloadHandler(
    filename = function() paste0("CTG_IC_Results_", Sys.Date(), ".xlsx"),
    content  = function(file) {
      req(rv$ic_summary)

      wb <- createWorkbook()

      # Shared header style
      hdr_style <- createStyle(
        fontColour     = "#FFFFFF",
        fgFill         = "#2c3e50",
        textDecoration = "bold",
        halign         = "center",
        border         = "Bottom",
        borderColour   = "#888888"
      )

      # ── Sheet 1: IC_Values (original linear units) ──────────────────────────
      addWorksheet(wb, "IC_Values")
      writeData(wb, "IC_Values", rv$ic_summary)
      addStyle(wb, "IC_Values", hdr_style,
               rows = 1, cols = seq_len(ncol(rv$ic_summary)), gridExpand = TRUE)
      setColWidths(wb, "IC_Values", cols = seq_len(ncol(rv$ic_summary)), widths = "auto")

      # Add a note row below the data explaining units
      note_row <- nrow(rv$ic_summary) + 3
      note_text <- paste0(
        "IC values are in the original linear concentration units entered ",
        "in the Concentrations tab (NOT log-transformed). ",
        "95% delta CI shown for IC50."
      )
      writeData(wb, "IC_Values",
                data.frame(Note = note_text),
                startRow = note_row, colNames = FALSE)
      note_style <- createStyle(fontColour = "#777777", textDecoration = "italic",
                                wrapText = TRUE)
      addStyle(wb, "IC_Values", note_style,
               rows = note_row, cols = 1, gridExpand = TRUE)

      # ── Sheet 2: IC_Unit_Converted (original + nM when µM) ──────────────────
      if (!is.null(rv$ic_converted) && nrow(rv$ic_converted) > 0) {
        addWorksheet(wb, "IC_Unit_Converted")
        writeData(wb, "IC_Unit_Converted", rv$ic_converted)
        addStyle(wb, "IC_Unit_Converted", hdr_style,
                 rows = 1, cols = seq_len(ncol(rv$ic_converted)), gridExpand = TRUE)
        setColWidths(wb, "IC_Unit_Converted",
                     cols = seq_len(ncol(rv$ic_converted)), widths = "auto")

        # Note row
        note_row2 <- nrow(rv$ic_converted) + 3
        note_text2 <- paste0(
          "IC columns are suffixed with their unit (e.g. IC50_\u00b5M). ",
          "For drugs entered in \u00b5M, additional _nM columns multiply by 1000. ",
          "Values are on the original linear concentration scale (NOT log-transformed)."
        )
        writeData(wb, "IC_Unit_Converted",
                  data.frame(Note = note_text2),
                  startRow = note_row2, colNames = FALSE)
        addStyle(wb, "IC_Unit_Converted", note_style,
                 rows = note_row2, cols = 1, gridExpand = TRUE)
      }

      # ── Sheet 3: Normalised_Data ─────────────────────────────────────────────
      if (!is.null(rv$norm_data)) {
        addWorksheet(wb, "Normalised_Data")
        writeData(wb, "Normalised_Data", rv$norm_data)
        addStyle(wb, "Normalised_Data", hdr_style,
                 rows = 1, cols = seq_len(ncol(rv$norm_data)), gridExpand = TRUE)
        setColWidths(wb, "Normalised_Data",
                     cols = seq_len(ncol(rv$norm_data)), widths = "auto")
      }

      # ── Sheet 4: About (provenance) ──────────────────────────────────────────
      add_about_sheet(wb,
                      ctg_file       = rv$ctg_filename,
                      datamodel_file = rv$datamodel_filename)

      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )

  output$dl_norm_xlsx <- downloadHandler(
    filename = function() paste0("CTG_Normalised_", Sys.Date(), ".xlsx"),
    content  = function(file) {
      req(rv$norm_data)
      wb <- createWorkbook()
      addWorksheet(wb, "Normalised_Data")
      writeData(wb, "Normalised_Data", rv$norm_data)
      add_about_sheet(wb,
                      ctg_file       = rv$ctg_filename,
                      datamodel_file = rv$datamodel_filename)
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )

  output$dl_all_plots <- downloadHandler(
    filename = function() paste0("CTG_All_Curves_", Sys.Date(), ".pdf"),
    content  = function(file) {
      req(rv$fit_results)
      pdf(file, width = 9, height = 6)
      for (key in names(rv$fit_results)) {
        parts <- strsplit(key, " \\| ")[[1]]
        p <- tryCatch(
          make_dr_plot(parts[1], parts[2], ctg_file = rv$ctg_filename),
          error = function(e) NULL
        )
        if (!is.null(p)) print(p)
      }
      dev.off()
    }
  )
}

# ---- Run ------------------------------------------------------------------
shinyApp(ui, server)
