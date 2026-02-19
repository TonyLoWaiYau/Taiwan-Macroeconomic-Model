# model_excel_workflow.R
library(readxl)
library(BVAR)
library(openxlsx)

# ---------------- helpers (same as yours) ----------------
lag4 <- function(x) {
  n <- length(x)
  if (n <= 4) return(rep(NA, n))
  c(rep(NA, 4), x[1:(n - 4)])
}
yoy_pct <- function(x) (x / lag4(x) - 1) * 100

get_forecast_matrix <- function(pred, ref_cols) {
  q <- pred[["quants"]]
  if (is.null(q)) stop("predict() output does not contain 'quants'.")
  
  if (is.array(q) && length(dim(q)) == 3) {
    qnames <- if (!is.null(dimnames(q))) dimnames(q)[[3]] else NULL
    idx <- 2
    if (!is.null(qnames)) {
      if ("0.5" %in% qnames) idx <- which(qnames == "0.5")[1]
      else if ("50%" %in% qnames) idx <- which(qnames == "50%")[1]
      else if ("median" %in% qnames) idx <- which(qnames == "median")[1]
      else idx <- ceiling(length(qnames) / 2)
    }
    q <- q[, , idx]
  }
  q <- as.matrix(q)
  colnames(q) <- ref_cols
  q
}

to_date <- function(x) {
  if (inherits(x, "Date")) return(x)
  if (inherits(x, "POSIXt")) return(as.Date(x))
  if (is.numeric(x)) return(as.Date(x, origin = "1899-12-30"))
  x <- as.character(x)
  d <- suppressWarnings(as.Date(x))
  if (all(is.na(d))) d <- suppressWarnings(as.Date(paste0("01/", x), format = "%d/%m/%Y"))
  d
}

create_styles <- function() {
  headerStyle <- createStyle(fontName="Times New Roman", fontSize=12,
                             textDecoration="bold", halign="center", border="Bottom")
  bodyStyle   <- createStyle(fontName="Times New Roman", fontSize=12, numFmt="0.0")
  predStyle   <- createStyle(fontName="Times New Roman", fontSize=12,
                             fontColour="#FF0000", numFmt="0.0")
  list(headerStyle=headerStyle, bodyStyle=bodyStyle, predStyle=predStyle)
}

write_result_raw_xlsx <- function(file, time_vec, result_raw, n_hist, horizon, styles) {
  n_obs <- nrow(result_raw)
  rounded_df <- result_raw
  num_cols <- vapply(rounded_df, is.numeric, logical(1))
  rounded_df[, num_cols] <- round(rounded_df[, num_cols], 1)
  out_df <- cbind(Time = time_vec, rounded_df)
  
  wb <- createWorkbook()
  addWorksheet(wb, "result_raw")
  writeData(wb, "result_raw", out_df, headerStyle = styles$headerStyle, keepNA = TRUE)
  addStyle(wb, "result_raw", styles$bodyStyle,
           rows = 2:(n_obs + 1), cols = 1:ncol(out_df), gridExpand = TRUE, stack = TRUE)
  
  first_fc_row <- n_hist + 1
  excel_fc_rows <- (first_fc_row:n_obs) + 1
  addStyle(wb, "result_raw", styles$predStyle,
           rows = excel_fc_rows, cols = 2:ncol(out_df), gridExpand = TRUE, stack = TRUE)
  setColWidths(wb, "result_raw", cols = 1:ncol(out_df), widths = "auto")
  freezePane(wb, "result_raw", firstRow = TRUE)
  
  addWorksheet(wb, "forecast_condition")
  writeData(wb, "forecast_condition", out_df, headerStyle = styles$headerStyle, keepNA = TRUE)
  addStyle(wb, "forecast_condition", styles$bodyStyle,
           rows = 2:(n_obs + 1), cols = 1:ncol(out_df), gridExpand = TRUE, stack = TRUE)
  addStyle(wb, "forecast_condition", styles$predStyle,
           rows = excel_fc_rows, cols = 2:ncol(out_df), gridExpand = TRUE, stack = TRUE)
  setColWidths(wb, "forecast_condition", cols = 1:ncol(out_df), widths = "auto")
  freezePane(wb, "forecast_condition", firstRow = TRUE)
  
  saveWorkbook(wb, file = file, overwrite = TRUE)
}

open_file <- function(path) {
  if (!file.exists(path)) stop("File not found: ", path)
  path <- normalizePath(path, winslash="\\", mustWork = TRUE)
  shell.exec(path)
}

# ---------------- PART 1: model + unconditional + create Excel conditioning file ----------------
run_part1 <- function(data_xlsx, sheet = "clean", out_dir = ".", seed = 123,
                      lags = 9, horizon = 9,
                      n_draw = 250000L, n_burn = 50000L, n_thin = 5L) {
  set.seed(seed)
  styles <- create_styles()
  
  data <- read_excel(data_xlsx, sheet = sheet)
  
  vars <- names(data)
  vars_to_yoy <- vars[vars != "t" & !grepl("^g_", vars) & vars != "unemploy"]
  yoy_list <- setNames(lapply(vars_to_yoy, function(v) yoy_pct(data[[v]])),
                       paste0("g_", vars_to_yoy))
  
  data_actual <- data.frame(
    t = data$t,
    as.data.frame(yoy_list),
    d_unemploy = data$unemploy - lag4(data$unemploy)
  )
  
  data_ae <- data_actual
  last_idx <- nrow(data_ae)
  gdp_components <- c("g_gdp", "g_pce", "g_gce", "g_gcf", "g_exports", "g_imports")
  gdp_components <- gdp_components[gdp_components %in% names(data_ae) & gdp_components %in% names(data)]
  data_ae[last_idx, gdp_components] <- data[last_idx, gdp_components]
  
  last_row_has_na <- any(is.na(data[nrow(data), ]))
  data_train <- if (!last_row_has_na) data_actual else data_ae
  if (nrow(data_train) > 4) data_train <- data_train[-(1:4), ]
  
  desired_order <- c("g_exports","g_gce","g_ip","g_pce","g_gcf",
                     "g_imports","d_unemploy","g_cpi","g_fx")
  desired_order <- desired_order[desired_order %in% names(data_train)]
  data_bvar <- data_train[, desired_order, drop = FALSE]
  
  soc <- bv_soc(mode = 1, sd = 1, min = 1e-04, max = 50)
  sur <- bv_sur(mode = 1, sd = 1, min = 1e-04, max = 50)
  priors <- bv_priors(hyper = "auto", mn = bv_mn(), soc = soc, sur = sur)
  irf <- bv_irf(horizon = horizon, fevd = FALSE)
  
  model <- bvar(data_bvar, lags = lags,
                n_draw = n_draw, n_burn = n_burn, n_thin = n_thin,
                priors = priors, verbose = TRUE, irf = irf)
  
  forecast_raw <- predict(model, horizon = horizon, conf_bands = 0.5)
  forecast_mat <- get_forecast_matrix(forecast_raw, colnames(data_bvar))
  
  # time vector
  t_train <- data_actual$t
  if (length(t_train) > 4) t_train <- t_train[-(1:4)]
  t_date <- to_date(t_train)
  time_vec_actual <- if (any(is.na(t_date))) as.character(t_train) else format(t_date, "%m/%Y")
  
  if (!any(is.na(t_date))) {
    future_dates <- seq.Date(from=t_date[length(t_date)], by="3 months", length.out=horizon+1)[-1]
    time_vec <- c(time_vec_actual, format(future_dates, "%m/%Y"))
  } else {
    time_vec <- c(time_vec_actual, paste0("F", 1:horizon))
  }
  
  result_raw <- data_bvar
  result_raw[(nrow(data_bvar) + 1):(nrow(data_bvar) + horizon), ] <- forecast_mat
  
  out_raw <- file.path(out_dir, "result_raw.xlsx")
  write_result_raw_xlsx(out_raw, time_vec, result_raw, nrow(data_bvar), horizon, styles)
  
  list(
    model = model,
    data = data,
    data_bvar = data_bvar,
    horizon = horizon,
    time_vec = time_vec,
    out_raw = out_raw,
    styles = styles
  )
}

# ---------------- PART 2: read edited forecast_condition + conditional + export exactly like you do ----------------
run_part2 <- function(part1_obj,
                      edited_result_raw_xlsx = part1_obj$out_raw,
                      out_dir = dirname(part1_obj$out_raw)) {
  
  model <- part1_obj$model
  data  <- part1_obj$data
  data_bvar <- part1_obj$data_bvar
  horizon <- part1_obj$horizon
  time_vec <- part1_obj$time_vec
  styles <- part1_obj$styles
  
  # ---- safety checks ----
  if (!file.exists(edited_result_raw_xlsx)) {
    stop("Cannot find result_raw.xlsx at: ", edited_result_raw_xlsx)
  }
  
  # read edited paths
  paths_df <- readxl::read_excel(edited_result_raw_xlsx, sheet = "forecast_condition")
  
  needed_cols <- c("Time", colnames(data_bvar))
  missing_cols <- setdiff(needed_cols, names(paths_df))
  if (length(missing_cols) > 0) {
    stop("The 'forecast_condition' sheet is missing required columns: ",
         paste(missing_cols, collapse = ", "),
         ". Do not rename columns.")
  }
  
  n_hist <- nrow(data_bvar)
  if (nrow(paths_df) < (n_hist + horizon)) {
    stop("'forecast_condition' does not have enough rows. Expected at least ",
         (n_hist + horizon), " rows, but found ", nrow(paths_df),
         ". Do not delete rows.")
  }
  
  future_block <- paths_df[(n_hist + 1):(n_hist + horizon), needed_cols, drop = FALSE]
  future_time <- future_block$Time
  future_num  <- future_block[, colnames(data_bvar), drop = FALSE]
  
  # Convert to numeric matrix (Excel edits sometimes become text)
  paths_mat <- as.matrix(data.frame(lapply(future_num, function(x) suppressWarnings(as.numeric(x)))))
  
  # ---- conditional forecast (catch and rethrow with guidance) ----
  conditional_forecast <- tryCatch(
    predict(model,
            horizon = horizon,
            cond_path = paths_mat,
            cond_var = 1:ncol(data_bvar),
            conf_bands = 0.5),
    error = function(e) {
      stop(
        "Conditional forecast failed. This is usually caused by impossible/unstable conditioning values ",
        "or a mismatch in the conditioning sheet.\n\nOriginal error: ",
        conditionMessage(e)
      )
    }
  )
  
  cond_mat <- get_forecast_matrix(conditional_forecast, colnames(data_bvar))
  
  result_conditional <- data_bvar
  result_conditional[(nrow(data_bvar) + 1):(nrow(data_bvar) + horizon), ] <- cond_mat
  
  # ---- write back "conditional forecast" sheet into result_raw.xlsx (exact formatting) ----
  n_obs <- nrow(result_conditional)
  rounded_cond <- result_conditional
  num_cols <- vapply(rounded_cond, is.numeric, logical(1))
  rounded_cond[, num_cols] <- round(rounded_cond[, num_cols], 1)
  out_cond <- cbind(Time = time_vec, rounded_cond)
  
  wb <- loadWorkbook(edited_result_raw_xlsx)
  if ("conditional forecast" %in% names(wb)) removeWorksheet(wb, "conditional forecast")
  addWorksheet(wb, "conditional forecast")
  writeData(wb, "conditional forecast", out_cond, headerStyle = styles$headerStyle, keepNA = TRUE)
  addStyle(wb, "conditional forecast", styles$bodyStyle,
           rows = 2:(n_obs + 1), cols = 1:ncol(out_cond), gridExpand = TRUE, stack = TRUE)
  
  first_fc_row <- nrow(data_bvar) + 1
  excel_fc_rows <- (first_fc_row:n_obs) + 1
  addStyle(wb, "conditional forecast", styles$predStyle,
           rows = excel_fc_rows, cols = 2:ncol(out_cond), gridExpand = TRUE, stack = TRUE)
  
  setColWidths(wb, "conditional forecast", cols = 1:ncol(out_cond), widths = "auto")
  freezePane(wb, "conditional forecast", firstRow = TRUE)
  
  saveWorkbook(wb, edited_result_raw_xlsx, overwrite = TRUE)
  
  # ---- Now reproduce your result_summary.xlsx logic exactly (copied with minimal edits) ----
  # Create level dataframe excluding variables starting with "g_"
  result_level <- data[, !grepl("^g_", names(data)), drop = FALSE]
  
  horizon_level <- nrow(result_conditional) - nrow(data_bvar)
  if (horizon_level > 0) {
    t_raw <- result_level$t
    t_date <- to_date(t_raw)
    if (!any(is.na(t_date))) {
      last_date <- t_date[length(t_date)]
      future_dates <- seq.Date(from = last_date, by = "3 months", length.out = horizon_level + 1)[-1]
      if (inherits(t_raw, "Date")) t_future <- future_dates
      else if (inherits(t_raw, "POSIXt")) t_future <- as.POSIXct(future_dates)
      else if (is.numeric(t_raw)) t_future <- as.numeric(future_dates - as.Date("1899-12-30"))
      else t_future <- format(future_dates, "%m/%Y")
    } else {
      t_future <- paste0("F", 1:horizon_level)
    }
    
    na_block <- as.data.frame(matrix(NA, nrow = horizon_level, ncol = ncol(result_level)))
    names(na_block) <- names(result_level)
    na_block[["t"]] <- t_future
    result_level <- rbind(result_level, na_block)
  }
  
  start_fill <- max(5, nrow(result_level) - (horizon_level + 9))
  end_fill <- nrow(result_level)
  level_vars <- setdiff(names(result_level), "t")
  
  for (v in level_vars) {
    gcol <- if (v == "unemploy") "d_unemploy" else paste0("g_", v)
    if (!gcol %in% names(result_conditional)) next
    
    for (r in start_fill:end_fill) {
      if (r < 5) next
      g_idx <- r - 4
      if (g_idx < 1 || g_idx > nrow(result_conditional)) next
      if (is.na(result_level[r, v])) {
        g_val <- result_conditional[g_idx, gcol]
        lag_val <- result_level[r - 4, v]
        if (!is.na(g_val) && !is.na(lag_val)) {
          if (v == "unemploy") result_level[r, v] <- lag_val + g_val
          else result_level[r, v] <- lag_val * (1 + g_val / 100)
        }
      }
    }
  }
  
  if ("gdp" %in% names(result_level)) {
    start_gdp <- max(1, nrow(result_level) - 9)
    end_gdp <- nrow(result_level)
    for (r in start_gdp:end_gdp) {
      if (is.na(result_level[r, "gdp"])) {
        comps <- c("pce", "gce", "gcf", "exports", "imports")
        if (all(comps %in% names(result_level))) {
          vals <- result_level[r, comps]
          if (all(!is.na(vals))) {
            result_level[r, "gdp"] <- vals[["pce"]] + vals[["gce"]] + vals[["gcf"]] +
              vals[["exports"]] - vals[["imports"]]
          }
        }
      }
    }
  }
  
  result_growth <- result_level
  growth_vars <- setdiff(names(result_level), c("t", "unemploy", "fx"))
  for (v in growth_vars) result_growth[[v]] <- yoy_pct(result_level[[v]])
  
  annual_sum_safe <- function(x) { x <- suppressWarnings(as.numeric(x)); if (all(is.na(x))) NA_real_ else sum(x, na.rm=TRUE) }
  annual_mean_safe <- function(x){ x <- suppressWarnings(as.numeric(x)); if (all(is.na(x))) NA_real_ else mean(x, na.rm=TRUE) }
  
  t_raw_level <- result_level$t
  t_date_level <- to_date(t_raw_level)
  if (any(is.na(t_date_level))) year_vec <- 1981 + ((0:(nrow(result_level) - 1)) %/% 4)
  else year_vec <- as.integer(format(t_date_level, "%Y"))
  annual_years <- unique(year_vec)
  
  gdp_vars <- intersect(c("gdp","pce","gce","gcf","exports","imports"), names(result_level))
  avg_vars <- intersect(c("ip","cpi","unemploy","fx"), names(result_level))
  other_vars <- setdiff(setdiff(names(result_level), c("t", gdp_vars)), avg_vars)
  
  annual_list <- lapply(annual_years, function(y) {
    rows <- year_vec == y
    out <- list(t = y)
    for (v in gdp_vars) out[[v]] <- annual_sum_safe(result_level[[v]][rows])
    for (v in avg_vars) out[[v]] <- annual_mean_safe(result_level[[v]][rows])
    for (v in other_vars) out[[v]] <- annual_sum_safe(result_level[[v]][rows])
    as.data.frame(out, check.names = FALSE)
  })
  result_annual_level <- do.call(rbind, annual_list)
  
  lag1 <- function(x) { n <- length(x); if (n <= 1) return(rep(NA, n)); c(NA, x[1:(n - 1)]) }
  annual_yoy_pct <- function(x) (x / lag1(x) - 1) * 100
  
  result_annual_growth <- result_annual_level
  annual_growth_vars <- setdiff(names(result_annual_level), c("t","unemploy","fx"))
  for (v in annual_growth_vars) result_annual_growth[[v]] <- annual_yoy_pct(result_annual_level[[v]])
  
  format_quarter <- function(x) {
    d <- to_date(x)
    if (all(is.na(d))) return(as.character(x))
    yr <- as.integer(format(d, "%Y"))
    mo <- as.integer(format(d, "%m"))
    q <- ((mo - 1) %/% 3) + 1
    paste0("Q", q, " ", yr)
  }
  format_year <- function(x) {
    if (is.numeric(x)) return(as.character(as.integer(round(x))))
    suppressWarnings(as.character(as.integer(x)))
  }
  
  fmt_sheet <- function(wb, sheet, df, time_col = "t", time_fmt = c("quarter","year","raw"), fc_rows = 0) {
    time_fmt <- match.arg(time_fmt)
    n_obs <- nrow(df)
    out_df <- df
    num_cols <- vapply(out_df, is.numeric, logical(1))
    out_df[, num_cols] <- round(out_df[, num_cols], 1)
    
    if (time_col %in% names(out_df)) {
      time_vec2 <- out_df[[time_col]]
      if (time_fmt == "quarter") time_vec2 <- format_quarter(time_vec2)
      if (time_fmt == "year") time_vec2 <- format_year(time_vec2)
      out_df <- cbind(Time = time_vec2, out_df[, setdiff(names(out_df), time_col), drop=FALSE])
    }
    
    addWorksheet(wb, sheetName = sheet)
    writeData(wb, sheet, out_df, headerStyle = styles$headerStyle, keepNA = TRUE)
    addStyle(wb, sheet, styles$bodyStyle,
             rows = 2:(n_obs + 1), cols = 1:ncol(out_df), gridExpand = TRUE, stack = TRUE)
    
    if (fc_rows > 0 && n_obs >= fc_rows) {
      fc_start <- n_obs - fc_rows + 1
      excel_fc_rows2 <- (fc_start:n_obs) + 1
      addStyle(wb, sheet, styles$predStyle,
               rows = excel_fc_rows2, cols = 2:ncol(out_df), gridExpand = TRUE, stack = TRUE)
    }
    
    setColWidths(wb, sheet, cols = 1:ncol(out_df), widths = "auto")
    freezePane(wb, sheet, firstRow = TRUE)
  }
  
  wb_out <- createWorkbook()
  fmt_sheet(wb_out, "result_level", result_level, "t", "quarter", fc_rows = 9)
  fmt_sheet(wb_out, "result_growth", result_growth, "t", "quarter", fc_rows = 9)
  fmt_sheet(wb_out, "result_annual_level", result_annual_level, "t", "year", fc_rows = 3)
  fmt_sheet(wb_out, "result_annual_growth", result_annual_growth, "t", "year", fc_rows = 3)
  
  out_summary <- file.path(out_dir, "result_summary.xlsx")
  saveWorkbook(wb_out, file = out_summary, overwrite = TRUE)
  
  list(out_raw = edited_result_raw_xlsx, out_summary = out_summary)
}