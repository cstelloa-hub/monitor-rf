# ============================================================
# GENERADOR DASHBOARD FONDO 0  (v2)
# Lee Monitor_F0.xlsm (macro v8) + copias diarias -> monitor_f0.html
# Requiere: openxlsx. JSON en R base (sin dependencias extra).
# v2: Historico 29 cols (SCORE LP/CP), Calidad en dos bloques,
#     sin Concentracion, template v2.
# ============================================================

# ----------------- CONFIG (ajustar rutas) -------------------
RUTA_XLSM      <- "D:/Profuturo/Claude/monitor_f0/Monitor_F0.xlsm"
CARPETA_COPIAS <- "D:/Profuturo/Claude/monitor_f0/copias"
RUTA_TEMPLATE  <- "D:/Profuturo/Claude/monitor_f0/dashboard/template_f0.html"
RUTA_SALIDA    <- "D:/Profuturo/Claude/monitor_f0/dashboard/monitor_f0.html"
MAX_DIAS_DETALLE <- 250     # cuantas copias diarias incluir como maximo
# ------------------------------------------------------------

suppressMessages(library(openxlsx))

# ----------------- utilitarios JSON (R base) ----------------
esc_json <- function(s) {
  s <- gsub("\\\\", "\\\\\\\\", s)
  s <- gsub('"', '\\\\"', s)
  s <- gsub("\n", "\\\\n", s)
  s <- gsub("\r", "", s)
  s <- gsub("\t", "\\\\t", s)
  s
}
val_json <- function(v) {
  if (length(v) == 0 || is.null(v)) return("null")
  if (is.na(v)) return("null")
  if (is.numeric(v)) {
    if (!is.finite(v)) return("null")
    return(format(v, scientific = FALSE, trim = TRUE, digits = 12))
  }
  paste0('"', esc_json(as.character(v)), '"')
}
fila_json <- function(fila) paste0("[", paste(vapply(fila, val_json, ""), collapse = ","), "]")
matriz_json <- function(m) {
  if (is.null(m) || nrow(m) == 0) return("[]")
  paste0("[", paste(vapply(seq_len(nrow(m)), function(i) fila_json(as.list(m[i, ])), ""), collapse = ","), "]")
}
lista_filas_json <- function(lf) {
  if (length(lf) == 0) return("[]")
  paste0("[", paste(vapply(lf, fila_json, ""), collapse = ","), "]")
}

num0 <- function(x) {
  v <- suppressWarnings(as.numeric(x))
  ifelse(is.na(v), NA_real_, v)
}
fecha_iso <- function(x) {
  if (inherits(x, "Date")) return(format(x, "%Y-%m-%d"))
  v <- suppressWarnings(as.numeric(x))
  if (!is.na(v) && v > 20000) return(format(as.Date(v, origin = "1899-12-30"), "%Y-%m-%d"))
  d <- suppressWarnings(as.Date(as.character(x), tryFormats = c("%d/%m/%Y", "%Y-%m-%d")))
  if (!is.na(d)) return(format(d, "%Y-%m-%d"))
  NA_character_
}
fecha_ddmm <- function(x) {
  iso <- fecha_iso(x)
  if (is.na(iso)) return(NA_character_)
  format(as.Date(iso), "%d/%m/%Y")
}

# ----------------- 1) HISTORICO -----------------------------
cat("Leyendo Historico...\n")
hist_raw <- read.xlsx(RUTA_XLSM, sheet = "Historico", colNames = FALSE,
                      startRow = 2, detectDates = TRUE, skipEmptyRows = TRUE)
if (is.null(hist_raw) || nrow(hist_raw) == 0) stop("Historico vacio: corre el backfill (macro v8).")
while (ncol(hist_raw) < 29) hist_raw[[ncol(hist_raw) + 1]] <- NA

hist_filas <- list()
for (i in seq_len(nrow(hist_raw))) {
  f <- fecha_iso(hist_raw[i, 1])
  if (is.na(f)) next
  fila <- list(f)
  for (j in 2:29) {
    v <- hist_raw[i, j]
    if (j == 17) fila[[j]] <- as.character(v)           # RATING LP (letra)
    else fila[[j]] <- num0(v)
  }
  hist_filas[[length(hist_filas) + 1]] <- fila
}
ord <- order(vapply(hist_filas, function(x) x[[1]], ""))
hist_filas <- hist_filas[ord]
cat("  ", length(hist_filas), "observaciones\n")

# ----------------- 2) COPIAS DIARIAS ------------------------
leer_hoja <- function(archivo, hoja, startRow) {
  out <- tryCatch(
    read.xlsx(archivo, sheet = hoja, colNames = FALSE, startRow = startRow,
              detectDates = FALSE, skipEmptyRows = TRUE),
    error = function(e) NULL)
  out
}

# Calidad v8: dos bloques con marcadores "LARGO PLAZO" y "CORTO PLAZO",
# cada uno con cabecera RATING/MONTO/PESO/PUNTAJE y filas debajo.
leer_calidad <- function(archivo) {
  raw <- leer_hoja(archivo, "Calidad", 1)
  vacio <- list(LP = list(), CP = list())
  if (is.null(raw) || nrow(raw) == 0) return(vacio)
  res <- vacio
  bloque <- ""
  for (i in seq_len(nrow(raw))) {
    c1 <- toupper(trimws(as.character(raw[i, 1])))
    if (is.na(c1) || c1 == "" || c1 == "NA") next
    if (grepl("^LARGO", c1))  { bloque <- "LP"; next }
    if (grepl("^CORTO", c1))  { bloque <- "CP"; next }
    if (c1 == "RATING" || grepl("^CALIDAD", c1)) next
    if (bloque == "") next
    mto <- num0(raw[i, 2])
    if (is.na(mto)) next
    pj_raw <- raw[i, 4]
    pj_num <- num0(pj_raw)
    pj <- if (!is.na(pj_num)) pj_num else as.character(pj_raw)   # "SIN MAPEO" pasa como texto
    fila <- list(as.character(raw[i, 1]), mto, num0(raw[i, 3]), pj)
    res[[bloque]][[length(res[[bloque]]) + 1]] <- fila
  }
  res
}

archivos <- list.files(CARPETA_COPIAS, pattern = "^F0_Monitor_\\d{8}\\.xlsx$", full.names = TRUE)
archivos <- sort(archivos, decreasing = TRUE)
if (length(archivos) > MAX_DIAS_DETALLE) archivos <- archivos[1:MAX_DIAS_DETALLE]

cat("Leyendo copias diarias...\n")
det_json_partes <- character(0)
for (arch in archivos) {
  fch <- sub("^F0_Monitor_(\\d{8})\\.xlsx$", "\\1", basename(arch))
  iso <- paste0(substr(fch, 1, 4), "-", substr(fch, 5, 6), "-", substr(fch, 7, 8))
  cat("  ", iso, "\n")

  # --- Cartera F0 (fila 2+): 19 cols v8; usamos 18 (DTS se calcula en JS) ---
  cart_raw <- leer_hoja(arch, "Cartera F0", 2)
  cart <- list()
  if (!is.null(cart_raw) && nrow(cart_raw) > 0) {
    for (i in seq_len(nrow(cart_raw))) {
      if (is.na(cart_raw[i, 3]) || cart_raw[i, 3] == "") next
      cart[[length(cart) + 1]] <- list(
        as.character(cart_raw[i, 3]),   # 0 codigo
        as.character(cart_raw[i, 4]),   # 1 emisor
        as.character(cart_raw[i, 6]),   # 2 categoria
        num0(cart_raw[i, 9]),           # 3 monto MM
        as.character(cart_raw[i, 10]),  # 4 rating
        fecha_ddmm(cart_raw[i, 11]),    # 5 vcto
        num0(cart_raw[i, 12]),          # 6 dias
        num0(cart_raw[i, 13]),          # 7 ytw
        num0(cart_raw[i, 14]),          # 8 spread
        num0(cart_raw[i, 15]),          # 9 duracion
        as.character(cart_raw[i, 16]),  # 10 en vector
        as.character(cart_raw[i, 17]),  # 11 deposito
        num0(cart_raw[i, 18])           # 12 peso
      )
    }
  }

  # --- Contribucion (fila 4+): 9 cols ---
  ctr_raw <- leer_hoja(arch, "Contribucion", 4)
  ctr <- list()
  if (!is.null(ctr_raw) && nrow(ctr_raw) > 0) {
    for (i in seq_len(nrow(ctr_raw))) {
      if (is.na(ctr_raw[i, 2]) || ctr_raw[i, 2] == "") next
      ctr[[length(ctr) + 1]] <- list(
        as.character(ctr_raw[i, 1]), as.character(ctr_raw[i, 2]),
        as.character(ctr_raw[i, 3]), num0(ctr_raw[i, 4]),
        num0(ctr_raw[i, 5]), num0(ctr_raw[i, 6]),
        num0(ctr_raw[i, 7]), num0(ctr_raw[i, 8]),
        as.character(ctr_raw[i, 9]))
    }
  }

  # --- Calidad v8: dos bloques ---
  cal <- leer_calidad(arch)

  # --- Vencimientos (fila 4+): anio, bonos, dep, papeles, total ---
  ven_raw <- leer_hoja(arch, "Vencimientos", 4)
  ven <- list()
  if (!is.null(ven_raw) && nrow(ven_raw) > 0) {
    for (i in seq_len(nrow(ven_raw))) {
      a <- suppressWarnings(as.integer(ven_raw[i, 1]))
      if (is.na(a)) next
      ven[[length(ven) + 1]] <- list(a, num0(ven_raw[i, 2]), num0(ven_raw[i, 3]),
                                     num0(ven_raw[i, 4]), num0(ven_raw[i, 5]))
    }
  }

  det_json_partes <- c(det_json_partes, paste0(
    '"', iso, '":{',
    '"cartera":', lista_filas_json(cart), ',',
    '"contrib":', lista_filas_json(ctr), ',',
    '"calLP":',  lista_filas_json(cal$LP), ',',
    '"calCP":',  lista_filas_json(cal$CP), ',',
    '"venc":',   lista_filas_json(ven), '}'))
}

# ----------------- 3) JSON + inyeccion ----------------------
json <- paste0(
  'const DATA = {',
  '"generado":"', format(Sys.time(), "%d/%m/%Y %H:%M"), '",',
  '"hist":', lista_filas_json(hist_filas), ',',
  '"det":{', paste(det_json_partes, collapse = ","), '}',
  '};')

cat("Inyectando template...\n")
tpl <- readChar(RUTA_TEMPLATE, file.size(RUTA_TEMPLATE), useBytes = TRUE)
Encoding(tpl) <- "UTF-8"
if (!grepl("/*__DATA__*/", tpl, fixed = TRUE)) stop("El template no tiene la marca /*__DATA__*/")
salida <- sub("/*__DATA__*/", json, tpl, fixed = TRUE)

con <- file(RUTA_SALIDA, open = "w", encoding = "UTF-8")
writeLines(salida, con, useBytes = FALSE)
close(con)
cat("Listo:", RUTA_SALIDA, "\n")
