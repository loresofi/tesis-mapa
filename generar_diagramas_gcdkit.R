# =============================================================================
# Genera diagramas geoquimicos con GCDkit — UNINORTE Tesis
# Uso: Rscript generar_diagramas_gcdkit.R <geochem.csv> <tipos.csv> <out_dir>
# =============================================================================
suppressPackageStartupMessages(library(GCDkit))

args <- commandArgs(trailingOnly = TRUE)
if (length(args) < 3) stop("Uso: Rscript generar_diagramas_gcdkit.R geochem.csv tipos.csv out_dir")

geo_csv  <- args[1]
tipo_csv <- args[2]
out_dir  <- args[3]
dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)

# -----------------------------------------------------------------------------
# Leer y preparar datos
# -----------------------------------------------------------------------------
geo   <- read.csv(geo_csv,  stringsAsFactors = FALSE, check.names = FALSE)
tipos <- read.csv(tipo_csv, stringsAsFactors = FALSE, check.names = FALSE)
geo   <- merge(geo, tipos, by = "UNINORTE_CODE", all.x = TRUE)
geo$TIPO_ROCA[is.na(geo$TIPO_ROCA)] <- "OTRO"
colnames(geo) <- gsub("Fe2O3\\(T\\)", "Fe2O3", colnames(geo))

oxidos <- c("SiO2","TiO2","Al2O3","Fe2O3","MnO","MgO","CaO","Na2O","K2O","P2O5")
trazas <- c("Ba","Sr","Y","Sc","Zr","V")

# Cargar subconjunto en GCDkit usando formato nativo .dat (tab-delim)
# Formato GCDkit:
#   Fila 1: nombre columnas  (Samp, luego oxidos/trazas)
#   Fila 2: unidades         (end para etiquetas, % o ppm para datos)
#   Filas 3+: datos
cargar_gcdkit <- function(sub) {
  if (nrow(sub) == 0) return(FALSE)

  cols_ox <- intersect(oxidos, colnames(sub))
  cols_tr <- intersect(trazas, colnames(sub))

  # Construir matrices numericas
  mat_ox <- matrix(0, nrow = nrow(sub), ncol = length(cols_ox),
                   dimnames = list(sub$UNINORTE_CODE, cols_ox))
  for (col in cols_ox) {
    mat_ox[, col] <- suppressWarnings(as.numeric(sub[[col]]))
    mat_ox[is.na(mat_ox[,col]), col] <- 0
  }

  mat_tr <- matrix(0, nrow = nrow(sub), ncol = max(1, length(cols_tr)),
                   dimnames = list(sub$UNINORTE_CODE,
                                   if(length(cols_tr)>0) cols_tr else "V"))
  for (col in cols_tr) mat_tr[,col] <- suppressWarnings(as.numeric(sub[[col]]))
  mat_tr[is.na(mat_tr)] <- 0

  # Archivo temporal con extension .dat (formato texto nativo GCDkit)
  tmp <- tempfile(fileext = ".dat")

  # Fila de encabezados
  header <- paste(c("Samp", cols_ox, cols_tr), collapse = "\t")
  # Fila de unidades: "end" para Samp, "%" para oxidos, "ppm" para trazas
  units  <- paste(c("end",
                    rep("%",   length(cols_ox)),
                    rep("ppm", length(cols_tr))), collapse = "\t")
  # Filas de datos
  data_lines <- sapply(sub$UNINORTE_CODE, function(s) {
    vals_ox <- as.character(round(mat_ox[s, ], 4))
    vals_tr <- if (length(cols_tr) > 0) as.character(round(mat_tr[s, ], 2)) else character(0)
    paste(c(s, vals_ox, vals_tr), collapse = "\t")
  })

  writeLines(c(header, units, data_lines), tmp)
  loadData(tmp)
  return(TRUE)
}

# Guardar diagrama como PNG
rutas <- list()
guardar <- function(key, expr, w = 1200, h = 950, res = 130) {
  ruta <- file.path(out_dir, paste0(key, ".png"))
  tryCatch({
    png(ruta, width = w, height = h, res = res, bg = "white", type = "cairo")
    eval(expr)
    dev.off()
    cat("  ->", basename(ruta), "\n")
    rutas[[key]] <<- ruta
  }, error = function(e) {
    try(dev.off(), silent = TRUE)
    cat("  FALLO", key, ":", conditionMessage(e), "\n")
  })
}

# =============================================================================
# IGNEAS
# =============================================================================
cat("\n== IGNEAS ==\n")
ig <- geo[toupper(geo$TIPO_ROCA) == "IGNEA", ]
if (nrow(ig) > 0 && cargar_gcdkit(ig)) {

  guardar("igneas_tas", quote({
    TAS()
    title("Diagrama TAS — Le Bas et al. (1986)", cex.main = 1.1)
  }))

  guardar("igneas_qapf", quote({
    QAPFPlut()
    title("QAPF Plutonico — Streckeisen (1976)", cex.main = 1.1)
  }))

  guardar("igneas_afm", quote({
    AFM()
    title("Diagrama AFM — Kuno (1968)", cex.main = 1.1)
  }))

  guardar("igneas_jensen", quote({
    Jensen()
    title("Diagrama de Jensen (1976)", cex.main = 1.1)
  }))

  guardar("igneas_shervais", quote({
    Shervais()
    title("Shervais (1982): V vs Ti", cex.main = 1.1)
  }))

  guardar("igneas_harker_mgo", quote({
    plot(WR[,"SiO2"], WR[,"MgO"],
         xlab = "SiO2 (%)", ylab = "MgO (%)",
         main = "Harker: SiO2 vs MgO",
         pch = 21, bg = "#e63946", cex = 1.5, col = "white",
         cex.lab = 1.1, cex.main = 1.1, las = 1)
    text(WR[,"SiO2"], WR[,"MgO"], labels = rownames(WR),
         pos = 3, cex = 0.72, col = "gray25")
    grid(col = "gray90")
  }))

  guardar("igneas_harker_al2o3", quote({
    plot(WR[,"SiO2"], WR[,"Al2O3"],
         xlab = "SiO2 (%)", ylab = "Al2O3 (%)",
         main = "Harker: SiO2 vs Al2O3",
         pch = 21, bg = "#e63946", cex = 1.5, col = "white",
         cex.lab = 1.1, cex.main = 1.1, las = 1)
    text(WR[,"SiO2"], WR[,"Al2O3"], labels = rownames(WR),
         pos = 3, cex = 0.72, col = "gray25")
    grid(col = "gray90")
  }))
}

# =============================================================================
# SEDIMENTARIAS
# =============================================================================
cat("\n== SEDIMENTARIAS ==\n")
sed <- geo[toupper(geo$TIPO_ROCA) == "SEDIMENTARIA", ]
if (nrow(sed) > 0 && cargar_gcdkit(sed)) {

  # CIA (calculado manualmente)
  guardar("sed_cia", quote({
    MW_v <- c(Al2O3=101.96, CaO=56.08, Na2O=61.98, K2O=94.20, P2O5=141.94)
    al  <- WR[,"Al2O3"] / MW_v["Al2O3"]
    cao <- WR[,"CaO"]   / MW_v["CaO"]
    na  <- WR[,"Na2O"]  / MW_v["Na2O"]
    k   <- WR[,"K2O"]   / MW_v["K2O"]
    p   <- if ("P2O5" %in% colnames(WR)) WR[,"P2O5"]/MW_v["P2O5"] else rep(0,nrow(WR))
    cao_corr <- pmax(0, cao - 10/3 * p)
    cia <- al / (al + cao_corr + na + k) * 100

    bp <- barplot(cia, names.arg = names(cia),
            col = ifelse(cia < 65, "#90caf9", ifelse(cia < 85, "#ffb74d", "#ef5350")),
            border = "white", ylim = c(0, 108),
            ylab = "CIA", main = "Indice CIA — Nesbitt & Young (1982)",
            cex.names = 0.9, las = 2, cex.lab = 1.1)
    abline(h = c(65, 85), lty = 2, col = c("steelblue","tomato"), lwd = 1.3)
    text(bp, cia + 1.8, round(cia, 1), cex = 0.85, col = "gray20", font = 2)
    legend("topright", fill = c("#90caf9","#ffb74d","#ef5350"), border = "white",
           legend = c("Poco alterada <65","Moderada 65-85","Intensa >85"),
           cex = 0.82, bty = "n")
    grid(nx = NA, ny = NULL, col = "gray90")
  }))

  guardar("sed_wf", quote({
    WinFloyd1()
    title("Winchester & Floyd (1977) — Sedimentarias", cex.main = 1.1)
  }))
}

# =============================================================================
# METAMORFICAS
# =============================================================================
cat("\n== METAMORFICAS ==\n")
met <- geo[toupper(geo$TIPO_ROCA) == "METAMORFICA", ]
if (nrow(met) > 0 && cargar_gcdkit(met)) {

  guardar("meta_wf1", quote({
    WinFloyd1()
    title("Winchester & Floyd (1977) — Metamorficas", cex.main = 1.1)
  }))

  guardar("meta_wf2", quote({
    WinFloyd2()
    title("Winchester & Floyd (1977) var. 2 — Metamorficas", cex.main = 1.1)
  }))

  guardar("meta_zrti", quote({
    tr_col <- if (!is.null(TRACE) && "Zr" %in% colnames(TRACE)) TRACE[,"Zr"] else rep(NA, nrow(WR))
    plot(WR[,"TiO2"], tr_col,
         xlab = "TiO2 (%)", ylab = "Zr (ppm)",
         main = "Zr vs TiO2 — Rocas Metamorficas",
         pch = 21, bg = "#2a9d3f", cex = 1.5, col = "white",
         cex.lab = 1.1, cex.main = 1.1, las = 1)
    text(WR[,"TiO2"], tr_col, labels = rownames(WR),
         pos = 3, cex = 0.72, col = "gray25")
    grid(col = "gray90")
  }))

  # ACF ternario
  guardar("meta_acf", quote({
    MW_v <- c(Al2O3=101.96, Fe2O3=159.69, K2O=94.20, Na2O=61.98,
              CaO=56.08, P2O5=141.94, MgO=40.30, MnO=70.94)
    al  <- WR[,"Al2O3"] / MW_v["Al2O3"]
    fe3 <- WR[,"Fe2O3"] / MW_v["Fe2O3"]
    k   <- WR[,"K2O"]   / MW_v["K2O"]
    na  <- WR[,"Na2O"]  / MW_v["Na2O"]
    cao <- WR[,"CaO"]   / MW_v["CaO"]
    p   <- if ("P2O5" %in% colnames(WR)) WR[,"P2O5"]/MW_v["P2O5"] else rep(0,nrow(WR))
    mg  <- WR[,"MgO"]   / MW_v["MgO"]
    mn  <- WR[,"MnO"]   / MW_v["MnO"]
    A   <- al + fe3 - k - na
    C   <- pmax(0, cao - 3.33 * p)
    F_v <- fe3 * 0.8998 + mg + mn
    tot <- A + C + F_v
    A   <- A/tot; C <- C/tot; F_v <- F_v/tot

    h <- sqrt(3)/2
    plot(0,0,type="n",xlim=c(-0.15,1.15),ylim=c(-0.15,1.0),asp=1,
         axes=FALSE,xlab="",ylab="",main="Diagrama ACF — Rocas Metamorficas",
         cex.main=1.1)
    polygon(c(0,1,0.5),c(0,0,h),border="black",lwd=2)
    text(0,-0.09,  "A\n(Al2O3+Fe2O3-K2O-Na2O)", cex=0.78, font=2)
    text(1,-0.09,  "F\n(FeO+MgO+MnO)",           cex=0.78, font=2)
    text(0.5,h+0.06,"C\n(CaO-3.33P2O5)",          cex=0.78, font=2)
    x_pts <- F_v + C*0.5
    y_pts <- C*h
    points(x_pts,y_pts,pch=21,bg="#2a9d3f",cex=1.7,col="white",lwd=1.5)
    text(x_pts,y_pts,rownames(WR),pos=3,cex=0.75,col="gray25")
  }))
}

# =============================================================================
# Exportar JSON de rutas para Python
# =============================================================================
cat("\n== Exportando rutas ==\n")
if (length(rutas) > 0) {
  json_lines <- mapply(function(k,v)
    paste0('"', k, '": "', gsub("\\\\","/",v), '"'),
    names(rutas), unlist(rutas))
  writeLines(paste0('{\n', paste(json_lines, collapse=',\n'), '\n}'),
             file.path(out_dir, "diagramas_gcdkit.json"))
  cat("Total diagramas:", length(rutas), "\n")
} else {
  cat("No se genero ningun diagrama.\n")
}
