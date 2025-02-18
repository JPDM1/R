# Cargar las librerías necesarias
rm(list = ls()) # Limpiar el environment
library(readxl)
library(dplyr)
library(openxlsx)
library(crayon)
setwd("~/Documents/Excel")
# Leer los archivos Excel
tabla1 <- read_excel("tabla1.xlsx")
tabla2 <- read_excel("tabla2.xlsx")

# Realizar el join por DNI y marcar las discrepancias
comparacion <- tabla1 %>%
  full_join(tabla2, by = "DNI", suffix = c("_tabla1", "_tabla2")) %>%
  mutate(tiene_discrepancia = `Nombre Completo_tabla1` != `Nombre Completo_tabla2`) #Nombre Completo es el nombre de la columna en común

# Crear un nuevo archivo Excel con formato
wb <- createWorkbook()
addWorksheet(wb, "Comparación")

# Escribir los datos
writeData(wb, "Comparación", comparacion)

# Definir el estilo para las filas con discrepancias
styleDiscrepancia <- createStyle(
  fgFill = "#90EE90",  # Verde claro
  textDecoration = "bold"
)

# Aplicar el formato condicional
filas_con_discrepancias <- which(comparacion$tiene_discrepancia)
addStyle(wb, "Comparación", 
         style = styleDiscrepancia, 
         rows = filas_con_discrepancias + 1,  # +1 porque Excel cuenta el encabezado
         cols = 1:ncol(comparacion), 
         gridExpand = TRUE)

# Ajustar el ancho de las columnas automáticamente
setColWidths(wb, "Comparación", cols = 1:ncol(comparacion), widths = "auto")

# Guardar el archivo
saveWorkbook(wb, "analisis_discrepancias.xlsx", overwrite = TRUE)

# Imprimir resumen
cat("Análisis completado:\n")
cat("Total de registros:", nrow(comparacion), "\n")
cat("Registros con discrepancias:", sum(comparacion$tiene_discrepancia), "\n")
