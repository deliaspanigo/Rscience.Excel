# Remove all
remove(list = ls())

library("openxlsx")
library("Rscience.base")
# library("Rscience.tools")
library("Rscience.Excel")

# Formatos internos de la funcion nueva

# Tablas varias
{


  tabla01 <- as.data.frame(psych::describe(mtcars))
  tabla02 <- rbind(as.data.frame(psych::describe(iris)), as.data.frame(psych::describe(iris)))
  tabla03 <- as.data.frame(summary(lm(mtcars))[[4]])
  tabla03 <- tabla03[c(1, nrow(tabla03)-1), ]
  tabla04 <- as.data.frame(summary(lm(iris))[[4]])
  tabla05 <- cars
  tabla07 <- tabla05
  tabla08 <- tabla05
  tabla09 <- tabla05

  titulo01 <- "Titulo01"
  titulo02 <- "Titulo02"
  titulo03 <- "Titulo03"
  titulo04 <- "Titulo04"
  titulo05 <- "Titulo05"
  titulo06 <- "Titulo06_grafico"
  titulo07 <- "Titulo07"
  titulo08 <- "Titulo08"
  titulo09 <- "Titulo09"
  titulo10 <- "Titulo10"
  titulo11 <- "Titulo11"
  titulo12 <- "Titulo12_grafico"
  titulo13 <- "Titulo13"

  grafico06 <- "plot(mtcars)"
}

# Les agregamos Metadatos a todos.
{
  titulo01 <- structure(titulo01, ExcelOutput = "Title01")
  titulo02 <- structure(titulo02, ExcelOutput = "Title01")
  titulo03 <- structure(titulo03, ExcelOutput = "Title01")
  titulo04 <- structure(titulo04, ExcelOutput = "Title01")
  titulo05 <- structure(titulo05, ExcelOutput = "Title01")
  titulo06 <- structure(titulo06, ExcelOutput = "Title01")
  titulo07 <- structure(titulo07, ExcelOutput = "Title01")
  titulo08 <- structure(titulo08, ExcelOutput = "Title01")
  titulo09 <- structure(titulo09, ExcelOutput = "Title01")
  titulo10 <- structure(titulo10, ExcelOutput = "Title01")
  titulo11 <- structure(titulo11, ExcelOutput = "Title01")
  titulo12 <- structure(titulo12, ExcelOutput = "Title01")
  titulo13 <- structure(titulo13, ExcelOutput = "Title01")


  tabla01 <- structure(tabla01, ExcelOutput = "DataTable")
  tabla02 <- structure(tabla02, ExcelOutput = "DataTable")
  tabla03 <- structure(tabla03, ExcelOutput = "DataTable")
  tabla04 <- structure(tabla04, ExcelOutput = "DataTable")
  tabla05 <- structure(tabla05, ExcelOutput = "DataTable")
  tabla07 <- structure(tabla07, ExcelOutput = "DataTable")
  tabla08 <- structure(tabla08, ExcelOutput = "DataTable")
  tabla09 <- structure(tabla09, ExcelOutput = "DataTable")
  tabla10 <- structure(tabla01, ExcelOutput = "DataTable")
  tabla11 <- structure(tabla02, ExcelOutput = "DataTable")
  tabla13 <- structure(tabla05, ExcelOutput = "DataTable")

  grafico06 <- structure(grafico06, ExcelOutput = "Graph_Sentence")
  grafico12 <- structure(grafico06, ExcelOutput = "Graph_Sentence")

  # Veamos uno
  # attributes(tabla01)
  # attributes(tabla01)$ExcelOutput
}


# El resto de los Inputs
{
  # Contenido General
  # El contenido general es una estructura de listas.
  # Cada lista contenida es una fila o una columna de lo que sera la salida
  # en Excel. El argumento 'dt_byrow' si es T sera por fila, si es False sera
  # por columnas.
  contenido_general <- list()
  contenido_general[[1]] <- Hmisc::llist(titulo01, tabla01, titulo02, tabla02, titulo06, grafico06, titulo05, tabla05)
  contenido_general[[2]] <- Hmisc::llist(titulo03, tabla03, titulo04, tabla04, titulo07, tabla07,   titulo08, tabla08, titulo09, tabla09)
  contenido_general[[3]] <- Hmisc::llist(titulo10, tabla10, titulo11, tabla11, titulo12, grafico12, titulo13, tabla13)
  names(contenido_general) <- paste0("contenido", 1:length(contenido_general))

  # contenido_general <- Hmisc::llist(columna01, columna02)



}
#
# contenido_general_mod <- contenido_general
#
# for(k_contenido in 1:length(contenido_general_mod)){
#   for(k_objeto in 1:length(contenido_general_mod[[k_contenido]])){
#
#     este_objeto <- NULL
#     este_objeto <- contenido_general_mod[[k_contenido]][[k_objeto]]
#
#     dt_transponer <- (is.data.frame(este_objeto) | is.matrix(este_objeto))
#
#     if(dt_transponer) {
#       contenido_general_mod[[k_contenido]][[k_objeto]] <- t(contenido_general_mod[[k_contenido]][[k_objeto]])
#       colnames(contenido_general_mod[[k_contenido]][[k_objeto]]) <- paste0("Col", 1:ncol(contenido_general_mod[[k_contenido]][[k_objeto]]))
#       rownames(contenido_general_mod[[k_contenido]][[k_objeto]]) <- paste0("Row", 1:nrow(contenido_general_mod[[k_contenido]][[k_objeto]]))
#       contenido_general_mod[[k_contenido]][[k_objeto]] <- as.data.frame(contenido_general_mod[[k_contenido]][[k_objeto]])
#
#       contenido_general_mod[[k_contenido]][[k_objeto]] <- structure(contenido_general_mod[[k_contenido]][[k_objeto]], ExcelOutput = "DataTable")
#     }
#   }
# }
#
# contenido_general <- contenido_general_mod

armado_especial <- Formato.ExcelOutput(contenido_general = contenido_general,
                                       dt_byrow  = T,
                                       vertical_space = 4,
                                       horizontal_space = 5,
                                       vertical_graph = 20,
                                       horizontal_graph = 7,
                                       tellme_formato = F)

folder_path <- "./"
file_name <- "PRUEBA_ESPECIAL"
complete_path <- paste0(folder_path, file_name, ".xlsx")

cat("Inicio de ", complete_path, "\n")

wb <- createWorkbook()




Fill.Excel.Sheet(wb = wb, sheet_name = "Tablas",
                 armado_especial = armado_especial,
                 new_sheet = T, start_row = 1, start_col = 1)

# save
saveWorkbook(wb, file =  complete_path, overwrite = T)

cat("Fin de ", complete_path, "\n")


openXL(complete_path)
