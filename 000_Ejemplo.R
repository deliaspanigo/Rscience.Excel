# Remove all
remove(list = ls())

library("stringr")
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
  tabla10 <- tabla01
  tabla11 <- tabla02
  tabla13 <- tabla05


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
  grafico12 <- grafico06

}




SpecialFormat.Generic(obj_name = c("titulo"),
                      new_attr = "Title01",
                      strict = FALSE,
                      execution = TRUE,
                      verbose = TRUE,
                      tellme_formato = FALSE)



SpecialFormat.Generic(obj_name = "tabla",
                      new_attr = "DataTable",
                      strict = FALSE,
                      execution = TRUE,
                      verbose = TRUE,
                      tellme_formato = FALSE)



SpecialFormat.Generic(obj_name = "grafico",
                      new_attr = "Graph_Sentence",
                      strict = FALSE,
                      execution = TRUE,
                      verbose = TRUE,
                      tellme_formato = FALSE)


# Armamos el contenido general
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


armado_especial <- Formato.ExcelOutput(contenido_general = contenido_general,
                                       dt_byrow  = F,
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




Fill.Excel.Sheet(wb = wb,
                 sheet_name = "Tablas",
                 armado_especial = armado_especial,
                 new_sheet = T,
                 start_row = 1,
                 start_col = 1)

# save
saveWorkbook(wb, file =  complete_path, overwrite = T)

cat("Fin de ", complete_path, "\n")


openXL(complete_path)
