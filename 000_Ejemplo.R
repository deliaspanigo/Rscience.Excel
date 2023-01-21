# Remove all
remove(list = ls())

library("stringr")
library("openxlsx")
library("Rscience.base")
# library("Rscience.tools")
library("Rscience.Excel")

# Funciones
{

  SpecialFormat.List.DataTable <- function(x,
                                           selected_pos = NULL,
                                           tellme_formato = FALSE){

    formatos_internos <- c("Title01",
                           "Title02",
                           "Title03",
                           "DataTable", "DataTable_Sentence", "DataTable_Path", "DataTable_Space",
                           "Table",     "Table_Sentence",     "Table_Path",     "Table_Space",
                           "Text",      "Text_Sentence",      "Text_Path",      "Text_Space",
                           "Graph_Object",     "Graph_Sentence",     "Graph_Path",     "Graph_Space",
                           "Fusion_Cell" # nuevo!
    )

    # Solo salida de formato que se usa...
    if(tellme_formato) return(formatos_internos)

    if(is.null(selected_pos)) if(is.list(x)) selected_pos <- 1:length(x)

    new_atrr = "DataTable"
    out <- x
    dt_general01 <- is.data.frame(out)
    dt_general02 <- is.list(out) && is.vector(out)

    if(dt_general01){
      attr(out, 'ExcelOutput') <- new_atrr
    } else
      if(dt_general02){
        if(length(out) == 1){
          attr(out[[1]], 'ExcelOutput') <- new_atrr
        } else
          if(length(out) > 1)
            for(k1 in selected_pos) {
              out[[k1]] <- SpecialFormat.List.DataTable(x = out[[k1]], tellme_formato = tellme_formato)
            }
      }


    return(out)

  }

  SpecialFormat.List.Graph_Sentence <- function(x,
                                                selected_pos = NULL,
                                                tellme_formato = FALSE){

    formatos_internos <- c("Title01",
                           "Title02",
                           "Title03",
                           "DataTable", "DataTable_Sentence", "DataTable_Path", "DataTable_Space",
                           "Table",     "Table_Sentence",     "Table_Path",     "Table_Space",
                           "Text",      "Text_Sentence",      "Text_Path",      "Text_Space",
                           "Graph_Object",     "Graph_Sentence",     "Graph_Path",     "Graph_Space",
                           "Fusion_Cell" # nuevo!
    )

    # Solo salida de formato que se usa...
    if(tellme_formato) return(formatos_internos)

    if(is.null(selected_pos)) if(is.list(x)) selected_pos <- 1:length(x)

    new_atrr = "Graph_Sentence"
    out <- x
    dt_general01 <- !is.list(out) && (is.null(ncol(out)) && is.null(nrow(out)))
    dt_general02 <- is.list(out) && is.vector(out)

    if(dt_general01){
      attr(out, 'ExcelOutput') <- new_atrr
    } else
      if(dt_general02){
        if(length(out) == 1){
          attr(out[[1]], 'ExcelOutput') <- new_atrr
        } else
          if(length(out) > 1)
            for(k1 in selected_pos) {
              out[[k1]] <- SpecialFormat.List.Graph_Sentence(x = out[[k1]], tellme_formato = tellme_formato)
            }
      }


    return(out)

  }


}



# DataTable
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

  objetos_DataTable <- grep(ls(), pattern = "tabla", value = TRUE)

  armado <- paste0(objetos_DataTable, collapse = ", ")
  armado <- paste0("Hmisc::llist(", armado, ")")

  all_DataTable <- eval(parse(text = armado))

  all_DataTable <- SpecialFormat.List.DataTable(x = all_DataTable,
                                                selected_pos = NULL,
                                                tellme_formato = FALSE)

}

# Titulos Tablas
{
  cantidad_digitos <- floor(log10(length(all_DataTable))) + 1

  new_title <- "title_DataTable"

  orden_DataTable <- str_pad(string = 1:length(all_DataTable),
                                 width = cantidad_digitos,
                                 side = "left", pad = "0")

  all_title_DataTable <- paste0(new_title, orden_DataTable)

  all_title_DataTable <- as.list(all_title_DataTable)
  names(all_title_DataTable) <- unlist(all_title_DataTable)
  all_title_DataTable <- SpecialFormat.List.Graph_Sentence(x =  all_title_DataTable,
                                                         selected_pos = NULL,
                                                         tellme_formato = FALSE)
}

# Graficos
{
  grafico01 <- "plot(mtcars)"
  grafico02 <- grafico01

  objetos_Graph_Sentence <- grep(ls(), pattern = "grafico", value = TRUE)

  armado <- paste0(objetos_Graph_Sentence, collapse = ", ")
  armado <- paste0("Hmisc::llist(", armado, ")")

  all_Graph_Sentence <- eval(parse(text = armado))

  all_Graph_Sentence <- SpecialFormat.List.DataTable(x = all_Graph_Sentence,
                                                      selected_pos = NULL,
                                                      tellme_formato = FALSE)
}

# Titulos Tablas
{
  cantidad_digitos <- floor(log10(length(all_Graph_Sentence))) + 1

  new_title <- "title_Graph_Sentence"

  orden_Graph_Sentence <- str_pad(string = 1:length(all_Graph_Sentence),
                             width = cantidad_digitos,
                             side = "left", pad = "0")

  all_title_Graph_Sentence  <- paste0(new_title, orden_Graph_Sentence)

  all_title_Graph_Sentence <- as.list(all_title_Graph_Sentence)
  names(all_title_Graph_Sentence) <- unlist(all_title_Graph_Sentence)
  all_title_Graph_Sentence <- SpecialFormat.List.Graph_Sentence(x =  all_title_Graph_Sentence,
                                                                selected_pos = NULL,
                                                                tellme_formato = FALSE)
}

# Contenido General
{

  row01 <- list(all_DataTable[[1]], all_DataTable[[3]], all_DataTable[[5]])
  names(row01) <- names(all_DataTable)[c(1,3,5)]

  row02 <- list(all_DataTable[[1]], all_DataTable[[3]], all_DataTable[[5]])
  names(row02) <- names(all_DataTable)[c(1,3,5)]

  row03 <- list(all_DataTable[[1]], all_DataTable[[3]], all_DataTable[[5]])
  names(row03) <- names(all_DataTable)[c(1,3,5)]

  contenido_general <- list(row01, row02, row03)
  names(contenido_general) <- paste0("contenido", 1:length(contenido_general))

}



if(FALSE){

  SpecialFormat.Generic(obj_name = "titulo",
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
}





# Armamos el contenido general
if(FALSE){
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





attributes(contenido_general$contenido1$titulo01)
attributes(contenido_general$contenido1$tabla01)
attributes(contenido_general$contenido1$grafico06)


# Todos los dataframe seran DataTable
contenido_general <- SpecialFormat.List.DataTable(x = contenido_general,
                                                  selected_pos = NULL,
                                                  tellme_formato = FALSE)

# Todos los vectores seran Graph_Sentence
contenido_general <- SpecialFormat.List.Graph_Sentence(x = contenido_general,
                                                       selected_pos = NULL,
                                                       tellme_formato = FALSE)

attributes(contenido_general$contenido1$titulo01)
attributes(contenido_general$contenido1$tabla01)
attributes(contenido_general$contenido1$grafico06)
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
