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

# Titulos Graficos
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


# Director de Orquesta
{

  Orquestador <- function(orquesta = NULL, ObjList = NULL, ObjPos = NA,
                          ObjName = NA, Pack = NULL, PackPos = NA){

    mis_columnas <- c("Orden", "ObjList", "ObjPos", "ObjName", "Pack", "PackPos")

    if(is.null(orquesta)){

      orquesta <- as.data.frame(matrix(NA, nrow= 0, ncol =length(mis_columnas)))
      colnames(orquesta) <- mis_columnas
      return(orquesta)
    }

    # original_args <- formals(Orquestador)
    original_args <- Hmisc::llist(orquesta, ObjList, ObjPos,
                                  ObjName, Pack, PackPos)
    arg_names <- names(original_args)
    pos_musicians <- 2:6
    the_musicians <- arg_names[pos_musicians]
    musicians_args <- sapply(the_musicians, function(x){
      aver <- original_args[[x]]
      aver
    }, USE.NAMES = T)




    # Control for argument 'orquesta'
    {
      if(!is.data.frame(orquesta)){
        cat("El argumento 'orquesta' debe ser un dataframe.")
        stop()
      }

      if(!identical(mis_columnas, colnames(orquesta))){
        cat("Las columnas de la 'orquesta' no son las esperadas. Revisar!")
        stop()
      }
    }

    # Control for NULLs arguments
    {
      if(sum(unlist(lapply(original_args, is.null))) > 0){

      vector_null <- unlist(lapply(original_args, is.null))
      count_null <- sum(vector_null)
      dt_global_null <- count_null > 0

      if(dt_global_null){
        text_part01 <- paste0("Hay ", count_null, " argumentos NULL: ")
        text_part02 <- paste0(arg_names, collapse = ", ")
        all_text <- paste0(text_part01, text_part02, ".\n")
        cat(all_text)
        stop()
      }
      }
    }

    # Control is.vector() for musicians
    {
      if(sum(!unlist(lapply(musicians_args, is.vector))) > 0){

        dt_non_vector <- !unlist(lapply(musicians_args, is.vector))
        count_non_vector <- sum(dt_non_vector)
        dt_global_non_vector <- count_non_vector > 0

        if(dt_global_non_vector){
          text_part01 <- paste0("Hay ", count_non_vector, " argumentos no vectores: ")
          text_part02 <- paste0(the_musicians, collapse = ", ")
          all_text <- paste0(text_part01, text_part02, ".")
          cat(all_text)
          stop()
        }
      }
    }


    # Control length() == 1,  for musicians
    {
      if(sum(!unlist(lapply(musicians_args, function(x){ length(x) == 1}))) > 0){

        dt_non_len1 <- !unlist(lapply(musicians_args, function(x){ length(x) == 1}))
        count_non_len1 <- sum(dt_non_len1)
        dt_global_non_len1 <- count_non_len1 > 0

        if(dt_global_non_len1){
          text_part01 <- paste0("Hay ", count_non_len1, " que no son vectores de longitud 1: ")
          text_part02 <- paste0(the_musicians, collapse = ", ")
          all_text <- paste0(text_part01, text_part02, ".")
          cat(all_text)
          stop()
        }
      }
    }

    # Control for 'ObjList'
    {
      ObjDetection <- function(all_obj, NamesObjList){

        dt_ObjList <- sapply(1:length(NamesObjList), function(x){
          sum(all_obj == NamesObjList[x]) > 0
        })

        dt_out <- sum(dt_ObjList) == length(dt_ObjList)


        if(!dt_out){
          text_part01 <- "Los objetos "
          mod_names <- paste0("'", ObjList[!dt_ObjList], "'")
          text_part02 <- paste0(mod_names, collapse = ", ")
          text_part03 <- " no son objetos del entorno o no existen."
          text_part04 <- "Verificar el argumento 'ObjList'."
          text_out <- paste0(text_part03, text_part02, text_part01, "\n",  text_part04, "\n")
          cat(text_out)
        }

        return(dt_out)
      }


      if(!ObjDetection(all_obj = ls(name = .GlobalEnv),
                      NamesObjList = original_args$ObjList)) stop()

    }


    # Control for ObjPos y ObjName
    {

      if(is.na(ObjPos) && is.na(ObjName)){
        cat("Los argumentos ObjPos y ObjName no pueden ser ambos simultaneamente NA.")
        stop()
      }

      if(is.null(ObjPos) && is.null(ObjName)){
        cat("Los argumentos ObjPos y ObjName no pueden ser ambos simultaneamente NULL.")
        stop()
      }

      if((ObjPos == "") && (ObjName == "")){
        cat("Los argumentos ObjPos y ObjName no pueden ser ambos simultaneamente ''.")
        stop()
      }

      if(!is.numeric(ObjPos) && !is.character(ObjName)){
        cat("El argumento ObjPos debe ser numerico y el argumento ObjName debe ser character.")
        stop()
      }

      if((length(ObjPos) == 0) && (length(ObjName) == 0)){
        cat("Los argumentos ObjPos y ObjName no pueden ser ambos de longitud cero.")
        stop()
      }

    }

    # Control for 'ObjPos'
    {

      # El objeto debe ser numerico
      if(!is.na(ObjPos) && !is.numeric(ObjPos)){
        cat("El argumento 'ObjPos' debe ser un vector numerico con un solo un elemento.\n")
        stop()
      }


      LengthList <- function(ObjList, ObjPos){
        sentencia <- 'length(_ObjList_)'
        sentencia <- gsub(pattern = "_ObjList_", replacement = ObjList, x = sentencia)

        cantidad_elementos <- eval(expr = parse(text = sentencia), envir = .GlobalEnv)
        dt_global <- ObjPos > cantidad_elementos

        if(dt_global){
          final_text <- paste0("La lista '", ObjList, "' posee ", cantidad_elementos,
                 " y estas intentando indexar la posicion ", ObjPos, ".")

          cat(final_text)
        }

        return(dt_global)
      }

      # No podemos tener un index superior a la cantidad de elementos de la lista
      if(!is.na(ObjPos) && LengthList(ObjList, ObjPos)) stop()

      # Si se tiene la posicion y no se tiene el nombre del objeto...
      if(!is.na(ObjPos) && is.na(ObjName)){
        sentencia <- 'names(_ObjList_)'
        sentencia <- gsub(pattern = "_ObjList_", replacement = ObjList, x = sentencia)

        nombre_originales <- eval(expr = parse(text = sentencia), envir = .GlobalEnv)
        ObjName <- nombre_originales[ObjPos]
      }
    }

    # Control for 'ObjName'
    {
      # El objeto debe ser tipo character
      if(!is.na(ObjName) && !is.character(ObjName)){
        cat("El argumento 'ObjName' debe ser un vector character con un solo un elemento.\n")
        stop()
      }


      NameDetectionList <- function(ObjList, ObjName){
        sentencia <- 'names(_ObjList_)'
        sentencia <- gsub(pattern = "_ObjList_", replacement = ObjList, x = sentencia)

        nombre_originales <- eval(expr = parse(text = sentencia), envir = .GlobalEnv)
        dt_global <- sum(nombre_originales == ObjName) != 1

        if(dt_global){
          final_text <- paste0("La lista '", ObjList, "' NO posee un objeto llamado '", ObjName, "'.")

          cat(final_text)
        }

        return(dt_global)
      }

      # El nombre del objeto debe estar en la lista que se sta usando
      if(!is.na(ObjName) && NameDetectionList(ObjList, ObjName)) stop()

      # Si se tiene el nombre pero no se tiene la posicon del objeto...
      if(is.na(ObjPos) && !is.na(ObjName)){
        sentencia <- 'names(_ObjList_)'
        sentencia <- gsub(pattern = "_ObjList_", replacement = ObjList, x = sentencia)

        nombre_originales <- eval(expr = parse(text = sentencia), envir = .GlobalEnv)
        dt_nombre <- nombre_originales == ObjName
        orden_nombre <- 1:length(dt_nombre)
        ObjPos <- orden_nombre[dt_nombre]
      }

    }

    # Control for 'Pack'
    {
      # El objeto debe ser numerico
      if(!is.numeric(Pack)){
        cat("El argumento 'Pack' debe ser un vector numerico con un solo un elemento.\n")
        stop()
      }
    }

    # Control for 'PackPos'
    {
      # El objeto debe ser numerico
      if(!is.numeric(PackPos)){
        cat("El argumento 'PackPos' debe ser un vector numerico con un solo un elemento.\n")
        stop()
      }
    }

    # Si llegamos hasta aca, es que todo esta OK
    # Agregamos los musicos a la orquesta!
    new_row <- nrow(orquesta) + 1
    new_vector <- c(new_row, ObjList, ObjPos, ObjName, Pack, PackPos)
    names(new_vector) <- colnames(orquesta)

    # Ordenamos todo
    orquesta[new_row, "Orden"]    <- new_row
    orquesta[new_row, "ObjList"]  <- ObjList
    orquesta[new_row, "ObjName"]  <- ObjName
    orquesta[new_row, "ObjPos"]   <- ObjPos
    orquesta[new_row, "Pack"]     <- Pack
    orquesta[new_row, "PackPos"]  <- PackPos

    return(orquesta)
  }



  orquesta <- Orquestador(orquesta = NULL)

  # # # Pack 1
  # titulo01, .
  # tabla01,
  # titulo02,
  # tabla02,
  # titulo06,
  # grafico06,
  # titulo05,
  # tabla05)

  number_pack <- 1
  # orquesta <- Orquestador(orquesta = orquesta, ObjList = "all_title_DataTable",
  #                         ObjPos = 1, ObjName = NA, Pack = number_pack, PackPos = 1)

  orquesta <- Orquestador(orquesta = orquesta, ObjList = "all_DataTable",
                          ObjPos = 1, ObjName = NA, Pack = number_pack, PackPos = 1)

  # orquesta <- Orquestador(orquesta = orquesta, ObjList = "all_title_DataTable",
  #                         ObjPos = 2, ObjName = NA, Pack = number_pack, PackPos = 3)

  orquesta <- Orquestador(orquesta = orquesta, ObjList = "all_DataTable",
                          ObjPos = 2, ObjName = NA, Pack = number_pack, PackPos = 2)

  # orquesta <- Orquestador(orquesta = orquesta, ObjList = "all_title_Graph_Sentence",
  #                         ObjPos = 1, ObjName = NA, Pack = number_pack, PackPos = 5)
  #
  # orquesta <- Orquestador(orquesta = orquesta, ObjList = "all_Graph_Sentence",
  #                         ObjPos = 1, ObjName = NA, Pack = number_pack, PackPos = 3)

  # orquesta <- Orquestador(orquesta = orquesta, ObjList = "all_title_DataTable",
  #                         ObjPos = 6, ObjName = NA, Pack = number_pack, PackPos = 6)

  orquesta <- Orquestador(orquesta = orquesta, ObjList = "all_DataTable",
                          ObjPos = 6, ObjName = NA, Pack = number_pack, PackPos = 3)

  # # # Pack 2
  number_pack <- 2
  orquesta <- Orquestador(orquesta = orquesta, ObjList = "all_DataTable", ObjPos = 1,
                          ObjName = "tabla01", Pack = number_pack, PackPos = 1)

  orquesta <- Orquestador(orquesta = orquesta, ObjList = "all_DataTable", ObjPos = NA,
                          ObjName = "tabla01", Pack = number_pack, PackPos = 2)

  orquesta <- Orquestador(orquesta = orquesta, ObjList = "all_DataTable", ObjPos = 1,
                          ObjName = NA, Pack = number_pack, PackPos = 3)

  Armador <- function(orquesta, packname = "armado", min_digits_name = 2){

    cantidad_grupos <- max(orquesta$Pack)
    vector_grupos <- 1:cantidad_grupos

    armado <- list()
    sentencia_general <- "_la_lista_[[_la_pos_]]"

    for(k1 in vector_grupos){

      dt_filas <- orquesta$Pack == k1

      mini_orquesta <- orquesta[dt_filas, ]

      sentencia_ejecucion <- c()
      for(k2 in 1:nrow(mini_orquesta)){

        este_objeto <- mini_orquesta$ObjList[k2]
        esta_pos <- mini_orquesta$ObjPos[k2]

        nueva_sentencia <- sentencia_general
        nueva_sentencia <- gsub("_la_lista_", este_objeto, nueva_sentencia)
        nueva_sentencia <- gsub("_la_pos_", esta_pos, nueva_sentencia)

        sentencia_ejecucion <- c(sentencia_ejecucion, nueva_sentencia)
      }

        sentencia_ejecucion <- paste0("list(", paste0(sentencia_ejecucion, collapse = ", "), ")")

        nombre_lista <- mini_orquesta$ObjName
        nombre_final <- nombre_lista

        for(k3 in 1:(length(nombre_lista)-1)){

          este_nombre <- nombre_lista[k3]
          los_siguientes <- (k3+1):length(nombre_lista)
          dt_idem <- nombre_lista[los_siguientes] == este_nombre
          orden_idem <- (1:length(dt_idem))[dt_idem]

          if(length(orden_idem) > 0){
            for(k4 in 1:length(orden_idem)){
              nombre_final[orden_idem[k4]] <- paste0(nombre_lista[orden_idem[k4]], ".", k4)
            }
          }
        }

        armado[[k1]] <- eval(expr = parse(text = sentencia_ejecucion), envir = .GlobalEnv)
        names(armado[[k1]]) <- nombre_final

    }

    digits <- floor(log10(length(armado))) + 1
    if(digits < min_digits_name) digits <- min_digits_name

    vector_orden <- 1:length(armado)
    numero_orden <- str_pad(string = vector_orden,
                            width = digits,
                            side = "left", pad = "0")

    names(armado) <- paste0(packname, numero_orden)

    return(armado)
  }


  contenido_general <- Armador(orquesta = orquesta)




}


# titulo03, tabla03, titulo04, tabla04, titulo07, tabla07,   titulo08, tabla08, titulo09, tabla09)
# titulo10, tabla10, titulo11, tabla11, titulo12, grafico12, titulo13, tabla13)



# Contenido General
if(FALSE){

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
