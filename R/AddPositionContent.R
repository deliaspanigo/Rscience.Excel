AddPositionContent <- function(the_position_content = NULL, ObjList = NULL, ObjPos = NA,
                               ObjName = NA, Pack = NULL, PackPos = NA){

  mis_columnas <- c("Orden", "ObjList", "ObjPos", "ObjName", "Pack", "PackPos")

  if(is.null(the_position_content)){

    the_position_content <- as.data.frame(matrix(NA, nrow= 0, ncol =length(mis_columnas)))
    colnames(the_position_content) <- mis_columnas
    return(the_position_content)
  }

  # original_args <- formals(AddPositionContent)
  original_args <- Hmisc::llist(the_position_content, ObjList, ObjPos,
                                ObjName, Pack, PackPos)
  arg_names <- names(original_args)
  pos_musicians <- 2:6
  the_musicians <- arg_names[pos_musicians]
  musicians_args <- sapply(the_musicians, function(x){
    aver <- original_args[[x]]
    aver
  }, USE.NAMES = T)




  # Control for argument 'the_position_content'
  {
    if(!is.data.frame(the_position_content)){
      cat("El argumento 'the_position_content' debe ser un dataframe.")
      stop()
    }

    if(!identical(mis_columnas, colnames(the_position_content))){
      cat("Las columnas de la 'the_position_content' no son las esperadas. Revisar!")
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
    if(is.na(PackPos)){

      PackPos <- 1
      dt_mini <- the_position_content$Pack == Pack

      if(sum(dt_mini) > 0){
        max_pos_pack <- max(the_position_content$PackPos[dt_mini])
        PackPos <- max_pos_pack + 1
      }
    }

    # El objeto debe ser numerico
    if(!is.numeric(PackPos)){
      cat("El argumento 'PackPos' debe ser un vector numerico con un solo un elemento.\n")
      stop()
    }
  }

  # Si llegamos hasta aca, es que todo esta OK
  # Agregamos los musicos a la the_position_content!
  new_row <- nrow(the_position_content) + 1
  new_vector <- c(new_row, ObjList, ObjPos, ObjName, Pack, PackPos)
  names(new_vector) <- colnames(the_position_content)

  # Ordenamos todo
  the_position_content[new_row, "Orden"]    <- new_row
  the_position_content[new_row, "ObjList"]  <- ObjList
  the_position_content[new_row, "ObjName"]  <- ObjName
  the_position_content[new_row, "ObjPos"]   <- ObjPos
  the_position_content[new_row, "Pack"]     <- Pack
  the_position_content[new_row, "PackPos"]  <- PackPos

  return(the_position_content)
}
