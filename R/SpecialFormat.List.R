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



SpecialFormat.List.Title03 <- function(x,
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

  new_atrr = "Title03"
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
            out[[k1]] <- SpecialFormat.List.Title03(x = out[[k1]], tellme_formato = tellme_formato)
          }
    }


  return(out)

}

SpecialFormat.List.Title <- function(x,
                                     selected_pos = NULL,
                                     tellme_formato = FALSE){

  out <- SpecialFormat.List.Title03(x = x,
                                    selected_pos = selected_pos,
                                    tellme_formato = tellme_formato)

  return(out)
}
