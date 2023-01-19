SpecialFormat.Generic <- function(obj_name,
                                  new_attr,
                                  strict = FALSE,
                                  execution = TRUE,
                                  verbose = TRUE,
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


  Ejecucion_DobleLista <- function(out){
    for(k1 in 1:length(out)){
      for(k2 in 1:length(out[[k1]])){
        eval(expr = parse(text = out[[k1]][[k2]]), envir = .GlobalEnv)
      }
    }
  }



  # Solo salida de formato que se usa...
  if(tellme_formato) return(formatos_internos)

  obj_name <- unique(obj_name)

  # The function
  if(length(obj_name) == 1){


    # Internal Objects
    specialformat_sentence <- "attr(_the_target_, 'ExcelOutput') <- '_new_attr_'"

    obj_name_mod <- obj_name
    if(strict) obj_name_mod <- paste0("^", obj_name_mod, "$")

    the_target <- grep(pattern = obj_name_mod,
                       x = c(ls(name = .GlobalEnv)),
                       value = TRUE)

    armado01 <- paste0("***Se encontraron ", length(the_target), " objetos relacionados a '", obj_name, "'***", "\n")
    armado02 <- paste0("***Se les asignara: '", new_attr, "'***", "\n")
    if(verbose) cat(armado01)
    if(verbose) cat(armado02)

    if(length(the_target) > 0){

      cantidad_digitos <- floor(log10(max(na.omit(length(the_target))))) + 1

      out <- list()
      out[[1]] <- list()
      names(out)[1] <- obj_name

      for(k1 in 1:length(the_target)){
        specialformat_mod <- specialformat_sentence

        specialformat_mod <- gsub(pattern = "_the_target_",
                                  replacement = the_target[k1],
                                  x = specialformat_mod)

        specialformat_mod <- gsub(pattern = "_new_attr_",
                                  replacement = new_attr,
                                  x = specialformat_mod)

        # position_number01 <- str_pad(string = k1,
        #                              width = cantidad_digitos,
        #                              side = "left",
        #                              pad = "0")



        out[[1]][[k1]] <- specialformat_mod
        names(out[[1]])[k1] <- the_target[k1]

      }


    } else cat("Not match for ", obj_name, ".")

  }else{
    out <- list()
    for(k2 in 1:length(obj_name)){
      out[[k2]] <- SpecialFormat.Generic(obj_name = obj_name[k2],
                                         new_attr = new_attr,
                                         strict =  strict,
                                         verbose = verbose,
                                         tellme_formato = tellme_formato)

    }

  }

  if(execution) Ejecucion_DobleLista(out) else return(out)



}
