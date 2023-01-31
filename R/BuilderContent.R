BuilderContent <- function(the_position_content, packname = "armado", min_digits_name = 2){

  cantidad_grupos <- max(the_position_content$Pack)
  vector_grupos <- 1:cantidad_grupos

  armado <- list()
  sentencia_general <- "_la_lista_[[_la_pos_]]"

  for(k1 in vector_grupos){

    dt_filas <- the_position_content$Pack == k1

    mini_the_position_content <- the_position_content[dt_filas, ]

    sentencia_ejecucion <- c()
    for(k2 in 1:nrow(mini_the_position_content)){

      este_objeto <- mini_the_position_content$ObjList[k2]
      esta_pos <- mini_the_position_content$ObjPos[k2]

      nueva_sentencia <- sentencia_general
      nueva_sentencia <- gsub("_la_lista_", este_objeto, nueva_sentencia)
      nueva_sentencia <- gsub("_la_pos_", esta_pos, nueva_sentencia)

      sentencia_ejecucion <- c(sentencia_ejecucion, nueva_sentencia)
    }

    sentencia_ejecucion <- paste0("list(", paste0(sentencia_ejecucion, collapse = ", "), ")")

    nombre_lista <- mini_the_position_content$ObjName
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
