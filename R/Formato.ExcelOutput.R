
Formato.ExcelOutput <- function(contenido_general,
                                dt_byrow = F,
                                general_align = T,
                                row_space = 4,
                                col_space = 5,
                                row_graph = 19,
                                col_graph = 7,
                                tellme_formato = F){




  formatos_internos <- c("Title01",
                         "Title02",
                         "Title03",
                         "DataTable", "DataTable_Sentence", "DataTable_Path", "DataTable_Space",
                         "Table",     "Table_Sentence",     "Table_Path",     "Table_Space",
                         "Text",      "Text_Sentence",      "Text_Path",      "Text_Space",
                         "Graph_Object",     "Graph_Sentence",     "Graph_Path",     "Graph_Space",
                         "Fusion_Cell" # nuevo!
  )

  # Para mas adelante... Podria ser que agregue como parte del paquete de R a una dataframe
  # que tiene todas las combinaciones posibles entre todos los elementos de listado anterior.
  # Y de esa forma se pueda saber cuantos espacio en filas y en columnas hay que dejar
  # para cada combinacion.
  # Por ejemplo... El espacio entre un titulo y una tabla es de 1 fila y 0 en columnas.
  # Es de una fila por que viene el titulo y despues la tabla.
  # Y es 0 en columna por que deben estar en la mismac columna.

  # combinaciones <- expand.grid(formatos_internos, formatos_internos)
  # combinaciones <- combinaciones[c(2,1)]
  # orden_combinaciones <- c(1:nrow(combinaciones))
  # nrow_space <- rep(NA, nrow(combinaciones))
  # ncol_space <- rep(NA, nrow(combinaciones))
  # combinaciones <- cbind.data.frame(orden_combinaciones, combinaciones, nrow_space, ncol_space)
  # write.xlsx(x = combinaciones, file = "CombinacionesOUT.xlsx")

  # Cargamos el espaciado
  # Luego esto sera parte del paquede directamente
  # matrix_espaciados <- read.xlsx(xlsxFile = "CombinacionesOUT.xlsx", sheet = 1)

  # Solo salida de formato que se usa...
  if(tellme_formato) return(formatos_internos)

  if (!dt_byrow){


    # Objetos Generales
    {

      # Nombre general de ubicacion en fila y columna
      nombre_ubicacion <- c("Row", "Col")

      # Nombre de los objetos ingresados
      names_contenido <- names(contenido_general)
      len_contenido <- length(contenido_general)
      len_interno <- unlist(lapply(contenido_general, length))
      max_len_interno <- max(len_interno)
      total_len_contenido <- sum(len_interno)

      # Box General
      dim_box <- c(max_len_interno, len_contenido)
      names(dim_box) <- nombre_ubicacion



      # Total de objetos ingresados
      total_len_box <- as.vector(dim_box[1]*dim_box[2])
      vector_order_box <- 1:total_len_box

      numeric_order_relleno <- 2
      formato_relleno <- nombre_ubicacion[2]
      forma_rellenado <- rep("Por columna", total_len_box)

    }


    # Matrices de Ordenamiento
    {
      # Es el conteo general del max box...
      # A cada objeto le damos un numero, comenzando desde el 1 hasta el n.
      matrix_ordenamiento01 <- matrix(data = vector_order_box,
                                      nrow = dim_box["Row"],
                                      ncol = dim_box["Col"],
                                      byrow = dt_byrow)

      # Es el conteo dentro de cada contenido
      matrix_ordenamiento02 <- sapply(1:len_contenido, function(x){

        generico_vacio <- rep(NA, max_len_interno)
        generico_vacio[1:len_interno[x]] <- 1:len_interno[x]
        generico_vacio
      }, simplify = F, USE.NAMES = F)
      matrix_ordenamiento02 <- do.call(cbind, matrix_ordenamiento02)


      # Nombre de cada elemento
      matrix_ordenamiento03 <- sapply(1:len_contenido, function(x){

        rejunte_info <-        sapply(1:len_interno[x], function(y){
          names(contenido_general[[x]])[y]
        })

        generico_vacio <- rep(NA, max_len_interno)
        generico_vacio[1:len_interno[x]] <- rejunte_info
        generico_vacio

      }, simplify = F, USE.NAMES = F)
      matrix_ordenamiento03 <- do.call(cbind, matrix_ordenamiento03)

      # Formato de cada elemento
      matrix_ordenamiento04 <- sapply(1:len_contenido, function(x){

        rejunte_info <-        sapply(1:len_interno[x], function(y){
          attributes(contenido_general[[x]][[y]])$ExcelOutput
        })

        generico_vacio <- rep(NA, max_len_interno)
        generico_vacio[1:len_interno[x]] <- rejunte_info
        generico_vacio

      }, simplify = F, USE.NAMES = F)
      matrix_ordenamiento04 <- do.call(cbind, matrix_ordenamiento04)



      # Se si es vacio o no cada celda
      matrix_ordenamiento05 <- !is.na(matrix_ordenamiento03)


      # Posicion en filas y columnas de cada celda
      secuencia_columnas <- rep(1:ncol(matrix_ordenamiento01), nrow(matrix_ordenamiento01))
      secuencia_filas <- rep(1:nrow(matrix_ordenamiento01), ncol(matrix_ordenamiento01))



      matrix_ordenamiento06 <- namel(names_vector = nombre_ubicacion, initial_value = NULL)
      matrix_ordenamiento06[["Row"]] <- matrix(secuencia_filas,
                                               nrow = nrow(matrix_ordenamiento01),
                                               ncol = ncol(matrix_ordenamiento01),
                                               byrow = dt_byrow)

      matrix_ordenamiento06[["Col"]] <- matrix(secuencia_columnas,
                                               nrow= nrow(matrix_ordenamiento01),
                                               ncol= ncol(matrix_ordenamiento01),
                                               byrow = !dt_byrow)




      ###########
      caso01 <- c("Title01", "Title02", "Title03", "Text")
      caso02 <- c("DataTable", "Table")
      caso03 <- c("DataTable_Path", "Table_Path", "Text_Path")
      caso04 <- c("DataTable_Sentence", "Table_Sentence", "Text_Sentence")
      caso05 <- c("Graph_Object", "Graph_Path", "Graph_Sentence", "Graph_Space")
      # Aqui esta la cantidad de filas y de columnas de cada objeto
      matrix_ordenamiento07 <- namel(names_vector = nombre_ubicacion, initial_value = NULL)

      matrix_ordenamiento07[["Row"]] <- sapply(1:nrow(matrix_ordenamiento03), function(fila_elegida){

        sapply(1:ncol(matrix_ordenamiento03), function(columna_elegida){

          nombre_objeto <- matrix_ordenamiento03[fila_elegida, columna_elegida]
          conteo_filas <- c()

          if(is.na(nombre_objeto)) conteo_filas[1] <- NA    else
            if(!is.na(nombre_objeto)){

              formato_exceloutput <- matrix_ordenamiento04[fila_elegida, columna_elegida]

              rejunte <- purrr::map(contenido_general, nombre_objeto)
              contenido_aislado <- rejunte[-which(sapply(rejunte, is.null))][[1]]

              dt01 <- sum(caso01 == formato_exceloutput) > 0 # Titulos - length()
              dt02 <- sum(caso02 == formato_exceloutput) > 0 # DataFrame y Tablas - nrow()
              dt03 <- sum(caso03 == formato_exceloutput) > 0 # Path - length()
              dt04 <- sum(caso04 == formato_exceloutput) > 0 # Sentence - length()
              dt05 <- sum(caso05 == formato_exceloutput) > 0 # Graph_Sentence
              dt06 <- "DataTable_Space" == formato_exceloutput
              dt07 <- "Table_Space" == formato_exceloutput
              dt08 <- "Text_Space" == formato_exceloutput

              # Cuando sacamos nrow() le sumamos 1 ya que nrow() no cuenta la fila
              # que ocupa el nombre de las columnas.

              if(dt02) conteo_filas[1] <-  nrow(contenido_aislado) + 1 else # DataFrame y Tablas - nrow()
                if(dt01 | dt03 | dt04) conteo_filas[1] <-  length(contenido_aislado) else
                  if(dt05) conteo_filas[1] <- row_graph else # "Graph" o "Graph_Space"
                    if(dt06 | dt07) conteo_filas[1] <- 5 else # "DataTable_Space"
                      if(dt08) conteo_filas[1] <- 1 # Text_Space


            }
        })

      })
      matrix_ordenamiento07[["Row"]] <- t(matrix_ordenamiento07[["Row"]])

      matrix_ordenamiento07[["Col"]] <- sapply(1:nrow(matrix_ordenamiento03), function(fila_elegida){

        sapply(1:ncol(matrix_ordenamiento03), function(columna_elegida){

          nombre_objeto <- matrix_ordenamiento03[fila_elegida, columna_elegida]
          conteo_columnas <- c()

          if(is.na(nombre_objeto)) conteo_columnas[1] <- NA    else
            if(!is.na(nombre_objeto)){

              formato_exceloutput <- matrix_ordenamiento04[fila_elegida, columna_elegida]

              rejunte <- purrr::map(contenido_general, nombre_objeto)
              contenido_aislado <- rejunte[-which(sapply(rejunte, is.null))][[1]]


              dt01 <- sum(caso01 == formato_exceloutput) > 0 # Titulos - length()
              dt02 <- sum(caso02 == formato_exceloutput) > 0 # DataFrame y Tablas - nrow()
              dt03 <- sum(caso03 == formato_exceloutput) > 0 # Path - length()
              dt04 <- sum(caso04 == formato_exceloutput) > 0 # Sentence - length()
              dt05 <- sum(caso05 == formato_exceloutput) > 0 # Graph_Sentence
              dt06 <- "DataTable_Space" == formato_exceloutput
              dt07 <- "Table_Space" == formato_exceloutput
              dt08 <- "Text_Space" == formato_exceloutput


              if(dt02) conteo_columnas[1] <-  ncol(contenido_aislado) else # DataFrame y Tablas - nrow()
                if(dt01 | dt03 | dt04) conteo_columnas[1] <-  1 else
                  if(dt05) conteo_columnas[1] <- col_graph else # "Graph" o "Graph_Space"
                    if(dt06 | dt07) conteo_columnas[1] <- 5 else # "DataTable_Space"
                      if(dt08) conteo_columnas[1] <- 1 # Text_Space


            }
        })

      })
      matrix_ordenamiento07[["Col"]] <- t(matrix_ordenamiento07[["Col"]])


      # orden_suma <- c(2,1)

      matrix_ordenamiento08 <- namel(names_vector = nombre_ubicacion, initial_value = NULL)
      matrix_ordenamiento08[["Row"]] <- sapply(1:ncol(matrix_ordenamiento07[["Row"]]), function(x){

        vector_aislado <- matrix_ordenamiento07[["Row"]][,x]
        vector_aislado[is.na(vector_aislado)] <- 0
        suma_acumulada <- cumsum(vector_aislado)
        suma_acumulada
        # diferencia_datos <- nrow(matrix_ordenamiento07[["Row"]]) - length(suma_acumulada)
        # salida <- c(suma_acumulada, rep(NA, diferencia_datos))
        #salida
      })


      matrix_ordenamiento08[["Col"]] <- sapply(1:nrow(matrix_ordenamiento07[["Col"]]), function(x){

        vector_aislado <- matrix_ordenamiento07[["Col"]][x,]
        vector_aislado[is.na(vector_aislado)] <- 0
        suma_acumulada <- cumsum(vector_aislado)
        suma_acumulada
      })
      matrix_ordenamiento08[["Col"]] <- t(matrix_ordenamiento08[["Col"]])






    }


    # Acomodamos todo un poco
    {

      # Inicio de cada objeto en cada columna
      matrix_posicion_fila <- matrix_ordenamiento08[["Row"]]
      matrix_posicion_fila <- matrix_posicion_fila + 1 # Essto es a donde empieza cada uno, excepto el primero
      matrix_posicion_fila <- rbind(rep(1, ncol(matrix_posicion_fila)), matrix_posicion_fila)
      matrix_posicion_fila <- matrix_posicion_fila[-nrow(matrix_posicion_fila), ]

      for (k1 in 1:nrow(matrix_ordenamiento05)){
        for (k2 in 1:ncol(matrix_ordenamiento05)){
          if(!matrix_ordenamiento05[k1, k2]) matrix_posicion_fila[k1, k2] <- NA
        }
      }

      ## Hasta aca, cada columna es independiente de otra columna.
      ## El detalle que sigue permite alinear el contenido entre columnas
      ## para que arranquen todos a la misma fila.

      # Inicio alineado de las filas
      if(general_align) matrix_posicion_fila <- t(apply(X = matrix_posicion_fila,
                                                        MARGIN = 1, # Por filas
                                                        FUN = function(x){
                                                          rep(max(na.omit(x)), ncol(matrix_posicion_fila))

                                                        }))
      for (k1 in 1:nrow(matrix_ordenamiento05)){
        for (k2 in 1:ncol(matrix_ordenamiento05)){
          if(!matrix_ordenamiento05[k1, k2]) matrix_posicion_fila[k1, k2] <- NA
        }
      }



      matrix_espacio_fila <- matrix(0, nrow = nrow(matrix_posicion_fila), ncol = ncol(matrix_posicion_fila))

      # Obviamos la primera fila, ya que arriba de la primera fila no hay que hacer espacio.
      # La idea del siguiente doble bucle es la siguiente...
      # Si un objeto es un titulo, debe alejarse del objeto que tiene arriba de el en la misma columna.
      # El alejamiento se hace agregando el espacio intefila al titulo que fue encontrado
      # y a todo lo que venga debajo del titulo en la misma columna, ya que tooodo debe bajar
      # la misma cantidad de espacios.

      # Entonces, si hay un titulo, agregara espacios inter_fila en la columna
      for(llave_columna in 1:ncol(matrix_ordenamiento04)){
        for(llave_fila in 2:nrow(matrix_ordenamiento04)){


          categorias <- c("Title01", "Title02", "Title03")
          este_tipo <- matrix_ordenamiento04[llave_fila, llave_columna]
          if(!is.na(este_tipo)){
            if(sum(categorias == este_tipo) > 0){
              espaciado_nuevo <- rep(0, nrow(matrix_ordenamiento04))
              espaciado_nuevo[llave_fila:nrow(matrix_ordenamiento04)] <- row_space
              matrix_espacio_fila[, llave_columna] <- matrix_espacio_fila[, llave_columna] + espaciado_nuevo
            }
          }
        }
      }
      matrix_posicion_fila <- matrix_posicion_fila + matrix_espacio_fila




      matrix_posicion_columna <- matrix_ordenamiento08[["Col"]]
      matrix_posicion_columna <- matrix_posicion_columna + 1 # Essto es a donde empieza cada uno, excepto el primero
      matrix_posicion_columna <- cbind(rep(1, nrow(matrix_posicion_columna)), matrix_posicion_columna)
      matrix_posicion_columna <- matrix_posicion_columna[,-ncol(matrix_posicion_columna) ]

      # Lo que tenemos hasta aca son las posiciones de columna, sin espacio entre los objetos.
      # O sea, termina uno y arranca el otro inemdiatamente al lado.
      # Hay que agregarle espacios.
      matrix_espacio_columna <- matrix(col_space, nrow = nrow(matrix_posicion_columna),
                                       ncol = ncol(matrix_posicion_columna))
      matrix_espacio_columna[,1] <- rep(0, nrow(matrix_espacio_columna))
      matrix_posicion_columna <- matrix_posicion_columna + matrix_espacio_columna

      for (k1 in 1:nrow(matrix_ordenamiento05)){
        for (k2 in 1:ncol(matrix_ordenamiento05)){
          if(!matrix_ordenamiento05[k1, k2]) matrix_posicion_columna[k1, k2] <- NA
        }
      }

      # Inicio alineado de las columnas
      if(general_align) matrix_posicion_columna <- apply(matrix_posicion_columna, 2, function(x){
        rep(max(na.omit(x)), nrow(matrix_posicion_columna))

      })

      for (k1 in 1:nrow(matrix_ordenamiento05)){
        for (k2 in 1:ncol(matrix_ordenamiento05)){
          if(!matrix_ordenamiento05[k1, k2]) matrix_posicion_columna[k1, k2] <- NA
        }
      }


      # Hacemos un cambio mas...
      # Y es que la posicion de columna de los titulos, tiene que ser la posicion
      # de columna de la tabla que lo acompania.
      #
      for(llave_fila in 1:(nrow(matrix_ordenamiento04)-1)){
        for(llave_columna in 1:ncol(matrix_ordenamiento04)){
          tipo_objeto <- matrix_ordenamiento04[llave_fila, llave_columna]

          if(!is.na(tipo_objeto)){
            if(tipo_objeto == "Title01") {
              matrix_posicion_columna[llave_fila, llave_columna] <- matrix_posicion_columna[llave_fila+1, llave_columna]
            }
          }
        }
      }

      # names(armado_box_completo) <- c("OrdenCelda", "Rellenado", "OrdenContenido", "ObjetoContenido",
      #                          "TipoContenido", "dt_contenido", "PosFilas", "PosColumna")

      # armado_box_completo <- do.call(cbind.data.frame, armado_box_completo)
      # armado_box_contenido <- na.omit(armado_box_completo)


    }


    # Salida
    {
      armado_especial <- list(
        as.vector(matrix_posicion_fila),
        as.vector(matrix_posicion_columna),
        as.vector(matrix_ordenamiento03),
        as.vector(matrix_ordenamiento04)
      )


      names(armado_especial) <- c("FilaInicio", "ColumnaInicio", "NombreObjeto", "TipoObjeto")
      armado_especial <- do.call(cbind.data.frame, armado_especial)
      armado_especial <- na.omit(armado_especial) # Fletamos a las celdas que no tienen objetos
      # armado_especial

      if(dt_byrow) cambio_orden <- order(armado_especial[,1])  else
        if(!dt_byrow) cambio_orden <- order(armado_especial[,2])

      armado_especial_ordenado <- armado_especial[cambio_orden, ]


    }

    # Return
    return(armado_especial_ordenado)

  }


  if(dt_byrow){

    out_primitivo <- Formato.ExcelOutput(contenido_general = contenido_general,
                                         dt_byrow = !dt_byrow,
                                         general_align = general_align,
                                         row_space = col_space,
                                         col_space = row_space,
                                         row_graph = col_graph,
                                         col_graph = row_graph,
                                         tellme_formato = tellme_formato)

    out_primitivo[c(1,2)] <- out_primitivo[c(2,1)]

    special_detection <- c("Title01", "Title02", "Title03")

    for(k in 1:nrow(out_primitivo)){

      if (sum(out_primitivo$TipoObjeto[k] == special_detection) > 0){
        if((k+1) <= nrow(out_primitivo)){
          out_primitivo$FilaInicio[k+1] <- out_primitivo$FilaInicio[k+1] + 1
          out_primitivo$ColumnaInicio[k+1] <- out_primitivo$ColumnaInicio[k+1] - 1
        }
      }
    }
    # colnames(out_primitivo)[c(1,2)] <- colnames(out_primitivo)[c(2,1)]

    return(out_primitivo)
  }












}
