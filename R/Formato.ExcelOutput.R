
Formato.ExcelOutput <- function(contenido_general,
                                dt_byrow = F,
                                vertical_space = 4,
                                horizontal_space = 5,
                                vertical_graph = 20,
                                horizontal_graph = 7,
                                tellme_formato = F){


  Agregar.NA <- function(matriz_original, matriz_logica){

    matrix_modificada <- matriz_original

    # Ahora... A los lugares de la matrix vacios, le colocamos NA
    for (k1 in 1:nrow(matriz_logica)){
      for (k2 in 1:ncol(matriz_logica)){
        if(!matriz_logica[k1, k2]){
          matrix_modificada[k1, k2] <- NA
        }
      }
    }

    return(matrix_modificada)
  }


  formatos_internos <- c("Title01",
                         "Title02",
                         "Title03",
                         "DataTable", "DataTable_Sentence", "DataTable_Path", "DataTable_Space",
                         "Table",     "Table_Sentence",     "Table_Path",     "Table_Space",
                         "Text",      "Text_Sentence",      "Text_Path",      "Text_Space",
                         "Graph_Object",     "Graph_Sentence",     "Graph_Path",     "Graph_Space",
                         "Fusion_Cell" # nuevo!
  )

  categorias_title <- c("Title01", "Title02", "Title03")
  categorias_table <- c("DataTable", "Table", "Graph_Sentence")

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
  # nvertical_space <- rep(NA, nrow(combinaciones))
  # nhorizontal_space <- rep(NA, nrow(combinaciones))
  # combinaciones <- cbind.data.frame(orden_combinaciones, combinaciones, nvertical_space, nhorizontal_space)
  # write.xlsx(x = combinaciones, file = "CombinacionesOUT.xlsx")

  # Cargamos el espaciado
  # Luego esto sera parte del paquede directamente
  # matrix_espaciados <- read.xlsx(xlsxFile = "CombinacionesOUT.xlsx", sheet = 1)

  # Solo salida de formato que se usa...
  if(tellme_formato) return(formatos_internos)


  # Si vas por columnas
  if (!dt_byrow){

      # Objetos Generales
      {

        # Nombre general de ubicacion en fila y columna
        nombre_ubicacion <- c("Row", "Col")
        detalle_relleno <- c("Por filas", "Por columnas")

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
        formato_relleno <- nombre_ubicacion[numeric_order_relleno]
        forma_rellenado <- rep(detalle_relleno[numeric_order_relleno], total_len_box)

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

        # La posicion en filas y columnas del contenido
        matrix_ordenamiento06 <- namel(names_vector = nombre_ubicacion, initial_value = NULL)
        matrix_ordenamiento06[["Row"]] <- matrix(secuencia_filas,
                                                 nrow = nrow(matrix_ordenamiento01),
                                                 ncol = ncol(matrix_ordenamiento01),
                                                 byrow = dt_byrow)

        matrix_ordenamiento06[["Col"]] <- matrix(secuencia_columnas,
                                                 nrow= nrow(matrix_ordenamiento01),
                                                 ncol= ncol(matrix_ordenamiento01),
                                                 byrow = !dt_byrow)


        matrix_ordenamiento06[["Row"]] <- Agregar.NA(matriz_original = matrix_ordenamiento06[["Row"]],
                                                     matriz_logica = matrix_ordenamiento05)

        matrix_ordenamiento06[["Col"]] <- Agregar.NA(matriz_original = matrix_ordenamiento06[["Col"]],
                                                     matriz_logica = matrix_ordenamiento05)

        ###########
        caso01 <- c("Title01", "Title02", "Title03", "Text")
        caso02 <- c("DataTable", "Table")
        caso03 <- c("DataTable_Path", "Table_Path", "Text_Path")
        caso04 <- c("DataTable_Sentence", "Table_Sentence", "Text_Sentence")
        caso05 <- c("Graph_Object", "Graph_Path", "Graph_Sentence", "Graph_Space")

        # Aqui esta la cantidad de filas y de columnas de cada objeto
        matrix_ordenamiento07 <- namel(names_vector = nombre_ubicacion, initial_value = NULL)

        matrix_ordenamiento07[["Row"]] <- matrix(NA, nrow=nrow(matrix_ordenamiento06[["Row"]]),
                                                 ncol = ncol(matrix_ordenamiento06[["Row"]]))

        matrix_ordenamiento07[["Col"]] <- matrix(NA, nrow=nrow(matrix_ordenamiento06[["Col"]]),
                                                 ncol = ncol(matrix_ordenamiento06[["Col"]]))

        for(columna_elegida in 1:ncol(matrix_ordenamiento07[["Row"]])){
          for(fila_elegida in 1:nrow(matrix_ordenamiento07[["Row"]])){

            nombre_objeto <- matrix_ordenamiento03[fila_elegida, columna_elegida]
            conteo_filas <- c()
            conteo_columnas <- c()

            if(is.na(nombre_objeto)){
              conteo_filas[1] <- NA
              conteo_columnas[1] <- NA
            }  else
              if(!is.na(nombre_objeto)){

                formato_exceloutput <- matrix_ordenamiento04[fila_elegida, columna_elegida]

                # rejunte <- purrr::map(contenido_general, nombre_objeto)
                # Modificado, para que si hay en diferentes filas objetos con el
                # mismo nombre, no haya problema.
                rejunte <- purrr::map(contenido_general, nombre_objeto)[[columna_elegida]]

                # if(length(rejunte) == 1) contenido_aislado <- rejunte[[1]] else
                #   if(length(rejunte) > 1) contenido_aislado <- rejunte[-which(sapply(rejunte, is.null))][[1]]
                # Lo silencie, por que ya no haria falta...
                contenido_aislado <- rejunte

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

                # Para las filas
                if(dt02) conteo_filas[1] <-  nrow(contenido_aislado) + 1 else # DataFrame y Tablas - nrow()
                  if(dt01 | dt03 | dt04) conteo_filas[1] <-  length(contenido_aislado) else
                    if(dt05) conteo_filas[1] <- vertical_graph else # "Graph" o "Graph_Space"
                      if(dt06 | dt07) conteo_filas[1] <- 5 else # "DataTable_Space"
                        if(dt08) conteo_filas[1] <- 1 # Text_Space

                # Para las columnas
                if(dt02) conteo_columnas[1] <-  ncol(contenido_aislado) + 1 else # DataFrame y Tablas - nrow()
                  if(dt01 | dt03 | dt04) conteo_columnas[1] <-  length(contenido_aislado) else
                    if(dt05) conteo_columnas[1] <- horizontal_graph else # "Graph" o "Graph_Space"
                      if(dt06 | dt07) conteo_columnas[1] <- 5 else # "DataTable_Space"
                        if(dt08) conteo_columnas[1] <- 1 # Text_Space

                matrix_ordenamiento07[["Row"]][fila_elegida, columna_elegida] <- conteo_filas
                matrix_ordenamiento07[["Col"]][fila_elegida, columna_elegida] <- conteo_columnas
          }
          }
        }


        matrix_ordenamiento07[["Row"]] <- Agregar.NA(matriz_original = matrix_ordenamiento07[["Row"]],
                                                     matriz_logica = matrix_ordenamiento05)

        matrix_ordenamiento07[["Col"]] <- Agregar.NA(matriz_original = matrix_ordenamiento07[["Col"]],
                                                     matriz_logica = matrix_ordenamiento05)





        # Suma de numero de filas y suma de numero de columnas
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
        # if(is.vector(matrix_ordenamiento08[["Row"]])) dim(matrix_ordenamiento08[["Row"]]) <- c(length(matrix_ordenamiento08[["Row"]]), 1)


        matrix_ordenamiento08[["Col"]] <- apply(matrix_ordenamiento07[["Col"]], 2, function(x){

          # vector_aislado <- matrix_ordenamiento07[["Col"]][x,]
          vector_aislado <- x
          vector_aislado[is.na(vector_aislado)] <- 0
          suma_acumulada <- cumsum(vector_aislado)
          suma_acumulada
        })


        matrix_ordenamiento08[["Row"]] <- Agregar.NA(matriz_original = matrix_ordenamiento08[["Row"]],
                                                     matriz_logica = matrix_ordenamiento05)

        matrix_ordenamiento08[["Col"]] <- Agregar.NA(matriz_original = matrix_ordenamiento08[["Col"]],
                                                     matriz_logica = matrix_ordenamiento05)





      }



    # Acomodamos todo un poco
    {

      ## Posicion en filas
      {
        matrix_cantidad_fila <- matrix_ordenamiento07[["Row"]]

        # Hasta aca, no tenemos la fila de inicio de cada objeto, sino la cantidad de filas acumuladas.
        # La posicion de inicio de cada objeto, sin contar espacios, es la fila siguiente a la cual termina
        # la tabla anterior. Por eso, vamos a hacer dos cosas:
        # Cosa 1) Sumarte a todos 1 fila
        # Con esto tendremos la fila de inicio de cada objeto, excepto del primero.
        # Cosa 2) Agregar una fila completa de 1, con esto tenemos la fila de inicio del primer objeto
        # Cosa 3) Retirar la info final, ya que tenemos hemos armado al ultimo
        # seria la fila de inicio de la tabla que sigue a la tabla final, y eso no puede ser
        # ya que no hay tabla

        # Tomamos la sumatoria de las filas.
        matrix_posicion_fila <- matrix_ordenamiento08[["Row"]]
        matrix_posicion_fila <- matrix_posicion_fila + 1 # Essto es a donde empieza cada uno, excepto el primero
        matrix_posicion_fila <- rbind(rep(1, ncol(matrix_posicion_fila)), matrix_posicion_fila)
        matrix_posicion_fila <- matrix_posicion_fila[-nrow(matrix_posicion_fila),]
        if(is.vector(matrix_posicion_fila)) dim(matrix_posicion_fila) <- c(length(matrix_posicion_fila), 1)
        matrix_posicion_fila <- Agregar.NA(matriz_original = matrix_posicion_fila,
                                           matriz_logica = matrix_ordenamiento05)

        inicio_fila <- matrix_posicion_fila
        fin_fila <- matrix_posicion_fila + matrix_cantidad_fila - 1


        # Espaciado Vertical
        espaciado_vertical <- matrix(0,
                                     nrow=nrow(matrix_posicion_fila),
                                     ncol=ncol(matrix_posicion_fila))


        for(k_columna in 1:ncol(matrix_ordenamiento04)){
          for(k_fila in 2:nrow(matrix_ordenamiento04)){

            anterior_tipo <- matrix_ordenamiento04[k_fila-1, k_columna]
            este_tipo <- matrix_ordenamiento04[k_fila, k_columna]
            el_espaciado <- c()


            if(!is.na(anterior_tipo) && !is.na(este_tipo)){
              dt_titulo_anterior <- sum(categorias_title == anterior_tipo) > 0
              dt_titulo_este <- sum(categorias_title == este_tipo) > 0


              if(dt_titulo_anterior && !dt_titulo_este) el_espaciado <- 0 else
                if(!dt_titulo_anterior && !dt_titulo_este) el_espaciado <- vertical_space else
                  if(!dt_titulo_anterior && dt_titulo_este) el_espaciado <- vertical_space else
                    if(dt_titulo_anterior && dt_titulo_este) el_espaciado <- vertical_space
            } else el_espaciado <- 0

            posiciones_fila <- k_fila:nrow(espaciado_vertical)
            espaciado_vertical[posiciones_fila, k_columna] <- espaciado_vertical[posiciones_fila, k_columna] + el_espaciado

          }
        }


        # Posicion Final por fila
        matrix_posicion_fila <- matrix_posicion_fila + espaciado_vertical



      }

      # Posicion Columna
      {
        # Parte 01 - OK!!!!
        matrix_cantidad_columna <- matrix_ordenamiento07[["Col"]]

        inicio_columna <- apply(matrix_cantidad_columna, 2, function(x){
            x[is.na(x)] <- 0
            rep(max(x), length(x))
          })

          # inicio_fila  <- Agregar.NA(matriz_original = inicio_fila ,
          #                            matriz_logica = matrix_ordenamiento05)


          # Toda la primer fila empieza en 1
          inicio_columna <- cbind(rep(1, nrow(inicio_columna)), inicio_columna)

          inicio_columna <- inicio_columna[,-ncol(inicio_columna)]

          inicio_columna <- Agregar.NA(matriz_original = inicio_columna ,
                                     matriz_logica = matrix_ordenamiento05)

          # El minimo de inicio columna como comienzo de cada objeto
          # es el numero de columna de contenido en el que esta
          for(k_columna in 1:ncol(inicio_columna)){
            dt_cambiaso <- inicio_columna[,k_columna ] < k_columna
            dt_cambiaso[is.na(dt_cambiaso)] <- FALSE
            inicio_columna[dt_cambiaso, k_columna] <- k_columna
          }

          for(k_fila in 1:nrow(inicio_columna)){
            fila_seleccionada <- inicio_columna[k_fila,]
            fila_seleccionada[is.na(fila_seleccionada)] <- 0
            vector_suma <- cumsum(fila_seleccionada)
            inicio_columna[k_fila,] <- vector_suma
          }


          inicio_columna <- Agregar.NA(matriz_original = inicio_columna ,
                                       matriz_logica = matrix_ordenamiento05)

          # Le sumamos 1 columna para marcar donde inicia en vez de
          # donde termina
          inicio_columna[,2:ncol(inicio_columna)] <- inicio_columna[,2:ncol(inicio_columna)] + 1

          # Hacemos la correccion para que toda la columna de contenido
          # comience en la misma columna Excel
          for(k_columna in 1:ncol(inicio_columna)){
            maximo_columna <- max(na.omit(inicio_columna[,k_columna ]))
            dt_cambiaso <- inicio_columna[,k_columna ] < maximo_columna
            dt_cambiaso[is.na(dt_cambiaso)] <- FALSE
            inicio_columna[dt_cambiaso, k_columna] <- maximo_columna
          }


          fin_columna <- inicio_columna + matrix_cantidad_columna - 1

          matrix_posicion_columna <- inicio_columna



          # Espaciado Vertical
          espaciado_horizontal <- matrix(NA,
                                         nrow=nrow(matrix_posicion_columna),
                                         ncol=ncol(matrix_posicion_columna))

          for(k_columna in 1:ncol(espaciado_horizontal)){
            espaciado_horizontal[,k_columna] <- horizontal_space*(k_columna-1)
          }

          # Posicion Final de Columnas
          matrix_posicion_columna <- matrix_posicion_columna + espaciado_horizontal







      }





    }


  }

  # Si lo quiere por filas...
  if (dt_byrow){

    # Objetos Generales
    {

      # Nombre general de ubicacion en fila y columna
      nombre_ubicacion <- c("Row", "Col")
      detalle_relleno <- c("Por filas", "Por columnas")

      # Nombre de los objetos ingresados
      names_contenido <- names(contenido_general)
      len_contenido <- length(contenido_general)
      len_interno <- unlist(lapply(contenido_general, length))
      max_len_interno <- max(len_interno)
      total_len_contenido <- sum(len_interno)

      # Box General
      dim_box <- c(max_len_interno, len_contenido)
      dim_box <- dim_box[c(2,1)]
      names(dim_box) <- nombre_ubicacion



      # Total de objetos ingresados
      total_len_box <- as.vector(dim_box[1]*dim_box[2])
      vector_order_box <- 1:total_len_box

      numeric_order_relleno <- 1
      formato_relleno <- nombre_ubicacion[numeric_order_relleno]
      forma_rellenado <- rep(detalle_relleno[numeric_order_relleno], total_len_box)

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
      matrix_ordenamiento02 <- do.call(rbind, matrix_ordenamiento02)


      # Nombre de cada elemento
      matrix_ordenamiento03 <- sapply(1:len_contenido, function(x){

        rejunte_info <-        sapply(1:len_interno[x], function(y){
          names(contenido_general[[x]])[y]
        })

        generico_vacio <- rep(NA, max_len_interno)
        generico_vacio[1:len_interno[x]] <- rejunte_info
        generico_vacio

      }, simplify = F, USE.NAMES = F)
      matrix_ordenamiento03 <- do.call(rbind, matrix_ordenamiento03)


      # Formato de cada elemento
      matrix_ordenamiento04 <- sapply(1:len_contenido, function(x){

        rejunte_info <-        sapply(1:len_interno[x], function(y){
          attributes(contenido_general[[x]][[y]])$ExcelOutput
        })

        generico_vacio <- rep(NA, max_len_interno)
        generico_vacio[1:len_interno[x]] <- rejunte_info
        generico_vacio

      }, simplify = F, USE.NAMES = F)
      matrix_ordenamiento04 <- do.call(rbind, matrix_ordenamiento04)




      # Se si es vacio o no cada celda
      matrix_ordenamiento05 <- !is.na(matrix_ordenamiento03)


      # Posicion en filas y columnas de cada celda
      secuencia_columnas <- rep(1:ncol(matrix_ordenamiento01), nrow(matrix_ordenamiento01))
      secuencia_filas <- rep(1:nrow(matrix_ordenamiento01), each = ncol(matrix_ordenamiento01))

      # La posicion en filas y columnas del contenido
      matrix_ordenamiento06 <- namel(names_vector = nombre_ubicacion, initial_value = NULL)
      matrix_ordenamiento06[["Row"]] <- matrix(secuencia_filas,
                                               nrow = nrow(matrix_ordenamiento01),
                                               ncol = ncol(matrix_ordenamiento01),
                                               byrow = dt_byrow)

      matrix_ordenamiento06[["Col"]] <- matrix(secuencia_columnas,
                                               nrow= nrow(matrix_ordenamiento01),
                                               ncol= ncol(matrix_ordenamiento01),
                                               byrow = dt_byrow)


      matrix_ordenamiento06[["Row"]] <- Agregar.NA(matriz_original = matrix_ordenamiento06[["Row"]],
                                                   matriz_logica = matrix_ordenamiento05)

      matrix_ordenamiento06[["Col"]] <- Agregar.NA(matriz_original = matrix_ordenamiento06[["Col"]],
                                                   matriz_logica = matrix_ordenamiento05)

      ###########
      caso01 <- c("Title01", "Title02", "Title03", "Text")
      caso02 <- c("DataTable", "Table")
      caso03 <- c("DataTable_Path", "Table_Path", "Text_Path")
      caso04 <- c("DataTable_Sentence", "Table_Sentence", "Text_Sentence")
      caso05 <- c("Graph_Object", "Graph_Path", "Graph_Sentence", "Graph_Space")

      # Aqui esta la cantidad de filas y de columnas de cada objeto
      matrix_ordenamiento07 <- namel(names_vector = nombre_ubicacion, initial_value = NULL)

      matrix_ordenamiento07[["Row"]] <- matrix(NA, nrow=nrow(matrix_ordenamiento06[["Row"]]),
                                               ncol = ncol(matrix_ordenamiento06[["Row"]]))

      matrix_ordenamiento07[["Col"]] <- matrix(NA, nrow=nrow(matrix_ordenamiento06[["Col"]]),
                                               ncol = ncol(matrix_ordenamiento06[["Col"]]))

      for(columna_elegida in 1:ncol(matrix_ordenamiento07[["Row"]])){
        for(fila_elegida in 1:nrow(matrix_ordenamiento07[["Row"]])){

          nombre_objeto <- matrix_ordenamiento03[fila_elegida, columna_elegida]
          conteo_filas <- c()
          conteo_columnas <- c()

          if(is.na(nombre_objeto)){
            conteo_filas[1] <- NA
            conteo_columnas[1] <- NA
          }  else
            if(!is.na(nombre_objeto)){

              formato_exceloutput <- matrix_ordenamiento04[fila_elegida, columna_elegida]

              rejunte <- purrr::map(contenido_general, nombre_objeto)

              if(length(rejunte) == 1) contenido_aislado <- rejunte[[1]] else
                if(length(rejunte) > 1) contenido_aislado <- rejunte[-which(sapply(rejunte, is.null))][[1]]

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

              # Para las filas
              if(dt02) conteo_filas[1] <-  nrow(contenido_aislado) + 1 else # DataFrame y Tablas - nrow()
                if(dt01 | dt03 | dt04) conteo_filas[1] <-  length(contenido_aislado) else
                  if(dt05) conteo_filas[1] <- vertical_graph else # "Graph" o "Graph_Space"
                    if(dt06 | dt07) conteo_filas[1] <- 5 else # "DataTable_Space"
                      if(dt08) conteo_filas[1] <- 1 # Text_Space

              # Para las columnas
              if(dt02) conteo_columnas[1] <-  ncol(contenido_aislado) + 1 else # DataFrame y Tablas - nrow()
                if(dt01 | dt03 | dt04) conteo_columnas[1] <-  length(contenido_aislado) else
                  if(dt05) conteo_columnas[1] <- horizontal_graph else # "Graph" o "Graph_Space"
                    if(dt06 | dt07) conteo_columnas[1] <- 5 else # "DataTable_Space"
                      if(dt08) conteo_columnas[1] <- 1 # Text_Space

              matrix_ordenamiento07[["Row"]][fila_elegida, columna_elegida] <- conteo_filas
              matrix_ordenamiento07[["Col"]][fila_elegida, columna_elegida] <- conteo_columnas
            }
        }
      }


      matrix_ordenamiento07[["Row"]] <- Agregar.NA(matriz_original = matrix_ordenamiento07[["Row"]],
                                                   matriz_logica = matrix_ordenamiento05)

      matrix_ordenamiento07[["Col"]] <- Agregar.NA(matriz_original = matrix_ordenamiento07[["Col"]],
                                                   matriz_logica = matrix_ordenamiento05)





      # Suma de numero de filas y suma de numero de columnas
      matrix_ordenamiento08 <- namel(names_vector = nombre_ubicacion, initial_value = NULL)
      matrix_ordenamiento08[["Row"]] <- sapply(1:nrow(matrix_ordenamiento07[["Row"]]), function(x){

        vector_aislado <- matrix_ordenamiento07[["Row"]][x,]
        vector_aislado[is.na(vector_aislado)] <- 0
        suma_acumulada <- cumsum(vector_aislado)
        suma_acumulada
        # diferencia_datos <- nrow(matrix_ordenamiento07[["Row"]]) - length(suma_acumulada)
        # salida <- c(suma_acumulada, rep(NA, diferencia_datos))
        #salida
      })
      matrix_ordenamiento08[["Row"]] <- t(matrix_ordenamiento08[["Row"]])

      # if(is.vector(matrix_ordenamiento08[["Row"]])) dim(matrix_ordenamiento08[["Row"]]) <- c(length(matrix_ordenamiento08[["Row"]]), 1)


      matrix_ordenamiento08[["Col"]] <- apply(matrix_ordenamiento07[["Col"]], 1, function(x){

        # vector_aislado <- matrix_ordenamiento07[["Col"]][x,]
        vector_aislado <- x
        vector_aislado[is.na(vector_aislado)] <- 0
        suma_acumulada <- cumsum(vector_aislado)
        suma_acumulada
      })
      matrix_ordenamiento08[["Col"]] <- t(matrix_ordenamiento08[["Col"]])

      matrix_ordenamiento08[["Row"]] <- Agregar.NA(matriz_original = matrix_ordenamiento08[["Row"]],
                                                   matriz_logica = matrix_ordenamiento05)

      matrix_ordenamiento08[["Col"]] <- Agregar.NA(matriz_original = matrix_ordenamiento08[["Col"]],
                                                   matriz_logica = matrix_ordenamiento05)






    }


    # Si va por columnas...


    # Acomodamos todo un poco
    {

      # Posicion Columna
      {
        # Cantidad de columnas
        matrix_cantidad_columna <- matrix_ordenamiento07[["Col"]]

        # Tomamos la sumatoria de las columnas.
        matrix_posicion_columna <- matrix_ordenamiento08[["Col"]]
        matrix_posicion_columna <- matrix_posicion_columna + 1 # Essto es a donde empieza cada uno, excepto el primero
        matrix_posicion_columna <- cbind(rep(1, nrow(matrix_posicion_columna)), matrix_posicion_columna)
        matrix_posicion_columna <- matrix_posicion_columna[,-ncol(matrix_posicion_columna)]
        if(is.vector(matrix_posicion_columna)) dim(matrix_posicion_columna) <- c(1, length(matrix_posicion_columna))
        matrix_posicion_columna <- Agregar.NA(matriz_original = matrix_posicion_columna,
                                           matriz_logica = matrix_ordenamiento05)

        inicio_columna <- matrix_posicion_columna
        fin_columna <- matrix_posicion_columna + matrix_cantidad_columna - 1






      }

      ## Posicion en filas
      {
        # Parte 01 - OK!!!!
        matrix_cantidad_fila <- matrix_ordenamiento07[["Row"]]


        inicio_fila <- apply(matrix_cantidad_fila, 2, function(x){
          x[is.na(x)] <- 0
          cumsum(x)
        })

        # inicio_fila  <- Agregar.NA(matriz_original = inicio_fila ,
        #                            matriz_logica = matrix_ordenamiento05)


        # Toda la primer fila empieza en 1
        inicio_fila <- rbind(rep(1, ncol(inicio_fila)), inicio_fila)

        inicio_fila <- inicio_fila[-nrow(inicio_fila),]

        inicio_fila  <- Agregar.NA(matriz_original = inicio_fila ,
                                           matriz_logica = matrix_ordenamiento05)

        # El minimo de inicio fila como comienzo de cada objeto
        # es el numero de fila de contenido en el que esta
        for(k_fila in 1:nrow(inicio_fila)){
          dt_cambiaso <- inicio_fila[k_fila, ] < k_fila
          dt_cambiaso[is.na(dt_cambiaso)] <- FALSE
          inicio_fila[k_fila, dt_cambiaso] <- k_fila
        }

        # Hacemos la correccion para que toda la fila de contenido
        # comience en la misma fila Excel
        for(k_fila in 1:nrow(inicio_fila)){
          maximo_fila <- max(na.omit(inicio_fila[k_fila, ]))
          dt_cambiaso <- inicio_fila[k_fila, ] < maximo_fila
          dt_cambiaso[is.na(dt_cambiaso)] <- FALSE
          inicio_fila[k_fila, dt_cambiaso] <- maximo_fila
        }


        fin_fila <- inicio_fila + matrix_cantidad_fila - 1

        matrix_posicion_fila <- inicio_fila









    }




      # Espaciado Horizontal y Vertiral
      {
        # Espaciado horizontal
        espaciado_horizontal <- matrix(0,
                                     nrow=nrow(matrix_posicion_columna),
                                     ncol=ncol(matrix_posicion_columna))

        for(k_fila in 1:nrow(matrix_ordenamiento04)){
          for(k_columna in 2:ncol(matrix_ordenamiento04)){


            anterior_tipo <- matrix_ordenamiento04[k_fila, k_columna-1]
            este_tipo <- matrix_ordenamiento04[k_fila, k_columna]
            el_espaciado <- c()


            if(!is.na(anterior_tipo) && !is.na(este_tipo)){
              dt_titulo_anterior <- sum(categorias_title == anterior_tipo) > 0
              dt_titulo_este <- sum(categorias_title == este_tipo) > 0


              if(dt_titulo_anterior && !dt_titulo_este) el_espaciado <- 0 else
                if(!dt_titulo_anterior && !dt_titulo_este) el_espaciado <- horizontal_space else
                  if(!dt_titulo_anterior && dt_titulo_este) el_espaciado <- horizontal_space else
                    if(dt_titulo_anterior && dt_titulo_este) el_espaciado <- horizontal_space
            } else el_espaciado <- 0

            posiciones_columna <- k_columna:ncol(espaciado_horizontal)
            espaciado_horizontal[k_fila, posiciones_columna] <- espaciado_horizontal[k_fila, posiciones_columna] + el_espaciado

          }
        }

        matrix_posicion_columna <- matrix_posicion_columna + espaciado_horizontal


        # Espaciado Vertital
        espaciado_vertical <- matrix(NA,
                                       nrow=nrow(matrix_posicion_columna),
                                       ncol=ncol(matrix_posicion_columna))
        for(k_fila in 1:nrow(espaciado_vertical)){
          espaciado_vertical[k_fila,] <- vertical_space*(k_fila-1)
        }

        matrix_posicion_fila <- matrix_posicion_fila + espaciado_vertical



      }

      # Reordenamiento especial
      {


        for(k_fila in 1:nrow(matrix_posicion_fila)){
          for(k_columna in 2:ncol(matrix_posicion_fila)){


            anterior_tipo <- matrix_ordenamiento04[k_fila, k_columna-1]
            este_tipo <- matrix_ordenamiento04[k_fila, k_columna]



            if(!is.na(anterior_tipo) && !is.na(este_tipo)){
              dt_titulo_anterior <- sum(categorias_title == anterior_tipo) > 0
              # dt_tabla <- sum(categorias_table == este_tipo) > 0
              dt_tabla <- sum(categorias_title != este_tipo) > 0
              dt_general <- (dt_titulo_anterior + dt_tabla) == 2

              if(dt_general){

                # Calculamos la nueva posicion de fila y columna
                nueva_fila <- matrix_posicion_fila[k_fila, k_columna] + 1
                nueva_columna <- matrix_posicion_columna[k_fila, k_columna] - 1

                # Lo asignamos
                matrix_posicion_fila[k_fila, k_columna] <- nueva_fila
                matrix_posicion_columna[k_fila, k_columna] <- nueva_columna


              }


          }
        }
        }



      }







}




}



  # Lo ultimo
  {

    MaxBoxRow <- max(na.omit(as.vector(fin_fila))) + vertical_space
    MaxBoxCol <- max(na.omit(as.vector(fin_columna))) + vertical_space

    vector_MBR <- rep(MaxBoxRow, total_len_box)
    vector_MBC <- rep(MaxBoxCol, total_len_box)

    vector_orden <- as.vector(matrix_ordenamiento06[[numeric_order_relleno]])
    cantidad_digitos <- floor(log10(max(na.omit(vector_orden)))) + 1
    if(cantidad_digitos < 2) cantidad_digitos <- 2

    vector_orden_mod <- str_pad(string = vector_orden,
                                width = cantidad_digitos,
                                side = "left", pad = "0")

    vector_pos_contenido <- paste0(nombre_ubicacion[numeric_order_relleno], vector_orden_mod)
    armado_especial <- list(
      as.vector(matrix_posicion_fila),
      as.vector(matrix_posicion_columna),
      as.vector(matrix_ordenamiento03),
      as.vector(matrix_ordenamiento04),
      vector_orden,
      vector_pos_contenido,
      vector_MBR,
      vector_MBC

    )


    names(armado_especial) <- c("FilaInicio", "ColumnaInicio",
                                "NombreObjeto", "TipoObjeto",
                                "OrdenContenido", "Contenido",
                                "MaxRowBox", "MaxColBox")
    armado_especial <- do.call(cbind.data.frame, armado_especial)
    armado_especial <- na.omit(armado_especial) # Fletamos a las celdas que no tienen objetos
    # armado_especial


    cambio_orden <- order(armado_especial$Contenido)

    armado_especial_ordenado <- armado_especial[cambio_orden, ]
    rownames(armado_especial_ordenado) <- c(1:nrow(armado_especial_ordenado))

  }

  # Return
  return(armado_especial_ordenado)

}


