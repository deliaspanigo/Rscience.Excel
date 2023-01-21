Fill.Excel.Sheet <- function(wb, sheet_name, armado_especial, new_sheet = T, start_row = 1, start_col = 1){

  if(new_sheet) addWorksheet(wb, sheet_name)

  for(llave_general in 1:nrow(armado_especial)) {

    nombre_objeto <- armado_especial$NombreObjeto[llave_general]
    tipo_objeto <- armado_especial$TipoObjeto[llave_general]
    fila_inicio <- armado_especial$FilaInicio[llave_general]
    columna_inicio <- armado_especial$ColumnaInicio[llave_general]
    orden_contenido <- armado_especial$OrdenContenido[llave_general]


    if(!is.na(tipo_objeto)){

      rejunte <- purrr::map(contenido_general, nombre_objeto)[[orden_contenido]]
      contenido_aislado <- rejunte
      # if(length(rejunte) == 1) contenido_aislado <- rejunte[[1]] else
      #   contenido_aislado <- rejunte[-which(sapply(rejunte, is.null))][[1]]

      if(tipo_objeto == "DataTable"){
        writeDataTable(wb,
                       sheet = sheet_name,
                       x = contenido_aislado,
                       startCol = columna_inicio,
                       startRow = fila_inicio,
                       tableStyle = "TablestyleMedium6")
      } else

        if(tipo_objeto == "Title01"){

          # A los titulos les tengo que aplicar as.vector() dentro de la funcion
          # esto es por que cuando le doy el atributo de ExcelOutput, dejan de ser vectores.
          # Y writeData necesita vectores o matrices.
          writeData(wb,
                    sheet = sheet_name,
                    x = as.vector(contenido_aislado),
                    startCol = columna_inicio,
                    startRow = fila_inicio
          )
        } else

          if(tipo_objeto == "Graph_Sentence"){

            # grafico01
            eval(parse(text = as.vector(contenido_aislado)))
            insertPlot(wb, sheet = sheet_name,
                       width = 6,
                       height = 4,
                       startCol = columna_inicio,
                       startRow = fila_inicio)

            # A los titulos les tengo que aplicar as.vector() dentro de la funcion
            # esto es por que cuando le doy el atributo de ExcelOutput, dejan de ser vectores.
            # Y writeData necesita vectores o matrices.

          }

    }
  }

}
