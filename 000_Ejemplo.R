# Remove all
remove(list = ls())

library("stringr")
library("openxlsx")
library("Rscience.base")
# library("Rscience.tools")
library("Rscience.Excel")


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

# Titles for Tables
{
  cantidad_digitos <- floor(log10(length(all_DataTable))) + 1

  new_title <- "title_DataTable"

  orden_DataTable <- str_pad(string = 1:length(all_DataTable),
                                 width = cantidad_digitos,
                                 side = "left", pad = "0")

  all_title_DataTable <- paste0(new_title, orden_DataTable)

  all_title_DataTable <- as.list(all_title_DataTable)
  names(all_title_DataTable) <- unlist(all_title_DataTable)


  all_title_DataTable <- SpecialFormat.List.Title(x =  all_title_DataTable,
                                                  selected_pos = NULL,
                                                  tellme_formato = FALSE)
}

# Graphs
{
  grafico01 <- "plot(mtcars)"
  grafico02 <- grafico01

  objetos_Graph_Sentence <- grep(ls(), pattern = "grafico", value = TRUE)

  armado <- paste0(objetos_Graph_Sentence, collapse = ", ")
  armado <- paste0("Hmisc::llist(", armado, ")")

  all_Graph_Sentence <- eval(parse(text = armado))

  all_Graph_Sentence <- SpecialFormat.List.Graph_Sentence(x = all_Graph_Sentence,
                                                      selected_pos = NULL,
                                                      tellme_formato = FALSE)
}

# Titles for graphs
{
  cantidad_digitos <- floor(log10(length(all_Graph_Sentence))) + 1

  new_title <- "title_Graph_Sentence"

  orden_Graph_Sentence <- str_pad(string = 1:length(all_Graph_Sentence),
                             width = cantidad_digitos,
                             side = "left", pad = "0")

  all_title_Graph_Sentence  <- paste0(new_title, orden_Graph_Sentence)

  all_title_Graph_Sentence <- as.list(all_title_Graph_Sentence)
  names(all_title_Graph_Sentence) <- unlist(all_title_Graph_Sentence)
  all_title_Graph_Sentence <- SpecialFormat.List.Title(x =  all_title_Graph_Sentence,
                                                       selected_pos = NULL,
                                                       tellme_formato = FALSE)
}


# Add Position Content - the_position_content
{




  the_position_content <- AddPositionContent(the_position_content = NULL)


  # # # Pack 1
  number_pack <- 1
  the_position_content <- AddPositionContent(the_position_content = the_position_content, ObjList = "all_title_DataTable",
                          ObjPos = 1, ObjName = NA, Pack = number_pack, PackPos = NA)

  the_position_content <- AddPositionContent(the_position_content = the_position_content, ObjList = "all_DataTable",
                          ObjPos = 1, ObjName = NA, Pack = number_pack, PackPos = NA)

  the_position_content <- AddPositionContent(the_position_content = the_position_content, ObjList = "all_title_DataTable",
                          ObjPos = 2, ObjName = NA, Pack = number_pack, PackPos = NA)

  the_position_content <- AddPositionContent(the_position_content = the_position_content, ObjList = "all_DataTable",
                          ObjPos = 2, ObjName = NA, Pack = number_pack, PackPos = NA)

  the_position_content <- AddPositionContent(the_position_content = the_position_content, ObjList = "all_title_Graph_Sentence",
                          ObjPos = 1, ObjName = NA, Pack = number_pack, PackPos = NA)

  the_position_content <- AddPositionContent(the_position_content = the_position_content, ObjList = "all_Graph_Sentence",
                          ObjPos = 1, ObjName = NA, Pack = number_pack, PackPos = NA)

  the_position_content <- AddPositionContent(the_position_content = the_position_content, ObjList = "all_title_DataTable",
                          ObjPos = 6, ObjName = NA, Pack = number_pack, PackPos = NA)

  the_position_content <- AddPositionContent(the_position_content = the_position_content, ObjList = "all_DataTable",
                          ObjPos = 6, ObjName = NA, Pack = number_pack, PackPos = NA)

  # # # Pack 2
  number_pack <- 2
  the_position_content <- AddPositionContent(the_position_content = the_position_content, ObjList = "all_DataTable", ObjPos = 1,
                          ObjName = "tabla01", Pack = number_pack, PackPos = NA)

  the_position_content <- AddPositionContent(the_position_content = the_position_content, ObjList = "all_DataTable", ObjPos = NA,
                          ObjName = "tabla01", Pack = number_pack, PackPos = NA)

  the_position_content <- AddPositionContent(the_position_content = the_position_content, ObjList = "all_DataTable", ObjPos = 1,
                          ObjName = NA, Pack = number_pack, PackPos = NA)







}


# Armado de contenido
the_builder_content <- BuilderContent(the_position_content = the_position_content)


the_exceloutput_position <- Formato.ExcelOutput(contenido_general = the_builder_content,
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

# length(sheets(wb))


Fill.Excel.Sheet(wb = wb,
                 sheet_name = "Tablas",
                 contenido_general = the_builder_content,
                 armado_especial = the_exceloutput_position,
                 start_row = 1,
                 start_col = 1)


Fill.Excel.Sheet(wb = wb,
                 sheet_name = "Tablas",
                 contenido_general = the_builder_content,
                 armado_especial = the_exceloutput_position,
                 start_row = 500,
                 start_col = 1)

Fill.Excel.Sheet(wb = wb,
                 sheet_name = "Tablas2",
                 contenido_general = the_builder_content,
                 armado_especial = the_exceloutput_position,
                 start_row = 1,
                 start_col = 1)

# save
saveWorkbook(wb, file =  complete_path, overwrite = T)

cat("Fin de ", complete_path, "\n")


openXL(complete_path)
