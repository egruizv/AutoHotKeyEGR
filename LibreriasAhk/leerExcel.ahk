#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

/*
Version 1.0.0 Ernesto Garcia 12/03/2024
    leer_Excel_completo_EGR(FilePathficheroExcelLectura)     
*/

/*
Esta funcion leer_Excel_completo_EGR(FilePathficheroExcelLectura) hace lo siguiente:
Recibe la direccion de un FilePathficheroExcel que sera un fichero Excel a leeer
Abre el fichero, y recorre todas  las filas  metiendo los datos en una variable RowArray y a su vez en Matriz_Excel
Devuelve Matriz_Excel que es un array de Arrays donde cada Matriz_Excel[i] corresponde a una fila
*/
leer_Excel_completo_EGR(FilePathficheroExcelLectura){
    ; Crea un objeto de Excel
    xlApp := ComObjCreate("Excel.Application")
    xlApp.Visible := false  ; Para que no se abra Excel visible
  
    ; Abre el archivo Excel
    FilePath := FilePathficheroExcelLectura
    xlBook := xlApp.Workbooks.Open(FilePath)
  
    ; Selecciona la primera hoja del libro
    xlSheet := xlBook.Sheets(1)
  
    ; Obtén el rango de celdas con datos
    xlRange := xlSheet.UsedRange
  
    ; Obtiene el número de filas y columnas en el rango
    Rows := xlRange.Rows.Count
    ;Rows := 11
    Cols := xlRange.Columns.Count  
  
    Matriz_Excel := [] ; Resultado
    numero_columna := 1
    numero_fila := 1
    ; Recorre el Excel por filas
    While, numero_fila <= Rows {
        RowArray := []  ; Inicializa un array para cada fila
        var_aux := 1
        numero_columna := 1
            While, numero_columna <= Cols {
                CellValue := xlSheet.Cells(numero_fila, A_Index).Value
                if (CellValue != "") {                
                RowArray.InsertAt(var_aux, CellValue) 
                }Else
                {
                    Break ;Si la celda esta vacia salimos del bucle
                }
                var_aux++
                ;fix
                numero_columna++
                
            }  
        ; Matriz_Excel.InsertAt(numero_fila, RowArray) solo metemos datos si no son vacios
        if(RowArray[1]!=null ){
            Matriz_Excel.InsertAt(numero_fila, RowArray) 
        }
        numero_fila := numero_fila +1
    }    
    ; Cierra el archivo Excel
    xlBook.Close(false)
  
    ; Cierra la aplicación de Excel
    xlApp.Quit()
  
    ; Libera los objetos de Excel de la memoria
    xlRange := ""
    xlSheet := ""
    xlBook := ""
    xlApp := ""
    Return Matriz_Excel
  }


  obtenerDatoFilaColumna(FilePathficheroExcelLectura,Fila,Columna){
     ; Crea un objeto de Excel
     xlApp := ComObjCreate("Excel.Application")
     xlApp.Visible := false  ; Para que no se abra Excel visible
   
     ; Abre el archivo Excel
     FilePath := FilePathficheroExcelLectura
     xlBook := xlApp.Workbooks.Open(FilePath)
   
     ; Selecciona la primera hoja del libro
     xlSheet := xlBook.Sheets(1)
   
     ; Obtén el rango de celdas con datos
     xlRange := xlSheet.UsedRange
   
     ; Obtiene el número de filas y columnas en el rango
     Rows := xlRange.Rows.Count
     ;Rows := 11
     Cols := xlRange.Columns.Count  
   
     ValorCelda :=  "" ;Resultado
     numero_columna := Columna
     numero_fila := Fila
     CellValue := xlSheet.Cells(numero_fila, numero_columna).Value
     if (CellValue != "") {   
        ValorCelda := CellValue
     }

       ; Cierra el archivo Excel
    xlBook.Close(false)
  
    ; Cierra la aplicación de Excel
    xlApp.Quit()
  
    ; Libera los objetos de Excel de la memoria
    xlRange := ""
    xlSheet := ""
    xlBook := ""
    xlApp := ""
    Return ValorCelda
  }