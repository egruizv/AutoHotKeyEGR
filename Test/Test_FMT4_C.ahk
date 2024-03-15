#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%
/******************************************************************************
********************* LIBRETIAS **************************************
*******************************************************************************
*/
#Include ../LibreriasAhk/ordenarExcelPorColumna.ahk

/******************************************************************************
********************* VARIABLES GLOBALES **************************************
*******************************************************************************
*/
FilePathficheroExcelLectura := "C:\Users\egarciar\Documents\CAIXA\AutoHotKeyEGR\Datos_Prueba\Excel_Filas_Desordenadas.xlsx"


/******************************************************************************
********************* LANZAMOS EL ROBOT ***************************************
*******************************************************************************
*/

numero_ColumnasFichero := numeroColumnasFichero(FilePathficheroExcelLectura)
;Array_TipoOrden := ["ASC", "DESC"] ; Array para ordenar de las dos formas
Array_TipoOrden := ["ASC"] ;Cogemos solo ASC para las pruebas
numero_Elementos_Array_TipoOrden := Array_TipoOrden.length() 
   controlWhile1 := 3 ;Cogemos la columna tres solo para las pruebas de ordenacion repetidas
   While, controlWhile1 <= numero_ColumnasFichero {
      Resultado := ""
      FilePathficheroExcelSalida := "C:\Users\egarciar\Documents\CAIXA\AutoHotKeyEGR\Datos_Prueba\Resultados\Test_FMT4\Excel_Filas_OrdenadasRepetidas"
      Ordenacion_Columna := controlWhile1
      ; Concatenar la variable Ordenacion_Columna con la cadena FilePathficheroExcelSalida
      FilePathficheroExcelSalida := FilePathficheroExcelSalida . "_" . Ordenacion_Columna 
      controlWhile2 := 1
      While, controlWhile2 <= numero_Elementos_Array_TipoOrden {
         TipoOrden := Array_TipoOrden[controlWhile2]
         FilePathficheroExcelSalidaA := FilePathficheroExcelSalida . "_" . TipoOrden ; MIRAR 
         FilePathficheroExcelSalidaB := FilePathficheroExcelSalidaA . ".xlsx"
         ordenar_Excel_Por_Columna_EGR(FilePathficheroExcelLectura,Ordenacion_Columna,FilePathficheroExcelSalidaB,TipoOrden)        
      controlWhile2++
      }
      break
      controlWhile1++
      ;MsgBox, %Reusltado%
   }
   MsgBox, Hemos Terminado 

/******************************************************************************
********************* FUNCIONES ***********************************************
*******************************************************************************
*/

/*
Funcion que recorre todo el excel y determina cuantas columnas tiene 
Luego lo usamos para generar un excel de cada una de las columnas 
Y asi tener ordenacion total un Excel ordenado por cada columna
*/
numeroColumnasFichero(FilePathficheroExcelLectura){
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
    ;Rows := xlRange.Rows.Count
    ;Rows := 11
    Cols := xlRange.Columns.Count 
    ; Cierra el archivo Excel
    xlBook.Close(false)
  
    ; Cierra la aplicación de Excel
    xlApp.Quit()
  
    ; Libera los objetos de Excel de la memoria
    xlRange := ""
    xlSheet := ""
    xlBook := ""
    xlApp := "" 
    Return Cols
 }