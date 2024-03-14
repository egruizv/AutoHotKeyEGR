#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%
/******************************************************************************
********************* LIBRETIAS **************************************
*******************************************************************************
*/
#Include ../LibreriasAhk/eliminarFilasRepetidasExcel.ahk

/******************************************************************************
********************* VARIABLES GLOBALES **************************************
*******************************************************************************
*/
;FilePathficheroExcelLectura := "C:\Users\egarciar\Documents\CAIXA\AutoHotKeyEGR\Datos_Prueba\Excel_Filas_Repertidas.xlsx"
;FilePathficheroExcelSalida := "C:\Users\egarciar\Documents\CAIXA\AutoHotKeyEGR\Datos_Prueba\Test_FMT3\Excel_Filas_SinRepetir.xlsx"

FilePathficheroExcelLectura := "C:\Users\egarciar\Documents\CAIXA\AutoHotKeyEGR\Datos_Prueba\Excel_Filas_Repertidas2.xlsx"
FilePathficheroExcelSalida := "C:\Users\egarciar\Documents\CAIXA\AutoHotKeyEGR\Datos_Prueba\Resultados\Test_FMT3\Excel_Filas_SinRepetir"

/******************************************************************************
********************* LANZAMOS EL ROBOT ***************************************
*******************************************************************************
*/
Resultado := ""
ArrayColumnasComparar := [] ;Eliminamos la fila 8 (Prueba 1)
ArrayColumnasComparar := [1,2] ;Eliminamos la fila 5, 6,8 (Prueba 2)
ArrayColumnasComparar := [1] ;Eliminamos la fila 5 y 6,8 (Prueba 3)
/*
Aqui el FilePathficheroExcelSalida tiene concatenado el nombre de las columna que aplican
Si el  ArrayColumnasComparar := [] Incluimos All 
*/
    if(ArrayColumnasComparar.length() <=0){
        FilePathficheroExcelSalida := FilePathficheroExcelSalida . "_All"
    }else {
        for index, value in ArrayColumnasComparar {
            Columna_Comparar := value
            FilePathficheroExcelSalida := FilePathficheroExcelSalida . "_" . Columna_Comparar
        }
    
    }
FilePathficheroExcelSalida := FilePathficheroExcelSalida . ".xlsx"
Resultado := eliminar_filas_repetidas_Excel_EGR(FilePathficheroExcelLectura,FilePathficheroExcelSalida,ArrayColumnasComparar)
MsgBox, %Resultado% 


/******************************************************************************
********************* FUNCIONES ***********************************************
*******************************************************************************
*/