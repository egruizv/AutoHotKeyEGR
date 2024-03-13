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
FilePathficheroExcelLectura := "C:\Users\egarciar\Documents\CAIXA\AutoHotKeyEGR\Datos_Prueba\Excel_Filas_Repertidas.xlsx"
FilePathficheroExcelSalida := "C:\Users\egarciar\Documents\CAIXA\AutoHotKeyEGR\Datos_Prueba\Excel_Filas_SinRepetir.xlsx"

/******************************************************************************
********************* LANZAMOS EL ROBOT ***************************************
*******************************************************************************
*/
Resultado := ""
Resultado := eliminar_filas_repetidas_Excel_EGR(FilePathficheroExcelLectura,FilePathficheroExcelSalida)
MsgBox, %Resultado% 


/******************************************************************************
********************* FUNCIONES ***********************************************
*******************************************************************************
*/