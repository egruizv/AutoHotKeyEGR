#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

/******************************************************************************
********************* LIBRETIAS **************************************
*******************************************************************************
*/

#Include ../LibreriasAhk/leerExcel.ahk


/******************************************************************************
********************* VARIABLES GLOBALES **************************************
*******************************************************************************
*/
FilePathficheroExcelLectura := "C:\Users\egarciar\Documents\CAIXA\AutoHotKeyEGR\Datos_Prueba\Excel_Prueba.xlsx"

/******************************************************************************
********************* LANZAMOS EL ROBOT ***************************************
*******************************************************************************
*/
Matriz_Excel := [] ; Resultado
Matriz_Excel := leer_Excel_completo_EGR(FilePathficheroExcelLectura)
MsgBox, Ha terminado todo correcto

/******************************************************************************
********************* FUNCIONES ***********************************************
*******************************************************************************
*/




