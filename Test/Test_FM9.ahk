#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%


/******************************************************************************
********************* LIBRETIAS **************************************
*******************************************************************************
*/

#Include ../LibreriasAhk/buscarFicheroRutaOrdenador.ahk


/******************************************************************************
********************* VARIABLES GLOBALES **************************************
*******************************************************************************
*/
; Unidad y archivo a buscar
CarpetaInicial := "C:\Users\egarciar\Documents\LibroDePetete"
ArchivoBuscar := "Querys_SQL_SERVER.sql"

/******************************************************************************
********************* LANZAMOS EL ROBOT ***************************************
*******************************************************************************
*/
; Iniciar la búsqueda

/******************************************************************************
********************* FUNCIONES ***********************************************
*******************************************************************************
*/

ArraySalida := BuscarArchivoEnC(CarpetaInicial,ArchivoBuscar)
longitud_ArraySalida := ArraySalida.length()
controlWhile1 := 1
While, controlWhile1 <= longitud_ArraySalida {
    OutputDebug, % ArraySalida[controlWhile1]
    controlWhile1++
}




