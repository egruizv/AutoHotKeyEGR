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
#Include ../LibreriasAhk/Funciones_Auxiliares.ahk

/******************************************************************************
********************* VARIABLES GLOBALES **************************************
*******************************************************************************
*/
; Unidad y archivo a buscar
CarpetaInicial := "C:\Users\egarciar\Documents"
ArchivoBuscar := "challenge.xlsx"
flagSoloRuta := 1 ; 1 = Solo Ruta, 0 = Ruta + nombre del archivo a buscar
;flagSoloRuta := 0 ; 1 = Solo Ruta, 0 = Ruta + nombre del archivo a buscar

/******************************************************************************
********************* LANZAMOS EL ROBOT ***************************************
*******************************************************************************
*/
; Iniciar la búsqueda

/******************************************************************************
********************* FUNCIONES ***********************************************
*******************************************************************************
*/
;Nota Se añade un flag, para determinar si solo queremos la ruta o la ruta + nombre del archivo a buscar
ArraySalida := BuscarArchivoEnC(CarpetaInicial,ArchivoBuscar,flagSoloRuta)
longitud_ArraySalida := ArraySalida.length()
controlWhile1 := 1
;ComparamosString y devolvemos las posiciones que son iguales
; Llamar a la función y almacenar el resultado en una variable
Repetidos := encontrarIndicesRepetidos(ArraySalida)
   
    ;Ahora creamos una Matriz_Final con sólo las filas no repetidas (Todas las filas menos "Repetidos")
    Matriz_Final := []
    Matriz_Final := filasSinRepetirMatriz(ArraySalida,Repetidos)

    longitud_Matriz_Final := Matriz_Final.length()
    controlWhile1 := 1
    While, controlWhile1 <= longitud_Matriz_Final {
        OutputDebug, % Matriz_Final[controlWhile1]
        controlWhile1++
    }
MsgBox,"Hemos Terminado" 




