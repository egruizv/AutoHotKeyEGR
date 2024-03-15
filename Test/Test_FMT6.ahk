#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%
/******************************************************************************
********************* LIBRETIAS **************************************
*******************************************************************************
*/
#Include ../LibreriasAhk/buscarDatoExcel.ahk

/******************************************************************************
********************* VARIABLES GLOBALES **************************************
*******************************************************************************
*/
FilePathficheroExcelLectura := "C:\Users\egarciar\Documents\CAIXA\AutoHotKeyEGR\Datos_Prueba\Excel_Buscar_elemento.xlsx"
DatoBuscado := "Luis"

/******************************************************************************
********************* LANZAMOS EL ROBOT ***************************************
*******************************************************************************
*/
MatrizPosiciones := []

MatrizPosiciones.InsertAt(1,[3,4])
MatrizPosiciones.InsertAt(2,[5,7])

;MatrizPosiciones := buscarDatoExcel(FilePathficheroExcelLectura,DatoBuscado)
;Pintamos MatrizPosiciones por consola usando AutoHotkey Debug
; Imprimir el contenido de la matriz
for index, value in MatrizPosiciones {
    ;MsgBox % "MatrizPosiciones[" . index . "] = [" . value[1] . ", " . value[2] . "]"    
    OutputDebug, % "MatrizPosiciones[" . index . "] = [" . value[1] . ", " . value[2] . "]"
}
/*
longitud_MatrizPosiciones := MatrizPosiciones.length()
controlWhile1 := 1
While, controlWhile1 <= longitud_MatrizPosiciones {
    valor := []
    valor :=  MatrizPosiciones[controlWhile1]
    longitud_valor := valor.length()
    controlWhile2 := 1
    While, controlWhile2 <= longitud_valor {
        Salida := valor[controlWhile2]       
        OutputDebug, "Posicion [%controlWhile1%,%controlWhile2%] = %Salida%"
        controlWhile2++
    }
    controlWhile1++
}
*/



