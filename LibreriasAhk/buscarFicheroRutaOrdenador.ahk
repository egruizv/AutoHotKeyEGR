#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

#Include ../LibreriasAhk/Funciones_Auxiliares.ahk


; Función para buscar el archivo en todas las carpetas y subcarpetas
BuscarArchivoEnC(carpetas, archivo,flagSoloRuta) {
    ;Creo un String con las rutas de todos los ficheros que hay en la carpeta
    StringAllFicheros := CreateString(carpetas)
    ; Recorro el StringAllFicheros y coloco todos los datos en una MatrizAuxiliar
    Matriz_Auxiliar := StrSplit(StringAllFicheros, ";")
    
    ;localizo donde esta el archivo, recorro MatrizAuxiliar y veo si archivo esta en algun MatrizAuxiliar[i]
    longitud_Matriz_Auxiliar := Matriz_Auxiliar.length()
    controlWhile1 := 1
    indiceMatrizSalida := 1    
    ArraySalida := []
    ArraySalidaAuxiliar := []
    indiceArraySalida := 1
    While, controlWhile1 <= longitud_Matriz_Auxiliar {
        ;Si el archivo esta en  MatrizAuxiliar[i] incluyo la direccion en ArraySalida[]
        ; Verificar si la subcadena está dentro de la cadena completa
        CadenaCompleta :=  Matriz_Auxiliar[controlWhile1]
        Subcadena := archivo
        if InStr(CadenaCompleta, Subcadena)
        {
            ArraySalidaAuxiliar.InsertAt(indiceArraySalida, Matriz_Auxiliar[controlWhile1])
            indiceArraySalida++
        }
        controlWhile1++
    }
    ; Llamamos a la funcion quitarCaracteres(ArraySalidaAuxiliar)
    if(flagSoloRuta = 1){
        ArraySalida:= quitarCaracteres(ArraySalidaAuxiliar,"\")
    }else{
        ArraySalida := ArraySalidaAuxiliar
    }
Return ArraySalida

}