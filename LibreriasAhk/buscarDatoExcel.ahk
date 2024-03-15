#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

#Include, ../LibreriasAhk/leerExcel.ahk  

/*
Version 1.0.0 Ernesto Garcia 15/03/2024
    Busca un elemento en el Excel y saca un informe de las posiciones donde se encuetra
    Recibe un documento Excel y una variable "String",  recorre el Excel y busca la variable en las celdas del Excel
    Devuelve un Matriz con las posiciones en el Excel de dicha variable MatrizPosiciones := [] 
    Por ejemplo MatrizPosiciones[1] := [1,3]  MatrizPosiciones[2] := [3,7]  [Fila,Columna]
    Dependencia con libreria 
        leerExcel.ahk (leer_Excel_completo_EGR)
   Funciones auxiliares en este script
        buscarStringMatriz(Matriz_Total,DatoBuscado) 
*/

buscarDatoExcel(FilePathficheroExcelLectura,DatoBuscado){
    Matriz_Total :=[]
    MatrizPosiciones := [] ;Matriz de Salida donde se van incluyendo las posiciones donde se enuentra el dato buscado
    Matriz_Total := leer_Excel_completo_EGR(FilePathficheroExcelLectura)
    longitud_Matriz_Total := Matriz_Total.length()
    if(longitud_Matriz_Total>0){
        MatrizPosiciones:= buscarStringMatriz(Matriz_Total,DatoBuscado) 
    }
    Return MatrizPosiciones
}

buscarStringMatriz(Matriz_Total,DatoBuscado) {
    MatrizSalida := []
    ;Recorremos la Matriz_Total y vamos buscando el DatoBuscado
    longitud_Matriz_Total := Matriz_Total.length()
    controlWhile1 := 1
    indiceMatrizSalida := 1    
    While, controlWhile1 <= longitud_Matriz_Total {
       RowArray := []  ; Inicializa un array para cada fila
       RowArray := Matriz_Total[controlWhile1] ; [A1,B1,C1,...,K1]
       longitud_RowArray:= RowArray.length()
       controlWhile2 := 1
       While, controlWhile2 <= longitud_RowArray {
            valor := RowArray[controlWhile2]
            if(valor = DatoBuscado){                
                MatrizSalida.InsertAt(indiceMatrizSalida, [controlWhile1,controlWhile2]) 
                indiceMatrizSalida++
            }
            controlWhile2++    
        }

       controlWhile1++
    }
    Return MatrizSalida
}

