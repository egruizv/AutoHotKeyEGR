#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%


quitarCaracteres(ArrayEntada,caracter){
    ; Buscamos el ultimo "\" que hay en el String ArrayEntada[i] y quitamos todo lo que hay detras de ese ultimo "\"
    controlWhile1 := 1
        indiceMatrizSalida := 1    
        ArraySalida := []
        ArraySalidaAuxiliar := []
        indiceArraySalida := 1
        longitud_ArrayEntada := ArrayEntada.length()
        While, controlWhile1 <= longitud_ArrayEntada {
            Cadena := ArrayEntada[controlWhile1]
            Partes := StrSplit(Cadena, caracter)
            ; Eliminar el último elemento
            Partes.Remove(Partes.MaxIndex())
            ; Unir las partes nuevamente con el carácter '\'        
            ; Concatenar las partes con '\'
            CadenaFinal := ""
            Loop, % Partes.MaxIndex()
            {
                CadenaFinal := CadenaFinal . Partes[A_Index] . "\"
            }
            ArraySalida.InsertAt(controlWhile1, CadenaFinal)
            controlWhile1++
        }
        Return ArraySalida
    }


    CreateString(Folder, Call=0)
    {
        global LoadIcons
        MatrizSalida := []
        Call++
        Loop, %Folder%\*.*, 1
        {
            ;Progress, %A_Index%, %A_LoopFileDir%
            /*
    
            If LoadIcons
                Icon := "`tIcon" GetIcon(A_LoopFileFullPath)
            
            */
            If InStr(FileExist(A_LoopFileFullPath), "D")
            {
                Loop, %Call%
                ;	String .= "`t"
                ;String .= A_LoopFileName . Icon "`n"
                String .= Folder . "\" . A_LoopFileName . ";"            
                String .= CreateString(A_LoopFileFullPath, Call)
            }
            Else
            {
                Loop, %Call%
                ;	Files .= "`t"
                ;Files .= A_LoopFileName . Icon "`n"
                Files .= Folder . "\" . A_LoopFileName . ";" 
            }
        }
        String .= Files
        Call--
        return String
    }

/*
 Función para encontrar índices repetidos excluyendo el índice actual
 Devuelve un Array con los indices repetidos excluyendo el índice de la fila que estamos mirando
 */
encontrarIndicesRepetidos(MatrizResultadoString) {
    Repetidos := []
    MaxIndex := MatrizResultadoString.MaxIndex()
    IndiceAux := 1
    
    ; Inicializar el índice del elemento actual
    Index := 1
    
    ; Iterar sobre cada elemento de la matriz
    while (Index <= MaxIndex)
    {
        ; Obtener el valor del elemento actual
        ValorActual := MatrizResultadoString[Index]
        
        ; Inicializar el índice del elemento comparado
        ComparadoIndex := Index + 1
        
        ; Iterar sobre los elementos restantes de la matriz
        while (ComparadoIndex <= MaxIndex)
        {
            ; Obtener el valor del elemento comparado
            ValorComparado := MatrizResultadoString[ComparadoIndex]
            
            ; Verificar si los valores son iguales y el índice no es el mismo
            if (ValorActual = ValorComparado && Index != ComparadoIndex)
            {
                ; Agregar el índice al arreglo de índices repetidos
                Repetidos.InsertAt(IndiceAux, Index) 
                IndiceAux++
            }
            
            ; Incrementar el índice del elemento comparado
            ComparadoIndex++
        }
        
        ; Incrementar el índice del elemento actual
        Index++
    }
    
    ; Devolver el arreglo de índices repetidos
    return Repetidos
}


/*
 Función que recibe  Matriz_Total de y los indices Repetidos
 Devuelve las filas del Matriz_Total Sin repetir en MatrizSalida
 */
 filasSinRepetirMatriz(Matriz_Total,Repetidos){
    MatrizSalida := []
    IndiceAux := 1
    Longitud_Matriz_Total := Matriz_Total.MaxIndex()
       
    ; Inicializar el índice del elemento actual
    Index := 1
    while (Index <= Longitud_Matriz_Total)
        {
            if(CompararNumeroEnListaNumeros(Index,Repetidos)){
                ;Al estar en la lista de Reperidos no se añade a la MatrizSalida
            }else{
                MatrizSalida.InsertAt(IndiceAux, Matriz_Total[Index]) 
                IndiceAux++
            }
            Index++
        }
    Return MatrizSalida    
    
}

CompararNumeroEnListaNumeros( Numero,ListaNumeros) {
    ; Iterar sobre la lista de números
    for index, num in ListaNumeros {
        ; Verificar si el número actual es igual al número dado
        if (num = Numero)
            return true  ; Devolver true si el número está en la lista
    }
    
    ; Devolver false si el número no está en la lista
    return false
}