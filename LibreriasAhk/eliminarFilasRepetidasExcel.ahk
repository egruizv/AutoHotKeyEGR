#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

#Include, ../LibreriasAhk/leerExcel.ahk  
#Include, ../LibreriasAhk/escribirExcel.ahk  
/*
Version 2.0.0 Ernesto Garcia 14/03/2024
    Se añade el eliminar filas repetidas indicando solo las columnas que tenemos que comparar
    Se envía un Array con el numero de las columnas que queremos comparar (Nota : si enviamos el Array vacio es que miramos todas)
    Cuando hay dos "filas iguales" Nos quedamos con la fila de mas arriba.
    eliminar_filas_repetidas_Excel_EGR(FilePathficheroExcelLectura,FilePathficheroExcelSalida,ArrayColumnasComparar) 
    Funciones auxiliares cambiadas
        string_datos_filas(Matriz_Total,ArrayColumnasComparar)

Version 1.0.0 Ernesto Garcia 13/03/2024
    eliminar_filas_repetidas_Excel_EGR(FilePathficheroExcelLectura,FilePathficheroExcelSalida) 
    Dependencia con libreria 
        leerExcel.ahk (leer_Excel_completo_EGR)
        escribirExcel.ahk (escribir_en_fichero_Excel_EGR)
    Funciones auxiliares en este script 
        string_datos_filas(Matriz_Total)
        encontrarIndicesRepetidos(MatrizResultadoString)    
        CompararNumeroEnListaNumeros(Index,Repetidos)
  Devolvemos un String de OK o KO       
*/

eliminar_filas_repetidas_Excel_EGR(FilePathficheroExcelLectura,FilePathficheroExcelSalida,ArrayColumnasComparar){
    Resultado := "" ; Devolvemos un String de OK o KO segun va Bien o Mal 
    ; Obtenemos los datos del EXCEL antes de empezar 
    Matriz_Total :=[]
    Matriz_Total := leer_Excel_completo_EGR(FilePathficheroExcelLectura)
    longitud_Matriz_Total := Matriz_Total.length()
   
    ;Ponemos las filas del excel en un String para luego ir comparando
    MatrizResultadoString := []
    MatrizResultadoString := string_datos_filas(Matriz_Total,ArrayColumnasComparar)
    longitud_Matriz_Resultado_String := MatrizResultadoString.length()
   
    ;ComparamosString y devolvemos las posiciones que son iguales
   ; Llamar a la función y almacenar el resultado en una variable
   Repetidos := encontrarIndicesRepetidos(MatrizResultadoString)
   
    ;Ahora creamos una Matriz_Final con sólo las filas no repetidas (Todas las filas menos "Repetidos")
    Matriz_Final := []
    Matriz_Final := filasSinRepetirExcel(Matriz_Total,Repetidos)
   
    ;Nota: Borramos el fichero FilePathficheroExcelSalida si existe antes de escribir en el 
    FileDelete, %FilePathficheroExcelSalida%   ; Eliminamos el fichero
    ;Escribimos en el fichero Excel fila a fila
    escribir_en_fichero_Excel_EGR(Matriz_Final,FilePathficheroExcelSalida)
   
    if(longitud_Matriz_Total = longitud_Matriz_Resultado_String){
       Resultado := "OK"
    }else{
        Resultado := "OK"
    }
    Return Resultado
   }


    /*
Esta funcion string_datos_filas(MatrizLeer) hace lo siguiente:
Recibe una Matriz RowArray :=[] que es un  Array de n posiciones y un array ArrayColumnasComparar de las columnas a comparar
Ejemplo : 
Recoge los datos de cada una de los RowArray[i] y los recorre 
RowArray[1], RowArray[2].....RowArray[n]
Concatena los datos de cada fila en un String (Solo las columnas que hay en ArrayColumnasComparar)
Devolvemos una Matriz formada por "Strings" de la concatenacion de las columnas de cada fila  (Solo las columnas que hay en ArrayColumnasComparar)
*/
string_datos_filas(MatrizTotal,ArrayColumnasComparar){
    ArrayColumnasElegidas := ArrayColumnasComparar 
    ; Si ArrayColumnasComparar (ArrayColumnasElegidas) esta en blanco tenemos que coger todas las columnas
    ;Variable cogerTodasLascolumnas
    cogerTodasLascolumnas := false
    if(ArrayColumnasElegidas = null or ArrayColumnasComparar.length() <=0){
        cogerTodasLascolumnas := true
    }

    
    ;La Matriz final de resultados
    MatrizResultado :=[]
    ; Calculamos la longitud de MatrizTotal
    longitud_MatrizTotal := MatrizTotal.MaxIndex()
    controlWhile1 := 1
    While, controlWhile1 <= longitud_MatrizTotal {
        MatizAuxiliar := MatrizTotal[controlWhile1]
        ; Calculamos la longitud de MatrizLeer
        longitud_MatizAuxiliar := MatizAuxiliar.MaxIndex()
        controlWhile2 := 1
        String_fila := ""
        While, controlWhile2 <= longitud_MatizAuxiliar {
            ;Inicializamos Concatenar
            Concatenar := ""
            encontrado := false      
            ; Itera sobre el array ArrayColumnasElegidas para ver si controlWhile2 pertenece a ArrayColumnasElegidas
            for index, value in ArrayColumnasElegidas {
                ; Comprueba si el valor actual es igual a vnumero
                if (value = controlWhile2) {
                    encontrado := true
                    break  ; Termina el bucle si encuentra el valor
                }
            }
            ;Solo cambiamos el valor de Concatenar en los ArrayColumnasElegidas
            if(cogerTodasLascolumnas or encontrado){
                Concatenar := MatizAuxiliar[controlWhile2]            
            }

            ;Solo añadimos Concatenar si es distinto de ""
            if(Concatenar!=""){
                ; si String_fila = "" entonves no agregamos un "|" antes de él (Mejora)
                If (String_fila !=null and String_fila != ""){
                    String_fila := String_fila . Concatenar
                }
                else{
                    String_fila := String_fila . "|" . Concatenar
                }
            }
            
            controlWhile2 := controlWhile2 + 1
        }
        MatrizResultado.InsertAt(controlWhile1, String_fila) 
        controlWhile1++
    }
   Return MatrizResultado
    
    
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
 Devuelve las filas del Excel Sin repetir 
 */
filasSinRepetirExcel(Matriz_Total,Repetidos){
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