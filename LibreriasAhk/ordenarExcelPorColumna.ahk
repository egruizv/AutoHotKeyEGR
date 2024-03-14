#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

#Include, ../LibreriasAhk/leerExcel.ahk  
#Include, ../LibreriasAhk/escribirExcel.ahk  
/*
Version 1.1.0 Ernesto Garcia 14/03/2024
    Se añade el orden ASC y DESC 
    Si es asc ordena los datos alfabéticamente (de la A a la Z) o mediante valores numéricos ascendentes
    Si es desc ordena los datos alfabéticamente (de la Z a la A) o mediante valores numéricos descendiente
    Cuando hay dos "filas iguales" Nos quedamos con la fila de mas arriba.
    Se envia una variable mas (TipoOrden) con dos posibles valores asc o desc.
    Nota: si se envia vacia o con un valor distinto a los dos posibles se considera orden ASC
    ordenar_Excel_Por_Columna_EGR(FilePathficheroExcelLectura,Ordenacion_Columna,FilePathficheroExcelSalida,TipoOrden)
    Funciones auxiliares cambiadas
        indices_filas_ordenadas_columna(Matriz_Total,Ordenacion_Columna,TipoOrden)

Version 1.0.0 Ernesto Garcia 13/03/2024
    ordenar_Excel_Por_Columna_EGR(FilePathficheroExcelLectura,Ordenacion_Columna,FilePathficheroExcelSalida)
    Dependencia con libreria 
        leerExcel.ahk (leer_Excel_completo_EGR)
        escribirExcel.ahk (escribir_en_fichero_Excel_EGR)
    Funciones auxiliares en este script 
        indices_filas_ordenadas_columna(Matriz_Total,Ordenacion_Columna)
            string_datos_filas_columna(Matriz_Total, columna)
            ordenar_String_Alfabeticamente(MatrizResultado,MatrizResultado2)
                obtener_indiceOriginal(MatrizAuxiliar,value) 
        filasOrdenadasExcel(Matriz_Total,InidcesFilasOrdenados)          
  Devolvemos un String de OK o KO       
*/

ordenar_Excel_Por_Columna_EGR(FilePathficheroExcelLectura,Ordenacion_Columna,FilePathficheroExcelSalida,TipoOrden){
    Resultado := "" ; Devolvemos un String de OK o KO segun va Bien o Mal 
    ; Obtenemos los datos del EXCEL antes de empezar 
    Matriz_Total :=[]
    Matriz_Total := leer_Excel_completo_EGR(FilePathficheroExcelLectura)
    longitud_Matriz_Total := Matriz_Total.length()
   
    ;Ponemos las filas del excel en un String para luego ir ordenando, solo cogemos la columna, y concatenamos la fila 
    InidcesFilasOrdenados := []
    InidcesFilasOrdenados := indices_filas_ordenadas_columna(Matriz_Total,Ordenacion_Columna,TipoOrden) ; 1xnumerofilas
    longitud_Inidces_Filas_Ordenados := InidcesFilasOrdenados.length()
      
    ;Ahora creamos una Matriz_Final con sólo las filas no repetidas (Todas las filas menos "Repetidos")
    Matriz_Final := []
    Matriz_Final := filasOrdenadasExcel(Matriz_Total,InidcesFilasOrdenados)
   
    ;Nota: Borramos el fichero FilePathficheroExcelSalida si existe antes de escribir en el 
    FileDelete, %FilePathficheroExcelSalida%   ; Eliminamos el fichero
    ;Escribimos en el fichero Excel fila a fila
    escribir_en_fichero_Excel_EGR(Matriz_Final,FilePathficheroExcelSalida)
    
    if(longitud_Matriz_Total = longitud_Inidces_Filas_Ordenados){
        Resultado := "OK" 
    }else{
        Resultado := "KO" 
    }    
     Return  Resultado
   }


   
 indices_filas_ordenadas_columna(Matriz_Total,Ordenacion_Columna,TipoOrden) {

    columna := Ordenacion_Columna
    MatrizResultado := []
    MatrizResultado2 := []
    MatrizResultado_F := []
    Indices_Filas_ordenadas := []
    Indices_Filas_ordenadas_Salida := []
 
    MatrizResultado := string_datos_filas_columna(Matriz_Total, columna) ;Array 1xfilas
    MatrizResultado2 := string_datos_filas_columna(Matriz_Total, columna) ;Array 1xfilas
 
    MatrizResultado_F:= ordenar_String_Alfabeticamente(MatrizResultado,MatrizResultado2)
    longitud_Matriz_Resultado_F := MatrizResultado_F.length()
 
    controlWhile1 := 1
    While, controlWhile1 <= longitud_Matriz_Resultado_F {
       valor := MatrizResultado_F[controlWhile1][1]
       Indices_Filas_ordenadas.InsertAt(controlWhile1, valor) 
       controlWhile1++
    }
    ;Aqui tenemos los indices ordenados de forma ASC. Si queremos que sea DESC llamamos a la funcion 
    if(TipoOrden = "desc" or TipoOrden = "DESC"){
        Indices_Filas_ordenadas_Salida := cambiarOrdenArray(Indices_Filas_ordenadas)
    }else{
        Indices_Filas_ordenadas_Salida := Indices_Filas_ordenadas
    }

    Return Indices_Filas_ordenadas_Salida
 
  }

  
/*
Datos de la columna x de todas las filas  
Datos que luego podemos usar para ordenar
*/
 string_datos_filas_columna(MatrizTotal, columna){

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
             if(controlWhile2 = columna){ 
                Concatenar := MatizAuxiliar[controlWhile2]
                ;Solo concateno un elemento no es necesario el control de primero, segundo...
                String_fila := String_fila . Concatenar               
          }
            controlWhile2 := controlWhile2 + 1
        }
        MatrizResultado.InsertAt(controlWhile1, String_fila) 
        controlWhile1++
    }
   Return MatrizResultado
    
    
 }
 
 /*
 ordenar_String_Alfabeticamente(MatrizResultado) Ordena un StringAlfabericamente
 Devuelve una Matriz con fila , Datos 
 Matriz_Salida[1]:= [fila, Dato]
 NO ORDENA BIEN 
 */
 ordenar_String_Alfabeticamente(MatrizEntrada,MatrizAuxiliar) {
    Matriz_Salida := [] ; Matriz_Salida[1]:= [fila, Datos]
    ArrayDesordenado := MatrizEntrada ; Este es el Array Desordenado
  
    ; Implementar el algoritmo de ordenación de burbuja
    Longitud := ArrayDesordenado.MaxIndex()
    Loop % Longitud {
        for index, value in ArrayDesordenado
        {
            if (index < Longitud) {
                if (ArrayDesordenado[index] > ArrayDesordenado[index + 1]) {
                    ; Intercambiar elementos si están en el orden incorrecto
                    Temp := ArrayDesordenado[index]
                    ArrayDesordenado[index] := ArrayDesordenado[index + 1]
                    ArrayDesordenado[index + 1] := Temp
                }
            }
        }
    }
 
    ; Iterar sobre los elementos ordenados
    for index, value in ArrayDesordenado
    {
        ; Insertar el par [fila, Dato] en Matriz_Salida
        indiceOriginalMatrizEntrada := obtener_indiceOriginal(MatrizAuxiliar,value) ; funcion que recorre MatrizEntrada y busca el value , cuando lo encuentre devuelve el indice
        Matriz_Salida.InsertAt(index, [indiceOriginalMatrizEntrada, value]) ; Usar index como el índice original
    }
 
    ; Devolver el Matriz_Salida
    return Matriz_Salida
 }
 
 obtener_indiceOriginal(MatrizEntrada,value){
    longitud_MatrizEntrada := MatrizEntrada.MaxIndex()
    controlWhile1 := 1
    While, controlWhile1 <= longitud_MatrizEntrada {
       if(MatrizEntrada[controlWhile1] = value){
          return controlWhile1
       }
       controlWhile1++ 
    }
 }


 filasOrdenadasExcel(Matriz_Total,InidcesFilasOrdenados) {
    Matriz_Salida := []
    ;Vamos recorriendo el Array InidcesFilasOrdenados y vamos incluyendo en Matriz_Salida los datos ya ordenados
    indice_mat_Salida := 1  
    longitud_Inidces_Filas_Ordenados := InidcesFilasOrdenados.MaxIndex()
    longitud_Matriz_Total := Matriz_Total.MaxIndex()
    controlWhile1 := 1
    While, controlWhile1 <= longitud_Inidces_Filas_Ordenados {
       var_aux := InidcesFilasOrdenados[controlWhile1]     
       Matriz_Salida.InsertAt(indice_mat_Salida, Matriz_Total[var_aux]) 
       indice_mat_Salida++
       controlWhile1++
    }
    Return Matriz_Salida
  }

  cambiarOrdenArray(Indices_Filas_ordenadas){
    ;Indice_Salida es un Array donde la primera posicion de Indices_Filas_ordenadas es la ultima
    ; la segunda posicion de Indices_Filas_ordenadas es la penúltima....
    Indice_Salida := []
    longitud_Indices_Filas_ordenadas := Indices_Filas_ordenadas.MaxIndex()
    controlWhile1 := 1
    While, controlWhile1 <= longitud_Indices_Filas_ordenadas {
        indiceAux := longitud_Indices_Filas_ordenadas - controlWhile1 +1 ; indice a donde va cada valor 
        Indice_Salida[indiceAux] := Indices_Filas_ordenadas[controlWhile1]
        controlWhile1++
    }
    Return Indice_Salida

  }
