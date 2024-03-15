#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

#Include, ../LibreriasAhk/leerExcel.ahk  
#Include, ../LibreriasAhk/escribirExcel.ahk 
#Include, ../LibreriasAhk/validacionDatos.ahk 
/*
Version 1.3.0 Ernesto Garcia 14/03/2024
    Solucionar Fix : Cuando en la columna hay datos repetidos no ordena bien 
     Funciones auxiliares cambiadas
        ordenar_String_Alfabeticamente(MatrizEntrada,MatrizAuxiliar)
        ordenar_Fecha(MatrizEntrada,MatrizAuxiliar)
        obtener_indiceOriginal(MatrizEntrada,value,controlIndicesRepetidos)
    Funciones auxiliares en este script 
         estaDatoEnArrayDatos(Dato,ArrayDatos)

Version 1.2.0 Ernesto Garcia 14/03/2024
    Valoramos que tipo de dato es y segun sea ordenamos Nota: Incluir ordenar fecha
    En la ordenacion valoramos el tipo de dato que es  la columna que vamos a ordenar (Mirar Ordenacion_Columna)
        Números: enteros y decimales. --> Aplicamos la ordenacion "normal" la actual
        Texto: cadenas de caracteres alfanuméricos. --> Aplicamos la ordenacion "normal" la actual 
        Fechas y horas: información de fecha y hora. --> Utilizaremos funcion especial de ordenacion Fechas 
        Valores lógicos: verdadero/falso o 0/1.  --> Utilizaremos funcion especial de ordenacion Valores Logicos
    Dependencia con libreria 
       validacionDatos.ahk ( IsValidDateEGR(dateString) )
       leerExcel.ahk  ( obtenerDatoFilaColumna(FilePathficheroExcelLectura,Fila,Columna))
    Funciones auxiliares cambiadas
        indices_filas_ordenadas_columna(Matriz_Total,Ordenacion_Columna,TipoOrden,TipoDatoColumna)
    Funciones auxiliares en este script 
        tipoDato(DatoColumna)
        ordenar_Fecha(MatrizEntrada,MatrizAuxiliar) 
        ordenar_Logico(MatrizEntrada,MatrizAuxiliar)
        Fecha_DDMMYYYY_a_YYYYMMDD(fecha_DDMMYYYY) 
        Fecha_YYYYMMDD_DDMMYYYY(fecha_YYYYMMDD)
        obtenerDatoFilaColumna()

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
    ; Aqui buscamos el dato de la Fila = 1 yColumna = Ordenacion_Columna
    FilaAux:= 1
    ColumnaAux := Ordenacion_Columna
    DatoColumna := ""
    DatoColumna := obtenerDatoFilaColumna(FilePathficheroExcelLectura,FilaAux,ColumnaAux)
    longitud_Matriz_Total := Matriz_Total.length()
   
    ;Ponemos las filas del excel en un String para luego ir ordenando, solo cogemos la columna, y concatenamos la fila 
    InidcesFilasOrdenados := []
    ; Mejora ordenacion segun el tipo de dato de la columna 
    TipoDatoColumna := tipoDato(DatoColumna)
    InidcesFilasOrdenados := indices_filas_ordenadas_columna(Matriz_Total,Ordenacion_Columna,TipoOrden,TipoDatoColumna) ; 1xnumerofilas    
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


   
 indices_filas_ordenadas_columna(Matriz_Total,Ordenacion_Columna,TipoOrden,TipoDatoColumna) {

    columna := Ordenacion_Columna
    MatrizResultado := []
    MatrizResultado2 := []
    MatrizResultado_F := []
    Indices_Filas_ordenadas := []
    Indices_Filas_ordenadas_Salida := []
    TipoDatoColumnaAux := TipoDatoColumna
    MatrizResultado := string_datos_filas_columna(Matriz_Total, columna) ;Array 1xfilas
    MatrizResultado2 := string_datos_filas_columna(Matriz_Total, columna) ;Array 1xfilas
    if(TipoDatoColumna = "Numero" or TipoDatoColumna = "Texto"){
        MatrizResultado_F:= ordenar_String_Alfabeticamente(MatrizResultado,MatrizResultado2)
    }else if(TipoDatoColumna = "Fecha"){
        MatrizResultado_F:= ordenar_Fecha(MatrizResultado,MatrizResultado2)
    }else if(TipoDatoColumna = "Logico"){
        MatrizResultado_F:= ordenar_Logico(MatrizResultado,MatrizResultado2)
    }else{ ;en el caso de que no se detecte o se envie un TipoDatoColumna lo consideramos como  la ordenacion "normal" la actual
        MatrizResultado_F:= ordenar_String_Alfabeticamente(MatrizResultado,MatrizResultado2)
    }

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
    controlIndicesRepetidos := []
    ; Iterar sobre los elementos ordenados
    for index, value in ArrayDesordenado
    {
        ; Insertar el par [fila, Dato] en Matriz_Salida
        ;Enviamos la variable controlIndicesRepetidos para controlar los valores repetidos
         ; funcion que recorre MatrizEntrada y busca el value , cuando lo encuentre devuelve el indice 
         ; si ese indice ya existe en controlIndicesRepetidos sigue buscando al siguiente
        indiceOriginalMatrizEntrada := obtener_indiceOriginal(MatrizAuxiliar,value,controlIndicesRepetidos) ; funcion que recorre MatrizEntrada y busca el value , cuando lo encuentre devuelve el indice
        controlIndicesRepetidos.InsertAt(index,indiceOriginalMatrizEntrada)
        Matriz_Salida.InsertAt(index, [indiceOriginalMatrizEntrada, value]) ; Usar index como el índice original
    }
 
    ; Devolver el Matriz_Salida
    return Matriz_Salida
 }
 
 obtener_indiceOriginal(MatrizEntrada,value,controlIndicesRepetidos){
    longitud_MatrizEntrada := MatrizEntrada.MaxIndex()
    controlWhile1 := 1
    bControlRepetido := true ; Si el valor es false es que no esta repetido, si es true es que esta repetido. Lo inicializamos a false para obligar a entrar en 
    While, controlWhile1 <= longitud_MatrizEntrada {
       if(MatrizEntrada[controlWhile1] = value){
        bControlRepetido:=  estaDatoEnArrayDatos(controlWhile1,controlIndicesRepetidos)
        if(!bControlRepetido){  ; Si coinciden MatrizEntrada[controlWhile1] = value y ademas no esta el indice en controlIndicesRepetidos lo incluimos 
            return controlWhile1
        }
       }
       controlWhile1++ 
    }
    ;Si llegamos aqui es que hay un error en la programacion 

 }


 estaDatoEnArrayDatos(Dato,ArrayDatos){    
    longitud_ArrayDatos:= ArrayDatos.length()
    controlWhile1 := 1
    While, controlWhile1 <= longitud_ArrayDatos {
        valor := ArrayDatos[controlWhile1]
        if(valor = Dato){                
            return true ; Si lo encuentra devuelve true
        }
        controlWhile1++    
    }
    return false ; Si no lo encuentra entonces devuelve false 
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

  tipoDato(DatoColumna){    
    TipoDato_Salida := "Numero"
    ;Las fechas las distinguimos por dd/mm/yyyy
    ;Usaremos la libreria validacionDatos.ahk
    if(IsValidDateEGR(DatoColumna)){
        TipoDato_Salida:= "Fecha"
    }
    Return TipoDato_Salida

  }



  ordenar_Fecha(MatrizEntrada,MatrizAuxiliar) {

    Matriz_Salida := [] ; Matriz_Salida[1]:= [fila, Datos]
    ArrayDesordenado := MatrizEntrada ; Este es el Array Desordenado
    ArrayDesordenado2 := [] ; Este es el Array Desordenado2 trasformando dd/mm/yyyy a yyyymmdd para poder ordenar alfabeticamente
    for index, value in ArrayDesordenado
    {
        ArrayDesordenado2[index] := Fecha_DDMMYYYY_a_YYYYMMDD(ArrayDesordenado[index])
    }
    ; Implementar el algoritmo de ordenación de burbuja
    Longitud := ArrayDesordenado2.MaxIndex()
    Loop % Longitud {
        for index, value in ArrayDesordenado2
        {
            if (index < Longitud) {
                if (ArrayDesordenado2[index] > ArrayDesordenado2[index + 1]) {
                    ; Intercambiar elementos si están en el orden incorrecto
                    Temp := ArrayDesordenado2[index]
                    ArrayDesordenado2[index] := ArrayDesordenado2[index + 1]
                    ArrayDesordenado2[index + 1] := Temp
                }
            }
        }
    }
 
    ;Volvemos a dejar el ArrayDesordenado2 que ya esta ordenado yyyymmdd a dd/mm/yyyy
    for index, value in ArrayDesordenado2
        {
            ArrayDesordenado2[index] := Fecha_YYYYMMDD_DDMMYYYY(ArrayDesordenado2[index])
        }
    controlIndicesRepetidos := []    
    ; Iterar sobre los elementos ordenados
    for index, value in ArrayDesordenado2
    {
        ; Insertar el par [fila, Dato] en Matriz_Salida
        ;Enviamos la variable controlIndicesRepetidos para controlar los valores repetidos
        ; funcion que recorre MatrizEntrada y busca el value , cuando lo encuentre devuelve el indice 
        ; si ese indice ya existe en controlIndicesRepetidos sigue buscando al siguiente
        indiceOriginalMatrizEntrada := obtener_indiceOriginal(MatrizAuxiliar,value,controlIndicesRepetidos) ; funcion que recorre MatrizEntrada y busca el value , cuando lo encuentre devuelve el indice
        controlIndicesRepetidos.InsertAt(index,indiceOriginalMatrizEntrada)
        Matriz_Salida.InsertAt(index, [indiceOriginalMatrizEntrada, value]) ; Usar index como el índice original
    }
 
    ; Devolver el Matriz_Salida
    return Matriz_Salida

    
  }
  ordenar_Logico(MatrizEntrada,MatrizAuxiliar) {
    ordenar_String_Alfabeticamente(MatrizEntrada,MatrizAuxiliar) 
  }

/*
llevamos dd/mm/yyyy a yyyymmdd
*/
  Fecha_DDMMYYYY_a_YYYYMMDD(fecha_DDMMYYYY) {
    ; Dividir la fecha en día, mes y año
    StringSplit, partes, fecha_DDMMYYYY, /

    ; Obtener día, mes y año
    dia := partes1
    mes := partes2
    anio := partes3

    ; Asegurarse de que el día y el mes tengan dos dígitos
    if (StrLen(dia) = 1)
        dia := "0" dia
    if (StrLen(mes) = 1)
        mes := "0" mes

    ; Formatear la fecha en el formato YYYYMMDD
    fecha_YYYYMMDD := anio mes dia

    ; Devolver la fecha formateada
    return fecha_YYYYMMDD
}
/*
llevamos yyyymmdd a dd/mm/yyyy
*/
Fecha_YYYYMMDD_DDMMYYYY(fecha_YYYYMMDD){
    ; SubStr(fecha_YYYYMMDD, posicion, longitud)
    fecha_DDMMYYYY :=  SubStr(fecha_YYYYMMDD, 7, 2)   . "/" . SubStr(fecha_YYYYMMDD, 5, 2)  . "/" SubStr(fecha_YYYYMMDD, 1, 4) 
    Return fecha_DDMMYYYY
}