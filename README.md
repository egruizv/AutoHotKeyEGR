# AutoHotKeyEGR
# Version 0.1 develop
# feature/FMT1 Trabajamos con Excel
    + leerExcel.ahk
# Mergeamos feature/FMT1 a develop 13/02/2024 
# feature/FMT2 Trabajamos con Validaciones de Datos
    + validacionDatos.ahk
# feature/FMT3  Trabajamos con Excel
    + escribirExcel.ahk
    + eliminarFilasRepetidasExcel.ahk
    * Test_FMT2.ahk (Modificamos para mejorar)
# feature/FMT4  Trabajamos con Excel
    + ordenarExcelPorColumna.ahk
# feature/FMT5  Trabajamos con Excel
    * eliminarFilasRepetidasExcel.ahk (Eliminamos filas repetidas segun una serie de columnas) (es decir eliminamos si A1 = A2 y B1 = B2 aunque los otros datos sean distintos quedandonos con la fila primera que se encuantra ) 
    * ordenarExcelPorColumna.ahk (Valoramos que tipo de dato es y segun sea ordenamos Nota: Incluir ordenar fecha)
    * ordenarExcelPorColumna.ahk (Ordenar Ascendente o Descendente)  
    - eliminamos ficheros prueba para ordenarlos
    + añadimos cartetas de Test para las pruebas
    * Test_FMT1.ahk
    * Test_FMT3.ahk
    - Test_FMT4.ahk
    + Test_FMT4_A.ahk
    + Test_FMT4_B.ahk
    * Test_FMT2.ahk
    * validacionDatos.ahk (Validamos email)
    * leerExcel.ahk (Funcion que te saca un elemento en concreto del excel)
# feature/FMT6  Trabajamos con Excel
    * ordenarExcelPorColumna.ahk (Solucionar cuando las filas de la columna tenga datos iguales) (es decir si A1=A2 no funciona bien)
    + buscarDatoExcel.ahk (Busca un elemento en el Excel y saca un informe de las posiciones donde se encuetra)
    + TEST_FMT6.ahk
    + Test_FMT4_C.ahk (Pruebas de ordenacion con datos repetidos en columnas)