#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

/*
Version 1.0.0 Ernesto Garcia 12/03/2024
    escribir_en_fichero_Excel_EGR(Matriz_Final,FilePathficheroExcelSalida) 
    Funciones auxiliares en este script 
        escribir_en_fichero_ExcelFila(Matriz_Final[Index],FilePathficheroExcelSalida)
        insertar_fila_En_excel(ArrayDatos,ex)
*/

/*
Recoge una Matriz con los datos de todas las filas que hay que incluir y la direccion del Fichero de Salida
donde incluiremos las filas
*/
escribir_en_fichero_Excel_EGR(Matriz_Final,FilePathficheroExcelSalida){
    Longitud_Matriz_Final:= Matriz_Final.MaxIndex()
       
    ; Inicializar el índice del elemento actual
    Index := 1
    while (Index <= Longitud_Matriz_Final)
        {
            escribir_en_fichero_ExcelFila(Matriz_Final[Index],FilePathficheroExcelSalida)
            Index++
        }
   
}


/*
escribir en un fichero Excel(Matriz_Final,FilePathficheroExcelSalida) hace lo siguiente:
Recibe un Array de Datos en Modo Fila con n datos 
Ejemplo Recibe una Matriz RowArray :=[] que es un  Array de n posiciones
Va incluyendo los datos en el Excel en la Hoja 1 posicion A1 , A2, ... An
NOTA: Primera version con control sobre el sobreescribir Excel NO pide confirmacion 
Nota: ArrayDatos es una fila 
*/
escribir_en_fichero_ExcelFila(ArrayDatos,FilePathficheroExcelSalida){


    longitud_ArrayDatos:= ArrayDatos.length()
    /*
    Comprobamos si existe el fichero FilePathficheroExcelSalida
    Si existe lo abrimos para tabajar con el , si no existe 
    creamos uno en blanco
    */
    ; Ruta del archivo a comprobar

    ; Comprobar si el archivo existe Ruta del archivo a comprobar FilePathficheroExcelSalida
    if (FileExist(FilePathficheroExcelSalida))
    {
        ;MsgBox, El archivo %FilePathficheroExcelSalida% existe.
        ;Si existe lo abrimos , no hay que crearlo de nuevo
        ex := ComObjCreate("Excel.Application")
        ex.visible := False
        ex.Workbooks.Open(FilePathficheroExcelSalida)
        insertar_fila_En_excel(ArrayDatos,ex)
        ;Automatizar el grabar cuando pide confirmacion  
        ; Desactivar las alertas de Excel para evitar el cuadro de diálogo de confirmación al guardar
        ex.DisplayAlerts := false   
        ; Guardar el libro de trabajo activo
        ex.ActiveWorkbook.Save
        ; Marcar el libro de trabajo como guardado para evitar la confirmación al cerrar
        ex.ActiveWorkbook.Saved := true
        
        ;Copiandolo en el Excel final a FilePathficheroExcelSalida
        ex.ActiveWorkbook.SaveAs(FilePathficheroExcelSalida)
        ; Volver a activar las alertas de Excel
        ex.DisplayAlerts := true
        


        ; Cerrar el archivo Excel
        ex.ActiveWorkbook.Close
        ; Cerrar la aplicación Excel
        ex.Quit
        


    }
    else
    {
        ;MsgBox, El archivo %FilePathficheroExcelSalida% no existe.
         ;Creamos un Excel en blanco
        ex := ComObjCreate("Excel.Application")
        ex.visible := True
        ex.Workbooks.Add
        insertar_fila_En_excel(ArrayDatos,ex)
        ;Automatizar el grabar cuando pide confirmacion  
        ; Desactivar las alertas de Excel para evitar el cuadro de diálogo de confirmación al guardar
        ex.DisplayAlerts := false   
        ; Guardar el libro de trabajo activo
        ex.ActiveWorkbook.Save
        ; Marcar el libro de trabajo como guardado para evitar la confirmación al cerrar
        ex.ActiveWorkbook.Saved := true

        ;Copiandolo en el Excel final a FilePathficheroExcelSalida
        ex.ActiveWorkbook.SaveAs(FilePathficheroExcelSalida)
        ; Cerrar el archivo Excel
        ex.ActiveWorkbook.Close
        ; Cerrar la aplicación Excel
        ex.Quit
        

    }  

}


/*
insertar_fila_En_excel(Datos_Incluir, ex_donde_incluir) hace lo siguiente:
Recibe un Array de Datos en Modo Fila con n datos y los incluye en la siguiente fila "vacia"
*/
insertar_fila_En_excel(Datos_Incluir, ex_donde_incluir)
{
    primera_fila_en_blanco := 1
    longitud_Datos_Incluir := Datos_Incluir.length()

    ; Buscamos la primera fila vacía 
    Loop, 1000
    {
        CellValue := ex_donde_incluir.Cells(A_Index, 1).Value
        if (CellValue != "") { 
            primera_fila_en_blanco := primera_fila_en_blanco + 1
        } else {
            Break
        }
    }
    
    ; Insertamos los datos en esa primera fila vacía
    longitud_Datos_Incluir := Datos_Incluir.MaxIndex()
    controlWhile := 1
    fila := primera_fila_en_blanco
    columna := 1
    While, controlWhile <= longitud_Datos_Incluir {
        Data := Datos_Incluir[controlWhile]
        
        ; Insertar datos en la fila actual y recorriendo la columna  
        ex_donde_incluir.Cells(fila, columna).Value := Data
        
        controlWhile := controlWhile + 1
        columna := columna + 1
    }
}