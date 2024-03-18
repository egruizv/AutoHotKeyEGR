#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

/******************************************************************************
********************* VARIABLES GLOBALES **************************************
*******************************************************************************
*/

url_Entrada  := "https://rpachallenge.com/"
tiempoEspera := 0
FilePathficheroExcelLectura := "C:\Users\egarciar\Documents\CAIXA\Practicas\Reto_1_rpachallenge\Datos\challenge_Prueba.xlsx"

/******************************************************************************
********************* LANZAMOS EL ROBOT ***************************************
*******************************************************************************
*/
abrir_pagina_web(url_Entrada,FilePathficheroExcelLectura)


/******************************************************************************
********************* FUNCIONES ***********************************************
*******************************************************************************
*/

/*
Esta funcion FUNCION PADRE abrir_pagina_web(url_Entrada,Fichero_plano,FilePathficheroExcelLectura) hace lo siguiente:
Recibe 
    url_Entrada --> La pagina web que tenemos que abrir    
    FilePathficheroExcelLectura --> Fichero Excel de donde vamos a obtener los datos a insertar

*/
abrir_pagina_web(url_Entrada,FilePathficheroExcelLectura)
{
    Url := url_Entrada

    web_browser := ComObjCreate("InternetExplorer.Application")
    web_browser.Visible := True

    web_browser.Navigate(Url)

    ;Mientras se esta cargando esperamos lo que decidamos cargando_pag_web(web_browser,tiempoEspera)
    cargando_pag_web(web_browser,500)

    ; Obtener el documento HTML de la página web actual
    ;Doc := ComObjCreate("InternetExplorer.Application").document

    ; Obtenemos los datos del EXCEL antes de empezar 
    Matriz_Total :=[]
    Matriz_Total := leer_Excel_completo(FilePathficheroExcelLectura)
    
    longitud_Matriz_Total := Matriz_Total.length()
    /*
    Paso 0 Damos a Start
    */
    presionar_Start(web_browser)
    ;Mientras se esta cargando esperamos lo que decidamos cargando_pag_web(web_browser,tiempoEspera)
    cargando_pag_web(web_browser,500)

    
    ;Tenemos que hacer el Paso 1 10 veces
    ;fix el numero de veces lo marcamos con la longitud de Matriz_Total
    ;Teniendo en cuenta que Matriz_Total tiene una fila mas que es la cabecera
    ;¿Cuantas veces tenemos que hacer el paso 1?  Respuesta Matriz_Total.Length -1
    bandera_situacion := 1
    repetir_loop := longitud_Matriz_Total -1
    if(repetir_loop<0)
        {
            repetir_loop:= 0
        }
    Loop, %repetir_loop%
    {
        /*
        Paso 1 Rellemanos los datos de la persona (bandera_situacion) y damos a Submit. persona 1, persona 2,... persona 10
        */
        ;1.1 Buscamos las ids de cada uno de los campos enviamos web_browser y elemento que indica que elemento hay que encontrar
            elemento := "labelFirstName"
            ID_first_name_input := buscar_ID_Campo(web_browser,elemento)
            elemento := "labelLastName"
            ID_last_name  := buscar_ID_Campo(web_browser,elemento)
            elemento := "labelCompanyName"
            ID_company_name  := buscar_ID_Campo(web_browser,elemento)
            elemento := "labelRole"
            ID_role_company  := buscar_ID_Campo(web_browser,elemento)
            elemento := "labelAddress"
            ID_address  := buscar_ID_Campo(web_browser,elemento)
            elemento := "labelEmail"
            ID_email  := buscar_ID_Campo(web_browser,elemento)
            elemento := "labelPhone"
            ID_phone_number  := buscar_ID_Campo(web_browser,elemento)        
        
            first_name_input := web_browser.document.getElementbyID(ID_first_name_input)
            last_name_input := web_browser.document.getElementbyID(ID_last_name)
            company_name_input := web_browser.document.getElementbyID(ID_company_name)
            role_company_input := web_browser.document.getElementbyID(ID_role_company)
            address_input := web_browser.document.getElementbyID(ID_address)
            email_input := web_browser.document.getElementbyID(ID_email)
            phone_number_input := web_browser.document.getElementbyID(ID_phone_number)

            ;1.2 Asociamos los elementos del Excel a las ids x_input
            /*
            Datos Obtenidos del EXCEL 
            */
            ; Aqui ya tenemos los datos del EXCEL en la variable Matriz_Total
            fila_actual := bandera_situacion +1 ; Ya que empliezan los datos que necesitamos del excel en la fila 2 
            Array_Datos := Matriz_Total[fila_actual]
            first_name_input.value := Array_Datos[1]
            last_name_input.value := Array_Datos[2]
            company_name_input.value := Array_Datos[3]
            role_company_input.value := Array_Datos[4]
            address_input.value := Array_Datos[5]
            email_input.value := Array_Datos[6]
            phone_number_input.value := Array_Datos[7]  
            Array_Datos := []
        bandera_situacion:= bandera_situacion+1
        ;1.3 Presionamos submit
        cargando_pag_web(web_browser,1000) ;Esperamos 2 segundos
        presionar_Submit(web_browser)
    }
    cargando_pag_web(web_browser,1000) ;Esperamos 1 segundos
    ;cerramos la web_browser 
    web_browser.Quit()
    ;MsgBox, % Doc.documentElement.outerHTML
}

/*
Esta funcion cargando_pag_web(web_browser,tiempoEspera) hace lo siguiente:
Recibe web_browser y el tiempo de espera que nosotros indiquemos 
Nota: 1000 = 1 segundo 
*/
cargando_pag_web(web_browser,tiempoEspera){
    while web_browser.busy
    {
        sleep tiempoEspera
    }
    ;Cuando termina de cargar esperamos 1000 ms = 1 segundo
    sleep tiempoEspera
}

/*
Esta funcion presionar_Start(web_browser) hace lo siguiente:
Recibe la variable web_browser y con ella obtener el documento HTML de la página web actual
Obtener todos los elementos de tipo botón 
Botones := Doc2.getElementsByTagName("button")
e Iterar sobre los botones para encontrar el que coincide con la clase específica
NOTA: En este caso uno concreto (Start) y realizamos la accion  hacer clic en el botón
*/
presionar_Start(web_browser){
     ; Esperar a que se cargue completamente el documento HTML
     /*
     while (web_browser.readyState != 4 || web_browser.document.readyState != "complete" || web_browser.busy){
        Sleep, 1000
     }
     */
      ; Obtener el documento HTML de la página web actual
      Doc2 := web_browser.document
      ; Verificar si se obtuvo correctamente el documento HTML
      if (IsObject(Doc2)) {
          ; Obtener todos los elementos de tipo botón
          Botones := Doc2.getElementsByTagName("button")
          
          ; Iterar sobre los botones para encontrar el que coincide con la clase específica
          BotonEncontrado := ""
          Loop, % Botones.length {
              Boton := Botones[A_Index-1]
              if (InStr(Boton.className, "waves-effect") && InStr(Boton.className, "col s12") && InStr(Boton.className, "m12") && InStr(Boton.className, "l12") && InStr(Boton.className, "btn-large") && InStr(Boton.className, "uiColorButton")) {
                  BotonEncontrado := Boton
                  break
              }
          }
  
          ; Verificar si se encontró el botón
          if (IsObject(BotonEncontrado)) {
              ; Hacer clic en el botón
              BotonEncontrado.click()
          } else {
              ;MsgBox, No se encontró el botón "Start".
          }
      } else {
          ;MsgBox, No se pudo obtener el documento HTML de la página web.
      }
  
}

/*
Esta funcion presionar_Submit(web_browser) hace lo siguiente:
Recibe la variable web_browser y con ella obtener el documento HTML de la página web actual
Obtener todos los elementos de tipo input 
Inputs := Doc3.getElementsByTagName("input")
e Iterar sobre los botones para encontrar el que coincide con la clase específica
NOTA: En este caso uno concreto (Submit) y realizamos la accion  hacer clic en el botón
*/
presionar_Submit(web_browser){

      ; Esperar a que se cargue completamente el documento HTML
    /*
     while (web_browser.readyState != 4 || web_browser.document.readyState != "complete" || web_browser.busy){
        Sleep, 1000
     }
     */

    ; Obtener el documento HTML de la página web actual    
        Doc3 := web_browser.document

        ; Verificar si se obtuvo correctamente el documento HTML
        if (IsObject(Doc3)) {
            ; Obtener todos los elementos de tipo input
            Inputs := Doc3.getElementsByTagName("input")
            
            ; Iterar sobre los inputs para encontrar el que es un botón de envío
            BotonEncontrado := ""
            Loop, % Inputs.length {
                Input := Inputs[A_Index-1]
                if (Input.type = "submit" && InStr(Input.className, "btn") && InStr(Input.className, "uiColorButton")) {
                    BotonEncontrado := Input
                    break
                }
            }

            ; Verificar si se encontró el botón
            if (IsObject(BotonEncontrado)) {
                ; Hacer clic en el botón
                BotonEncontrado.click()
            } else {
            ; MsgBox, No se encontró el botón de envío.
            }
        } else {
            ;MsgBox, No se pudo obtener el documento HTML de la página web.
        }

    


}
/*
Esta funcion buscar_ID_Campo(web_browser,elemento) hace lo siguiente:
Recibe la variable web_browser y con ella obtener el documento HTML de la página web actual

Obtener todos los elementos de tipo input y con ng-reflect-name = elemento 
e Iterar sobre los elemento para encontrar el que coincide con la clase específica
NOTA: En este caso uno concreto (first_name) y realizamos la accion  de obtener su ID
*/
buscar_ID_Campo(web_browser,elemento) {
    
    ; Esperar a que se cargue completamente el documento HTML
    /*
     while (web_browser.readyState != 4 || web_browser.document.readyState != "complete" || web_browser.busy){
        Sleep, 1000
     }
     */

    ; Obtener el documento HTML de la página web actual
    Doc := web_browser.document

    ; Verificar si se obtuvo correctamente el documento HTML
    if (IsObject(Doc)) {
        ; Obtener todos los elementos de tipo input
        Inputs := Doc.getElementsByTagName("input")
        
        ; Iterar sobre los inputs para encontrar el que tiene el atributo ng-reflect-name="labelFirstName"
        Loop, % Inputs.length {
            Input := Inputs[A_Index-1]
            if (Input.getAttribute("ng-reflect-name") = elemento) {
                ; Obtener el ID del elemento de entrada
                ID := Input.id
                
                ; Devolver el ID
                Return ID
            }
        }
        
        ; Si no se encuentra ningún elemento, mostrar un mensaje
       ; MsgBox, No se encontró el elemento de entrada con ng-reflect-name="labelFirstName".
    } else {
      ;  MsgBox, No se pudo obtener el documento HTML de la página web.
    }
}

/*
Esta funcion lleer_Excel_completo(FilePathficheroExcelLectura) hace lo siguiente:
Recibe la direccion de un FilePathficheroExcel que sera un fichero Excel a leeer
Abre el fichero, y recorre todas  las filas  metiendo los datos en una variable RowArray y a su vez en Matriz_Excel
Devuelve Matriz_Excel que es un array de Arrays donde cada Matriz_Excel[i] corresponde a una persona
*/
leer_Excel_completo(FilePathficheroExcelLectura){
    ; Crea un objeto de Excel
    xlApp := ComObjCreate("Excel.Application")
    xlApp.Visible := false  ; Para que no se abra Excel visible
  
    ; Abre el archivo Excel
    FilePath := FilePathficheroExcelLectura
    xlBook := xlApp.Workbooks.Open(FilePath)
  
    ; Selecciona la primera hoja del libro
    xlSheet := xlBook.Sheets(1)
  
    ; Obtén el rango de celdas con datos
    xlRange := xlSheet.UsedRange
  
    ; Obtiene el número de filas y columnas en el rango
    Rows := xlRange.Rows.Count
    ;Rows := 11
    Cols := xlRange.Columns.Count  
  
    Matriz_Excel := [] ; Resultado
    numero_columna := 1
    numero_fila := 1
    ; Recorre el Excel por filas
    While, numero_fila <= Rows {
        RowArray := []  ; Inicializa un array para cada fila
        var_aux := 1
        numero_columna := 1
            While, numero_columna <= Cols {
                CellValue := xlSheet.Cells(numero_fila, A_Index).Value
                if (CellValue != "") {
                ;CellValue := xlSheet.Cells(A_Index, A_LoopField).Value
                ;RowArray.Push(CellValue)  ; Agrega el valor de la celda al array de la fila
                RowArray.InsertAt(var_aux, CellValue) 
                }Else
                {
                    Break ;Si la celda esta vacia salimos del bucle
                }
                var_aux++
                ;fix
                numero_columna++
                
            }  
        ; Aquí puedes hacer lo que necesites con los datos de la fila, por ejemplo, mostrarlos en un MsgBox
        ; Incluimos la fila en un Array Matriz_Excel :=[] Matriz_Excel[1] := RowArray 
        ; Matriz_Excel.InsertAt(numero_fila, RowArray) 
        Matriz_Excel.InsertAt(numero_fila, RowArray) 
        numero_fila := numero_fila +1
    }    
    ; Cierra el archivo Excel
    xlBook.Close(false)
  
    ; Cierra la aplicación de Excel
    xlApp.Quit()
  
    ; Libera los objetos de Excel de la memoria
    xlRange := ""
    xlSheet := ""
    xlBook := ""
    xlApp := ""
    Return Matriz_Excel
  }