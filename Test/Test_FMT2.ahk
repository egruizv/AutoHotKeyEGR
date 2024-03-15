#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%
/******************************************************************************
********************* LIBRETIAS **************************************
*******************************************************************************
*/
#Include ../LibreriasAhk/validacionDatos.ahk

/******************************************************************************
********************* VARIABLES GLOBALES **************************************
*******************************************************************************
*/

; Ejemplo de uso IsValidDateEGR (formato dd/mm/yyyy)
;fecha := "27/19/2025" ; incorrecto
fecha := "27/09/2025" ; correcto
; Ejemplo de uso IsValidDNIEGR
;dni := "12345678-A" ;incorrecto
dni := "00000003-A" ;correcto
; Ejemplo de uso isValidEmailEGR
emailToCheck := "nombre.apellido@@dominio.com" ;incorrecto dos arrobas
; Ejemplo de uso IsValidEsBooleanEGR
Dato := true ;correcto
;Dato := 10  ;incorrecto

; Ejemplo de uso IsValidEsNumeroEGR
;Dato2:= "casa" ;incorrecto
;Dato2 := 10 ;correcto
;Dato2 := 3.14 ;correcto
Dato2 := 3,17 ;correcto


;emailToCheck := "nombre.apellido@dominio.com" ;correcto

/******************************************************************************
********************* LANZAMOS EL ROBOT ***************************************
*******************************************************************************
*/
; Ejemplo de uso IsValidDateEGR (formato dd/mm/yyyy)
pruebaIsValidDateEGR(fecha)
; Ejemplo de uso IsValidDNIEGR
pruebaIsValidDNIEGR(dni)
; Ejemplo de uso isValidEmailEGR
pruebaisValidEmailEGR(emailToCheck)
; Ejemplo de uso IsValidEsBooleanEGR
pruebaIsValidEsBooleanEGR(Dato)
;MsgBox, Ha terminado todo correcto
; Ejemplo de uso IsValidEsNumeroEGR
pruebaIsValidEsNumeroEGR(Dato2)


/******************************************************************************
********************* FUNCIONES ***********************************************
*******************************************************************************
*/
pruebaIsValidDateEGR(fecha){
    if (IsValidDateEGR(fecha)) {
        MsgBox "La fecha " . %fecha% . " es valida."
    } else {
        MsgBox "La fecha " . %fecha% . " no es valida."
    }
}

pruebaIsValidDNIEGR(dni){
    if (IsValidDNIEGR(dni)) {
        MsgBox "El DNI " . %dni% . " es valido."
    } else {
        MsgBox "El DNI " . %dni% . " no es valido."
    }
}


pruebaisValidEmailEGR(emailStr){
isEmailValid := isValidEmailEGR(emailStr)
if(isEmailValid){
    MsgBox, "El correo" . %emailStr% . " es valido."
}else{
    MsgBox, "El correo" . %emailStr% . " NO es valido."
}

}

pruebaIsValidEsBooleanEGR(Dato){
    isBooleano := IsValidEsBooleanEGR(Dato)
    if(isBooleano){
        MsgBox, "El dato" . %Dato% . " es Boolean."
    }else{
        MsgBox, "El dato" . %Dato% . " NO es Boolean."
    }
}


pruebaIsValidEsNumeroEGR(Dato){
    isNumero := IsValidEsNumeroEGR(Dato)
    if(isNumero){
        MsgBox, "El dato" . %Dato% . " es Numero."
    }else{
        MsgBox, "El dato" . %Dato% . " NO es Numero."
    }
}
