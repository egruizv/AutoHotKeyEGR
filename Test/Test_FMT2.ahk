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

/******************************************************************************
********************* LANZAMOS EL ROBOT ***************************************
*******************************************************************************
*/
; Ejemplo de uso IsValidDateEGR (formato dd/mm/yyyy)
pruebaIsValidDateEGR(fecha)
; Ejemplo de uso IsValidDNIEGR
pruebaIsValidDNIEGR(dni)
;MsgBox, Ha terminado todo correcto

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