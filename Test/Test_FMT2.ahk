#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

#Include ../LibreriasAhk/validacionDatos.ahk

; Ejemplo de uso IsValidDateEGR (formato dd/mm/yyyy)
;fecha := "27/19/2025" ; incorrecto
fecha := "27/09/2025" ; correcto
if (IsValidDateEGR(fecha)) {
    MsgBox "La fecha " . %fecha% . " es valida."
} else {
    MsgBox "La fecha " . %fecha% . " no es valida."
}


; Ejemplo de uso IsValidDNIEGR
;dni := "12345678-A" ;incorrecto
dni := "00000003-A" ;correcto
if (IsValidDNIEGR(dni)) {
    MsgBox "El DNI " . %dni% . " es valido."
} else {
    MsgBox "El DNI " . %dni% . " no es valido."
}