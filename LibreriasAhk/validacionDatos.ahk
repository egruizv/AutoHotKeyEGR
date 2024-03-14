#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

/*
Version 1.1.0 Ernesto Garcia 13/03/2024 
    isValidEmailEGR(emailStr)
Version 1.0.0 Ernesto Garcia 12/03/2024 
    IsValidDateEGR(dateString)
    IsValidDNIEGR(dniString)
*/
/*
Una funcion que reciba un string y que valide si es una fecha valida de tipo dd/mm/yyyy 
dode dd es el dia , mm es el mes y yyyy es el año
En esta función:
Primero, verificamos si la cadena tiene el formato correcto utilizando una expresión regular.
Luego, extraemos el día, mes y año de la cadena.
Finalmente, comprobamos si los valores son válidos 
(por ejemplo, el día está entre 1 y 31, el mes entre 1 y 12, y el año entre 1900 y 9999).
*/
; Función para validar una fecha en formato dd/mm/yyyy
IsValidDateEGR(dateString) {
    ; Verificar si la cadena tiene el formato correcto
    if (RegExMatch(dateString, "^\d{2}/\d{2}/\d{4}$")) {
        ; Extraer día, mes y año
        day := SubStr(dateString, 1, 2)
        month := SubStr(dateString, 4, 2)
        year := SubStr(dateString, 7, 4)

        ; Verificar si los valores son válidos
        if (day >= 1 && day <= 31 && month >= 1 && month <= 12 && year >= 1900 && year <= 9999) {
            ; La fecha es válida
            return true
        }
    }
    ; La fecha no es válida
    return false
}


/*
En esta función:
Verificamos si la cadena tiene el formato correcto (8 dígitos seguidos de un guión y una letra).
Extraemos el número y la letra del DNI.
Calculamos la letra esperada según el número del DNI y comparamos con la letra proporcionada.
*/
; Función para validar un DNI Español
; Función para validar un DNI Español
IsValidDNIEGR(dniString) {
    ; Verificar si la cadena tiene el formato correcto (8 dígitos seguidos de una letra)
    if (RegExMatch(dniString, "^\d{8}-[A-Za-z]$")) {
        ; Extraer el número y la letra del DNI
        dniNumber := SubStr(dniString, 1, 8)
        dniLetter := SubStr(dniString, 10)

        ; Calcular la letra esperada según el número del DNI
        expectedLetter := "TRWAGMYFPDXBNJZSQVHLCKE"
        calculatedIndex := Mod(dniNumber, 23)
        expectedLetter := SubStr(expectedLetter, calculatedIndex + 1, 1)

        ; Convertir la letra del DNI a mayúsculas
        dniLetterASCII := Asc(dniLetter)
        if (dniLetterASCII >= 97 && dniLetterASCII <= 122) {
            dniLetter := Chr(dniLetterASCII - 32)
        }

        ; Verificar si la letra coincide con la esperada
        if (dniLetter = expectedLetter) {
            ; El DNI es válido
            return true
        }
    }
    ; El DNI no es válido
    return false
}

; Función para validar un correo electrónico
isValidEmailEGR(emailStr) {
    ; Obtener la longitud del string
    emailStrLen := StrLen(emailStr)
    
    ; Eliminar espacios en blanco (AutoTrim)
    emailStr := Trim(emailStr)
    
    ; Convertir a minúsculas
    StringLower, emailStr, emailStr
    
    ; Dividir en parte local y dominio
    StringGetPos, atPos, emailStr, @, R
    If ErrorLevel
        Return false ; No hay @ en el string
    
    StringLeft, localPart, emailStr, %atPos%
    StringRight, domainPart, emailStr, % emailStrLen - atPos - 1
    
    ; Permitir comentarios en partes de dominio
    domainPart := RegExReplace(domainPart, "\([^()]*\)", "")
    
    ; Asegurarse de que no haya más de un @ (ya hemos sanitizado los válidos)
    If InStr(localPart, "@")
        Return false ; Demasiados @ en el string
    
    ; Verificar que el nombre de usuario tenga al menos 1 caracter
    If StrLen(localPart) = 0
        Return false ; Falta el nombre de usuario
    
    ; Dividir la parte de dominio en componentes
    StringSplit, domainComponents, domainPart, `. 
    
    ; Verificar que haya al menos 2 componentes de dominio
    If domainComponents0 < 2
        Return false ; No hay suficientes componentes de dominio
    
    ; Verificar cada componente de dominio
    Loop %domainComponents0%
    {
        domainComponent := domainComponents%A_Index%
        If (StrLen(domainComponent) > 0)
        {
            StringLeft, firstChar, domainComponent, 1
            StringRight, lastChar, domainComponent, 1
            If RegExMatch(firstChar, "[\.-]")
                Return false ; Carácter inválido al principio
        }
    }
    
    ; El correo es válido
    Return true
}

