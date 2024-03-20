#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%


; Función para buscar el archivo en todas las carpetas y subcarpetas
BuscarArchivoEnC(carpetas, archivo) {
    ;Creo un String con las rutas de todos los ficheros que hay en la carpeta
    StringAllFicheros := CreateString(carpetas)
    ; Recorro el StringAllFicheros y coloco todos los datos en una MatrizAuxiliar
    Matriz_Auxiliar := StrSplit(StringAllFicheros, ";")
    
    ;localizo donde esta el archivo, recorro MatrizAuxiliar y veo si archivo esta en algun MatrizAuxiliar[i]
    longitud_Matriz_Auxiliar := Matriz_Auxiliar.length()
    controlWhile1 := 1
    indiceMatrizSalida := 1    
    ArraySalida := []
    indiceArraySalida := 1
    While, controlWhile1 <= longitud_Matriz_Auxiliar {
        ;Si el archivo esta en  MatrizAuxiliar[i] incluyo la direccion en ArraySalida[]
        ; Verificar si la subcadena está dentro de la cadena completa
        CadenaCompleta :=  Matriz_Auxiliar[controlWhile1]
        Subcadena := archivo
        if InStr(CadenaCompleta, Subcadena)
        {
            ArraySalida.InsertAt(indiceArraySalida, Matriz_Auxiliar[controlWhile1])
            indiceArraySalida++
        }
        controlWhile1++
    }
Return ArraySalida

}




CreateString(Folder, Call=0)
{
	global LoadIcons
	MatrizSalida := []
	Call++
	Loop, %Folder%\*.*, 1
	{
		;Progress, %A_Index%, %A_LoopFileDir%
        /*

		If LoadIcons
			Icon := "`tIcon" GetIcon(A_LoopFileFullPath)
        
        */
		If InStr(FileExist(A_LoopFileFullPath), "D")
		{
			Loop, %Call%
			;	String .= "`t"
			;String .= A_LoopFileName . Icon "`n"
            String .= Folder . "\" . A_LoopFileName . ";"            
			String .= CreateString(A_LoopFileFullPath, Call)
		}
		Else
		{
			Loop, %Call%
			;	Files .= "`t"
			;Files .= A_LoopFileName . Icon "`n"
            Files .= Folder . "\" . A_LoopFileName . ";" 
		}
	}
	String .= Files
	Call--
	return String
}