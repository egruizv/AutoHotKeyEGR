#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%
#Persistent
#Include ../LibreriasAhk/Chrome.ahk
linkUrl := "https://rpachallenge.com/assets/downloadFiles/challenge.xlsx"
global url := "https://rpachallenge.com"
global localPath := A_MyDocuments . "\challenge.xlsx"
;#Include, ../../libs/OutputWindow.ahk


; Download File
Req := ComObjCreate("Msxml2.XMLHTTP.6.0")
Req.Open("GET", linkUrl, true)
Req.onreadystatechange := Func("Ready").Bind(Req, localPath)
Req.Send()

; when file ready launch
Ready(Req, filePath) {
   if !(Req.readyState = 4 && Req.status = 200)
      Return
   
   Arr := Req.responseBody
   pData := NumGet(ComObjValue(Arr) + 8 + A_PtrSize)
   len := Arr.MaxIndex() + 1
   FileOpen(filePath, "w").RawWrite(pData + 0, len)
   MsgBox, File downloaded

   Start()
}

Start(){
    ChromeInst := LaunchChrome_start("https://rpachallenge.com/")
    FillForms(ChromeInst)
    getResults(ChromeInst)
}

getResults(ChromeInst){
    message1 := ChromeInst.Evaluate("document.querySelector(``.message1``).innerText").value
    message2 := ChromeInst.Evaluate("document.querySelector(``.message2``).innerText").value

    ChromeInst.Call("Browser.close")
    MsgBox, % message1 . "`n" . message2
    ExitApp
}

LaunchChrome_start(url){
    FileCreateDir, ChromeProfile
    ChromeInst := new Chrome("ChromeProfile")
    ; --- Connect to the page/open browser ---
    if !(ChromeInst := ChromeInst.GetPage())
    {
        MsgBox, Could not retrieve page!
        ChromeInst.Kill()
    }
    else
    {
        ; --- Navigate to the desired URL ---        
        ChromeInst.WaitForLoad()
        ChromeInst.Call("Page.navigate", {url : url})
        ChromeInst.WaitForLoad()
    }

    ChromeInst.Evaluate("document.querySelector('button').click()")
    
    return ChromeInst
}

FillForms(ChromeInst){
    
    xls := ComObjCreate("Excel.Application")
    ; xls.Visible := True
    xls.Workbooks.Open(localPath)
    
    ; starting in line 2, evading columnName
    rowStartIndex := 2
    try{

        ; loop over the 10 rows
        Loop, 10 {
            Sleep, 50
            
            ; loop over every col
            while (A_Index < 8)
            {
                ; get the name of the column
                colNameParsed := parseColName(xls.Range(colParser(A_Index, 1)).text)
                
                ; get the value
                cellValue := xls.Range(colParser(A_Index, rowStartIndex)).text
                
                fillInputs(ChromeInst, colNameParsed, cellValue )
            }
            rowStartIndex+=1
            
            
            ChromeInst.Evaluate("document.querySelector('input[type=submit]').click()")
        }
    }
    Catch, e
    {
        MsgBox, % "Exception encountered in " e.What ":`n`n"
        . e.Message "`n`n"
        . "Specifically:`n`n"
        . Chrome.Jxon_Dump(Chrome.Jxon_Load(e.Extra), "`t")

        ExitApp
    }
}

fillInputs(ChromeInst, searchInputValue, insertValue ){
    try {
        ; OutPutDebugg, "test"
        ; OutputWindow(searchInputValue . "`n" . insertValue . "`n--------------------------------------------`n")
        selectorStr := "document.querySelector('div input[ng-reflect-name=""" . searchInputValue . """]').value='" . insertValue . "'"  
        ChromeInst.Evaluate(selectorStr)    
    }
    Catch, e 
    {

    }
}

colParser(colIndex, rowIndex){

    objectModelParser := { "1": "A"
        , "2": "B"
        , "3": "C"
        , "4": "D"
        , "5": "E"
        , "6": "F"
        , "7": "G"}

    return objectModelParser["" colIndex] . "" rowIndex
}


parseColName(colName){
    ; parse on needed inputs
    if(InStr(colname, "Phone" ))
        colNameParsed := "labelPhone"
    else if(InStr(colname, "Role" ))
        colNameParsed := "labelRole"
    else
    colNameParsed := "label" . StrReplace(colName, A_Space, "")

    return colNameParsed
}