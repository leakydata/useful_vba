Private Sub PrintDocsInFolder()
    'https://www.mrexcel.com/board/threads/print-word-document-from-excel-file-using-vba.977562/
    Dim objWord
    Set objWord = CreateObject("Word.Application")
    
    ' Hidden window!
    objWord.Visible = False
    
    ' Save the original printer, otherwise you will reset the system default!
    objWord.ActivePrinter = "\\SPT-FPS\KONICA MINOLTA C654e"
    Debug.Print "CURRENT ACTIVE PRINTER:" & objWord.ActivePrinter
    
    
    ' Find and Loop through Files in Folder
    Dim Path As String
    Dim FName As String
    
    Path = "X:\Deposits\AR Collections\TEST\"
    FName = Dir(Path & "*.docx")
    Do While FName <> ""
        Debug.Print Path & FName
        Dim objDoc
        Set objDoc = objWord.Documents.Open(Path & FName)
        objDoc.PrintOut
        objDoc.Close

        
    FName = Dir()
    Loop
    
    objWord.Quit
    
    
End Sub
