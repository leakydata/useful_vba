Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
    'https://www.mrexcel.com/board/threads/turn-a-cell-into-a-button.488172/
    Dim fn As String
    Dim wb As Workbook
    Dim TargetFolder As String
    Dim filepath As String
    
    Application.ScreenUpdating = False
    
    filepath = "X:\Deposits\AR Collections\02-10-2023OPTION1\"

    Debug.Print Target.Address
    Debug.Print Target.Value
    Debug.Print Target.Column
    Debug.Print VarType(Target)
    Debug.Print Alpha_Column(Target)
    
    If Alpha_Column(Target) = "J" Then
        Debug.Print "YOU ARE IN COLUMN J"
        Debug.Print Target.Offset(0, -8).Value
        
        TargetFolder = filepath & Target.Offset(0, -8).Value
    
        If Right(TargetFolder, 1) <> Application.PathSeparator Then
            TargetFolder = TargetFolder & Application.PathSeparator
        End If
        
        Debug.Print TargetFolder
        
        fn = Dir(TargetFolder & "*.docx") ' the first file name in the folder
        PrintDocsInFolder
        
        
    End If

End Sub

Private Function Alpha_Column(Cell_Add As Range) As String
    Dim No_of_Rows As Integer
    Dim No_of_Cols As Integer
    Dim Num_Column As Integer
    No_of_Rows = Cell_Add.Rows.Count
    No_of_Cols = Cell_Add.Columns.Count
    If ((No_of_Rows <> 1) Or (No_of_Cols <> 1)) Then
        Alpha_Column = ""
        Exit Function
    End If
     Num_Column = Cell_Add.Column
    If Num_Column < 26 Then
        Alpha_Column = Chr(64 + Num_Column)
    Else
    
        Alpha_Column = Chr(Int(Num_Column / 26) + 64) & Chr((Num_Column Mod 26) + 64)
    End If
End Function

