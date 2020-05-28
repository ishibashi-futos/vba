Sub DeleteDefinedNames()
 
    Dim n As Name
    For Each n In ActiveWorkbook.Names
        On Error Resume Next  ' エラーを無視。
        n.Delete
    Next
 
End Sub
