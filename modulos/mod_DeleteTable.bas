Attribute VB_Name = "mod_DeleteTable"
Sub DeleteContTable(cTabla As String)
Dim strquery As String

strquery = "delete * from " & cTabla
dbdiet.Execute (strquery)

End Sub
