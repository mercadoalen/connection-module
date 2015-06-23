Public DB As New ADODB.Connection
Public rs As New ADODB.Recordset
Public Sub DCN()
    If DB.State = adStateOpen Then DB.Close
    DB.CursorLocation = adUseClient
    DB.Open "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;Database=myDB;User=user;PASSWORD=mypass;"
    Exit Sub
End Sub
Public Sub rsconn(ByVal table As String)
    If rs.State = adStateOpen Then rs.Close
    rs.Open table, DB, 3, 3
End Sub
