Attribute VB_Name = "Module1"
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sqlStr As String

Function GetConnection() As ADODB.Connection

conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\USER\Documents\SAPI.mdb;Persist Security Info=False"

End Function


