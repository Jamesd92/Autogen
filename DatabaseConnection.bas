Attribute VB_Name = "DatabaseConnection"
Option Compare Database
Public Function Connection_Query(query As String) As ADODB.Recordset
    'Database Open Connection and Query
    ' Inputs Query as String
    ' Outputs RecordSet
    Dim cn As New ADODB.connection
    Dim rs As New ADODB.Recordset
    
    'Opening Database Connection
    cn.Open "DRIVER={SQL Server};SERVER=W6ZGQ9M2;" & "trusted_connection=yes;DATABASE=Infrastructure"
    rs.Open query, cn, adOpenStatic
    Set Connection_Query = rs
End Function

