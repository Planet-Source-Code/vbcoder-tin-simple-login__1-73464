Attribute VB_Name = "ModConnection"
Option Explicit
Public ConConnection As New ADODB.Connection
Public RstRecord As New ADODB.Recordset

'ESTABLISHING CONNECTION IN DATABASE
Sub DBConnection()
Set ConConnection = New ADODB.Connection
    ConConnection.CursorLocation = adUseClient
    ConConnection.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\database\Sample.mdb"
End Sub
