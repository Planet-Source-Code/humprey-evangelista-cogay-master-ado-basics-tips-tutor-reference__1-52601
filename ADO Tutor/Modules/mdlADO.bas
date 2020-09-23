Attribute VB_Name = "mdlADO"
Public strConn As String
Public strSQL As String

'-----------------------------------------------------
'--                   ADO Objects                   --
'-----------------------------------------------------
    
    Public objADOConn As New ADODB.Connection
    Public objADORs As New ADODB.Recordset
    Public objADORec As New ADODB.Record
    Public objADOCmd As New ADODB.Command
    Public objADOPar As New ADODB.Parameter
    Public objADOStream As New ADODB.Stream

