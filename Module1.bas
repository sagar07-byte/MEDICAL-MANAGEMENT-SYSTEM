Attribute VB_Name = "Module1"
Public CC As ADODB.Connection
Public R As ADODB.Recordset
Public R2 As ADODB.Recordset
Public R3 As ADODB.Recordset
Public S As String
Public T As String
Public U As String
Public c1, I As Integer
Public Function path()
Set CC = New ADODB.Connection
CC.Open "Provider=MSDAORA.1;User ID=medicine/mmp;persist security info=true"
Set R = New ADODB.Recordset
Set R2 = New ADODB.Recordset
End Function





