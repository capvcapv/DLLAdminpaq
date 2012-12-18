VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Agente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function obtenerContrasena(codigo As String) As String

   Dim conn As ADODB.Connection
   Dim rst As ADODB.Recordset
   Set conn = ConectarBDAdminpaq
   conn.Open
   Set rst = conn.Execute("select CTEXTOEX01 from MGW10001 where CCODIGOA01='" & codigo & "'")
   rst.MoveFirst
   obtenerContrasena = Trim(rst(0))

End Function
