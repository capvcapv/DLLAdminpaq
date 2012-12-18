VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function obtenerContrasena(cliente As String) As String
    
   Dim conn As ADODB.Connection
   Dim rst As ADODB.Recordset
   Set conn = ConectarBDAdminpaq
   conn.Open
   Set rst = conn.Execute("select CTEXTOEX01 from MGW10002 where CRFC='" & cliente & "'")
   rst.MoveFirst
   obtenerContrasena = Trim(rst(0))
    
End Function
    
Public Function obtenerClave(cliente As String) As String
    
   Dim conn As ADODB.Connection
   Dim rst As ADODB.Recordset
   Set conn = ConectarBDAdminpaq
   conn.Open
   Set rst = conn.Execute("select CCODIGOC01 from MGW10002 where CRFC='" & cliente & "'")
   rst.MoveFirst
   obtenerClave = Trim(rst(0))
    
End Function

Public Function obtenerTodos(nombre As String) As String

   Dim conn As ADODB.Connection
   Dim rst As ADODB.Recordset
   Dim json As String
   Dim ch As String
   
   Set conn = ConectarBDAdminpaq
   conn.Open
   Set rst = conn.Execute("select CCODIGOC01,CRAZONSO01 from MGW10002 where CTIPOCLI01=1 and CRAZONSO01 like '" & nombre & "%'")
   rst.MoveFirst
   
   json = "["
      
   While Not rst.EOF
    
    json = json & "{" & Chr(34) & "codigo" & Chr(34) & ":" & Chr(34) & rst(0) & Chr(34) & "," & Chr(34) & "nombre" & Chr(34) & ":" & Chr(34) & rst(1) & Chr(34) & "},"
    
    rst.MoveNext
   Wend
   
   json = Mid$(json, 1, Len(json) - 1)
   json = json & "]"
   obtenerTodos = json

End Function

Public Function obtenerDetalle(codigo As String) As String
    
   Dim conn As ADODB.Connection
   Dim rst As ADODB.Recordset
   Dim json As String
   Dim ch As String
   
   Set conn = ConectarBDAdminpaq
   conn.Open
   
   Set rst = conn.Execute("select CIDCLIEN01,CRAZONSO01 from MGW10002 where CCODIGOC01='" & codigo & "'")
   
   Dim idcliente As String
   Dim nombre As String
   Dim saldo As String
   Dim cargo, abono As Double
   
   idcliente = rst(0)
   nombre = rst(1)
   
   Set rst = conn.Execute("select sum(ctotal) from mgw10008 where cnatural01=0 and cidclien01=" & idcliente)
   rst.MoveFirst
   cargo = CDbl(rst(0))
   
   Set rst = conn.Execute("select sum(ctotal) from mgw10008 where cnatural01=1 and cidclien01=" & idcliente)
   If Not rst.EOF Then
   abono = CDbl(rst(0))
   Else
   abono = 0
   End If
   
   saldo = cargo - abono
   
   Set rst = conn.Execute("select CNOMBREC01,CNUMEROE01,CNUMEROI01,CCOLONIA,CCODIGOP01,CTELEFONO1,CEMAIL,CCIUDAD,CESTADO from MGW10011 where CIDCATAL01=" & idcliente & " and CTIPOCAT01=1 and CTIPODIR01=0")
   rst.MoveFirst
      
   Dim direccion As String
   direccion = rst(0) & " N." & rst(1) & " Int." & rst(2) & " Col. " & rst(3) & " C.P. " & rst(4) & rst(7) & " " & rst(8)
      
   json = "{" & Chr(34) & "nombre" & Chr(34) & ":" & Chr(34) & nombre & Chr(34) & "," & Chr(34) & "direccion" & Chr(34) & ":" & Chr(34) & direccion & Chr(34) & "," & Chr(34) & "telefono" & Chr(34) & ":" & Chr(34) & rst(5) & Chr(34) & "," & Chr(34) & "correo" & Chr(34) & ":" & Chr(34) & rst(6) & Chr(34) & "," & Chr(34) & "saldo" & Chr(34) & ":" & Chr(34) & saldo & Chr(34) & "}"
   
   obtenerDetalle = json
   
End Function
