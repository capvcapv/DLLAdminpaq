VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Producto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function obtenerProductos(nombre As String) As String

   Dim conn As ADODB.Connection
   Dim rst As ADODB.Recordset
   Dim json As String
   Dim ch As String
   
   Set conn = ConectarBDAdminpaq
   conn.Open
   Set rst = conn.Execute("select CCODIGOP01,CNOMBREP01 from MGW10005 where CTIPOPRO01=1 and CNOMBREP01 LIKE '" & nombre & "%'")
   rst.MoveFirst
   
   json = "["
   
   While Not rst.EOF
   
         json = json & "{" & Chr(34) & "codigo" & Chr(34) & ":" & Chr(34) & rst(0) & Chr(34) & "," & Chr(34) & "nombre" & Chr(34) & ":" & Chr(34) & rst(1) & Chr(34) & "},"
    
        rst.MoveNext
   Wend

   json = Mid$(json, 1, Len(json) - 1)
   json = json & "]"
   obtenerProductos = json

End Function

Public Function obtenerDetalle(codigo As String) As String

   Dim conn As ADODB.Connection
   Dim rst As ADODB.Recordset
   Dim json As String
   
   Set conn = ConectarBDAdminpaq
   conn.Open
            
    
   Set rst = conn.Execute("select CIDPRODU01,CNOMBREP01,CPRECIO1,CPRECIO2,CPRECIO3 from MGW10005 where CTIPOPRO01=1 and CCODIGOP01='" & codigo & "'")
   
   Dim idproducto, nombre, precio1, precio2, precio3 As String
   
   idproducto = rst(0)
   nombre = rst(1)
   precio1 = rst(2)
   precio2 = rst(3)
   precio3 = rst(4)
   
   Dim mes As Integer
   
   mes = Month(Now)
   
   mes = mes + 1
   
   If mes < 10 Then
        Set rst = conn.Execute("select CENTRADA0" & mes & ",CSALIDAS0" & mes & " FROM MGW10030 WHERE CIDALMACEN=1 AND CIDPRODU01=" & idproducto)
   Else
        Set rst = conn.Execute("select CENTRADA" & mes & ",CSALIDAS" & mes & " FROM MGW10030 WHERE CIDALMACEN=1 AND CIDPRODU01=" & idproducto)
   End If
   
   Dim existencia As Double
   
   existencia = CDbl(rst(0)) - CDbl(rst(1))
   
   json = "{" & Chr(34) & "nombre" & Chr(34) & ":" & Chr(34) & nombre & Chr(34) & "," & Chr(34) & "precio1" & Chr(34) & ":" & Chr(34) & precio1 & Chr(34) & "," & Chr(34) & "precio2" & Chr(34) & ":" & Chr(34) & precio2 & Chr(34) & "," & Chr(34) & "precio3" & Chr(34) & ":" & Chr(34) & precio3 & Chr(34) & "," & Chr(34) & "existencia" & Chr(34) & ":" & Chr(34) & existencia & Chr(34) & "}"
   
   obtenerDetalle = json

End Function

