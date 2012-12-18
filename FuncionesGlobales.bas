Attribute VB_Name = "FuncionesConexion"
Option Explicit

Public Function ConectarBDAdminpaq() As ADODB.Connection

    Dim conn As ADODB.Connection
    
    Set conn = New ADODB.Connection

    Dim ruta As String

    ruta = "C:\Compacw\Empresas\puta"
        
    conn = "Provider=MSDASQL.1; Presist Security Info=FALSE;Extended Properties=Driver={Microsoft Visual FoxPro Driver};UID=;SourceDB=" & ruta & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
    
    Set ConectarBDAdminpaq = conn
    
    
End Function


