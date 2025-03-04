Attribute VB_Name = "clsModulo"
Option Explicit
' En ADO, se usa el objeto Connection para abrir las bases de datos
Public cnn As ADODB.Connection
Public cnnxls As ADODB.Connection
' Necesitamos los eventos si queremos controlar algunas cosillas
Public rst As ADODB.Recordset
Public rstXls As ADODB.Recordset

'TODO:CAMBIAR la obtencion del path de la base

'Public Const sPathBase As String = "C:\Documents and Settings\paolettime\Escritorio\luc\Conciliacion Facturacion\concilaciones.mdb"
Public Const sPathBase As String = "L:\Prueba\Conciliacion Facturacion\concilaciones.mdb"

Public Sub AbrirConexionDB(ByRef cnn As ADODB.Connection)

    Set cnn = New ADODB.Connection
    
    ' Crear la conexión manualmente
    ' Usar "Provider=Microsoft.Jet.OLEDB.3.51;" para bases de Access 97
    ' Usar "Provider=Microsoft.Jet.OLEDB.4.0;"  para bases de Access 2000
    With cnn
        .ConnectionString = _
            "Provider=Microsoft.Jet.OLEDB.3.51;" & _
            "Data Source=" & sPathBase & ";"
        .Open
    End With
    
End Sub


Public Sub AbrirConexionXLS(archivoXLS As String, ByRef cnnxls As ADODB.Connection)
    
    Set cnnxls = New ADODB.Connection
    
    cnnxls.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & archivoXLS & _
               ";Extended Properties=""Excel 8.0;HDR=Yes;"""
               
End Sub


Public Sub CopyPasteGrid(grilla As MSFlexGrid)

 Dim s As String
    Dim i As Integer, j As Integer
    
    s = ""
    With grilla
    'almacenamos la region seleecionada a un string
    'para que funcione en el formato de excel las colummas
    'van separadas por tab's y los renglones por CrLf
    For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
            If j > 0 Then s = s & vbTab
            s = s & grilla.TextMatrix(i, j)
        Next j
    s = s & vbCrLf
    Next i
    
    
    Clipboard.Clear
    Clipboard.SetText s
    
    'si quieres pasar toda la grid modifica el for por
    
    End With



End Sub


Public Function ObtenerFechaMod(Path As String) As Date
  
  
    'Variable de tipo FileSystemObject y File
  
    Dim o_Fso As New FileSystemObject
    Dim Archivo As File
  
    ' Lee las propiedades del archivo mediante GetFile
    Set Archivo = o_Fso.GetFile(Path)
  
       
    'Visualiza el resultado: Creación ,acceso y modificado etc..
'    MsgBox "Fecha de creación del archivo: " & Format(Archivo.DateCreated), vbInformation
'    MsgBox "Fecha de modificación : " & Format(Archivo.DateLastModified), vbInformation
'    MsgBox "Fecha de del último acceso: " & Format(Archivo.DateLastAccessed), vbInformation
'    MsgBox "Tamaño del archivo : " & Format(Archivo.Size) & " Bytes", vbInformation
'    MsgBox "Tipo de archivo : " & Format(Archivo.Type), vbInformation
'
    Dim fecha As Date
    
    fecha = Format(Archivo.DateLastModified, "DD/MM/YYYY HH:MM:SS")
       
    ' Elimina las variables de objeto
    Set Archivo = Nothing
    Set o_Fso = Nothing
  
    ObtenerFechaMod = fecha
    
End Function


Public Function obtenerUltimaFilaExcel(Path As String) As Long

    Dim conexion As ADODB.Connection
    Dim rs As ADODB.Recordset
  
    Set conexion = New ADODB.Connection
       
    conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=" & Path & _
                  ";Extended Properties=""Excel 8.0;HDR=Yes;"""
  
         
    ' Nuevo recordset
    Set rs = New ADODB.Recordset
       
    With rs
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
    End With
    Dim i As Long
    Dim rango As String
    Dim hoja As String
    
    
    For i = 1 To 65536
       rango = "a" & i & ":" & "a" & (i + 1)
       hoja = "Base de datos" & "$" & rango
    
        rs.Open "SELECT * FROM [" & hoja & "]", conexion, , , adCmdText
           
        ' Mostramos los datos en el datagrid
        'Set DataGrid1.DataSource = rs
        
        If Not rs.EOF Then
            If rs(0).Value = "" Or IsNull(rs(0).Value) Then
                Exit For
            Else
                rs.Close
            End If
        Else
            Exit For
        End If
        
        
    Next

    rs.Close
    conexion.Close
    
    Set rs = Nothing
    Set conexion = Nothing

    obtenerUltimaFilaExcel = i

End Function

Public Function obtenerUltimaFilaExcelPosicion(Path As String, inicial As Long) As Long

    Dim conexion As ADODB.Connection
    Dim rs As ADODB.Recordset
  
    Set conexion = New ADODB.Connection
       
    conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=" & Path & _
                  ";Extended Properties=""Excel 8.0;HDR=Yes;"""
  
         
    ' Nuevo recordset
    Set rs = New ADODB.Recordset
       
    With rs
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
    End With
    Dim i As Long
    Dim rango As String
    Dim hoja As String
    
    
    For i = inicial To 65536
       rango = "a" & i & ":" & "a" & (i + 1)
       hoja = "evaluadores" & "$" & rango
    
        rs.Open "SELECT * FROM [" & hoja & "]", conexion, , , adCmdText
           
        ' Mostramos los datos en el datagrid
        'Set DataGrid1.DataSource = rs
        
        If Not rs.EOF Then
            If rs(0).Value = "" Or IsNull(rs(0).Value) Then
                Exit For
            Else
                rs.Close
            End If
        Else
            Exit For
        End If
        
        
    Next

    rs.Close
    conexion.Close
    
    Set rs = Nothing
    Set conexion = Nothing

    obtenerUltimaFilaExcelPosicion = i

End Function


Public Function ObtenerSoloDir(pathCompleto As String) As String

    Dim rutanorte As String
    Dim completa As String
    Dim controlillo As Boolean
    Dim posi As Integer
    Dim equis As Integer
    
    
    rutanorte = ""
    
    completa = pathCompleto
    
    controlillo = True
    posi = 1
    While controlillo
        equis = InStr(posi, completa, "\")
        If equis = 0 Then
            rutanorte = Mid(completa, 1, posi - 2)
            controlillo = False
        Else
            posi = equis + 1
        End If
    Wend

    ObtenerSoloDir = rutanorte


End Function


Public Function LimpiarFlx(flxgrid As MSFlexGrid)
    Dim i As Integer
    flxgrid.Clear
    For i = 1 To (flxgrid.Rows - 2)
        flxgrid.RemoveItem (1)
    Next
    
    If flxgrid.Rows = 2 Then
        flxgrid.FixedRows = 1 ' o x, las que tengamos
        flxgrid.Rows = flxgrid.FixedRows  ' desaparecen todas las líneas que no son fijas,
    End If

End Function

Public Function LimpiarStrSql(expresion As String) As String
    Dim strResult As String
    
    strResult = expresion
    
    strResult = Replace(strResult, "'", "")
    strResult = Replace(strResult, "--", "")
    strResult = Replace(strResult, """", "")
    strResult = Replace(strResult, "*", "")
    strResult = Replace(strResult, "/", "")
    If (expresion = Null) Then
        expresion = ""
    End If
    
    LimpiarStrSql = strResult
    
End Function

