VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportacionExcel 
   Caption         =   "Evaluaciones Ingresadas Via Excel"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   Icon            =   "frmImportacionExcel.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   10005
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog comDialog 
      Left            =   8280
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLevantarXLS 
      Caption         =   "Levantar XLS Marcados"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton cmdRelevarXLS 
      Caption         =   "Relevar XLS"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox txtlista 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   720
      Width           =   9855
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "..."
      Height          =   255
      Left            =   6240
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtDir 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Text            =   "L:\Prueba\Conciliacion Facturacion\Base de datos 2011"
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label lblError 
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   9375
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccione Directorio"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmImportacionExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdLevantarXLS_Click()
    
    Dim con As ADODB.Connection
    Dim conxls As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim rstXls As ADODB.Recordset
    
    Dim sql As String
    Dim rta As Boolean
    Dim pathCompleto As String
    Dim strFecha As String
    
    
    'Call AbrirConexionDB(con)
    Set rst = New ADODB.Recordset
    Call AbrirConexionDB(con)
    
     rst.Open "SELECT pathCompleto  FROM Archivo where cambio=true", con, adOpenDynamic, adLockOptimistic
     
        Do While Not rst.EOF
            'Borrar evaluaciones Pre existentes
            pathCompleto = rst.Fields("pathCompleto")
            sql = ""
            sql = sql & "DELETE from evaluacion where pathCompleto='" & pathCompleto & "'"
            Call con.Execute(sql)
            
            Dim ultFila As Long
            ultFila = obtenerUltimaFilaExcel(rst.Fields("pathcompleto"))
            MsgBox "Se importarán " & ultFila & vbCrLf & " Archivo : " & pathCompleto, vbInformation + vbOKOnly, "Atencion"
            
            
            Call AbrirConexionXLS(rst.Fields("pathcompleto"), conxls)
            Set rstXls = New ADODB.Recordset
            With rstXls
             .CursorLocation = adUseClient
             .CursorType = adOpenStatic
             .LockType = adLockOptimistic
            End With
            
            '------------------------------------------------------Alta de Evaluacion
            rstXls.Open "SELECT * FROM [Base de datos$A1:AO" & ultFila & "]", conxls, adOpenDynamic, adLockOptimistic
            Do While Not rstXls.EOF
                    'Armado del Insert Into
                    sql = ""
                    sql = sql & "INSERT INTO Evaluacion ("
                    sql = sql & "pathCompleto, nro, Accion, PersonaNombre, DocumentoTipo, DocumentoNro, CUIL, Direccion, Localidad, Provincia, ProvinciaCodigo, Municipio, MunicipioCodigo, PostalCodigo, Telefono, NacimientoFecha, Nacionalidad, NacimientoPais, Sexo, NivelEducativo, RazonSocial, EvaluacionFecha, EvaluacionLugar, EvaluacionProvincia, EvaluacionLocalidad, EvaluacionDireccion, NormaTitulo, NormaCodigo, EvaluacionResultado, EvaluadorNombe, EvaluadorCuil, FechaCertif, FuenteFinanciamiento, Protocolo, Convenio, Tutor, Registro, IronMountain, Caja, Observaciones, EntregadoMteyss"
                    sql = sql & ") values ("
                    sql = sql & "'" & rst.Fields("pathcompleto") & "',"
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(0)) & "',"     'Nro
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(1)) & "',"   'Accion
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(2)) & "',"   'PersonaNombre
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(3)) & "',"   'DocumentoTipo
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(4) & "") & "',"  'DocumentoNro
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(5) & "") & "',"  'CUIL
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(6) & "") & "',"  'Direccion
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(7) & "") & "'," 'Localidad
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(8)) & "',"   'Provincia
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(9)) & "',"   'ProvinciaCodigo
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(10) & "") & "'," 'Municipio
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(11)) & "',"  'MunicipioCodigo
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(12) & "") & "'," 'PostalCodigo
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(13) & "") & "',"  'Telefono
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(14) & "") & "'," 'NacimientoFecha
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(15)) & "',"  'Nacionalidad,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(16)) & "',"  'NacimientoPais,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(17)) & "',"  'Sexo,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(18) & "") & "'," 'NivelEducativo,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(19)) & "',"  'RazonSocial,
                    
                    If IsNull(rstXls.Fields(20)) Then
                        strFecha = Format(Now, "MM/DD/YYYY HH:MM:SS")
                    Else
                        If Not IsDate(rstXls.Fields(20)) Then
                            strFecha = "1/1/1900"
                        Else
                            strFecha = Format(rstXls.Fields(20), "MM/DD/YYYY HH:MM:SS")
                        End If
                    End If
                    
                    sql = sql & "#" & strFecha & "#,"  'EvaluacionFecha,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(21)) & "',"   'EvaluacionLugar,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(22)) & "',"  'EvaluacionProvincia,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(23)) & "',"  'EvaluacionLocalidad,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(24) & "") & "'," 'EvaluacionDireccion,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(25) & "") & "'," 'NormaTitulo,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(26) & "") & "'," 'NormaCodigo,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(27) & "") & "'," 'EvaluacionResultado,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(28) & "") & "'," 'EvaluadorNombe,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(29) & "") & "'," 'EvaluadorCuil,
                    
                    If IsNull(rstXls.Fields(30)) Then
                        strFecha = "1/1/1900"
                    Else
                         If Not IsDate(rstXls.Fields(30)) Then
                            strFecha = "1/1/1900"
                        Else
                            strFecha = Format(rstXls.Fields(30), "MM/DD/YYYY HH:MM:SS")
                        End If
                        
                        
                    End If
                    sql = sql & "#" & strFecha & "#,"  'FechaCertif,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(31) & "") & "',"  'FuenteFinanciamiento,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(32) & "") & "'," 'Protocolo,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(33) & "") & "'," 'Convenio,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(34) & "") & "'," 'Tutor,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(35) & "") & "'," 'Registro,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(36) & "") & "'," 'IronMountain,
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(37) & "") & "'," 'Caja,
                    'sql = sql & "'" & rstxls.Fields(38) & "',"  'NADA
                    sql = sql & "'" & rstXls.Fields(39) & "',"  'Observaciones,
                    
                    'Problemas Con campo EntregadoMteyss
'                    If rstXls.Fields.Count > 39 Then
'                        If IsNull(rstXls.Fields(40)) Then
'                            sql = sql & "'" & "'"    'EntregadoMteyss
'                        Else
'                            sql = sql & "'" & LimpiarStrSql(rstXls.Fields(40)) & "'"  'EntregadoMteyss
'                         End If
'                    Else
'                        sql = sql & "'" & "'"    'EntregadoMteyss
'                    End If
                    sql = sql & "'" & "'"    'EntregadoMteyss
                    
                    sql = sql & ")"
                    Call con.Execute(sql)
                    
                rstXls.MoveNext
            Loop
            rstXls.Close
            '------------------------------------------------------'Alta de Evaluacion
            rst.MoveNext
            Call MarcarArchivo(pathCompleto, ObtenerFechaMod(pathCompleto))
            lblError.Caption = "Importado exitosamente " + pathCompleto
            DoEvents
            
            
            
        Loop
        
        lblError.Caption = "La Importacion Ha sido Exitosa!!!"
        
End Sub



Private Sub cmdRelevarXLS_Click()
 Dim MyFile As String, Sep As String
    Dim mypath As String
    
    Me.txtlista.Text = ""
    mypath = txtDir.Text
    
    ' Test for Windows or Macintosh platform. Make the directory request.
    Sep = "\"
    
    MyFile = Dir(mypath & Sep & "*.xls")
    
    
    Do While MyFile <> ""
       'MsgBox mypath & Sep & MyFile
       Dim pathCompleto As String
       pathCompleto = mypath & Sep & MyFile
       
       Dim rta As Boolean
       rta = RelevoArchivo(pathCompleto, ObtenerFechaMod(pathCompleto))
       
       Me.txtlista.Text = Me.txtlista.Text & vbCrLf & pathCompleto & "Marcado=" & rta
       
       
       MyFile = Dir()
    Loop
    
End Sub

Private Sub cmdSeleccionar_Click()
    With comDialog
        .InitDir = "C:\"
        .DialogTitle = " Seleccionar archivo Excel para cargar"
        .Filter = "Archivos XLS|*.xls"
        .ShowOpen
        
        If .FileName = "" Then Exit Sub

        Me.txtDir.Text = ObtenerSoloDir(.FileName)
    End With
    
End Sub






