VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportacionFacturacion 
   Caption         =   "Importacion Facturacion"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImport 
      Caption         =   "Importar"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy PASTE"
      Height          =   735
      Left            =   6720
      TabIndex        =   6
      Top             =   2280
      Width           =   735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ObtenerDifEvaluadores"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "frmImportacionFacturacion.frx":0000
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7435
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      AllowUserResizing=   3
   End
   Begin VB.TextBox Text1 
      DataField       =   "evaluadorcuil"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   7335
   End
   Begin MSComDlg.CommonDialog comDialog 
      Left            =   6960
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtDir 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "..."
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblError 
      Caption         =   "Status:"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   6600
      Width           =   7215
   End
   Begin VB.Label Label2 
      Caption         =   "Log"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccione Archivo .xls"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmImportacionFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImport_Click()
      
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
    
            
            Dim ultFila As Long
            ultFila = obtenerUltimaFilaExcelPosicion(txtDir.Text, 5)
            MsgBox "Se importarán " & ultFila & vbCrLf & " Archivo : " & txtDir.Text, vbInformation + vbOKOnly, "Atencion"
            
            
            Call AbrirConexionXLS(txtDir.Text, conxls)
            Set rstXls = New ADODB.Recordset
            With rstXls
             .CursorLocation = adUseClient
             .CursorType = adOpenStatic
             .LockType = adLockOptimistic
            End With
            
            '------------------------------------------------------Alta de Evaluacion
            rstXls.Open "SELECT * FROM [evaluadores$A5:m5" & ultFila & "]", conxls, adOpenDynamic, adLockOptimistic
            Do While Not rstXls.EOF
                    'Armado del Insert Into
                    sql = ""
                    sql = sql & "INSERT INTO Facturacion ("
                    sql = sql & "nroFactura, Cantidad, CUIL, Mesl, aniol, fechaing"
                    sql = sql & ") values ("
                    
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(3)) & "',"   'N° FACTURA
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(8)) & "',"   'EVALUACIONES REALIZADAS
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(1)) & "',"   'CUIL
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(5)) & "',"   'mesl
                    sql = sql & "'" & LimpiarStrSql(rstXls.Fields(6)) & "',"   'aniol
                    
                    
                    If IsNull(rstXls.Fields(2)) Then
                        strFecha = "1/1/1900"
                    Else
                         If Not IsDate(rstXls.Fields(2)) Then
                            strFecha = "1/1/1900"
                        Else
                            strFecha = Format(rstXls.Fields(2), "MM/DD/YYYY HH:MM:SS")
                        End If
                    End If
                    
                    sql = sql & "'" & strFecha & "'"   'aniol
                    
                    sql = sql & ")"
                    Call con.Execute(sql)
                    
                rstXls.MoveNext
            Loop
            rstXls.Close
            '------------------------------------------------------'Alta de Evaluacion
            
            lblError.Caption = "Importado exitosamente " + pathCompleto
            DoEvents
            
            
            
        
        
        lblError.Caption = "La Importacion Ha sido Exitosa!!!"
End Sub

Private Sub cmdSelect_Click()
    With comDialog
        .InitDir = "L:\Prueba\Conciliacion Facturacion"
        .DialogTitle = " Seleccionar archivo Excel para cargar"
        .Filter = "Archivos XLS|*.xls"
        .ShowOpen
        
        If .FileName = "" Then Exit Sub

        Me.txtDir.Text = .FileName
    End With
End Sub

Private Sub Command1_Click()
    Dim s As String
    Dim i As Integer, j As Integer
    
    s = ""
    With MSFlexGrid1
    'almacenamos la region seleecionada a un string
    'para que funcione en el formato de excel las colummas
    'van separadas por tab's y los renglones por CrLf
    For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
            If j > 0 Then s = s & vbTab
            s = s & MSFlexGrid1.TextMatrix(i, j)
        Next j
    s = s & vbCrLf
    Next i
    
    
    Clipboard.Clear
    Clipboard.SetText s
    
    'si quieres pasar toda la grid modifica el for por
    
    End With

End Sub

Private Sub Form_Load()
    Me.Data1.DatabaseName = sPathBase
End Sub
