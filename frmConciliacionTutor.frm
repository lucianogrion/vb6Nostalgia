VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConciliacionTutor 
   Caption         =   "Conciliacion Tutoria-Facturacion"
   ClientHeight    =   10665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10665
   ScaleWidth      =   12240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCUIL 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5520
      MaxLength       =   15
      TabIndex        =   7
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   8040
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton CopiarFact 
      Caption         =   "<Ctrl>+C"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton cmdCopyEval 
      Caption         =   "<Ctrl>+C"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   10200
      Width           =   855
   End
   Begin VB.CommandButton cmdCopyConciliacion 
      Caption         =   "<Ctrl>+C"
      Height          =   375
      Left            =   11040
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox cmbTutor 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin MSFlexGridLib.MSFlexGrid flxEvaluaciones 
      Height          =   3495
      Left            =   1080
      TabIndex        =   4
      Top             =   7080
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   1
      Cols            =   13
      FixedCols       =   0
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid flxFacturacion 
      Height          =   2175
      Left            =   1080
      TabIndex        =   5
      Top             =   4800
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxConciliacion 
      Height          =   3375
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5953
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblSelec 
      Caption         =   "Seleccione"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Seleccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Facturado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Evaluado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label lblNombre 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   600
      Width           =   3735
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   11880
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label lblError 
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
   Begin VB.Line Line2 
      X1              =   7920
      X2              =   7920
      Y1              =   4560
      Y2              =   4800
   End
   Begin VB.Line Line3 
      X1              =   11040
      X2              =   11040
      Y1              =   4560
      Y2              =   6960
   End
   Begin VB.Line Line4 
      X1              =   11040
      X2              =   11400
      Y1              =   6960
      Y2              =   6480
   End
   Begin VB.Line Line5 
      X1              =   11040
      X2              =   10560
      Y1              =   6960
      Y2              =   6480
   End
End
Attribute VB_Name = "frmConciliacionTutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbTutor_Change()
 'Buscar y mostrar cuil
    
    Dim intX As Integer
    
    Dim splited() As String
    
    splited = Split(Me.cmbTutor.Text, "-", -1, vbTextCompare)
    
    If UBound(splited) <> 0 Then
       For intX = 0 To UBound(splited)
        Me.txtCUIL.Text = splited(1)
        Me.lblNombre.Caption = splited(0)
       Next
    End If
End Sub


Private Sub cmbTutor_Click()
 'Buscar y mostrar cuil
    
    Dim intX As Integer
    
    Dim splited() As String
    
    splited = Split(Me.cmbTutor.Text, "-", -1, vbTextCompare)
    
    If UBound(splited) <> 0 Then
       For intX = 0 To UBound(splited)
        Me.txtCUIL.Text = splited(1)
        Me.lblNombre.Caption = splited(0)
       Next
    End If
End Sub

    

Private Sub cmdBuscar_Click()
  If (txtCUIL.Text <> "") Then
        Dim strTutor As String
        strTutor = Replace(txtCUIL.Text, "-", "")
        
        Me.lblNombre.Caption = CargarflxTutor(strTutor, flxEvaluaciones, 0, 0)
        
        Call CargarflxFacturacionTutor(strTutor, flxFacturacion, 0, 0)
        'Call CargarflxConciliacionTutor(strTutor, flxConciliacion)
    Else
        lblError.Caption = "Debe Seleccionar Tutor"
    End If
End Sub

Private Sub Form_Load()
    Call CargarComboTutor(Me.cmbTutor)
    
End Sub
