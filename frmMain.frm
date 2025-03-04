VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Sistema Auxiliar de Conciliacion De Facturacion"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5520
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFactTutor 
      Caption         =   "Importar Fact Tutor"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdImportacionFact 
      Caption         =   "Importar Fact Evaluador"
      Height          =   735
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdConciliarTutores 
      Caption         =   "Control Tutores"
      Height          =   735
      Left            =   2880
      TabIndex        =   3
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton cmdEvaluaciones 
      Caption         =   "Importar Evaluaciones"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5295
   End
   Begin VB.CommandButton cmdConciliacion 
      Caption         =   "Control Evaluadores"
      Height          =   735
      Left            =   2880
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton cmdFacturacion 
      Caption         =   "Ingreso  Manual Facturacion"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5520
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5400
      Y1              =   2280
      Y2              =   2280
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConciliacion_Click()
    frmConciliacion.Show vbModal
    
End Sub

Private Sub cmdConciliarTutores_Click()
    frmConciliacionTutor.Show vbModal
    
End Sub

Private Sub cmdEvaluaciones_Click()
    frmImportacionExcel.Show vbModal
End Sub

Private Sub cmdFacturacion_Click()
    frmFacturacion.Show vbModal
End Sub

Private Sub cmdImportacionFact_Click()
    frmImportacionFacturacion.Show vbModal
    
End Sub

Private Sub Form_Load()
    
    'Dim con As ADODB.Connection
    'Call AbrirConexionDB(con)
        
        'Dim rst As ADODB.Recordset
        'Set rst = New ADODB.Recordset
        
        
        'rst.Open "SELECT max(FechaIng) as maxFecha FROM Facturacion", con, adOpenDynamic, adLockOptimistic
        'Me.lblUltimaFac.Caption = rst.Fields("maxFecha") & ""
        'rst.Close
        'Set rst = Nothing
        
        'rst.Open "SELECT max(isnull(fechaIngreso,0)) as maxFecha FROM Evaluacion", cnn, adOpenDynamic, adLockOptimistic
        'Me.lblUltimaEval.Caption = rst.Fields("maxFecha") & ""
        'rst.Close
        
        
    'con.Close
    'Set con = Nothing
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rst = Nothing
    Set cnn = Nothing
End Sub

Private Sub Label2_Click()

End Sub
