VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFacturacion 
   Caption         =   "Facturacion"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   Icon            =   "frmFacturacion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox txtCUIL 
      Height          =   375
      Left            =   1200
      TabIndex        =   13
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      AllowPrompt     =   -1  'True
      MaxLength       =   13
      Mask            =   "##-########-#"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.Frame fraIngreso 
      Caption         =   "Ingreso de Facturacion"
      Height          =   5655
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar Factura/s"
         Height          =   735
         Left            =   6360
         TabIndex        =   18
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtFactura 
         Height          =   285
         Left            =   960
         TabIndex        =   15
         Top             =   840
         Width           =   5055
      End
      Begin VB.ComboBox cmbAnio 
         Height          =   315
         ItemData        =   "frmFacturacion.frx":030A
         Left            =   2280
         List            =   "frmFacturacion.frx":032C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cmbMes 
         Height          =   315
         ItemData        =   "frmFacturacion.frx":034E
         Left            =   960
         List            =   "frmFacturacion.frx":0379
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtCantidad 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "5"
         Top             =   1530
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Agregar"
         Height          =   615
         Left            =   4560
         TabIndex        =   1
         Top             =   1200
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid flxFacturacion 
         Height          =   3135
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
      End
      Begin VB.Label lblfact 
         Caption         =   "Factura"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Facturacion Registrada"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label lblRol 
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Rol"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Periodo"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   2040
         X2              =   2160
         Y1              =   1440
         Y2              =   1200
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   855
      End
   End
   Begin VB.Label lblError 
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "CUIL"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Persona"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblPersona 
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   840
      Width           =   5295
   End
End
Attribute VB_Name = "frmFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdBuscar_Click()
    Dim con As ADODB.Connection
    Dim cuil As String
    
    Set rst = New ADODB.Recordset
      
    Call AbrirConexionDB(con)
        
        cuil = Replace(txtCUIL.Text, "-", "")
        rst.Open "SELECT evaluadornombe,evaluadorcuil  FROM evaluacion where evaluadorcuil='" + cuil + "'", con, adOpenDynamic, adLockOptimistic
        
        If Not rst.EOF Then
            Me.lblPersona.Caption = rst.Fields("evaluadornombe") + " "
            Me.lblRol.Caption = "Evaluador"
            fraIngreso.Visible = True
            Call CargarflxFacturacion(cuil, Me.flxFacturacion, 0, 0)
        Else
            Me.lblPersona.Caption = "CUIL NO ENCONTRADO"
            fraIngreso.Visible = False
        End If
        
    con.Close
    Set con = Nothing
    
End Sub

Private Sub fraPersonales_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub cmdQuitar_Click()
    'todo: Hacer que borre
    MsgBox "Todavia no implementado"
    
End Sub

Private Sub Command2_Click()
    Dim con As ADODB.Connection
    Dim sql As String
    Dim cuil As String
    
    'TODO: Agregar que valide que este todo ingresado
    
    Call AbrirConexionDB(con)
    cuil = Replace(txtCUIL.Text, "-", "")
    
    sql = ""
    sql = sql & " Insert Into Facturacion (NroFactura,FechaIng,Cantidad,Cuil,mesl,aniol) "
    sql = sql & " Values ('" & txtFactura.Text & "',now()," & txtCantidad.Text & ",'" & cuil & "'," & (cmbMes.ListIndex + 1) & "," & cmbAnio & ")"
    con.Execute sql
      
    con.Close
    Set con = Nothing
    
    cuil = Replace(txtCUIL.Text, "-", "")
    Call CargarflxFacturacion(cuil, Me.flxFacturacion, 0, 0)
    
End Sub

