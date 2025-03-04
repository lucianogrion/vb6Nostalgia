VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Contador 
   Caption         =   "Sistema de Control Accesorio"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin MSACAL.Calendar Calendar1 
      Height          =   2415
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   5055
      _Version        =   524288
      _ExtentX        =   8916
      _ExtentY        =   4260
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2010
      Month           =   8
      Day             =   26
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   2
      GridCellEffect  =   0
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog commonDialog 
      Left            =   5280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox lstPersonas 
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular Cantidad"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Ingrese Persona"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese Fecha Alta"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Contador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnFox As ADODB.Connection
Dim rsFox As ADODB.Recordset



Private Sub Command1_Click()

Screen.MousePointer = 11


Set cnFox = New ADODB.Connection
cnFox.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver={Microsoft Visual FoxPro Driver};UID=;SourceDB=" & App.Path & "\;SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"

Dim strConexion As String
strConexion = strConexion + "[ODBC]"
strConexion = strConexion + "" + "DRIVER=Microsoft dBase Driver (*.dbf)"
 strConexion = strConexion + ";" + "UID = admin"
 strConexion = strConexion + ";" + "UserCommitSync = Yes"
 strConexion = strConexion + ";" + "Threads = 3"
 strConexion = strConexion + ";" + "Statistics = 0"
 strConexion = strConexion + ";" + "SafeTransactions = 0"
strConexion = strConexion + ";" + "PageTimeout = 5"
strConexion = strConexion + ";" + "MaxScanRows = 8"
strConexion = strConexion + ";" + "MaxBufferSize = 2048"
strConexion = strConexion + ";" + "FIL=dBase 5.0"
strConexion = strConexion + ";" + "DriverId = 533"
strConexion = strConexion + ";" + "Deleted = 0"
strConexion = strConexion + ";" + "DefaultDir=O:\RUN"
strConexion = strConexion + ";" + "DBQ=O:\RUN\certifica.DBF"
strConexion = strConexion + ";" + "CollatingSequence = AscII"

'cnFox.ConnectionString = strConexion

cnFox.Open

Set rsFox = New ADODB.Recordset
rsFox.Open "select * from Certifica", cnFox, adOpenForwardOnly, adLockReadOnly



Set cnExcel = New ADODB.Connection
cnExcel.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\TPruebaXLS.xls;Persist Security Info=False; Extended Properties=Excel 8.0;"
cnExcel.Open


Set rsExcel = New ADODB.Recordset
rsExcel.CursorLocation = adUseClient
rsExcel.Open "select * from [Data$] where CUIL = ''", cnExcel, adOpenKeyset, adLockOptimistic



While Not rsFox.EOF
    rsExcel.AddNew

    For i = 0 To rsExcel.Fields.Count - 1
        rsExcel(rsFox(i).Name) = rsFox(i)
    Next i
    rsExcel.Update
    rsFox.MoveNext
Wend


' Utiliza este codigo para eliminar los datos de la tabla en fox

'If rsFox.RecordCount > 0 Then
'  rsFox.MoveFirst
' While Not rsFox.EOF
'  rsFox.Delete
'  rsFox.MoveNext
' Wend
'End If



rsExcel.Close
Set rsExcel = Nothing

rsFox.Close
Set rsFox = Nothing

cnExcel.Close
Set cnExcel = Nothing

cnFox.Close
Set cnFox = Nothing

Screen.MousePointer = 0

MsgBox ("Datos Transferidos"), vbInformation

End Sub

End Sub

Private Sub Form_Load()

End Sub
