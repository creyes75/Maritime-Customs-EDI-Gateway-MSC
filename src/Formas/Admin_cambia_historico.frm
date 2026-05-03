VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Admin_cambia_historico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambia Conexión a base histórica"
   ClientHeight    =   825
   ClientLeft      =   4230
   ClientTop       =   2640
   ClientWidth     =   4455
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   825
   ScaleWidth      =   4455
   Begin VB.Frame Fra_linea 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00000080&
      Height          =   540
      Left            =   105
      TabIndex        =   3
      Top             =   60
      Width           =   2820
      Begin MSDataListLib.DataCombo Cmb_linea 
         Height          =   315
         Left            =   165
         TabIndex        =   0
         Top             =   210
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc data_linea 
         Height          =   330
         Left            =   1770
         Top             =   255
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Histoico 
         Caption         =   "Histórico"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   105
         TabIndex        =   4
         Top             =   0
         Width           =   825
      End
   End
   Begin Threed.SSCommand btn_sgte 
      Height          =   495
      Left            =   3030
      TabIndex        =   1
      ToolTipText     =   "Continuar"
      Top             =   120
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   78
      Picture         =   "Admin_cambia_historico.frx":0000
   End
   Begin Threed.SSCommand btn_elim 
      Height          =   495
      Left            =   3570
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   495
      _Version        =   65536
      _ExtentX        =   868
      _ExtentY        =   868
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Admin_cambia_historico.frx":0452
   End
End
Attribute VB_Name = "Admin_cambia_historico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_elim_Click()
    Unload Me
End Sub

'===================================================================================
Private Sub btn_sgte_Click()
On Error GoTo Errores
If Cmb_linea.MatchedWithList Then
  Call cambio_base(Cmb_linea.BoundText)
End If
Exit Sub
Errores:
    MsgBox "ERROR " & Err.Description
    Exit Sub
    Resume 0
End Sub

Private Sub Cmb_linea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub Form_Load()
On Error GoTo Errores
    'carga_datacombo data_linea, Cmb_linea, "Select * from docu_base_historica", "hist_codigo", "hist_descripcion"
    carga_datacombo data_linea, Cmb_linea, "Select * from master..sysdatabases where name like '" & Mid(Gstr_base_doc, 1, 4) & "%'", "name", "name"
    If Not data_linea.Recordset.EOF Then data_linea.Recordset.MoveFirst
    Exit Sub
Errores:
    MsgBox "Ocurrio el Error " & Err.Description, vbCritical
    Exit Sub
End Sub
