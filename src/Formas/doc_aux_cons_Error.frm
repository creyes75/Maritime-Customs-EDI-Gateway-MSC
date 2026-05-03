VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form doc_aux_cons_Error 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Consulta de Errores de CAE"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "doc_aux_cons_Error.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2355
   ScaleWidth      =   5415
   Begin VB.TextBox Txt_error 
      Height          =   285
      Left            =   1515
      TabIndex        =   0
      Top             =   195
      Width           =   1290
   End
   Begin Threed.SSCommand cmd_consulta 
      Default         =   -1  'True
      Height          =   525
      Left            =   3120
      TabIndex        =   2
      Top             =   15
      Width           =   555
      _Version        =   65536
      _ExtentX        =   979
      _ExtentY        =   926
      _StockProps     =   78
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      AutoSize        =   1
      Picture         =   "doc_aux_cons_Error.frx":628A
   End
   Begin Threed.SSCommand cmd_salir 
      Cancel          =   -1  'True
      Height          =   525
      Left            =   3750
      TabIndex        =   3
      Top             =   15
      Width           =   555
      _Version        =   65536
      _ExtentX        =   979
      _ExtentY        =   926
      _StockProps     =   78
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      Picture         =   "doc_aux_cons_Error.frx":66DC
   End
   Begin VB.Frame Frame1 
      Height          =   1770
      Left            =   60
      TabIndex        =   5
      Top             =   525
      Width           =   5370
      Begin VB.TextBox Txt_descripcion 
         Height          =   660
         Left            =   195
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   555
         Width           =   4950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción del Error"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   195
         TabIndex        =   7
         Top             =   285
         Width           =   1470
      End
      Begin VB.Label Lbl_tipo 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Error"
         Height          =   195
         Left            =   225
         TabIndex        =   6
         Top             =   1395
         Width           =   4785
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Código de Error"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   270
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "doc_aux_cons_Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_consulta_Click()
Dim reg_error As New ADODB.Recordset
If Trim(Txt_error.Text) <> "" And IsNumeric(Txt_error.Text) Then
    sql = "Select * from docu_error Where error_codigo = '" & Trim(CDbl(Txt_error.Text)) & "'"
    reg_error.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    If Not reg_error.EOF Then
        Txt_descripcion.Text = Trim(reg_error("error_descripcion"))
        Lbl_tipo.Caption = Trim(reg_error("error_tipo"))
    Else
        Txt_descripcion.Text = ""
        Lbl_tipo.Caption = ""
    End If
    reg_error.Close
End If
End Sub

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub Txt_error_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
