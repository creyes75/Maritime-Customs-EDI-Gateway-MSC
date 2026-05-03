VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form Seguridad 
   Caption         =   "Seguridad de Documentación"
   ClientHeight    =   1950
   ClientLeft      =   3240
   ClientTop       =   3180
   ClientWidth     =   3495
   Icon            =   "Doc_SEGUR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1950
   ScaleWidth      =   3495
   Begin Threed.SSPanel SSPanel1 
      Height          =   1905
      Left            =   30
      TabIndex        =   4
      Top             =   15
      Width           =   3435
      _Version        =   65536
      _ExtentX        =   6059
      _ExtentY        =   3360
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton btn_continuar 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   285
         TabIndex        =   2
         Top             =   1410
         Width           =   1095
      End
      Begin VB.CommandButton btn_cancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2115
         TabIndex        =   3
         Top             =   1410
         Width           =   1095
      End
      Begin VB.TextBox txt_usuario 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1590
         MaxLength       =   20
         TabIndex        =   0
         Text            =   " "
         Top             =   315
         Width           =   1605
      End
      Begin VB.TextBox txt_password 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1575
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   855
         Width           =   1620
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "Doc_SEGUR.frx":2072
         Top             =   240
         Width           =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   45
         X2              =   3390
         Y1              =   1290
         Y2              =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   660
         TabIndex        =   6
         Top             =   375
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   630
         TabIndex        =   5
         Top             =   885
         Width           =   915
      End
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   105
      Stretch         =   -1  'True
      Top             =   75
      Width           =   840
   End
End
Attribute VB_Name = "Seguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intcont As Integer

Private Sub BTN_CANCELAR_Click()
    End
End Sub

Private Sub BTN_CANCELAR_LostFocus()
    txt_usuario.SetFocus
End Sub

Private Sub BTN_CONTINUAR_Click()
Dim segur        As String
Dim reg          As New ADODB.Recordset
Dim rc, rc2      As Integer

    On Error GoTo LError
    'intcont = 0
    If Trim$(txt_usuario) = "" Then
        MsgBox "Por favor ingrese el Usuario", 48, Me.Caption:        txt_usuario.SetFocus:        Exit Sub
    End If
'    If Trim$(txt_password) = "" Then
'        MsgBox "Por favor ingrese el Password", 48, Me.Caption:        txt_password.SetFocus:      Exit Sub
'    End If
    Gstr_user = Trim$(txt_usuario.Text)
    Gstr_psw = Trim$(txt_password.Text)
    Screen.MousePointer = vbHourglass
    
salir:
    Screen.MousePointer = vbDefault
    Unload Me
    Exit Sub
    
LError:
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " " & Err.Description, vbInformation, App.Title
    txt_usuario.SelStart = 0
    txt_usuario.SelLength = Len(txt_usuario)
    txt_usuario.SetFocus
    Exit Sub
    Resume 0
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 0
   Me.Top = (Screen.Height - Me.Height) / 2
   Me.Left = (Screen.Width - Me.Width) / 2
   Ver_msg = True
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub txt_password_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case 38, 37 'KEY_UP, KEY_LEFT
         txt_usuario.SetFocus
    Case 40, 39 'KEY_DOWN, KEY_RIGHT
         btn_continuar.SetFocus
    End Select
    
End Sub

Private Sub txt_password_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then
        btn_continuar.SetFocus
    End If
End Sub

Private Sub txt_usuario_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case 38, 37 'KEY_UP, KEY_LEFT
         KeyCode = 0
    Case 40, 39 'KEY_DOWN, KEY_RIGHT
         txt_password.SetFocus
    End Select
    
End Sub

Private Sub txt_usuario_KeyPress(KeyAscii As Integer)
    'KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then
        txt_password.SetFocus
    End If
End Sub
