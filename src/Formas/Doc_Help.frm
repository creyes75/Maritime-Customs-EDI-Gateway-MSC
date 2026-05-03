VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form doc_HELP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda de ..."
   ClientHeight    =   4200
   ClientLeft      =   3105
   ClientTop       =   2565
   ClientWidth     =   4545
   ClipControls    =   0   'False
   Icon            =   "Doc_Help.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4545
   Begin VB.CommandButton Cmd_canc 
      Height          =   510
      Left            =   3975
      Picture         =   "Doc_Help.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cancelar"
      Top             =   3645
      Width           =   510
   End
   Begin VB.CommandButton Cmd_acep 
      Cancel          =   -1  'True
      Height          =   510
      Left            =   3405
      Picture         =   "Doc_Help.frx":052C
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Aceptar"
      Top             =   3645
      Width           =   510
   End
   Begin VB.Frame Fra_busq 
      Caption         =   "Busqueda"
      ForeColor       =   &H000000C0&
      Height          =   2820
      Left            =   105
      TabIndex        =   5
      Top             =   765
      Width           =   4380
      Begin MSGrid.Grid Grd_help 
         Height          =   2550
         Left            =   60
         TabIndex        =   6
         Top             =   210
         Width           =   4275
         _Version        =   65536
         _ExtentX        =   7541
         _ExtentY        =   4498
         _StockProps     =   77
         BackColor       =   16777215
         FixedCols       =   0
      End
   End
   Begin VB.CommandButton Cmd_busq 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4110
      Picture         =   "Doc_Help.frx":07AE
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Buscar"
      Top             =   270
      Width           =   300
   End
   Begin VB.Frame Fra_crit 
      Caption         =   "Criterios de Búsqueda"
      ForeColor       =   &H000000C0&
      Height          =   600
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   2385
      Begin VB.OptionButton Opt_desc 
         Caption         =   "Descripción"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   255
         Width           =   1200
      End
      Begin VB.OptionButton Opt_codi 
         Caption         =   "Código"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   75
         TabIndex        =   1
         Top             =   255
         Width           =   915
      End
   End
   Begin VB.TextBox Txt_crit 
      Height          =   285
      Left            =   2535
      TabIndex        =   3
      Top             =   270
      Width           =   1515
   End
End
Attribute VB_Name = "doc_HELP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_acep_Click()
   Grd_help.Col = 0
   
   If Trim$(Grd_help.Text) <> "" And Grd_help.Row <> 0 Then
      Gtyp_ayuda.valor_devuelto = Grd_help.Text
   Else
      Gtyp_ayuda.valor_devuelto = ""
   End If
   Unload Me
End Sub

Private Sub Cmd_busq_Click()
Dim sql      As String
Dim lr_cons  As New ADODB.Recordset
Dim Li_ind   As Integer
'-----------------------------------------------------------------'
Li_ind = 1
Grd_help.Rows = 2
Grd_help.Col = 0: Grd_help.Text = "Código":      Grd_help.ColWidth(0) = 1000
Grd_help.Col = 1: Grd_help.Text = "Descripción": Grd_help.ColWidth(1) = 3000
Grd_help.Row = 1: Grd_help.Col = 0: Grd_help.Text = "": Grd_help.Col = 1: Grd_help.Text = ""
Grd_help.Row = 0

If Trim$(Txt_crit.Text) = "" Then
  Gtyp_ayuda.criterio = ""
  Call Form_Load
End If
If Opt_codi.Value Then
  sql = "Select " & Gtyp_ayuda.campo1 & ", " & Gtyp_ayuda.campo2 & " From " & Gtyp_ayuda.tabla & " Where " & Gtyp_ayuda.campo1 & " like '" & Trim$(Txt_crit.Text) & "%'"
ElseIf Opt_desc.Value Then
  sql = "Select " & Gtyp_ayuda.campo1 & ", " & Gtyp_ayuda.campo2 & " From " & Gtyp_ayuda.tabla & " Where " & Gtyp_ayuda.campo2 & " like '" & Trim$(Txt_crit.Text) & "%'"
End If
If Trim$(Gtyp_ayuda.criterio) <> "" Then
   sql = sql & " and " & Gtyp_ayuda.criterio
End If
If Gtyp_ayuda.Serv_rol Then
   'Set lr_cons = Gcon_ROL.OpenRecordset(sql, rdOpenKeyset)
   lr_cons.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
Else
   'Set lr_cons = Gcon_APG.OpenRecordset(sql, rdOpenKeyset)
   lr_cons.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
End If
If lr_cons.EOF Then
  Exit Sub
End If
Do While Not lr_cons.EOF
   Grd_help.Rows = Li_ind + 1
   Grd_help.Row = Li_ind
   Grd_help.Col = 0: Grd_help.Text = lr_cons(0)
   Grd_help.Col = 1: Grd_help.Text = lr_cons(1)
   Li_ind = Li_ind + 1
   lr_cons.MoveNext
Loop
If Grd_help.Enabled Then Grd_help.SetFocus
End Sub

Private Sub Cmd_canc_Click()
  Gtyp_ayuda.valor_devuelto = ""
  Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo errores_ayuda
Dim sql     As String
Dim lr_cons As New ADODB.Recordset
Dim Li_ind  As Integer

Call Centra_Pantalla(doc_HELP)
Me.Caption = "Ayuda de " & Trim$(Gtyp_ayuda.nombre)
Opt_codi.Value = True
Li_ind = 1
Grd_help.Rows = 2
Grd_help.Row = 0
Grd_help.Col = 0: Grd_help.Text = "Código":      Grd_help.ColWidth(0) = 1000
Grd_help.Col = 1: Grd_help.Text = "Descripción": Grd_help.ColWidth(1) = 3000

sql = "Select " & Gtyp_ayuda.campo1 & ", " & Gtyp_ayuda.campo2 & " from " & Gtyp_ayuda.tabla
If Trim$(Gtyp_ayuda.criterio) <> "" Then
   sql = sql & " Where " & Gtyp_ayuda.criterio
End If
'If Gtyp_ayuda.Serv_rol Then
   'Set lr_cons = Gcon_docu.OpenRecordset(sql, rdOpenKeyset)
   lr_cons.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
'Else
'   Set lr_cons = Gcon_APG.OpenRecordset(sql, rdOpenKeyset)
'End If

If lr_cons.EOF Then
  Exit Sub
End If
Do While Not lr_cons.EOF
   Grd_help.Rows = Li_ind + 1
   Grd_help.Row = Li_ind
   Grd_help.Col = 0: Grd_help.Text = lr_cons(0)
   Grd_help.Col = 1: Grd_help.Text = lr_cons(1)
   Li_ind = Li_ind + 1
   lr_cons.MoveNext
Loop
Exit Sub
errores_ayuda:
  MsgBox "Ocurrió el error " & Err.Description, vbExclamation, App.Title
  Exit Sub
  Resume 0
End Sub

Private Sub Grd_help_DblClick()
Call Cmd_acep_Click
End Sub

Private Sub Grd_help_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Cmd_acep_Click
End Sub

Private Sub Txt_crit_KeyPress(KeyAscii As Integer)
   'If Chr(KeyAscii) = "*" Then KeyAscii = 0
   If KeyAscii = 13 Then Cmd_busq_Click
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
