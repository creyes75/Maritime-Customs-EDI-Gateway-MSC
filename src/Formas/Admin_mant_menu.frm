VERSION 5.00
Begin VB.Form Admin_mant_menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Permisos de Menús a usuarios "
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   4155
   Begin VB.CheckBox Chk_todos 
      Caption         =   "Seleccionar Todos"
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   1845
   End
   Begin VB.CommandButton Cmd_modificar 
      Caption         =   "Ok"
      Height          =   660
      Left            =   2490
      Picture         =   "Admin_mant_menu.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Crea Bls con Cnts vacíos"
      Top             =   4335
      Width           =   750
   End
   Begin VB.CommandButton Cmd_salir 
      Caption         =   "Cancel"
      Height          =   660
      Left            =   3255
      Picture         =   "Admin_mant_menu.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   4350
      Width           =   750
   End
   Begin VB.TextBox Txt_user 
      Height          =   330
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   1710
   End
   Begin VB.ListBox Lst_menu 
      Height          =   3660
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   600
      Width           =   3915
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario :"
      Height          =   285
      Left            =   165
      TabIndex        =   1
      Top             =   165
      Width           =   705
   End
End
Attribute VB_Name = "Admin_mant_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Chk_todos_Click()
For i = 0 To Lst_menu.ListCount - 1
    Lst_menu.Selected(i) = Chk_todos.Value
Next i
End Sub

Private Sub cmd_modificar_Click()
sql = "delete from docu_menu where mnu_usuario = '" & Txt_user.Text & "'"
Gcon_docu.Execute sql

For i = 0 To Lst_menu.ListCount - 1
    Lst_menu.ListIndex = i
    If Lst_menu.Selected(i) Then
        sql = "insert into docu_menu values ('" & Txt_user.Text & "','" & Lst_menu.Text & "')"
        Gcon_docu.Execute sql
    End If
Next i
MsgBox "Permisos para menu creados ", vbInformation
End Sub

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Lst_menu.AddItem "mnu_mant"
Lst_menu.AddItem "mnu_mant_linea"
Lst_menu.AddItem "mnu_mant_puerto"
Lst_menu.AddItem "mnu_mant_buque"

Lst_menu.AddItem "mnu_mant_almacen"
Lst_menu.AddItem "mnu_mant_cond_cont"
Lst_menu.AddItem "mnu_mant_cond"
Lst_menu.AddItem "mnu_mant_cons"
Lst_menu.AddItem "mnu_mant_depo"
Lst_menu.AddItem "mnu_mant_emba"
Lst_menu.AddItem "mnu_mant_emp_anav"
Lst_menu.AddItem "mnu_mant_emp_resp"
Lst_menu.AddItem "mnu_mant_impexp"
Lst_menu.AddItem "mnu_mant_indic_llen_vac"
Lst_menu.AddItem "mnu_mant_meto_pago"
Lst_menu.AddItem "mnu_mant_moneda"
Lst_menu.AddItem "mnu_mant_nacio"
Lst_menu.AddItem "mnu_mant_natura_carga"
Lst_menu.AddItem "mnu_mant_imo"
Lst_menu.AddItem "mnu_mant_imo_onu"
Lst_menu.AddItem "mnu_mant_prov_equip"
Lst_menu.AddItem "mnu_mant_regimen"
Lst_menu.AddItem "mnu_mant_requer_serv_trans"
Lst_menu.AddItem "mnu_mant_tamano_equip"
Lst_menu.AddItem "mnu_mant_tipo_carga"
Lst_menu.AddItem "mnu_mant_tipo_deposito"
Lst_menu.AddItem "mnu_mant_tipo_doc"
Lst_menu.AddItem "mnu_mant_tipo_envio"
Lst_menu.AddItem "mnu_mant_tipo_Equipo"
Lst_menu.AddItem "mnu_mant_tipo_salida"

Lst_menu.AddItem "mnu_proceso"
Lst_menu.AddItem "mnu_proc_viaje"
Lst_menu.AddItem "mnu_proc_bl_status"
Lst_menu.AddItem "mnu_proc_impo"
Lst_menu.AddItem "mnu_proc_expo"
Lst_menu.AddItem "mnu_proc_archivo"
Lst_menu.AddItem "mnu_proc_respuesta"
Lst_menu.AddItem "mnu_proc_cons_error"
Lst_menu.AddItem "mnu_proc_dtx"

Lst_menu.AddItem "mnu_reportes"
Lst_menu.AddItem "mnu_repo_manifiesto"
Lst_menu.AddItem "mnu_repo_blWord"
Lst_menu.AddItem "mnu_repo_correc"
Lst_menu.AddItem "contimpo"
Lst_menu.AddItem "mnu_repo_cont_sin_bl"
Lst_menu.AddItem "mnu_repo_aviso"
Lst_menu.AddItem "mnu_repo_flete_impo"
Lst_menu.AddItem "mnu_repo_blWord_expo"
Lst_menu.AddItem "mnu_repo_mani_expo"
Lst_menu.AddItem "mnu_repo_bl_export"
Lst_menu.AddItem "mnu_repo_cont_expo"
Lst_menu.AddItem "mnu_repo_cont_sin_bl_expo"
Lst_menu.AddItem "mnu_repo_flete_Expo"

Lst_menu.AddItem "mnu_adm"
End Sub

Private Sub Txt_user_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_user_LostFocus()
Dim sql As String
Dim reg As New ADODB.Recordset

    For i = 0 To Lst_menu.ListCount - 1
        Lst_menu.Selected(i) = False
    Next i
    
    If Txt_user.Text <> "" Then
        sql = "Select * from docu_menu where mnu_usuario = '" & Txt_user.Text & "'"
        reg.Open sql, Gcon_docu
        Do While Not reg.EOF
            For i = 0 To Lst_menu.ListCount - 1
                Lst_menu.ListIndex = i
                If Trim(Lst_menu.Text) = Trim(reg("mnu_menu")) Then
                    Lst_menu.Selected(i) = True
                    Exit For
                End If
            Next i
            reg.MoveNext
        Loop
    End If
End Sub
