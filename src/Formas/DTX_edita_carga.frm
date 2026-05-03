VERSION 5.00
Begin VB.Form DTX_edita_carga 
   Caption         =   "Edita Carga DTX"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   3375
      Begin VB.OptionButton Opt_ninguno 
         Caption         =   "Ningun Pto"
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Opt_trasb 
         Caption         =   "Solo Trasbordo"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Opt_todos 
         Caption         =   "Todos Ptos"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Opt_gye 
         Caption         =   "Solo GYE"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Cmd_canc 
      Height          =   510
      Left            =   2985
      Picture         =   "DTX_edita_carga.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancelar"
      Top             =   2280
      Width           =   510
   End
   Begin VB.CommandButton Cmd_acep 
      Cancel          =   -1  'True
      Height          =   510
      Left            =   2400
      Picture         =   "DTX_edita_carga.frx":00EA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Aceptar"
      Top             =   2280
      Width           =   510
   End
   Begin VB.ComboBox Cmb_Columna 
      Height          =   315
      ItemData        =   "DTX_edita_carga.frx":036C
      Left            =   1440
      List            =   "DTX_edita_carga.frx":037C
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txt_reemplazo 
      Height          =   285
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txt_valor 
      Height          =   285
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Columna :"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Reemplazar con :"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Valor :"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   450
   End
End
Attribute VB_Name = "DTX_edita_carga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_acep_Click()
Dim pto_descarga As String
Dim pto_dest_final As String
Select Case Me.Cmb_Columna.ListIndex
    Case 0: Col = 6
    Case 1: Col = 7
    Case 2: Col = 8 'pto descarga
    Case 3: Col = 9
End Select
For i = 1 To DTX_PROCESO.Spr_dtx.MaxRows
        DTX_PROCESO.Spr_dtx.Row = i
        DTX_PROCESO.Spr_dtx.Col = Col
        If Trim(DTX_PROCESO.Spr_dtx.Text) = Trim(txt_valor.Text) Then
           DTX_PROCESO.Spr_dtx.Text = Trim(txt_reemplazo.Text)
        End If
Next i
'se desmarcan todos los check
For i = 1 To DTX_PROCESO.Spr_dtx.MaxRows
    DTX_PROCESO.Spr_dtx.Row = i
    DTX_PROCESO.Spr_dtx.Col = 1
    DTX_PROCESO.Spr_dtx.Value = 0
Next i
'solo ECGYE
If Me.Opt_gye.Value Then
    For i = 1 To DTX_PROCESO.Spr_dtx.MaxRows
            DTX_PROCESO.Spr_dtx.Row = i
            DTX_PROCESO.Spr_dtx.Col = 8: pto_descarga = Trim(DTX_PROCESO.Spr_dtx.Text)
            DTX_PROCESO.Spr_dtx.Col = 9: pto_dest_final = Trim(DTX_PROCESO.Spr_dtx.Text)
            If Trim(pto_descarga) = "ECGYE" And Trim(pto_dest_final) = "ECGYE" And Trim(pto_descarga) <> "" And Trim(pto_dest_final) <> "" Then
               DTX_PROCESO.Spr_dtx.Col = 1
               DTX_PROCESO.Spr_dtx.Value = 1
               st_detOK = True
            ElseIf Trim(DTX_PROCESO.Spr_dtx.Text) <> "ECGYE" And Trim(DTX_PROCESO.Spr_dtx.Text) <> "" Then
               st_detOK = False
            Else
                If st_detOK Then
                    DTX_PROCESO.Spr_dtx.Col = 1
                    DTX_PROCESO.Spr_dtx.Value = 1
                End If
            End If
    Next i
End If
'SOLO TRASBORDO
If Me.Opt_trasb.Value Then
If DTX_PROCESO.Opt_tipo_impo Then
    For i = 1 To DTX_PROCESO.Spr_dtx.MaxRows
            DTX_PROCESO.Spr_dtx.Row = i
            DTX_PROCESO.Spr_dtx.Col = 8: pto_descarga = Trim(DTX_PROCESO.Spr_dtx.Text)
            DTX_PROCESO.Spr_dtx.Col = 9: pto_dest_final = Trim(DTX_PROCESO.Spr_dtx.Text)
            If Trim(pto_descarga) = "ECGYE" And Trim(pto_dest_final) <> "ECGYE" And Trim(pto_descarga) <> "" And Trim(pto_dest_final) <> "" Then
               DTX_PROCESO.Spr_dtx.Col = 1
               DTX_PROCESO.Spr_dtx.Value = 1
               st_detOK = True
            ElseIf (Trim(pto_descarga) <> "ECGYE" And Trim(pto_descarga) <> "") Or (Trim(pto_descarga) = "ECGYE" And Trim(pto_dest_final) = "ECGYE") Then
               st_detOK = False
            Else
                If st_detOK Then
                    DTX_PROCESO.Spr_dtx.Col = 1
                    DTX_PROCESO.Spr_dtx.Value = 1
                End If
            End If
    Next i
Else 'EXPORTACION
    For i = 1 To DTX_PROCESO.Spr_dtx.MaxRows
            DTX_PROCESO.Spr_dtx.Row = i
            DTX_PROCESO.Spr_dtx.Col = 8: pto_descarga = Trim(DTX_PROCESO.Spr_dtx.Text)
            DTX_PROCESO.Spr_dtx.Col = 9: pto_dest_final = Trim(DTX_PROCESO.Spr_dtx.Text)
            If Trim(pto_descarga) <> Trim(pto_dest_final) And Trim(pto_descarga) <> "" And Trim(pto_dest_final) <> "" Then
               DTX_PROCESO.Spr_dtx.Col = 1
               DTX_PROCESO.Spr_dtx.Value = 1
               st_detOK = True
            ElseIf Trim(pto_descarga) = Trim(pto_dest_final) And Trim(pto_descarga) <> "" And Trim(pto_dest_final) <> "" Then
               st_detOK = False
            Else
                If st_detOK Then
                    DTX_PROCESO.Spr_dtx.Col = 1
                    DTX_PROCESO.Spr_dtx.Value = 1
                End If
            End If
    Next i

End If
End If






If Me.Opt_todos.Value Then
    For i = 1 To DTX_PROCESO.Spr_dtx.MaxRows
        DTX_PROCESO.Spr_dtx.Row = i
        DTX_PROCESO.Spr_dtx.Col = 1
        DTX_PROCESO.Spr_dtx.Value = 1
    Next i
End If


End Sub

Private Sub Cmd_canc_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Opt_gye.Value = True
End Sub
