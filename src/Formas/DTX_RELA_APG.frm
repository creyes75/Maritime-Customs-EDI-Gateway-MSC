VERSION 5.00
Begin VB.Form DTX_RELA_APG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relación de Códigos CAE - Linea"
   ClientHeight    =   2295
   ClientLeft      =   3285
   ClientTop       =   2655
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4410
   Begin VB.OptionButton Opt_nuevo 
      Caption         =   "nuevo registro (S/N)"
      Height          =   240
      Left            =   120
      TabIndex        =   20
      Top             =   1995
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.CommandButton Cmd_cancel 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   3300
      TabIndex        =   9
      Top             =   1905
      Width           =   975
   End
   Begin VB.CommandButton Cmd_grabar 
      Caption         =   "&Grabar"
      Height          =   330
      Left            =   2295
      TabIndex        =   8
      Top             =   1905
      Width           =   975
   End
   Begin VB.Frame fra_tipo_contein 
      Caption         =   " Datos de la Tipo de Contenedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1770
      Left            =   90
      TabIndex        =   48
      Top             =   45
      Visible         =   0   'False
      Width           =   4200
      Begin VB.CommandButton CMD_TIPO_CONT_HELP 
         Caption         =   "..."
         Height          =   240
         Left            =   3345
         TabIndex        =   59
         Top             =   1500
         Width           =   270
      End
      Begin VB.TextBox txt_tipo_cont_desc 
         Height          =   285
         Left            =   1815
         TabIndex        =   52
         Top             =   1140
         Width           =   1620
      End
      Begin VB.TextBox txt_tipo_cont_size 
         Height          =   285
         Left            =   1815
         TabIndex        =   50
         Top             =   525
         Width           =   1050
      End
      Begin VB.TextBox txt_tipo_cont_apg 
         Height          =   285
         Left            =   1830
         TabIndex        =   53
         Top             =   1440
         Width           =   1425
      End
      Begin VB.TextBox txt_tipo_cont_type 
         Height          =   285
         Left            =   1815
         TabIndex        =   51
         Top             =   825
         Width           =   1065
      End
      Begin VB.TextBox txt_tipo_cont_linea 
         Height          =   285
         Left            =   1815
         TabIndex        =   49
         Top             =   225
         Width           =   1050
      End
      Begin VB.Label Label25 
         Caption         =   "Descripción :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   315
         TabIndex        =   58
         Top             =   1170
         Width           =   1275
      End
      Begin VB.Label Label24 
         Caption         =   "Size :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   315
         TabIndex        =   57
         Top             =   555
         Width           =   1395
      End
      Begin VB.Label Label23 
         Caption         =   "CAE :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   315
         TabIndex        =   56
         Top             =   1455
         Width           =   1155
      End
      Begin VB.Label Label22 
         Caption         =   "Type :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   315
         TabIndex        =   55
         Top             =   855
         Width           =   1275
      End
      Begin VB.Label Label21 
         Caption         =   "Línea Naviera :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   315
         TabIndex        =   54
         Top             =   255
         Width           =   1395
      End
   End
   Begin VB.Frame fra_cond_contein 
      Caption         =   " Datos de la Condición del Contenedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1770
      Left            =   90
      TabIndex        =   39
      Top             =   45
      Visible         =   0   'False
      Width           =   4200
      Begin VB.CommandButton cmd_cond_cont_apg 
         Caption         =   "..."
         Height          =   240
         Left            =   3285
         TabIndex        =   60
         Top             =   1335
         Width           =   270
      End
      Begin VB.TextBox txt_cond_cont_linea 
         Height          =   285
         Left            =   1815
         TabIndex        =   40
         Top             =   300
         Width           =   1050
      End
      Begin VB.TextBox txt_cond_cont_dest 
         Height          =   285
         Left            =   1815
         TabIndex        =   44
         Top             =   960
         Width           =   1620
      End
      Begin VB.TextBox txt_cond_cont_APG 
         Height          =   285
         Left            =   1815
         TabIndex        =   46
         Top             =   1290
         Width           =   1425
      End
      Begin VB.TextBox txt_cond_cont_orig 
         Height          =   285
         Left            =   1815
         TabIndex        =   42
         Top             =   630
         Width           =   1050
      End
      Begin VB.Label Label20 
         Caption         =   "Línea Naviera :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   315
         TabIndex        =   47
         Top             =   330
         Width           =   1395
      End
      Begin VB.Label Label19 
         Caption         =   "Destino :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   315
         TabIndex        =   45
         Top             =   990
         Width           =   1275
      End
      Begin VB.Label Label18 
         Caption         =   "CAE :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   300
         TabIndex        =   43
         Top             =   1290
         Width           =   1155
      End
      Begin VB.Label Label17 
         Caption         =   "Origen :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   315
         TabIndex        =   41
         Top             =   660
         Width           =   1395
      End
   End
   Begin VB.Frame fra_embalaje 
      Caption         =   " Datos del Embalaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1770
      Left            =   90
      TabIndex        =   30
      Top             =   45
      Visible         =   0   'False
      Width           =   4200
      Begin VB.CommandButton cmd_emba_apg 
         Caption         =   "..."
         Height          =   240
         Left            =   3285
         TabIndex        =   61
         Top             =   1335
         Width           =   270
      End
      Begin VB.TextBox txt_emba_codi 
         Height          =   285
         Left            =   1815
         TabIndex        =   32
         Top             =   630
         Width           =   1050
      End
      Begin VB.TextBox txt_emba_apg 
         Height          =   285
         Left            =   1815
         TabIndex        =   34
         Top             =   1290
         Width           =   1425
      End
      Begin VB.TextBox txt_emba_desc 
         Height          =   285
         Left            =   1815
         TabIndex        =   33
         Top             =   960
         Width           =   1620
      End
      Begin VB.TextBox txt_emba_linea 
         Height          =   285
         Left            =   1815
         TabIndex        =   31
         Top             =   300
         Width           =   1050
      End
      Begin VB.Label Label16 
         Caption         =   "Código :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   315
         TabIndex        =   38
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label Label15 
         Caption         =   "CAE :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   315
         TabIndex        =   37
         Top             =   1305
         Width           =   1155
      End
      Begin VB.Label Label14 
         Caption         =   "Descripción :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   315
         TabIndex        =   36
         Top             =   990
         Width           =   1275
      End
      Begin VB.Label Label13 
         Caption         =   "Línea Naviera :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   315
         TabIndex        =   35
         Top             =   330
         Width           =   1395
      End
   End
   Begin VB.Frame fra_moneda 
      Caption         =   " Datos de la Moneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1770
      Left            =   90
      TabIndex        =   21
      Top             =   45
      Visible         =   0   'False
      Width           =   4200
      Begin VB.CommandButton cmd_mone_apg 
         Caption         =   "..."
         Height          =   240
         Left            =   3300
         TabIndex        =   62
         Top             =   1335
         Width           =   270
      End
      Begin VB.TextBox Txt_mone_linea 
         Height          =   285
         Left            =   1815
         TabIndex        =   22
         Top             =   300
         Width           =   1050
      End
      Begin VB.TextBox txt_mone_desc 
         Height          =   285
         Left            =   1815
         TabIndex        =   26
         Top             =   960
         Width           =   1620
      End
      Begin VB.TextBox txt_mone_apg 
         Height          =   285
         Left            =   1815
         TabIndex        =   28
         Top             =   1290
         Width           =   1425
      End
      Begin VB.TextBox txt_mone_codi 
         Height          =   285
         Left            =   1815
         TabIndex        =   24
         Top             =   630
         Width           =   1050
      End
      Begin VB.Label Label12 
         Caption         =   "Línea Naviera :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   315
         TabIndex        =   29
         Top             =   330
         Width           =   1395
      End
      Begin VB.Label Label11 
         Caption         =   "Descripción :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   315
         TabIndex        =   27
         Top             =   990
         Width           =   1275
      End
      Begin VB.Label Label10 
         Caption         =   "CAE :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   315
         TabIndex        =   25
         Top             =   1305
         Width           =   1155
      End
      Begin VB.Label Label9 
         Caption         =   "Código :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   315
         TabIndex        =   23
         Top             =   660
         Width           =   1395
      End
   End
   Begin VB.Frame fra_puerto 
      Caption         =   " Datos del Puerto "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1770
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   4200
      Begin VB.CommandButton cmd_puer_apg 
         Caption         =   "..."
         Height          =   240
         Left            =   3315
         TabIndex        =   63
         Top             =   1320
         Width           =   270
      End
      Begin VB.TextBox Txt_puer_codi 
         Height          =   285
         Left            =   1815
         TabIndex        =   3
         Top             =   630
         Width           =   1050
      End
      Begin VB.TextBox Txt_puer_apg 
         Height          =   285
         Left            =   1815
         TabIndex        =   5
         Top             =   1290
         Width           =   1425
      End
      Begin VB.TextBox Txt_puer_desc 
         Height          =   285
         Left            =   1815
         TabIndex        =   4
         Top             =   960
         Width           =   1620
      End
      Begin VB.TextBox Txt_puer_linea 
         Height          =   285
         Left            =   1815
         TabIndex        =   2
         Top             =   300
         Width           =   1050
      End
      Begin VB.Label Label7 
         Caption         =   "Código :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   315
         TabIndex        =   10
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "CAE :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   315
         TabIndex        =   7
         Top             =   1305
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   315
         TabIndex        =   6
         Top             =   990
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Línea Naviera :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   315
         TabIndex        =   1
         Top             =   330
         Width           =   1395
      End
   End
   Begin VB.Frame fra_imo 
      Caption         =   " Datos de la Peligrosidad "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1770
      Left            =   90
      TabIndex        =   11
      Top             =   45
      Visible         =   0   'False
      Width           =   4200
      Begin VB.CommandButton cmd_imo_apg 
         Caption         =   "..."
         Height          =   240
         Left            =   3285
         TabIndex        =   64
         Top             =   1320
         Width           =   270
      End
      Begin VB.TextBox Txt_imo_linea 
         Height          =   285
         Left            =   1815
         TabIndex        =   12
         Top             =   300
         Width           =   1050
      End
      Begin VB.TextBox txt_imo_desc 
         Height          =   285
         Left            =   1815
         TabIndex        =   16
         Top             =   960
         Width           =   1620
      End
      Begin VB.TextBox txt_imo_apg 
         Height          =   285
         Left            =   1815
         TabIndex        =   18
         Top             =   1290
         Width           =   1425
      End
      Begin VB.TextBox Txt_imo_codi 
         Height          =   285
         Left            =   1815
         TabIndex        =   14
         Top             =   630
         Width           =   1050
      End
      Begin VB.Label Label8 
         Caption         =   "Línea Naviera :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   315
         TabIndex        =   19
         Top             =   330
         Width           =   1395
      End
      Begin VB.Label Label6 
         Caption         =   "Descripción :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   315
         TabIndex        =   17
         Top             =   990
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "CAE :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   315
         TabIndex        =   15
         Top             =   1305
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Código :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   315
         TabIndex        =   13
         Top             =   660
         Width           =   1395
      End
   End
End
Attribute VB_Name = "DTX_RELA_APG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_cancel_Click()
   Unload Me
End Sub

Private Sub cmd_cond_cont_apg_Click()
txt_cond_cont_APG.Text = Ayuda("docu_cond_contenedor", "cond_cont_codigo", "cond_cont_descripcion", "cond_cont_linea='CAE' and cond_cont_status ='A'", "Condición de Contenedor CAE")
End Sub

Private Sub cmd_emba_apg_Click()
txt_emba_apg.Text = Ayuda("docu_embalaje", "emba_codigo", "emba_descripcion", "emba_linea ='CAE' and emba_status ='A'", "Embalajes de Carga CAE")
End Sub

Private Sub Cmd_grabar_Click()
On Error GoTo Errores

If Opt_nuevo.Value Then
   Select Case Tipo_rela_APG
      Case "DTX_PUERTO":       sql = "INSERT INTO DTX_PUERTO       VALUES ('" & Gstr_Linea & "','" & Txt_puer_codi.Text & "', '" & Txt_puer_desc.Text & "', '" & Txt_puer_apg.Text & "')"
      Case "DTX_IMO":          sql = "INSERT INTO DTX_IMO          VALUES ('" & Gstr_Linea & "','" & Txt_imo_codi.Text & "', '" & txt_imo_desc.Text & "', '" & txt_imo_apg.Text & "')"
      Case "DTX_COND_CONTEIN": sql = "INSERT INTO DTX_COND_CONTEIN VALUES ('" & Gstr_Linea & "','" & txt_cond_cont_orig.Text & "','" & txt_cond_cont_dest.Text & "','" & txt_cond_cont_APG.Text & "')"
      Case "DTX_EMBALAJE":     sql = "INSERT INTO DTX_EMBALAJE     VALUES ('" & Gstr_Linea & "','" & Mid$(txt_emba_codi.Text, 1, 3) & "', '" & txt_emba_desc.Text & "', '" & txt_emba_apg.Text & "')"
      Case "DTX_MONEDA":       sql = "INSERT INTO DTX_MONEDA       VALUES ('" & Gstr_Linea & "','" & txt_mone_codi.Text & "', '" & txt_mone_desc.Text & "', '" & txt_mone_apg.Text & "')"
      Case "DTX_TIPO_CONTEIN": sql = "INSERT INTO DTX_TIPO_CONTEIN VALUES ('" & Gstr_Linea & "','" & txt_tipo_cont_size.Text & "','" & txt_tipo_cont_type.Text & "','" & txt_tipo_cont_desc.Text & "','" & txt_tipo_cont_apg.Text & "')"
   End Select
Else
   Select Case Tipo_rela_APG
      Case "DTX_PUERTO":       sql = "UPDATE DTX_PUERTO       SET descripcion = '" & Txt_puer_desc.Text & "',      apg = '" & Txt_puer_apg.Text & "'      WHERE linea = '" & Gstr_Linea & "' and codigo = '" & Txt_puer_codi.Text & "'"
      Case "DTX_IMO":          sql = "UPDATE DTX_IMO          SET descripcion = '" & txt_imo_desc.Text & "' ,      apg = '" & txt_imo_apg.Text & "'       WHERE linea = '" & Gstr_Linea & "' and codigo = '" & Txt_imo_codi.Text & "'"
      Case "DTX_COND_CONTEIN": sql = "UPDATE DTX_COND_CONTEIN SET                                                  apg = '" & txt_cond_cont_APG.Text & "' WHERE linea = '" & Gstr_Linea & "' and origen = '" & txt_cond_cont_orig.Text & "' and destino = '" & txt_cond_cont_dest.Text & "'"
      Case "DTX_EMBALAJE":     sql = "UPDATE DTX_EMBALAJE     SET descripcion = '" & txt_emba_desc.Text & "',      apg = '" & txt_emba_apg.Text & "'      WHERE linea = '" & Gstr_Linea & "' and codigo = '" & txt_emba_codi.Text & "'"
      Case "DTX_MONEDA":       sql = "UPDATE DTX_MONEDA       SET descripcion = '" & txt_mone_desc.Text & "',      apg = '" & txt_mone_apg.Text & "'      WHERE linea = '" & Gstr_Linea & "' and codigo = '" & txt_mone_codi.Text & "'"
      Case "DTX_TIPO_CONTEIN": sql = "UPDATE DTX_TIPO_CONTEIN SET descripcion = '" & txt_tipo_cont_desc.Text & "', apg = '" & txt_tipo_cont_apg.Text & "' WHERE linea = '" & Gstr_Linea & "' and size   = '" & txt_tipo_cont_size.Text & "' and type = '" & txt_tipo_cont_type.Text & "'"
   End Select
End If

Gcon_docu.Execute sql, dbFailOnError
Unload Me
Exit Sub

Errores:
MsgBox "Ocurrio el Error " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0
End Sub

Private Sub cmd_imo_apg_Click()
txt_imo_apg.Text = Ayuda("docu_imo", "imo_codigo", "imo_descripcion", "imo_linea ='CAE' and imo_status ='A'", "Peligrosidad CAE")
End Sub

Private Sub cmd_mone_apg_Click()
txt_mone_apg.Text = Ayuda("docu_moneda", "mone_codigo", "mone_descripcion", "mone_linea ='CAE' and mone_status ='A'", "Monedas CAE")
End Sub

Private Sub cmd_puer_apg_Click()
Txt_puer_apg.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", "puer_linea ='CAE' and puer_status ='A'", "Puertos CAE")
End Sub

Private Sub CMD_TIPO_CONT_HELP_Click()
txt_tipo_cont_apg.Text = Ayuda("DOCU_TAMANO_EQUIP", "tama_equip_codigo", "tama_equip_descripcion", "tama_equip_linea ='CAE'", "Tipos de Contenedor CAE")
End Sub

Private Sub Form_Load()
Dim reg As New ADODB.Recordset
Dim sql As String

Select Case UCase(Tipo_rela_APG)
    Case "DTX_PUERTO":       fra_puerto.Visible = True
    Case "DTX_IMO":          Fra_IMO.Visible = True
    Case "DTX_COND_CONTEIN": fra_cond_contein.Visible = True
    Case "DTX_EMBALAJE":     fra_embalaje.Visible = True
    Case "DTX_MONEDA":       Fra_moneda.Visible = True
    Case "DTX_TIPO_CONTEIN": fra_tipo_contein.Visible = True
End Select

If Not Opt_nuevo.Value Then
   Select Case UCase(Tipo_rela_APG)
      Case "DTX_PUERTO":       sql = "Select * from DTX_PUERTO       where Linea = '" & Gstr_Linea & "' and  codigo = '" & Txt_puer_codi.Text & "'"
      Case "DTX_IMO":          sql = "Select * from DTX_IMO          Where linea = '" & Gstr_Linea & "' and  codigo = '" & Txt_imo_codi.Text & "'"
      Case "DTX_COND_CONTEIN": sql = "Select * from DTX_COND_CONTEIN Where linea = '" & Gstr_Linea & "' and  Origen = '" & txt_cond_cont_orig.Text & "' and Destino = '" & txt_cond_cont_dest.Text & "'"
      Case "DTX_EMBALAJE":     sql = "Select * from DTX_EMBALAJE     Where linea = '" & Gstr_Linea & "' and  codigo = '" & txt_emba_codi.Text & "'"
      Case "DTX_MONEDA":       sql = "Select * from DTX_MONEDA       Where linea = '" & Gstr_Linea & "' and  codigo = '" & txt_mone_codi.Text & "'"
      Case "DTX_TIPO_CONTEIN": sql = "Select * from DTX_TIPO_CONTEIN Where linea = '" & Gstr_Linea & "' and  size   = '" & txt_tipo_cont_size.Text & "' and type    = '" & txt_tipo_cont_type.Text & "'"
   End Select
      
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   
   If Not reg.EOF Then
      Select Case UCase(Tipo_rela_APG)
         Case "DTX_PUERTO":       Txt_puer_desc.Text = reg("descripcion")
         Case "DTX_IMO":          txt_imo_desc.Text = reg("descripcion")
         Case "DTX_COND_CONTEIN": ' no existe descripcion
         Case "DTX_EMBALAJE":     txt_emba_desc.Text = reg("descripcion")
         Case "DTX_MONEDA":       txt_mone_desc.Text = reg("descripcion")
         Case "DTX_TIPO_CONTEIN": txt_tipo_cont_desc.Text = reg("descripcion")
      End Select
   End If
End If

End Sub
