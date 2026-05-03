VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form DTX_PROCESO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proceso de Carga de Archivos DTX."
   ClientHeight    =   4470
   ClientLeft      =   4995
   ClientTop       =   2280
   ClientWidth     =   8295
   Icon            =   "DTX_PROCESO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   8295
   Begin VB.CommandButton Cmd_edita_carga 
      Caption         =   "Edicion"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Cmd_Siguiente 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Anterior 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   4545
      TabIndex        =   11
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_procesar 
      Caption         =   "&Procesar"
      Height          =   375
      Left            =   5775
      TabIndex        =   13
      ToolTipText     =   "Procesar archivo DTX"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_salir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6990
      TabIndex        =   12
      ToolTipText     =   "Salir del Proceso"
      Top             =   3960
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog Dlg_DTX 
      Left            =   4380
      Top             =   2145
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Abrir Archivo DTX"
   End
   Begin VB.Frame Fra_confirma 
      Caption         =   " Confirmación de Datos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   8055
      Begin FPSpread.vaSpread Spr_dtx 
         Height          =   3330
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   7860
         _Version        =   131077
         _ExtentX        =   13864
         _ExtentY        =   5874
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "DTX_PROCESO.frx":030A
      End
   End
   Begin VB.Frame Fra_proceso 
      Caption         =   "  Procesar  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   8055
      Begin VB.Frame Frame4 
         Height          =   435
         Left            =   2520
         TabIndex        =   28
         Top             =   960
         Width           =   2700
         Begin VB.Label Lbl_avan 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   75
            TabIndex        =   30
            Top             =   180
            Width           =   2595
         End
         Begin VB.Label Lbl_prog 
            BackColor       =   &H00800000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   30
            TabIndex        =   29
            Top             =   120
            Visible         =   0   'False
            Width           =   2625
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Por favor espere"
         Height          =   375
         Left            =   480
         TabIndex        =   32
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Lbl_proceso_BL 
         Height          =   375
         Left            =   480
         TabIndex        =   31
         Top             =   1680
         Width           =   4095
      End
   End
   Begin VB.Frame Fra_DTX 
      Caption         =   " Archivo DTX "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   8055
      Begin VB.Frame Frame3 
         Height          =   525
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   5055
         Begin VB.OptionButton Opt_tipo_expo 
            Caption         =   "Arch. Exportación"
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
            Height          =   225
            Left            =   2565
            TabIndex        =   1
            Top             =   240
            Width           =   2205
         End
         Begin VB.OptionButton Opt_tipo_impo 
            Caption         =   "Arch. Importación"
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
            Height          =   195
            Left            =   420
            TabIndex        =   0
            Top             =   240
            Width           =   2115
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2625
         Left            =   120
         TabIndex        =   15
         Top             =   750
         Width           =   5055
         Begin VB.CheckBox Chk_rela_fija 
            Caption         =   "Relación de códigos Fija"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   2280
            Width           =   2055
         End
         Begin VB.TextBox Txt_viaje 
            Height          =   285
            Left            =   1140
            MaxLength       =   5
            TabIndex        =   4
            Top             =   1125
            Width           =   1515
         End
         Begin VB.OptionButton Opt_khs 
            Caption         =   "Formato KHS"
            Height          =   195
            Left            =   2415
            TabIndex        =   21
            Top             =   2295
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.OptionButton Opt_msc 
            Caption         =   "Formato MSC"
            Height          =   195
            Left            =   420
            TabIndex        =   20
            Top             =   2310
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.Frame Fra_linea 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   360
            Left            =   255
            TabIndex        =   18
            Top             =   225
            Width           =   3675
            Begin MSDataListLib.DataCombo Cmb_linea 
               Height          =   315
               Left            =   885
               TabIndex        =   2
               Top             =   15
               Width           =   2430
               _ExtentX        =   4286
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin MSAdodcLib.Adodc data_linea 
               Height          =   330
               Left            =   2445
               Top             =   45
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
            Begin VB.Label Label5 
               Caption         =   "Linea"
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
               Left            =   120
               TabIndex        =   19
               Top             =   0
               Width           =   825
            End
         End
         Begin VB.Frame Fra_buque 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   345
            Left            =   255
            TabIndex        =   16
            Top             =   690
            Width           =   3975
            Begin MSDataListLib.DataCombo Cmb_buque 
               Height          =   315
               Left            =   885
               TabIndex        =   3
               Top             =   0
               Width           =   2430
               _ExtentX        =   4286
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin MSAdodcLib.Adodc data_buque 
               Height          =   330
               Left            =   2460
               Top             =   -30
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
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Buque"
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
               Height          =   195
               Left            =   120
               TabIndex        =   17
               Top             =   -15
               Width           =   555
            End
         End
         Begin VB.TextBox Txt_DTX 
            Height          =   285
            Left            =   1830
            TabIndex        =   6
            Top             =   1935
            Width           =   2715
         End
         Begin VB.CommandButton Cmd_DTX 
            Caption         =   "..."
            Height          =   285
            Left            =   4560
            TabIndex        =   7
            ToolTipText     =   "Seleccionar Archivo a Procesar"
            Top             =   1950
            Width           =   285
         End
         Begin VB.PictureBox RPT 
            Height          =   480
            Left            =   3840
            ScaleHeight     =   420
            ScaleWidth      =   1140
            TabIndex        =   34
            Top             =   1245
            Width           =   1200
         End
         Begin MSMask.MaskEdBox Msk_atraque 
            Height          =   285
            Left            =   2175
            TabIndex        =   5
            Top             =   1530
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Viaje :"
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
            Height          =   255
            Left            =   375
            TabIndex        =   24
            Top             =   1170
            Width           =   1740
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha de Atraque :"
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
            Height          =   255
            Left            =   375
            TabIndex        =   23
            Top             =   1530
            Width           =   1740
         End
         Begin VB.Label Label6 
            Caption         =   "Archivo DTX :"
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
            Height          =   180
            Left            =   375
            TabIndex        =   22
            Top             =   1950
            Width           =   1365
         End
      End
   End
   Begin VB.Menu MNU_DTX 
      Caption         =   "DTX"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnu_TIPO 
         Caption         =   "DTX &Tipo"
      End
      Begin VB.Menu mnu_detalle 
         Caption         =   "DTX &Detalle"
      End
      Begin VB.Menu mnu_linea 
         Caption         =   "&Línea Naviera"
      End
      Begin VB.Menu mnu_param 
         Caption         =   "&Parámetro"
      End
      Begin VB.Menu mnu_form_deta 
         Caption         =   "&Formato Detalle"
      End
      Begin VB.Menu mnu_migra 
         Caption         =   "migra puertos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_raya 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_mant_puer 
         Caption         =   "Mant. Puertos"
      End
      Begin VB.Menu mnu_mant_emba 
         Caption         =   "Mant. Embalaje"
      End
      Begin VB.Menu mnu_mant_tipo_cont 
         Caption         =   "Mant. Tipo Contein"
      End
      Begin VB.Menu mnu_mant_cond_cont 
         Caption         =   "Mant. Cond Contein"
      End
      Begin VB.Menu Mnu_mant_IMO 
         Caption         =   "Mant. IMO"
      End
      Begin VB.Menu mnu_raya2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_repo_puerto 
         Caption         =   "Reporte Puerto"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_repo_emba 
         Caption         =   "Reporte Embalaje"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_repo_tipo_cont 
         Caption         =   "Reporte Tipo Contein"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_repo_cond_cont 
         Caption         =   "Reporte Cond Contein"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_repo_imo 
         Caption         =   "Reporte IMO"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_repo_estruc 
         Caption         =   "Reporte de Estructuras"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_raya3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_acerca 
         Caption         =   "Acerca de ..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_salir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "DTX_PROCESO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num_tab As Integer
Private Sub Chk_rela_fija_Click()
   Gbol_relacion_fija = CBool(Chk_rela_fija.Value)
End Sub

Private Sub Cmb_buque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmb_linea_Change()
Dim reg As New ADODB.Recordset
Dim sql As String

If Cmb_linea.MatchedWithList Then

    carga_datacombo data_buque, Cmb_buque, "Select * from docu_buque Where buqu_linea = '" & Cmb_linea.BoundText & "' and buqu_status ='A'", "buqu_codigo", "buqu_descripcion"
    
    'Gstr_Pto_desc = "ECGYE"   ' este seteo debe estar parametrizado ' comentario 22/sept 2002
    If Cmb_linea.BoundText = "CAE" Then
        Opt_khs.Value = True: Opt_khs.Enabled = True
        Opt_msc.Value = True: Opt_msc.Enabled = True
    Else
        Gstr_Linea_Nav_APG = Cmb_linea.BoundText
        If Cmb_linea.BoundText = "MSC" Then
            Opt_msc.Enabled = True: Opt_msc.Value = True
            Opt_khs.Enabled = False: Opt_khs.Value = False
        Else
            Opt_khs.Enabled = True: Opt_khs.Value = True
            Opt_msc.Enabled = False: Opt_msc.Value = False
        End If
    End If
        
End If
End Sub

Private Sub Cmd_Anterior_Click()
Select Case num_tab
    Case 2: num_tab = 1
    Case 3: num_tab = 2
End Select
Call Carga_frame

End Sub

Private Sub Cmd_DTX_Click()
   Dlg_DTX.Filter = "Todos los Archivos (*.*)|*.*|Archivos DTX (*.dtx)|*.dtx"
   Dlg_DTX.ShowOpen
   Txt_DTX.Text = Dlg_DTX.FileName
End Sub

Private Sub Cmd_edita_carga_Click()
    DTX_edita_carga.Show 1
End Sub

Private Sub Cmd_procesar_Click()
Dim reg As New ADODB.Recordset


'Call procesa_archivo_DTX_XML(Txt_DTX)

GLinea_act_Grid = 0

If Cmb_linea.BoundText = "CAE" Then
    Select Case Gstr_Linea
        Case "MSC"
           If verifica_version_DTX(Txt_DTX) = "55" Then  ' se verifica que la version sea la correcta
                Call procesar_arch_v55(Txt_DTX)                   'proceso de carga normal
                If MsgBox("Desea procesar en este momento datos de contenedores vacíos?", vbQuestion + vbYesNo) = vbNo Then GoTo fin
                GLinea_act_Grid = 0
                Call procesar_arch_MSC_cont_vacios_V55(Txt_DTX)  'proceso de carga para contened. vacios
           Else
                Gstr_Version_MSC = "54"
                Call procesar_arch(Txt_DTX)                  'proceso de carga normal
                If MsgBox("Desea procesar en este momento datos de contenedores vacíos?", vbQuestion + vbYesNo) = vbNo Then GoTo fin
                GLinea_act_Grid = 0
                Call procesar_arch_MSC_cont_vacios(Txt_DTX)  'proceso de carga para contened. vacios
           End If
        Case "KHS"
           Call procesar_arch_KHS(Txt_DTX)
    End Select
ElseIf Cmb_linea.BoundText = "MSC" Then
    If verifica_version_DTX(Txt_DTX) = "55" Then ' se verifica que la version sea la correcta
                Call procesar_arch_v55(Txt_DTX)                   'proceso de carga normal
                If MsgBox("Desea procesar en este momento datos de contenedores vacíos?", vbQuestion + vbYesNo) = vbNo Then GoTo fin
                Call procesar_arch_MSC_cont_vacios_V55(Txt_DTX)  'proceso de carga para contened. vacios
           Else
                Call procesar_arch(Txt_DTX)                  'proceso de carga normal
                If MsgBox("Desea procesar en este momento datos de contenedores vacíos?", vbQuestion + vbYesNo) = vbNo Then GoTo fin
                Call procesar_arch_MSC_cont_vacios(Txt_DTX)  'proceso de carga para contened. vacios
           End If
ElseIf Cmb_linea.BoundText = "KHS" Then
    MsgBox "OPCION NO PERMITIDA", vbInformation
    'Call procesar_arch_KHS(Txt_DTX)
End If

fin:

    'TL Procedimiento que alerta los casos en que los contenedores estan en el detalle del Bl y no en la tabla de contenedores
    Call valida_consist_cont(Cmb_linea.BoundText, Cmb_buque.BoundText, Txt_viaje.Text)
    'TL Procedimiento que alerta los casos en que las cabeceras de bls suben sin detalles
    Call valida_consist_cabecera(Cmb_linea.BoundText, Cmb_buque.BoundText, Txt_viaje.Text)
    
    MsgBox "Fin de Proceso de Carga de Archivos DTX!", vbInformation, App.Title
    Lbl_proceso_BL = "Proceso Finalizado"
    DTX_PROCESO.Lbl_avan.ForeColor = &H404040
    DTX_PROCESO.Lbl_avan.Caption = ""
    DTX_PROCESO.Lbl_prog.Width = 0
    DTX_PROCESO.Lbl_prog.Visible = False
End Sub
Private Function valida_Datos() As Boolean
    valida_Datos = False
    If Not IsDate(Msk_atraque.Text) Then
       MsgBox "La fecha de Atraque no existe o no tiene el Formato esperado.", vbInformation, App.Title
       Exit Function
    End If
    
    If Trim$(Txt_DTX.Text) = "" Then
       MsgBox "Ingrese la Ruta del Archivo a Procesar.", vbInformation, App.Title
       Exit Function
    End If
    
    If Dir(Txt_DTX, vbArchive) = "" Then
       MsgBox "La Ruta del Archivo a Procesar, No es una RUTA Válida.", vbInformation, App.Title
       Exit Function
    End If
    valida_Datos = True
End Function
Private Function valida_Datos2() As Boolean
On Error GoTo ERROR_VALIDA
Dim reg As New ADODB.Recordset
    valida_Datos2 = False
    For i = 1 To Spr_dtx.MaxRows
        Spr_dtx.Row = i
        Spr_dtx.Col = 1
        If CBool(Spr_dtx.Value) Then
            Spr_dtx.Col = 2
            If Spr_dtx.Text <> "" Then
                Spr_dtx.Col = 6 '***** pto de embarque
                If Spr_dtx.Text = "" Then
                    MsgBox "Debe ingresar un valor valido para el puerto de Embarque, en la linea " & i, vbInformation
                    Exit Function
                Else
                    reg.Open "select * from docu_puerto where puer_codigo = '" & Spr_dtx.Text & "'", Gcon_docu
                    If reg.EOF Then
                        MsgBox "Debe ingresar un codigo valido para el puerto de Embarque, en la linea " & i, vbInformation
                        Exit Function
                    End If
                    reg.Close
                End If
                
                Spr_dtx.Col = 7
                If Spr_dtx.Text = "" Then
                    MsgBox "Debe ingresar un valor valido para el puerto de Origen, en la linea " & i, vbInformation
                    Exit Function
                Else
                    reg.Open "select * from docu_puerto where puer_codigo = '" & Spr_dtx.Text & "'", Gcon_docu
                    If reg.EOF Then
                        MsgBox "Debe ingresar un codigo valido para el puerto de Origen, en la linea " & i, vbInformation
                        Exit Function
                    End If
                    reg.Close
                End If
                
                Spr_dtx.Col = 8
                If Spr_dtx.Text = "" Then
                    MsgBox "Debe ingresar un valor valido para el puerto de Descarga, en la linea " & i, vbInformation
                    Exit Function
                Else
                    reg.Open "select * from docu_puerto where puer_codigo = '" & Spr_dtx.Text & "'", Gcon_docu
                    If reg.EOF Then
                        MsgBox "Debe ingresar un codigo valido para el puerto de Descarga, en la linea " & i, vbInformation
                        Exit Function
                    End If
                    reg.Close
                End If
                
                Spr_dtx.Col = 9
                If Spr_dtx.Text = "" Then
                    MsgBox "Debe ingresar un valor valido para el puerto de Destino Final, en la linea " & i, vbInformation
                    Exit Function
                Else
                    reg.Open "select * from docu_puerto where puer_codigo = '" & Spr_dtx.Text & "'", Gcon_docu
                    If reg.EOF Then
                        MsgBox "Debe ingresar un codigo valido para el puerto de Destino Final, en la linea " & i, vbInformation
                        Exit Function
                    End If
                    reg.Close
                End If
                
                
                
            End If
        End If
    Next i
    
    valida_Datos2 = True
    Exit Function
ERROR_VALIDA:
    MsgBox "Ocurrio el error " & Err.Description
    Screen.MousePointer = 0
    valida_Datos2 = False
End Function


Private Sub cmd_salir_Click()
   Unload Me
End Sub

Private Sub Cmd_Siguiente_Click()
Select Case num_tab
    Case 1:
        Screen.MousePointer = 11
        If Not valida_Datos Then Exit Sub
        carga_parametros_DTX Gstr_Linea, Cmb_linea.BoundText
        carga_parametros Cmb_linea.BoundText
        verifica_version_DTX (Txt_DTX.Text)
        carga_archivo_DTX_grid (Txt_DTX.Text)
        Screen.MousePointer = 0
        Me.Cmd_edita_carga.Visible = True
        num_tab = 2
    Case 2:
        Screen.MousePointer = 11
        If Not valida_Datos2 Then Exit Sub
        Screen.MousePointer = 0
        Me.Cmd_edita_carga.Visible = False
        num_tab = 3
        
    
End Select
Call Carga_frame

End Sub


Private Sub Form_Load()
Dim reg As New ADODB.Recordset
Dim sql As String

On Error GoTo Errores
num_tab = 1
Carga_frame
   Lbl_prog.Width = 0
   DTX_PROCESO.Caption = "Proceso de Carga de Archivos DTX. (" & Gstr_Linea & ")"
   Opt_tipo_impo.Value = True
   Opt_msc.Value = True
   Chk_rela_fija.Value = 1 ' si esta true no va a preguntar si desea establecer la relacion, solo setea el valor por defecto.
   Gbol_relacion_fija = True
   
   carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
   
   SPR_Carga_Cabecera
   
   Exit Sub
Errores:
   MsgBox "Ocurrio el Error " & Err.Description, vbInformation, App.Title
   Exit Sub
   Resume
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 Then
     PopupMenu MNU_DTX
  End If
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
 If Button = 2 Then
     PopupMenu MNU_DTX
  End If
End Sub

Private Sub mnu_acerca_Click()
'DTX_About.Show 1
End Sub

Private Sub mnu_detalle_Click()
 DTX_mant_detalle.Show
End Sub

Private Sub mnu_form_deta_Click()
DTX_mant_formato_det.Show
End Sub

Private Sub mnu_linea_Click()
DTX_mant_linea.Show
End Sub

Private Sub mnu_mant_cond_cont_Click()
   Dtx_mant_cond_contein.Show
End Sub

Private Sub mnu_mant_emba_Click()
   DTX_mant_embal.Show
End Sub

Private Sub Mnu_mant_IMO_Click()
   DTX_mant_imo.Show
End Sub

Private Sub mnu_mant_puer_Click()
   DTX_mant_puer.Show
End Sub

Private Sub mnu_mant_tipo_cont_Click()
   DTX_mant_tipo_cont.Show
End Sub

Private Sub mnu_migra_Click()
   Dim sql As String
   Dim reg As New ADODB.Recordset
   Dim reg2 As New ADODB.Recordset
   Dim i As Integer
   
   sql = "Select * from Puerto Where linea ='" & Gstr_Linea & "'"
   Set reg = Base_MSC.OpenRecordset(sql)
   i = 0
   Do While Not reg.EOF
      sql = "Select * from puerto where puerto = '" & reg("codigo") & "'"
      Set reg2 = Base_APG.OpenRecordset(sql)
      If Not reg2.EOF Then
         sql = "UPDATE PUERTO SET APG='" & reg2("PUERTO") & "' WHERE LINEA ='" & Gstr_Linea & "' AND CODIGO = '" & reg("CODIGO") & "'"
         Base_MSC.Execute sql, dbFailOnError
         i = i + 1
      End If
      reg2.Close
      reg.MoveNext
   Loop
   MsgBox "MIGRACION OK, No. REGISTROS " & i
End Sub

Private Sub mnu_param_Click()
   DTX_mant_parametro.Show
End Sub

Private Sub mnu_repo_cond_cont_Click()
'RPT.ReportFileName = Gstr_ruta_reportes & "dtx_cond_cont.rpt"
'RPT.SelectionFormula = "{dtx_cond_contein.linea} = '" & Gstr_Linea & "'"
'RPT.Connect = Gstr_conex_docu
'RPT.Formulas(1) = "cia ='" & Gstr_cia & "'"
'RPT.Action = 1
End Sub

Private Sub mnu_repo_emba_Click()
'RPT.ReportFileName = Gstr_ruta_reportes & "dtx_embal.rpt"
'RPT.SelectionFormula = "{dtx_embalaje.linea} = '" & Gstr_Linea & "'"
'RPT.Connect = Gstr_conex_docu
'RPT.Formulas(1) = "cia ='" & Gstr_cia & "'"
'RPT.Action = 1
End Sub

Private Sub mnu_repo_estruc_Click()
'RPT.ReportFileName = Gstr_ruta_reportes & "dtx_estruc_arch.rpt"
'RPT.Connect = Gstr_conex_docu
'RPT.SelectionFormula = ""
'RPT.Formulas(1) = "cia ='" & Gstr_cia & "'"
'RPT.Action = 1
'falta rpt
End Sub

Private Sub mnu_repo_imo_Click()
'RPT.ReportFileName = Gstr_ruta_reportes & "dtx_imo.rpt"
'RPT.SelectionFormula = "{dtx_imo.linea} = '" & Gstr_Linea & "'"
'RPT.Connect = Gstr_conex_docu
'RPT.Formulas(1) = "cia ='" & Gstr_cia & "'"
'RPT.Action = 1
End Sub

Private Sub mnu_repo_puerto_Click()
'RPT.ReportFileName = Gstr_ruta_reportes & "dtx_puer.rpt"
'RPT.SelectionFormula = "{dtx_puerto.linea} = '" & Gstr_Linea & "'"
'RPT.Connect = Gstr_conex_docu
'RPT.Formulas(1) = "cia ='" & Gstr_cia & "'"
'RPT.Action = 1
End Sub

Private Sub mnu_repo_tipo_cont_Click()
'RPT.ReportFileName = Gstr_ruta_reportes & "dtx_tipo_cont.rpt"
'RPT.SelectionFormula = "{dtx_tipo_contein.linea} = '" & Gstr_Linea & "'"
'RPT.Connect = Gstr_conex_docu
'RPT.Formulas(1) = "cia ='" & Gstr_cia & "'"
'RPT.Action = 1
End Sub

Private Sub mnu_salir_Click()
   Unload Me
End Sub

Private Sub mnu_TIPO_Click()
  DTX_mant_TIPO.Show
End Sub

Private Sub Msk_atraque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Opt_khs_Click()
    Gstr_Linea = "KHS"
End Sub

Private Sub Opt_msc_Click()
    Gstr_Linea = "MSC"
End Sub

Private Sub Opt_tipo_expo_Click()
If Opt_tipo_expo.Value Then Gstr_Tipo = "E"
End Sub

Private Sub Opt_tipo_impo_Click()
If Opt_tipo_impo.Value Then Gstr_Tipo = "I"
End Sub



Public Sub SPR_Carga_Cabecera()

  Spr_dtx.MaxCols = 10
  Spr_dtx.MaxRows = 0
  Spr_dtx.Row = 0:
  Spr_dtx.Col = 1: Spr_dtx.Text = "St"
  Spr_dtx.Col = 2: Spr_dtx.Text = "Embarq"
  Spr_dtx.Col = 3: Spr_dtx.Text = "Origen"
  Spr_dtx.Col = 4: Spr_dtx.Text = "Descarga"
  Spr_dtx.Col = 5: Spr_dtx.Text = "Dest Final"
  
  Spr_dtx.Col = 6: Spr_dtx.Text = "Embarq"
  Spr_dtx.Col = 7: Spr_dtx.Text = "Origen"
  Spr_dtx.Col = 8: Spr_dtx.Text = "Descarga"
  Spr_dtx.Col = 9: Spr_dtx.Text = "Dest Final"
  
  
  Spr_dtx.Col = 10: Spr_dtx.Text = "BL"
  
  
  
  Spr_dtx.ColWidth(1) = 5
  Spr_dtx.ColWidth(2) = 8
  Spr_dtx.ColWidth(3) = 8
  Spr_dtx.ColWidth(4) = 8
  Spr_dtx.ColWidth(5) = 8
  Spr_dtx.ColWidth(6) = 8
  Spr_dtx.ColWidth(7) = 8
  Spr_dtx.ColWidth(8) = 8
  Spr_dtx.ColWidth(9) = 8
  Spr_dtx.ColWidth(10) = 15
  
  Spr_dtx.ColsFrozen = 1
  'Spr_dtx.Col = 23: Spr_dtx.ColHidden = True
  
End Sub


Private Sub Carga_frame()
    
    Select Case num_tab
    Case 1:
        Fra_DTX.Visible = True: Fra_confirma.Visible = False: Fra_proceso.Visible = False
        Cmd_Anterior.Enabled = False
        Cmd_Siguiente.Enabled = True
        Cmd_Procesar.Enabled = False
        Cmd_Siguiente.Visible = True
        Cmd_Procesar.Visible = False
        Me.Cmd_edita_carga.Visible = False
    Case 2:
        Fra_DTX.Visible = False: Fra_confirma.Visible = True: Fra_proceso.Visible = False
        Cmd_Anterior.Enabled = True
        Cmd_Siguiente.Enabled = True
        Cmd_Procesar.Enabled = False
        Cmd_Siguiente.Visible = True
        Cmd_Procesar.Visible = False
        Me.Cmd_edita_carga.Visible = True
    Case 3:
        Fra_DTX.Visible = False: Fra_confirma.Visible = False: Fra_proceso.Visible = True
        Cmd_Anterior.Enabled = True
        Cmd_Siguiente.Enabled = False
        Cmd_Procesar.Enabled = True
        Cmd_Siguiente.Visible = False
        Cmd_Procesar.Visible = True
        Me.Cmd_edita_carga.Visible = False
    End Select
End Sub



'Funcion : setea formato spread
' lo que esta funcion realiza es setear  el formato de cada celda del spread
' 0 : date
' 1 : Edit
' 2 : float
' 3 : Integer
' 5 : Static Cell
' 6 : Time
' 7 : Button
' 8 : Combo
' 9 : Picture
'10 : check_box
'---------------------------------------------------
Public Sub Setea_Formato_Spread_pto(ByVal fila As Integer)
  If fila < 0 Then fila = 1

  Spr_dtx.Row = fila
  Spr_dtx.Col = 1: Spr_dtx.CellType = 10: Spr_dtx.TypeCheckCenter = True: Spr_dtx.Value = True
  Spr_dtx.Col = 2: Spr_dtx.CellType = 5
  Spr_dtx.Col = 3: Spr_dtx.CellType = 5
  Spr_dtx.Col = 4: Spr_dtx.CellType = 5
  Spr_dtx.Col = 5: Spr_dtx.CellType = 5
  Spr_dtx.Col = 6: Spr_dtx.CellType = 1
  Spr_dtx.Col = 7: Spr_dtx.CellType = 1
  Spr_dtx.Col = 8: Spr_dtx.CellType = 1
  Spr_dtx.Col = 9: Spr_dtx.CellType = 1
  Spr_dtx.Col = 10: Spr_dtx.CellType = 5
  
'  Spr_dtx.Col = 4: Spr_dtx.CellType = 7:  Spr_dtx.TypeButtonText = "?"
'  Spr_dtx.Col = 5: Spr_dtx.CellType = 5
'  Spr_dtx.Col = 7: Spr_dtx.CellType = 7:  Spr_dtx.TypeButtonText = "?"
'  Spr_dtx.Col = 18: Spr_dtx.CellType = 8:  Spr_dtx.TypeComboBoxList = Lstr_spr_Tipo_Cont
'  Spr_dtx.Col = 23: Spr_dtx.CellType = 8:  Spr_dtx.TypeComboBoxList = Lstr_spr_Tipo_Cont_cod
'  'Spr_dtx.Col = 2: Spr_dtx.CellType = 2: Spr_dtx.TypeFloatDecimalPlaces = 2: Spr_adicionales.TypeFloatMin = "0.00": Spr_adicionales.TypeFloatMax = "999.99": Spr_adicionales.TypeFloatSeparator = True: Spr_adicionales.TypeFloatDecimalChar = Asc("."): Spr_adicionales.TypeFloatSepChar = Asc(",")
'  'Spr_dtx.Col = 3: Spr_dtx.CellType = 2: Spr_dtx.TypeFloatDecimalPlaces = 2: Spr_adicionales.TypeFloatMin = "0.00": Spr_adicionales.TypeFloatMax = "9,999,999.99": Spr_adicionales.TypeFloatSeparator = True: Spr_adicionales.TypeFloatDecimalChar = Asc("."): Spr_adicionales.TypeFloatSepChar = Asc(",")
'  'Spr_dtx.Col = 4: Spr_dtx.CellType = 1
'
  
  Spr_dtx.Col = 2: Spr_dtx.BackColor = &H80FF80
  Spr_dtx.Col = 3: Spr_dtx.BackColor = &H80FF80
  Spr_dtx.Col = 4: Spr_dtx.BackColor = &H80FF80
  Spr_dtx.Col = 5: Spr_dtx.BackColor = &H80FF80
  
End Sub

Public Sub Setea_Formato_Spread_bl(ByVal fila As Integer)
  If fila < 0 Then fila = 1

  Spr_dtx.Row = fila
  Spr_dtx.Col = 1: Spr_dtx.CellType = 10: Spr_dtx.TypeCheckCenter = True: Spr_dtx.Value = True
  Spr_dtx.Col = 2: Spr_dtx.CellType = 5
  
  
  Spr_dtx.Col = 3: Spr_dtx.CellType = 5
  Spr_dtx.Col = 4: Spr_dtx.CellType = 5
  Spr_dtx.Col = 5: Spr_dtx.CellType = 5
  Spr_dtx.Col = 6: Spr_dtx.CellType = 5
  Spr_dtx.Col = 7: Spr_dtx.CellType = 5
  Spr_dtx.Col = 8: Spr_dtx.CellType = 5
  Spr_dtx.Col = 9: Spr_dtx.CellType = 5
  
  
  Spr_dtx.Col = 2: Spr_dtx.BackColor = &HC0FFFF
  Spr_dtx.Col = 3: Spr_dtx.BackColor = &HC0FFFF
  Spr_dtx.Col = 4: Spr_dtx.BackColor = &HC0FFFF
  Spr_dtx.Col = 5: Spr_dtx.BackColor = &HC0FFFF
  Spr_dtx.Col = 6: Spr_dtx.BackColor = &HC0FFFF
  Spr_dtx.Col = 7: Spr_dtx.BackColor = &HC0FFFF
  Spr_dtx.Col = 8: Spr_dtx.BackColor = &HC0FFFF
  Spr_dtx.Col = 9: Spr_dtx.BackColor = &HC0FFFF
  Spr_dtx.Col = 10: Spr_dtx.BackColor = &HC0FFFF
  
End Sub

