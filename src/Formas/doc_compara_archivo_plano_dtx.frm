VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form doc_archivo_plano 
   Caption         =   "Carga de información de Archivo Plano"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   1845
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   8715
   Begin VB.Frame Frame2 
      Height          =   630
      Left            =   5490
      TabIndex        =   16
      Top             =   1050
      Width           =   3030
      Begin VB.OptionButton Opt_msc 
         Caption         =   "Formato MSC"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   255
         Width           =   1455
      End
      Begin VB.OptionButton Opt_khs 
         Caption         =   "Formato KHS"
         Height          =   195
         Left            =   1650
         TabIndex        =   17
         Top             =   225
         Width           =   1245
      End
   End
   Begin FPSpread.vaSpread Spr_plano 
      Height          =   2490
      Left            =   90
      TabIndex        =   9
      Top             =   1740
      Width           =   8340
      _Version        =   131077
      _ExtentX        =   14711
      _ExtentY        =   4392
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
      SpreadDesigner  =   "doc_archivo_plano.frx":0000
   End
   Begin VB.CommandButton Cmd_buscar 
      Caption         =   "Consulta"
      Height          =   660
      Left            =   5565
      Picture         =   "doc_archivo_plano.frx":00FD
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Consulta información del Plano"
      Top             =   150
      Width           =   750
   End
   Begin VB.CommandButton Cmd_salir 
      Caption         =   "Salir"
      Height          =   660
      Left            =   7095
      Picture         =   "doc_archivo_plano.frx":0407
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   150
      Width           =   750
   End
   Begin VB.CommandButton Cmd_modificar 
      Caption         =   "Crear Bls"
      Height          =   660
      Left            =   6330
      Picture         =   "doc_archivo_plano.frx":0711
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Crea Bls con Cnts vacíos"
      Top             =   150
      Width           =   750
   End
   Begin VB.Frame Frame1 
      Height          =   1680
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   5355
      Begin VB.Frame Fra_buque 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   180
         TabIndex        =   13
         Top             =   1005
         Width           =   3090
         Begin MSDataListLib.DataCombo Cmb_buque 
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_buque 
            Height          =   330
            Left            =   1920
            Top             =   240
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
            TabIndex        =   15
            Top             =   0
            Width           =   555
         End
      End
      Begin VB.Frame Fra_viaje 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   3300
         TabIndex        =   10
         Top             =   1005
         Width           =   1965
         Begin MSDataListLib.DataCombo Cmb_viaje 
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_viaje 
            Height          =   330
            Left            =   795
            Top             =   240
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
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Viaje"
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
            Left            =   135
            TabIndex        =   12
            Top             =   0
            Width           =   435
         End
      End
      Begin VB.CommandButton Cmd_DTX 
         Caption         =   "..."
         Height          =   285
         Left            =   4590
         TabIndex        =   5
         ToolTipText     =   "Seleccionar Archivo a Procesar"
         Top             =   690
         Width           =   285
      End
      Begin VB.TextBox Txt_DTX 
         Height          =   285
         Left            =   1815
         TabIndex        =   4
         Top             =   690
         Width           =   2715
      End
      Begin MSDataListLib.DataCombo Cmb_linea 
         Height          =   315
         Left            =   1815
         TabIndex        =   1
         Top             =   225
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc data_linea 
         Height          =   330
         Left            =   4440
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
         Caption         =   "Adodc1"
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
      Begin MSComDlg.CommonDialog Dlg_DTX 
         Left            =   4770
         Top             =   585
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Abrir Archivo DTX"
      End
      Begin VB.Label Label6 
         Caption         =   "Archivo Plano :"
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
         Left            =   270
         TabIndex        =   6
         Top             =   690
         Width           =   1365
      End
      Begin VB.Label Label2 
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
         Left            =   270
         TabIndex        =   2
         Top             =   300
         Width           =   825
      End
   End
End
Attribute VB_Name = "doc_archivo_plano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lstr_spr_Tipo_Cont      As String
Dim Lstr_spr_Tipo_Cont_cod  As String


Private Sub Cmb_buque_Change()
   Dim sql As String
   sql = "Select * from docu_buque_viaje " & _
         " Where buqv_linea ='" & Cmb_linea.BoundText & "'" & _
         "   And buqv_buque ='" & Cmb_buque.BoundText & "'" & _
         "   And buqv_status = 'A'"
   
   carga_datacombo data_viaje, Cmb_viaje, sql, "buqv_viaje", "buqv_viaje"
End Sub

Private Sub Cmb_linea_Change()
Dim reg As New ADODB.Recordset
   carga_datacombo data_buque, Cmb_buque, "Select * from docu_buque Where buqu_linea = '" & Cmb_linea.BoundText & "' and buqu_status ='A'", "buqu_codigo", "buqu_descripcion"
   
   'CARGA DATOS PARA EL COMBO DE TIPOS DE CONTENEDOR
   sql = "Select tama_equip_codigo, tama_equip_descripcion From docu_tamano_equip Where tama_equip_linea ='" & Cmb_linea.BoundText & "' and tama_equip_status = 'A' Order by tama_equip_descripcion"
   carga_combo_Spread sql, "tama_equip_codigo", "tama_equip_descripcion", Lstr_spr_Tipo_Cont, Lstr_spr_Tipo_Cont_cod
   
End Sub

Private Sub Cmd_buscar_Click()
Dim sql             As String
Dim reg_plano_det   As New ADODB.Recordset
Dim PLANO_LINEA     As String
Dim linea           As String
Dim Columna         As Integer
' datos del archivo
Dim contenedor      As String
Dim status          As String
Dim pto_emb         As String
Dim pto_desc        As String
Dim pto_final       As String
Dim cont_size       As String
Dim peso            As Double
Dim volumen         As Double
Dim Tipo_cont       As String
' se valida que exista la ruta del archivo
If Dir(Txt_DTX.Text, vbArchive) = "" Then
    MsgBox "El archivo seleccionado no existe", vbInformation
    Exit Sub
End If
If Opt_msc.Value Then
    PLANO_LINEA = "MSC"
ElseIf Opt_khs.Value Then
    PLANO_LINEA = "KHS"
End If
Me.MousePointer = 11
Call carga_parametros_plano(PLANO_LINEA, Cmb_linea.BoundText)
Spr_plano.MaxRows = 0

sql = "Select * from plano_detalle Where plano_det_linea = '" & PLANO_LINEA & "'"
reg_plano_det.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
If reg_plano_det.EOF Then Exit Sub

x = FreeFile
Open Txt_DTX.Text For Input As #x
Do While Not EOF(x)
    Line Input #x, linea
    reg_plano_det.MoveFirst
    contenedor = "":    status = "":    cont_size = ""
    pto_emb = "":       pto_desc = "":  pto_final = ""
    If Mid(linea, 1, 1) = "$" Then GoTo siguiente ' esto es para las lineas iniciales que viene con signo dolar y no traen mas que inf general
    Do While Not reg_plano_det.EOF
        Select Case Trim(reg_plano_det("plano_det_codigo"))
            Case "CONTENEDOR":  ' se hace validacion para verificar si la linea leida es de un contenedor es una cabecera o resumen de algo
                    contenedor = Mid(linea, reg_plano_det("plano_det_Desde"), reg_plano_det("plano_det_longitud"))
                    If Not IsNumeric(Mid(contenedor, 5)) Then Exit Do
            Case "STATUS":      status = Mid(linea, reg_plano_det("plano_det_Desde"), reg_plano_det("plano_det_longitud"))
            Case "CONT_SIZE":   cont_size = Mid(linea, reg_plano_det("plano_det_Desde"), reg_plano_det("plano_det_longitud"))
            Case "PTO_EMB":     pto_emb = Mid(linea, reg_plano_det("plano_det_Desde"), reg_plano_det("plano_det_longitud"))
            Case "PTO_DESC":    pto_desc = Mid(linea, reg_plano_det("plano_det_Desde"), reg_plano_det("plano_det_longitud"))
            Case "PTO_FINAL":   pto_final = Mid(linea, reg_plano_det("plano_det_Desde"), reg_plano_det("plano_det_longitud"))
        End Select
        reg_plano_det.MoveNext
    Loop
    'poner estos valores en una tabla de parametros para los planos 'esto es mas para exportacion que para importacion
    If Trim(status) = "E" And pto_emb = "GYE" Then
        Spr_plano.MaxRows = Spr_plano.MaxRows + 1
        Setea_Formato_Spread Spr_plano.MaxRows
        If (PLANO_LINEA = "MSC" And Mid(cont_size, 3, 1) = "2") Or _
           (PLANO_LINEA = "KHS" And Mid(cont_size, 1, 1) = "2") Then
                peso = Gstr_param_plano.PESO_CONT20
                volumen = Gstr_param_plano.Volumen20
        Else
                peso = Gstr_param_plano.PESO_CONT40
                volumen = Gstr_param_plano.Volumen40
        End If
        
        Spr_plano.Row = Spr_plano.MaxRows
        'Spr_plano.Col = 4: Spr_plano.Text = status
        'Spr_plano.Col = 5: Spr_plano.Text = pto_emb
        Spr_plano.Col = 1: Spr_plano.Value = True
        Spr_plano.Col = 2: Spr_plano.Text = pto_desc
        ' si el codigo de pto es opcional se pone el codigo equivalente por defecto
        If pto_desc = "OPT" Then
           Spr_plano.Col = 3: Spr_plano.Text = Gstr_param_plano.pto_opt
        End If
        ' si el pto final esta en blanco o es --- se le pone el mismo pto de descarga
        If pto_final = "---" Or Trim(pto_final) = "" Then
            Spr_plano.Col = 5: Spr_plano.Text = pto_desc
            pto_final = pto_desc
        Else
            Spr_plano.Col = 5: Spr_plano.Text = pto_final
        End If
        ' si el codigo de pto es opcional se pone el codigo equivalente por defecto
        If pto_final = "OPT" Then
           Spr_plano.Col = 6: Spr_plano.Text = Gstr_param_plano.pto_opt
        End If
        
        Spr_plano.Col = 8: Spr_plano.Text = "BLTemp_" & pto_desc & "_" & IIf(pto_final <> "---" And Trim(pto_final) <> "" And pto_final <> pto_desc, pto_final & "_", "") & "001"
        Spr_plano.Col = 9: Spr_plano.Text = Gstr_param_plano.consignatario
        Spr_plano.Col = 10: Spr_plano.Text = Gstr_param_plano.Direc_consignatario
        Spr_plano.Col = 11: Spr_plano.Text = Gstr_param_plano.notificador
        Spr_plano.Col = 12: Spr_plano.Text = Gstr_param_plano.Direc_notificador
        Spr_plano.Col = 13: Spr_plano.Text = Gstr_param_plano.embarcador
        Spr_plano.Col = 14: Spr_plano.Text = Gstr_param_plano.Direc_Embarcador
        Spr_plano.Col = 15: Spr_plano.Text = "1"
        Spr_plano.Col = 16: Spr_plano.Text = Gstr_param_plano.embalaje
        Spr_plano.Col = 17: Spr_plano.Text = contenedor
        Spr_plano.Col = 23: Tipo_cont = Codigo_equivalente_PLANO("plano_tipo_cont_linea_proc", "plano_tipo_cont_linea", Cmb_linea.BoundText, PLANO_LINEA, "PLANO_TIPO_CONT", cont_size, "PLANO_TIPO_CONT_CODIGO", "PLANO_TIPO_CONT_EQ")
        For j = 0 To Spr_plano.TypeComboBoxCount
            Spr_plano.TypeComboBoxCurSel = j
            If Trim(Spr_plano.Text) = Tipo_cont Then Exit For
        Next j
        Spr_plano.Col = 18: Spr_plano.TypeComboBoxCurSel = j
        Spr_plano.Col = 19: Spr_plano.Text = Gstr_param_plano.Sello
        Spr_plano.Col = 20: Spr_plano.Text = peso
        Spr_plano.Col = 21: Spr_plano.Text = volumen
        Spr_plano.Col = 22: Spr_plano.Text = Gstr_param_plano.Imo
       
    End If
siguiente:
Loop
Close #x
Me.MousePointer = 0

'  Spr_plano.Col = 2: Spr_plano.Text = "Pto Desc"
'  Spr_plano.Col = 3: Spr_plano.Text = "Pto Desc Eq"
'  Spr_plano.Col = 4: Spr_plano.Text = "" 'boton de ayuda
'  Spr_plano.Col = 5: Spr_plano.Text = "Pto Final"
'  Spr_plano.Col = 6: Spr_plano.Text = "Pto Final Eq"
'  Spr_plano.Col = 7: Spr_plano.Text = "" 'boton de ayuda
'  Spr_plano.Col = 8: Spr_plano.Text = "BL"
'  Spr_plano.Col = 9: Spr_plano.Text = "Consignatario"
'  Spr_plano.Col = 10: Spr_plano.Text = "Direc Consig"
'  Spr_plano.Col = 11: Spr_plano.Text = "Notificador"
'  Spr_plano.Col = 12: Spr_plano.Text = "Direc Notif"
'  Spr_plano.Col = 13: Spr_plano.Text = "Embarcador"
'  Spr_plano.Col = 14: Spr_plano.Text = "Direc Embarcador"
'  Spr_plano.Col = 15: Spr_plano.Text = "Bultos"
'  Spr_plano.Col = 16: Spr_plano.Text = "Embalaje"
'  Spr_plano.Col = 17: Spr_plano.Text = "Contenedor"
'  Spr_plano.Col = 18: Spr_plano.Text = "Tipo Cont"
'  Spr_plano.Col = 19: Spr_plano.Text = "Sello"
'  Spr_plano.Col = 20: Spr_plano.Text = "Peso"
'  Spr_plano.Col = 21: Spr_plano.Text = "Volumen"
'  Spr_plano.Col = 22: Spr_plano.Text = "IMO"
'  Spr_plano.Col = 23: Spr_plano.Text = "Cod TipoCont"


End Sub

Private Sub Cmd_DTX_Click()
   Dlg_DTX.Filter = "Todos los Archivos (*.*)|*.*|Archivos DTX (*.dtx)|*.dtx"
   Dlg_DTX.ShowOpen
   Txt_DTX.Text = Dlg_DTX.FileName
End Sub

Private Sub cmd_modificar_Click()
crea_bls
End Sub

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
If Not data_linea.Recordset.EOF Then data_linea.Recordset.MoveFirst
Opt_msc.Value = True
Carga_Cabecera
'Spr_plano.MaxRows = 1
'Setea_Formato_Spread Spr_plano.MaxRows

End Sub

Public Sub Carga_Cabecera()

  Spr_plano.MaxCols = 23
  Spr_plano.MaxRows = 0
  Spr_plano.Row = 0:
  Spr_plano.Col = 1: Spr_plano.Text = "Cargar? "
  Spr_plano.Col = 2: Spr_plano.Text = "Pto Desc"
  Spr_plano.Col = 3: Spr_plano.Text = "Pto Desc Eq"
  Spr_plano.Col = 4: Spr_plano.Text = "?" 'boton de ayuda
  Spr_plano.Col = 5: Spr_plano.Text = "Pto Final"
  Spr_plano.Col = 6: Spr_plano.Text = "Pto Final Eq"
  Spr_plano.Col = 7: Spr_plano.Text = "?" 'boton de ayuda
  Spr_plano.Col = 8: Spr_plano.Text = "BL"
  Spr_plano.Col = 9: Spr_plano.Text = "Consignatario"
  Spr_plano.Col = 10: Spr_plano.Text = "Direc Consig"
  Spr_plano.Col = 11: Spr_plano.Text = "Notificador"
  Spr_plano.Col = 12: Spr_plano.Text = "Direc Notif"
  Spr_plano.Col = 13: Spr_plano.Text = "Embarcador"
  Spr_plano.Col = 14: Spr_plano.Text = "Direc Embarcador"
  Spr_plano.Col = 15: Spr_plano.Text = "Bultos"
  Spr_plano.Col = 16: Spr_plano.Text = "Embalaje"
  Spr_plano.Col = 17: Spr_plano.Text = "Contenedor"
  Spr_plano.Col = 18: Spr_plano.Text = "Tipo Cont"
  Spr_plano.Col = 19: Spr_plano.Text = "Sello"
  Spr_plano.Col = 20: Spr_plano.Text = "Peso"
  Spr_plano.Col = 21: Spr_plano.Text = "Volumen"
  Spr_plano.Col = 22: Spr_plano.Text = "IMO"
  Spr_plano.Col = 23: Spr_plano.Text = "Cod TipoCont"
  
  Spr_plano.ColWidth(1) = 5
  Spr_plano.ColWidth(2) = 8
  Spr_plano.ColWidth(3) = 10
  Spr_plano.ColWidth(4) = 3
  Spr_plano.ColWidth(5) = 8
  Spr_plano.ColWidth(6) = 10
  Spr_plano.ColWidth(7) = 3
  Spr_plano.ColWidth(8) = 12
  Spr_plano.ColWidth(9) = 12
  Spr_plano.ColWidth(10) = 15
  Spr_plano.ColWidth(11) = 12
  Spr_plano.ColWidth(12) = 15
  Spr_plano.ColWidth(13) = 12
  Spr_plano.ColWidth(14) = 15
  Spr_plano.ColWidth(15) = 6
  Spr_plano.ColWidth(16) = 5
  Spr_plano.ColWidth(17) = 11
  Spr_plano.ColWidth(18) = 18
  Spr_plano.ColWidth(19) = 5
  Spr_plano.ColWidth(20) = 5
  Spr_plano.ColWidth(21) = 5
  Spr_plano.ColWidth(22) = 5
  Spr_plano.ColWidth(23) = 5
  Spr_plano.ColsFrozen = 1
  Spr_plano.Col = 23: Spr_plano.ColHidden = True
  
End Sub


'Funcion        : Codigo Equivalente_PLANO
'Objetivo       : Devolver el código equivalente en CAE, MSC o KHS en base al código del PLANO
'                 de los archivos DTX
'Fecha Creacion : LUNES 3 de Marzo del 2003 - 17:30:00(carnaval)
'Creado por     : Carlos Reyes Atiencia
'Parámetros     : Linea_proc.-          nombre del campo de la linea a la que se va a subir la informacion puede ser CAE, MSC y KHS
'                 Linea.-               nombre del campo de la linea de la que es el formato del archivo  puede ser unicamente MSC y KHS
'                 codigo_Linea_proc.-   valor de la linea a la que se va a subir la informacion puede ser CAE, MSC y KHS
'                 codigo_Linea.-        valor de la linea de la que es el formato del archivo  puede ser unicamente MSC y KHS
'                 Tabla .-              es el nombre de la tabla intermedia que tiene ambos codigos
'                 codigo_buscado .-     es el valor en si que se esta buscando
'                 campo .-              es el nombre del campo que contiene el código plano
'                 campo_Eq .-           es el nombre del campo que contiene el código equivalente al codigo del plano es el que se devuelve
'                 condicion.-           condicion adicional opcional
'===============================================================================

Public Function Codigo_equivalente_PLANO(linea_proc As String, linea As String, codigo_Linea_proc As String, codigo_Linea As String, tabla As String, ByVal codigo_buscado As String, campo As String, campo_eq As String, Optional condicion As String) As String

On Error GoTo error_busca_cod_apg
   Dim sql As String
   Dim REGISTRO   As New ADODB.Recordset
   Dim reg        As New ADODB.Recordset
   Dim Response   As Integer
   Dim ENCONTRADO As Boolean
   Dim ST_NUEVO   As Boolean
   
   sql = "select " & Trim$(campo_eq) & " from " & Trim$(tabla) & _
         " where " & linea_proc & "='" & codigo_Linea_proc & "' " & _
         "   And " & linea & "     ='" & codigo_Linea & "'      " & _
         "   And " & Trim$(campo) & " = '" & Trim$(codigo_buscado) & "'"
      
   If Not IsMissing(condicion) And Trim$(condicion) <> "" Then
      sql = sql + " and " & condicion
   End If
      
   REGISTRO.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   
   ENCONTRADO = True:   ST_NUEVO = True
   If REGISTRO.EOF Then
      ENCONTRADO = False
   Else
      If Trim(REGISTRO.Fields(campo_eq)) = "" Or IsNull(REGISTRO.Fields(campo_eq)) Then
         ENCONTRADO = False
      End If
   End If
   
   If Not ENCONTRADO Then
      Codigo_equivalente_PLANO = "NO"
   Else
      Codigo_equivalente_PLANO = Trim(REGISTRO.Fields(campo_eq))
   End If
   REGISTRO.Close
   
   ' se busca en la tabla maestra para verificar si existe en la tabla de maestra de su correspondiente
'   Select Case UCase(tabla)
'            Case "DTX_EMBALAJE":
'                  sql = "Select * from docu_embalaje where emba_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "' and emba_codigo ='" & Codigo_equivalente_DTX & "' and emba_status = 'A'"
'                  reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
'                  If reg.EOF Then Codigo_equivalente_DTX = Gstr_Embalaje
'            Case "DTX_PUERTO":
'                  sql = "Select * from docu_puerto where puer_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "' and puer_codigo ='" & Codigo_equivalente_DTX & "'"
'                  reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
'                  If reg.EOF Then
'                        MsgBox "El código de puerto " & codigo_buscado & " NO existe o esta mal relacionado, por favor verifique la relación, por el momento se pondra el codigo por defecto " & Gstr_Puerto, vbCritical
'                        Codigo_equivalente_DTX = Gstr_Puerto
'                  End If
'            Case "DTX_IMO":
'                  sql = "Select * from docu_imo where imo_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "' and imo_codigo ='" & Codigo_equivalente_DTX & "'"
'                  reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
'                  If reg.EOF Then Codigo_equivalente_DTX = Gstr_IMO
'            Case "DTX_COND_CONTEIN":
'                  sql = "Select * from docu_cond_contenedor where cond_cont_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "' and cond_cont_codigo ='" & Codigo_equivalente_DTX & "' and cond_cont_status ='A'"
'                  reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
'                  If reg.EOF Then Codigo_equivalente_DTX = Gstr_Cond_Contein
'            Case "DTX_MONEDA":
'                  sql = "Select * from docu_moneda where mone_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "' and mone_codigo ='" & Codigo_equivalente_DTX & "'"
'                  reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
'                  If reg.EOF Then Codigo_equivalente_DTX = Gstr_moneda
'            Case "DTX_TIPO_CONTEIN":
'                  sql = "Select * from docu_tamano_equip where tama_equip_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "' and tama_equip_codigo ='" & Codigo_equivalente_DTX & "' and tama_equip_status ='A'"
'                  reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
'                  If reg.EOF Then
'                    If codigo_buscado = "20" Or codigo_buscado = "0" Or codigo_buscado = "1" Or codigo_buscado = "2" Then
'                       Codigo_equivalente_DTX = Gstr_Tipo_Contein    ' de 20
'                    Else
'                       Codigo_equivalente_DTX = Gstr_Tipo_Contein_40 ' de 40
'                    End If
'                  End If
'            Case "DOCU_DEBTOR":
'                  sql = "Select * from docu_debtor where debtor_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "' and debtor_codigo ='" & Codigo_equivalente_DTX & "'"
'                  reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
'                  If reg.EOF Then Codigo_equivalente_DTX = Gstr_Debtor
'            Case "DOCU_CARGO":
'                  sql = "Select * from docu_cargo where cargo_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "' and cargo_codigo ='" & Codigo_equivalente_DTX & "'"
'                  reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
'                  If reg.EOF Then Codigo_equivalente_DTX = Gstr_Cargo
'   End Select
'   reg.Close

   Exit Function
   
error_busca_cod_apg:
   MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
   Exit Function
   Resume 0
End Function


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
Private Sub Setea_Formato_Spread(ByVal fila As Integer)
  If fila < 0 Then fila = 1

  Spr_plano.Row = fila
  Spr_plano.Col = 1: Spr_plano.CellType = 10: Spr_plano.TypeCheckCenter = True
  Spr_plano.Col = 2: Spr_plano.CellType = 5
  Spr_plano.Col = 4: Spr_plano.CellType = 7:  Spr_plano.TypeButtonText = "?"
  Spr_plano.Col = 5: Spr_plano.CellType = 5
  Spr_plano.Col = 7: Spr_plano.CellType = 7:  Spr_plano.TypeButtonText = "?"
  Spr_plano.Col = 18: Spr_plano.CellType = 8:  Spr_plano.TypeComboBoxList = Lstr_spr_Tipo_Cont
  Spr_plano.Col = 23: Spr_plano.CellType = 8:  Spr_plano.TypeComboBoxList = Lstr_spr_Tipo_Cont_cod
  'Spr_plano.Col = 2: Spr_plano.CellType = 2: Spr_plano.TypeFloatDecimalPlaces = 2: Spr_adicionales.TypeFloatMin = "0.00": Spr_adicionales.TypeFloatMax = "999.99": Spr_adicionales.TypeFloatSeparator = True: Spr_adicionales.TypeFloatDecimalChar = Asc("."): Spr_adicionales.TypeFloatSepChar = Asc(",")
  'Spr_plano.Col = 3: Spr_plano.CellType = 2: Spr_plano.TypeFloatDecimalPlaces = 2: Spr_adicionales.TypeFloatMin = "0.00": Spr_adicionales.TypeFloatMax = "9,999,999.99": Spr_adicionales.TypeFloatSeparator = True: Spr_adicionales.TypeFloatDecimalChar = Asc("."): Spr_adicionales.TypeFloatSepChar = Asc(",")
  'Spr_plano.Col = 4: Spr_plano.CellType = 1
  
  
  Spr_plano.Col = 3: Spr_plano.BackColor = &HC0FFFF
  Spr_plano.Col = 6: Spr_plano.BackColor = &HC0FFFF
  Spr_plano.Col = 17: Spr_plano.BackColor = &HC0FFFF
  
End Sub

Private Sub Form_Resize()
If Me.Height - 2385 > 0 Then Spr_plano.Height = Me.Height - 2385
If Me.Width - 465 > 0 Then Spr_plano.Width = Me.Width - 465
End Sub

Private Sub Spr_plano_Change(ByVal Col As Long, ByVal Row As Long)
Dim pto_origen As String
Dim pto_eq     As String

Dim pto_1      As String
Dim pto_2      As String
Dim bl         As String

Spr_plano.Col = Col
 Spr_plano.Row = Row
 Select Case Col
    ' el pto eq se setea igual para todos las filas en donde el codigo de pto original(pto_desc, pto_final) sea igual al de la fila cambiada
    Case 3, 6  ' pto descarga,  pto final
        pto_eq = Spr_plano.Text
        Spr_plano.Col = Col - 1: pto_origen = Spr_plano.Text
        If MsgBox("Desea aplicar este valor a todos los Puertos con el mismo codigo original", vbYesNo + vbQuestion) = vbYes Then
            For i = 1 To Spr_plano.MaxRows
                Spr_plano.Row = i
                Spr_plano.Col = Col - 1
                If Spr_plano.Text = pto_origen Then
                    Spr_plano.Col = Col
                    Spr_plano.Text = pto_eq
                End If
            Next i
        End If
     Case 8 ' para  cambiar el codigo del bl de todos los registros con el mismo pto desc y pto final
        bl = Spr_plano.Text
        Spr_plano.Col = 2: pto_1 = Spr_plano.Text
        Spr_plano.Col = 5: pto_2 = Spr_plano.Text
        If MsgBox("Desea aplicar este valor a todos los Bls con el mismo codigo de Pto Descarga y Pto Final", vbYesNo + vbQuestion) = vbYes Then
            For i = 1 To Spr_plano.MaxRows
                Spr_plano.Row = i
                Spr_plano.Col = 2
                If Spr_plano.Text = pto_1 Then
                    Spr_plano.Col = 5
                    If Spr_plano.Text = pto_2 Then
                       Spr_plano.Col = Col
                       Spr_plano.Text = bl
                    End If
                End If
            Next i
        End If
 End Select
 Spr_plano.SetFocus
End Sub

Private Sub Spr_plano_Click(ByVal Col As Long, ByVal Row As Long)
 Spr_plano.Col = Col
 Spr_plano.Row = Row
 Select Case Col
    Case 4:
         If Row = 0 Then Exit Sub
         If Spr_plano.BlockMode Then Exit Sub
         
         pto = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Descarga")
         Spr_plano.Col = 2: Spr_plano.Text = pto
         'Call Spr_adicionales_Change(1, Row)
         SendKeys "{ENTER}"
    Case 7:
         If Row = 0 Then Exit Sub
         If Spr_plano.BlockMode Then Exit Sub
         
         pto = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Destino Final")
         Spr_plano.Col = 5: Spr_plano.Text = pto
         'Call Spr_adicionales_Change(1, Row)
         SendKeys "{ENTER}"
'   Case 5: Call Spr_adicionales_Change(Col, Row)
'   Case 6: Call Spr_adicionales_Change(Col, Row)
 End Select
End Sub

Private Sub crea_bls()
Dim sql                 As String
Dim reg                 As New ADODB.Recordset
Dim pto_desc            As String
Dim pto_final           As String
Dim pto_emb             As String
Dim bl                  As String
Dim cont                As String
Dim fecha_zarpe         As String
Dim consignatario       As String
Dim Direc_consignatario As String
Dim notificador         As String
Dim Direc_notificador   As String
Dim embarcador          As String
Dim Direc_Embarcador    As String

Dim bultos              As Integer
Dim embalaje            As String
Dim Tipo_cont           As String
Dim Sello               As String
Dim peso                As Double
Dim volumen             As Double
Dim Imo                 As String

Dim marca_num           As String
Dim desc_carga          As String

'TOTALES DE ITEMS CREADOS
Dim No_bls                  As Integer
Dim no_cont                 As Integer
Dim No_bl_det               As Integer

On Error GoTo errores_bl

' se realizan validaciones principales
If Not Cmb_linea.MatchedWithList Then
    MsgBox "No existe ninguna Linea seleccionada", vbInformation, Me.Caption
    Exit Sub
End If
If Not Cmb_buque.MatchedWithList Then
    MsgBox "No existe ningun Buque seleccionado", vbInformation, Me.Caption
    Exit Sub
End If
If Not Cmb_viaje.MatchedWithList Then
    MsgBox "No existe ningun Viaje seleccionado", vbInformation, Me.Caption
    Exit Sub
End If

Me.MousePointer = 11
' se inicia el proceso de subir informacion de los bls
carga_parametros Cmb_linea.BoundText
fecha_zarpe = Format(Now, GSTR_FORMATO_FECHA)
desc_carga = Gstr_param_plano.desc_carga
marca_num = Gstr_param_plano.marcas_num
'if expo then
pto_emb = Gstr_param.pto_descarga ' Gstr_param.pto_descarga contiene la informacion del puerto local (ECGYE) que para importacion es descarga y para expo es embarque
No_bls = 0
no_cont = 0
No_bl_det = 0

For i = 1 To Spr_plano.MaxRows
    Spr_plano.Row = i
    Spr_plano.Col = 1
    If CBool(Spr_plano.Value) Then  ' si esta marcado, significa que se lo debe subir, si no NO.
        Spr_plano.Col = 3: pto_desc = Spr_plano.Text
        Spr_plano.Col = 6: pto_final = Spr_plano.Text
        Spr_plano.Col = 8: bl = Spr_plano.Text
        Spr_plano.Col = 17: cont = Spr_plano.Text
        Spr_plano.Col = 9:  consignatario = Spr_plano.Text
        Spr_plano.Col = 10: Direc_consignatario = Spr_plano.Text
        Spr_plano.Col = 11: notificador = Spr_plano.Text
        Spr_plano.Col = 12: Direc_notificador = Spr_plano.Text
        Spr_plano.Col = 13: embarcador = Spr_plano.Text
        Spr_plano.Col = 14: Direc_Embarcador = Spr_plano.Text
  'detalles
        Spr_plano.Col = 15: bultos = Spr_plano.Text
        Spr_plano.Col = 16: embalaje = Spr_plano.Text
        Spr_plano.Col = 17: cont = Spr_plano.Text
        Spr_plano.Col = 18: Tipo_cont = Spr_plano.TypeComboBoxCurSel
        Spr_plano.Col = 23: Spr_plano.TypeComboBoxCurSel = CInt(Tipo_cont)
                            Tipo_cont = Spr_plano.Text
        Spr_plano.Col = 19: Sello = Spr_plano.Text
        Spr_plano.Col = 20: peso = CDbl(Spr_plano.Text)
        Spr_plano.Col = 21: volumen = CDbl(Spr_plano.Text)
        Spr_plano.Col = 22: Imo = Spr_plano.Text
  
        
        
                
        If pto_desc <> "" And pto_final <> "" And bl <> "" Then ' solo si todo esta completo se crea el bl
            ' se inserta el manifiesto si no existe
            If Not Existe_Manifiesto("E", Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, pto_emb, pto_desc) Then
                sql = "Insert into docu_manifiesto_exp "
                'If Gstr_Tipo = "I" Then sql = sql & "docu_manifiesto " Else sql = sql & "docu_manifiesto_exp "
                sql = sql & " (mani_linea,    mani_buque,    mani_viaje, " & _
                        " mani_puer_emb, mani_puer_des, mani_capitan, " & _
                        " mani_puer_ini, mani_fecha_zarp_ini , mani_emp_prop, mani_emp_nave) " & _
                        " Values('" & _
                        Cmb_linea.BoundText & "','" & _
                        Cmb_buque.BoundText & "','" & _
                        Cmb_viaje.BoundText & "','" & _
                        pto_emb & "','" & _
                        pto_desc & "', '" & _
                        "Desconocido" & "', '" & _
                        Gstr_param.pto_descarga & "','" & _
                        fecha_zarpe & "','" & _
                        Gstr_param.Linea_Nav & "','" & _
                        Gstr_param.naviera & "')"
                Gcon_docu.Execute sql
            End If 'manifiesto
            
            If Not Existe_BL("E", Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, pto_emb, pto_desc, Gstr_param.naviera, bl) Then
                 sql = "Insert into docu_blcab_exp "
                 'If Gstr_Tipo = "I" Then sql = sql & " docu_Blcab " Else sql = sql & " docu_blcab_exp "
                 sql = sql & "(blcab_linea,         blcab_buque,            blcab_viaje,        blcab_puer_emb," & _
                             " blcab_puer_des,      blcab_emp_resp,         blcab_cono_emb,     blcab_emp_anav," & _
                             " blcab_tipo_doc,      blcab_tipo_ide_cons,    blcab_ide_cons,     blcab_consig,  " & _
                             " blcab_direc_consig,  blcab_tipo_ide_notif,   blcab_ide_notif,    blcab_notif,   " & _
                             " blcab_direc_notif,   blcab_embarcador,       blcab_direc_embarc, blcab_puer_origen," & _
                             " blcab_almac_temp,    blcab_regimen_adu,      blcab_pto_dest_final,blcab_metod_pago," & _
                             " blcab_moneda,        blcab_fecha_emb,        blcab_user_ing,     blcab_fecha_ing,  " & _
                             " blcab_hora_ing)"
                 sql = sql & " Values('" & _
                     Cmb_linea.BoundText & "','" & _
                     Cmb_buque.BoundText & "','" & _
                     Cmb_viaje.BoundText & "','" & _
                     pto_emb & "','" & _
                     pto_desc & "','" & _
                     Gstr_param.naviera & "','" & _
                     bl & "','" & _
                     Gstr_param.Linea_Nav & "','" & _
                     Gstr_param.tipo_documento & "', '" & _
                     Gstr_param.tipo_id & "','" & _
                     Gstr_param_plano.id_consignatario & "','"
                 'If Gstr_Tipo = "I" Then ' si es exportacion los campos se setean a como debe ser
                 '    sql = sql & _
                     Trim(chk_str(Gstr_param_plano.consignatario)) & "','" & _
                     Trim(chk_str(Gstr_param_plano.Direc_consignatario)) & "','"
                 'Else         ' si es exportacion los campos se invierten
                     sql = sql & _
                     Trim(chk_str(embarcador)) & "','" & _
                     Trim(chk_str(Direc_Embarcador)) & "','"
                 'End If
                 sql = sql & _
                     Gstr_param.tipo_id & "','" & _
                     Gstr_param_plano.id_notificador & "','" & _
                     Trim(chk_str(notificador)) & "','" & _
                     Trim(chk_str(Direc_notificador)) & "','"
                 'If Gstr_Tipo = "E" Then
                     sql = sql & _
                     Trim(chk_str(consignatario)) & "','" & _
                     Trim(chk_str(Direc_consignatario)) & "','"
                 'Else
                 '    sql = sql & _
                     Trim(chk_str(Gstr_param_plano.embarcador)) & "','" & _
                     Trim(chk_str(Gstr_param_plano.DIREC_EMBARCADOR)) & "','"
                 'End If
                 
                 'If Gstr_Tipo = "E" Then
                 sql = sql & pto_emb & "','"
                 'Else
                 'sql = sql & pto_origen & "','" ' este valor debe ser establecido posteriormente debe ser igual al puerto de embarque
                 'End If
                 sql = sql & Gstr_param.almacen_temporal & "','"
                 'If Gstr_Tipo = "E" Then
                 sql = sql & Gstr_param.regimen_adua_expo & "','"
                 'Else
                 'sql = sql & Gstr_param.regimen_adua & "','"
                 'End If
                 sql = sql & pto_final & "','" & _
                     Gstr_param.metodo_pago & "','" & _
                     Gstr_param.MONEDA & "','" & _
                     Format(Now, GSTR_FORMATO_FECHA) & "', '" & _
                     "PLANO" & "','" & _
                     Format(Now, GSTR_FORMATO_FECHA) & "', '" & _
                     Format(Time, "hh:mm:ss") & "')"
                Gcon_docu.Execute sql
                No_bls = No_bls + 1
            End If 'bl
            '*****************************************************************************************************************************************
            If Not Existe_Contenedor("E", Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, pto_emb, pto_desc, cont) Then        ' Gstr_param.pto_descarga contiene la informacion del puerto local (ECGYE) que para importacion es descarga y para expo es embarque
                sql = "Insert Into docu_cont_exp "
                'If Gstr_Tipo = "I" Then sql = sql & " docu_cont " Else sql = sql & " docu_cont_exp "
                sql = sql & "(cont_linea,           cont_buque,         cont_viaje,             cont_puer_emb,      cont_puer_des,       " & _
                            " cont_contenedor,      cont_tam_equip,     cont_condicion,         cont_tara,          cont_tipo_equip,     " & _
                            " cont_proved_equip,    cont_estad_equip,   cont_indic_llen_vac,    cont_cond_contrato, cont_reg_serv_transp," & _
                            " cont_natu_carg,       cont_peso,          cont_sello1,            cont_tipo_bien,     cont_sello2,         " & _
                            " cont_sello3,          cont_sello4,        cont_temp_dur_almac,    cont_temp_dur_trans,cont_temp_dur_manip, " & _
                            " cont_temp_min_almac,  cont_temp_max_almac,cont_temp_min_trans,    cont_temp_max_trans,cont_temp_min_manip, " & _
                            " cont_temp_max_manip)"
                sql = sql & " Values('" & Cmb_linea.BoundText & _
                                    "','" & Cmb_buque.BoundText & _
                                    "','" & Cmb_viaje.BoundText & _
                                    "','" & pto_emb & _
                                    "','" & pto_desc & _
                                    "','" & cont & _
                                    "','" & Tipo_cont & _
                                    "','" & Gstr_param_plano.cond_cont & _
                                    "','" & peso & _
                                    "','" & Gstr_param.tipo_equipo & _
                                    "','" & Gstr_param.proveed_eq & _
                                    "','" & Gstr_param.estado_equip_expo & _
                                    "','" & Gstr_param_plano.Indic_vacio & _
                                    "','" & Gstr_param.cond_contrato & _
                                    "','" & Gstr_param.requer_serv & _
                                    "','" & Gstr_param.naturaleza_carga & _
                                    "','" & peso & _
                                    "','" & Sello & "', '', '','','',0,0,0,0,0,0,0,0,0)"
                Gcon_docu.Execute sql
                no_cont = no_cont + 1
            End If ' contenedor
            '*****************************************************************************************************************************************
            If Not Existe_BL_det("E", Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, pto_emb, pto_desc, Gstr_param.naviera, bl, cont) Then
                SECUENCIA = Max_Secuencia_BL_det("E", Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, pto_emb, pto_desc, Gstr_param.naviera, bl)
                'If Gstr_Tipo = "I" Then
                '    sql = "Insert into docu_bldet "
                'ElseIf Gstr_Tipo = "E" Then
                    sql = "Insert into docu_bldet_exp "
                'End If
                sql = sql & "    (bldet_linea,      bldet_buque,    bldet_viaje,    bldet_puer_emb,  " & _
                                " bldet_puer_des,   bldet_emp_resp, bldet_cono_emb, bldet_secuencia, " & _
                                " bldet_tipo_carga, bldet_peso,     bldet_bultos,   bldet_embalaje, " & _
                                " bldet_marcas_num, bldet_desc_carg,bldet_volumen,  bldet_contenedor," & _
                                " bldet_IMO)" & _
                                " Values('" & _
                                Cmb_linea.BoundText & "','" & _
                                Cmb_buque.BoundText & "','" & _
                                Cmb_viaje.BoundText & "','" & _
                                pto_emb & "','" & _
                                pto_desc & "','" & _
                                Gstr_param.naviera & "','" & _
                                bl & "','" & _
                                SECUENCIA & "','" & _
                                Gstr_param_plano.Tipo_carga_cont_vacio & "',"
                sql = sql & CDbl(peso) & ", '" & _
                                bultos & "', '" & _
                                embalaje & "','" & _
                                chk_str(marca_num) & "','" & _
                                chk_str(desc_carga) & "', " & _
                                CDbl(volumen) & ",'" & _
                                cont & "','" & _
                                Imo & "')"
                Gcon_docu.Execute sql
                No_bl_det = No_bl_det + 1
                Call actualiza_totales_bl_expo(Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, pto_emb, pto_desc, Gstr_param.naviera, bl)
                Call actualiza_totales_manifiesto_expo(Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, pto_emb, pto_desc, Gstr_param.naviera)
                
            End If 'bl_det
        End If
    End If
Next i
MsgBox "Proceso de creacion de Bls finalizado con éxito" & Chr(13) & " Se crearon " & No_bls & " Bls, " & no_cont & " Contenedores, " & No_bl_det & " Detalles de Bls.", vbInformation
Me.MousePointer = 0
Exit Sub
errores_bl:
    MsgBox "Ocurrio el Error " & Err.Description, vbInformation
    Me.MousePointer = 0
    Exit Sub
    Resume 0
End Sub
