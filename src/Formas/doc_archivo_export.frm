VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form doc_archivo_Expo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación Archivo Export"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "doc_archivo_export.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6210
   Begin VB.Frame Frame1 
      Height          =   4185
      Left            =   45
      TabIndex        =   8
      Top             =   0
      Width           =   6030
      Begin VB.CheckBox Chk_todos 
         Caption         =   "Todos BLs"
         Height          =   255
         Left            =   225
         TabIndex        =   24
         Top             =   2565
         Width           =   1410
      End
      Begin VB.ListBox Lst_bl 
         Height          =   1635
         Left            =   1785
         Style           =   1  'Checkbox
         TabIndex        =   22
         Top             =   2190
         Width           =   3930
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   90
         TabIndex        =   19
         Top             =   1560
         Width           =   1500
         Begin MSDataListLib.DataCombo Cmb_pto_descarga 
            Height          =   315
            Left            =   135
            TabIndex        =   20
            Top             =   210
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "3"
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_pto_descarga 
            Height          =   330
            Left            =   480
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
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "Puerto Descarga"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   105
            TabIndex        =   21
            Top             =   -15
            Width           =   1200
         End
      End
      Begin VB.Frame Fra_Date 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   585
         Left            =   195
         TabIndex        =   18
         Top             =   5145
         Width           =   2025
      End
      Begin VB.Frame Fra_pto_descarga 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   1665
         TabIndex        =   15
         Top             =   1530
         Width           =   1590
         Begin VB.TextBox Txt_pto_emb 
            Height          =   285
            Left            =   120
            MaxLength       =   5
            TabIndex        =   3
            Top             =   270
            Width           =   975
         End
         Begin VB.CommandButton Cmd_hlp_pto_emb 
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
            Left            =   1140
            Picture         =   "doc_archivo_export.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar"
            Top             =   255
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Puerto Embarque"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   15
            Width           =   1230
         End
      End
      Begin VB.Frame Fra_viaje 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   3255
         TabIndex        =   13
         Top             =   840
         Width           =   2145
         Begin MSDataListLib.DataCombo Cmb_viaje 
            Height          =   315
            Left            =   120
            TabIndex        =   2
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
            Left            =   1050
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
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   135
            TabIndex        =   14
            Top             =   0
            Width           =   345
         End
      End
      Begin VB.Frame Fra_buque 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   135
         TabIndex        =   11
         Top             =   840
         Width           =   3090
         Begin MSDataListLib.DataCombo Cmb_buque 
            Height          =   315
            Left            =   120
            TabIndex        =   1
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
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   0
            Width           =   465
         End
      End
      Begin VB.Frame Fra_linea 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   120
         TabIndex        =   9
         Top             =   165
         Width           =   2865
         Begin MSDataListLib.DataCombo Cmb_linea 
            Height          =   315
            Left            =   135
            TabIndex        =   0
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_linea 
            Height          =   330
            Left            =   1665
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
         Begin VB.Label Label5 
            Caption         =   "Linea"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   165
            TabIndex        =   10
            Top             =   0
            Width           =   825
         End
      End
      Begin MSDataListLib.DataCombo Cmb_oper_carga 
         Height          =   315
         Left            =   3375
         TabIndex        =   5
         Top             =   1755
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc data_oper_carga 
         Height          =   330
         Left            =   4800
         Top             =   1755
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bill of Lading"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   225
         TabIndex        =   23
         Top             =   2235
         Width           =   900
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Agen. Oper. de Carg"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3405
         TabIndex        =   17
         Top             =   1515
         Width           =   1455
      End
   End
   Begin Threed.SSCommand cmd_Proceso 
      Height          =   375
      Left            =   3705
      TabIndex        =   6
      Top             =   4275
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Procesar"
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
   End
   Begin Threed.SSCommand cmd_salir 
      Height          =   375
      Left            =   4905
      TabIndex        =   7
      Top             =   4290
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Salir"
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
   End
   Begin MSMAPI.MAPIMessages mapMess 
      Left            =   660
      Top             =   4245
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession mapSess 
      Left            =   0
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "doc_archivo_Expo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bNewSession     As Boolean ' utilizada para saber si se dio logon o no a la conexion de correo de outlook
Dim mail_to         As String
Dim mail_subject    As String
Dim mail_text       As String

Private Sub Cmb_buque_Change()
   Dim sql As String
   sql = "Select * from docu_buque_viaje " & _
         " Where buqv_linea ='" & Cmb_linea.BoundText & "'" & _
         "   And buqv_buque ='" & Cmb_buque.BoundText & "'" & _
         "   And buqv_status = 'A'"
   
   carga_datacombo data_viaje, Cmb_viaje, sql, "buqv_viaje", "buqv_viaje"
End Sub

Private Sub cmb_linea_Change()
carga_datacombo data_buque, Cmb_buque, "Select * from docu_buque Where buqu_linea = '" & Cmb_linea.BoundText & "' and buqu_status ='A'", "buqu_codigo", "buqu_descripcion"
carga_datacombo data_oper_carga, Cmb_oper_carga, "Select * from docu_emp_resp Where emp_resp_linea = '" & Cmb_linea.BoundText & "' and emp_resp_status ='A'", "emp_resp_codigo", "emp_resp_descripcion"
'carga_datacombo data_tipo_transaccion, Cmb_tipo_transaccion, "Select * from docu_tipo_transaccion WHERE tipo_tran_aplica ='S'", "tipo_tran_codigo", "tipo_tran_descripcion"
'carga_datacombo data_tipo_envio, Cmb_tipo_envio, "Select * from docu_tipo_envio Where tipo_env_linea ='" & Cmb_linea.BoundText & "'", "tipo_env_codigo", "tipo_env_descripcion"

carga_parametros Cmb_linea.BoundText

Txt_pto_emb.Text = Gstr_param.pto_descarga
Cmb_oper_carga.BoundText = Gstr_param.naviera

End Sub

Private Sub Cmb_pto_descarga_Change()
Dim reg As New ADODB.Recordset

'If Opt_bl.Value Then
    sql = "Select blcab_cono_emb " & _
               " from docu_blcab_exp " & _
               " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'" & _
               "   and blcab_puer_des    ='" & Cmb_pto_descarga.BoundText & "'"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Lst_bl.Clear
    Do While Not reg.EOF
        Lst_bl.AddItem reg("blcab_cono_emb")
        reg.MoveNext
    Loop
'End If

End Sub

Private Sub Cmb_viaje_Change()
   Dim sql As String
   sql = "Select * from docu_manifiesto_exp " & _
         " Where mani_linea ='" & Cmb_linea.BoundText & "'" & _
         "   And mani_buque ='" & Cmb_buque.BoundText & "'" & _
         "   And mani_viaje ='" & Cmb_viaje.BoundText & "'"
         '"   And buqv_status = 'A'"
   
   carga_datacombo data_pto_descarga, Cmb_pto_descarga, sql, "mani_puer_des", "mani_puer_des"

End Sub

Private Sub Cmd_hlp_pto_emb_Click()
   Txt_pto_emb.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Embarque")
End Sub

Private Sub cmd_Proceso_Click()

If Cmb_linea.MatchedWithList And Cmb_buque.MatchedWithList And Cmb_viaje.MatchedWithList And Cmb_pto_descarga.MatchedWithList And _
      Txt_pto_emb.Text <> "" And Cmb_oper_carga.MatchedWithList Then
'        If Not Valida_datos_CAE(Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, Cmb_pto_emb.BoundText) Then
'            If MsgBox("Existen errores en los datos, por lo cual se pueden presentar problemas con la CAE, Desea continuar de todas maneras?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
'        End If
        'Call actualiza_totales_manifiesto(Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, Cmb_pto_emb.BoundText, Txt_pto_descarga.Text, Cmb_oper_carga.BoundText) ' se pone para que siempre esten actualizados los totales antes del envio
        If Cmb_linea.BoundText = "KHS" Then
            Call exporta_datos(Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, Txt_pto_emb.Text, Cmb_pto_descarga.BoundText, Cmb_oper_carga.BoundText)
        Else
            MsgBox "Esta opción no esta disponible para otra Línea Naviera"
        End If
   Else
        MsgBox "No estan seleccionados todos los parámetros para la generación del archivo", vbInformation
   End If
End Sub

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub Chk_todos_Click()
Dim valor As Boolean
valor = Chk_todos.Value
For i = 0 To Lst_bl.ListCount - 1
    Lst_bl.Selected(i) = valor
Next i
End Sub

Private Sub Form_Load()
    carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_codigo= 'KHS' and linea_status ='A'", "linea_codigo", "linea_descripcion"
    If Not data_linea.Recordset.EOF Then data_linea.Recordset.MoveFirst
    'Carga_parametros_MAIL
    
End Sub


Public Sub exporta_datos(linea As String, buque As String, viaje As String, pto_emb As String, pto_descarga As String, oper_carga As String)
Dim sql         As String
Dim reg_mani    As New ADODB.Recordset
Dim reg_bls     As New ADODB.Recordset
Dim reg_blcab   As New ADODB.Recordset
Dim reg_bldet   As New ADODB.Recordset
Dim reg_blflete As New ADODB.Recordset
Dim reg_cont    As New ADODB.Recordset
Dim reg_cont_dt As New ADODB.Recordset
Dim REG_viaje   As New ADODB.Recordset
Dim reg_dtx_deta As New ADODB.Recordset
Dim emp_anav    As String
Dim sep         As String
Dim del         As String
Dim x
Dim cadena      As String '* 72
Dim No_bls      As Integer
Dim descrip_marcas As String
Dim descrip_carga As String
Dim viaje_trunc   As String

Dim flete_marcas        As String
Dim cont_sello          As String
Dim cont_size           As String
Dim cont_loading_serv   As String
Dim cont_discharge_serv As String
Dim cont_type           As String
Dim cont_sign_from      As String
Dim cont_temp_from      As String
Dim cont_sign_to        As String
Dim cont_temp_to        As String
Dim cont_peso_cy        As Double
Dim cont_vol_cy         As Double
Dim cont_peso_cfs       As Double
Dim cont_vol_cfs        As Double

On Error GoTo Errores

carga_parametros linea
Gstr_Linea = "KHS"

If Me.Lst_bl.SelCount = 0 Then
        If MsgBox("No se ha seleccionado ningun bl, desea generar todo el Puerto.", vbYesNo + vbQuestion) = vbYes Then
           Chk_todos.Value = 1
        Else
            Exit Sub
        End If
'   Else
'        MsgBox "No se ha seleccionado ningun bl", vbQuestion
'        Exit Sub
   End If


    bls = ""
    For i = 0 To Lst_bl.ListCount - 1
        If Lst_bl.Selected(i) Then
           Lst_bl.ListIndex = i
           bls = bls & "'" & Lst_bl.Text & "', "
        End If
    Next i
    bls = Mid(bls, 1, Len(Trim(bls)) - 1) ' se quita la ultima coma

     x = FreeFile
     If Dir(Gstr_Ruta_files & "\" & buque, vbDirectory) = "" Then MkDir (Gstr_Ruta_files & "\" & buque)
     If Dir(Gstr_Ruta_files & "\" & buque & "\" & viaje, vbDirectory) = "" Then MkDir (Gstr_Ruta_files & "\" & buque & "\" & viaje)
     If Dir(Gstr_Ruta_files & "\" & buque & "\" & viaje & "\" & pto_descarga, vbDirectory) = "" Then MkDir (Gstr_Ruta_files & "\" & buque & "\" & viaje & "\" & pto_descarga)
     
     If Dir(Gstr_Ruta_files & "TMP\", vbArchive) <> "" Then Kill Gstr_Ruta_files & "TMP\*.*"
    
     sql = "Select * from " & _
                    " docu_buque_viaje, docu_buque " & _
                    " Where buqv_linea    ='" & linea & "'" & _
                    "   and buqv_buque    ='" & buque & "'" & _
                    "   and buqv_viaje    ='" & viaje & "'" & _
                    "   and buqv_linea    = buqu_linea" & _
                    "   and buqv_buque    = buqu_codigo"
     REG_viaje.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
     
     sql = "Select * from " & _
                    " docu_manifiesto_exp " & _
                    " Where mani_linea    ='" & linea & "'" & _
                    "   and mani_buque    ='" & buque & "'" & _
                    "   and mani_viaje    ='" & viaje & "'" & _
                    "   and mani_puer_emb ='" & pto_emb & "'" & _
                    "   and mani_puer_des ='" & pto_descarga & "'"
     
     reg_mani.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
     
     
     'Open Gstr_Ruta_files & "TMP\export.txt" For Output As #x
     Open Gstr_Ruta_files & buque & "\" & viaje & "\" & Trim(pto_descarga) & "\EXPORT.TXT" For Output As #x
     '************************* V0 ********************************************
      cadena = "V0"
      sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
      reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
      Do While Not reg_dtx_deta.EOF
        Select Case Trim(reg_dtx_deta("campo"))
            Case "1": cadena = cadena & Completa_Spacios(viaje, reg_dtx_deta("longitud"))
            Case "2": cadena = cadena & Devuelve_Formato("V0", "2", reg_mani("mani_fecha_zarp_ini")) ' reg_mani("mani_fecha_salida") & Space(reg_dtx_deta("longitud") - Len(reg_mani("mani_fecha_salida")))
            Case "3": cadena = cadena & Devuelve_Formato("V0", "3", "1")                             ' este valor "EXCHANGE RATE" debe ser un parametro o estar en alguna pantalla
        End Select
        reg_dtx_deta.MoveNext
      Loop
      Print #x, cadena
      reg_dtx_deta.Close
            
      sql = "Select * from " & _
                    " docu_blcab_exp " & _
                    " Where blcab_linea      ='" & linea & "'" & _
                    "   and blcab_buque    ='" & buque & "'" & _
                    "   and blcab_viaje    ='" & viaje & "'" & _
                    "   and blcab_puer_emb ='" & pto_emb & "'" & _
                    "   and blcab_puer_des ='" & pto_descarga & "'" & _
                    "   and blcab_emp_resp ='" & oper_carga & "'" & _
                    "   and blcab_cono_emb in (" & bls & ")" & _
            "Order by  blcab_cono_emb"
     reg_blcab.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
     No_bls = 0
     Do While Not reg_blcab.EOF
        No_bls = No_bls + 1
        '************************* A0 ********************************************
        cadena = "A0"
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        Do While Not reg_dtx_deta.EOF
            Select Case Trim(reg_dtx_deta("campo"))
                Case "1": tama_bl = Len(Trim(reg_blcab("blcab_cono_emb")))
                        If tama_bl > reg_dtx_deta("longitud") Then tama_bl = reg_dtx_deta("longitud")
                        cadena = cadena & reg_blcab("blcab_cono_emb") & Space(reg_dtx_deta("longitud") - tama_bl)
                ' no se usa aparentemente 22 enero 2003
                Case "2": cadena = cadena & Completa_Spacios("", reg_dtx_deta("longitud"))                          'service contract number
            End Select
            reg_dtx_deta.MoveNext
        Loop
        Print #x, cadena
        reg_dtx_deta.Close
        '************************* A1 ********************************************
        cadena = "A1"
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        Do While Not reg_dtx_deta.EOF
            Select Case Trim(reg_dtx_deta("campo"))
                Case "1": cadena = cadena & Completa_Spacios(reg_blcab("blcab_consig"), reg_dtx_deta("longitud"))
            End Select
            reg_dtx_deta.MoveNext
        Loop
        Print #x, cadena
        reg_dtx_deta.Close
        '************************* A2 ********************************************
        cadena = "A2"
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        If Not reg_dtx_deta.EOF Then
            If Len(Trim(reg_blcab("blcab_direc_consig"))) > reg_dtx_deta("longitud") Then
                temp = Trim(reg_blcab("blcab_direc_consig"))
                i = 1:
                Do While temp <> "" And i <= 5
                    cadena = "A2" & Mid(temp, 1, reg_dtx_deta("longitud"))
                    temp = Mid(temp, reg_dtx_deta("longitud") + 1)
                    Print #x, cadena
                    i = i + 1
                Loop
            Else
                cadena = cadena & Completa_Spacios(reg_blcab("blcab_direc_consig"), reg_dtx_deta("longitud"))
                Print #x, cadena
            End If
        End If
        reg_dtx_deta.Close
        '************************* A3 ********************************************
        cadena = "A3"
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        Do While Not reg_dtx_deta.EOF
            Select Case Trim(reg_dtx_deta("campo"))
                Case "1": cadena = cadena & Completa_Spacios(reg_blcab("blcab_embarcador"), reg_dtx_deta("longitud"))
            End Select
            reg_dtx_deta.MoveNext
        Loop
        Print #x, cadena
        
        reg_dtx_deta.Close
        '************************* A4 ********************************************
        cadena = "A4"
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        If Not reg_dtx_deta.EOF Then
            If Len(Trim(reg_blcab("blcab_direc_embarc"))) > reg_dtx_deta("longitud") Then
                temp = Trim(reg_blcab("blcab_direc_embarc"))
                i = 1
                Do While temp <> "" And i <= 3
                    cadena = "A4" & Mid(temp, 1, reg_dtx_deta("longitud"))
                    temp = Mid(temp, reg_dtx_deta("longitud") + 1)
                    Print #x, cadena:       i = i + 1
                Loop
            Else
                cadena = cadena & Completa_Spacios(reg_blcab("blcab_direc_embarc"), reg_dtx_deta("longitud"))
                Print #x, cadena
            End If
        End If
        reg_dtx_deta.Close
        '************************* A5 ********************************************
        cadena = "A5"
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        Do While Not reg_dtx_deta.EOF
            Select Case Trim(reg_dtx_deta("campo"))
                Case "1": cadena = cadena & Completa_Spacios(reg_blcab("blcab_notif"), reg_dtx_deta("longitud"))
            End Select
            reg_dtx_deta.MoveNext
        Loop
        Print #x, cadena
        reg_dtx_deta.Close
        '************************* A6 ********************************************
        cadena = "A6"
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        If Not reg_dtx_deta.EOF Then
            If Len(Trim(reg_blcab("blcab_direc_notif"))) > reg_dtx_deta("longitud") Then
                temp = Trim(reg_blcab("blcab_direc_notif"))
                i = 1:
                Do While temp <> "" And i <= 5 ' maximo 5 lineas
                    cadena = "A6" & Mid(temp, 1, reg_dtx_deta("longitud"))
                    temp = Mid(temp, reg_dtx_deta("longitud") + 1)
                    i = i + 1
                    Print #x, cadena
                Loop
            Else
                cadena = cadena & Completa_Spacios(reg_blcab("blcab_direc_notif"), reg_dtx_deta("longitud"))
                Print #x, cadena
            End If
        End If
        reg_dtx_deta.Close
        '************************* A9 ********************************************
        '(no se usa 22 enero 2003)
        '************************* S1 ********************************************
        'totales por contenedores CY y CFS
        sql = "Select convert(char(1),cont_condicion) , sum(bldet_peso), sum(bldet_volumen) from " & _
                        " docu_bldet_exp, docu_cont_exp  " & _
                        " Where bldet_linea      ='" & linea & "'" & _
                        "   and bldet_buque    ='" & buque & "'" & _
                        "   and bldet_viaje    ='" & viaje & "'" & _
                        "   and bldet_puer_emb ='" & pto_emb & "'" & _
                        "   and bldet_puer_des ='" & pto_descarga & "'" & _
                        "   and bldet_emp_resp ='" & oper_carga & "'" & _
                        "   and bldet_cono_emb ='" & reg_blcab("blcab_cono_emb") & "'" & _
                        "   and bldet_linea    = cont_linea " & _
                        "   and bldet_buque    = cont_buque " & _
                        "   and bldet_viaje    = cont_viaje " & _
                        "   and bldet_puer_emb = cont_puer_emb " & _
                        "   and bldet_puer_des = cont_puer_des " & _
                        "   and bldet_contenedor = cont_contenedor" & _
                " Group by  convert(char(1),cont_condicion)"
        reg_cont.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        
        cadena = "S1"
        cont_peso_cy = 0:         cont_vol_cy = 0
        cont_peso_cfs = 0:        cont_vol_cfs = 0
        Do While Not reg_cont.EOF
            If reg_cont(0) = "1" Then
                cont_peso_cy = reg_cont(1)
                cont_vol_cy = reg_cont(2)
            ElseIf reg_cont(0) = "2" Then
                cont_peso_cfs = reg_cont(1)
                cont_vol_cfs = reg_cont(2)
            End If
            reg_cont.MoveNext
        Loop
        cadena = cadena & Devuelve_Formato("S1", "1", cont_vol_cy)
        cadena = cadena & Devuelve_Formato("S1", "2", cont_vol_cfs)
        cadena = cadena & Devuelve_Formato("S1", "3", cont_peso_cy)
        cadena = cadena & Devuelve_Formato("S1", "4", cont_peso_cfs)
        Print #x, cadena
        reg_cont.Close
        '************************* FIN S1 ****************************************
        
        '************************* F1 ********************************************
        cadena = "F1"
        sql = "Select * from " & _
                        " docu_blflete_exp " & _
                        " Where flete_linea      ='" & linea & "'" & _
                        "   and flete_buque    ='" & buque & "'" & _
                        "   and flete_viaje    ='" & viaje & "'" & _
                        "   and flete_puer_emb ='" & pto_emb & "'" & _
                        "   and flete_puer_des ='" & pto_descarga & "'" & _
                        "   and flete_emp_resp ='" & oper_carga & "'" & _
                        "   and flete_cono_emb ='" & reg_blcab("blcab_cono_emb") & "'" & _
                " Order by  flete_cono_emb, flete_secuencia "
        reg_blflete.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        
        Do While Not reg_blflete.EOF
            cadena = "F1"
            reg_dtx_deta.MoveFirst
            If Not reg_dtx_deta.EOF Then flete_marcas = IIf(Not IsNull(reg_blflete("flete_marcas")), reg_blflete("flete_marcas"), "")
            Do While Not reg_dtx_deta.EOF
                Select Case Trim(reg_dtx_deta("campo"))
                    Case "1": cadena = cadena & Completa_Spacios(reg_blflete("flete_cargo"), reg_dtx_deta("longitud"))  ' codigo del cargo
                    Case "2": cadena = cadena & Devuelve_Formato("F1", "2", reg_blflete("flete_valor_unit"))            ' valor unitario
                    Case "3": cadena = cadena & Completa_Spacios(reg_blflete("flete_por"), reg_dtx_deta("longitud"))    ' codigo del nivel de cargo (por)
                    Case "4": cadena = cadena & Completa_Spacios(reg_blflete("flete_moneda"), reg_dtx_deta("longitud")) ' codigo de la moneda
                    Case "5": cadena = cadena & Devuelve_Formato("F1", "5", reg_blflete("flete_valor"))                 ' valor total (amount)
                    Case "6": cadena = cadena & Trim(reg_blflete("flete_debtor"))                                       ' debtor prepaid/colect
                    Case "7": cadena = cadena & Devuelve_Formato("F1", "7", reg_blflete("flete_cantidad"))              ' cantidad
                    Case "8":
                            If Not IsNull(reg_blflete("flete_lugar")) Then
                                cadena = cadena & Completa_Spacios(reg_blflete("flete_lugar"), reg_dtx_deta("longitud")) ' lugar de pago
                            Else
                                cadena = cadena & Space(reg_dtx_deta("longitud"))  ' lugar de pago
                            End If
                End Select
                reg_dtx_deta.MoveNext
            Loop
            Print #x, cadena
            reg_blflete.MoveNext
        Loop
        reg_dtx_deta.Close
        '************************* F2 ********************************************
        ' es una observacion del F1
        cadena = "F2"
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        
        'reg_blflete.MoveFirst
        cadena = cadena & Completa_Spacios(flete_marcas, reg_dtx_deta("longitud"))   ' marcas del flete
        Print #x, cadena
        reg_dtx_deta.Close
        reg_blflete.Close
        '*************************************************************************
                
        ' **************  SELECT A LOS DETALLES DEL ACTUAL BL   ******************
         sql = "Select * from " & _
                        " docu_bldet_exp  " & _
                        " Where bldet_linea      ='" & linea & "'" & _
                        "   and bldet_buque    ='" & buque & "'" & _
                        "   and bldet_viaje    ='" & viaje & "'" & _
                        "   and bldet_puer_emb ='" & pto_emb & "'" & _
                        "   and bldet_puer_des ='" & pto_descarga & "'" & _
                        "   and bldet_emp_resp ='" & oper_carga & "'" & _
                        "   and bldet_cono_emb ='" & reg_blcab("blcab_cono_emb") & "'" & _
                " Order by  bldet_cono_emb, bldet_secuencia "
         reg_bldet.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
         If Not reg_bldet.EOF Then
            '************************* B1 *****************************************************************
            '**** POR LA DESCRIPCION DE LA CARGA SE PUEDEN REPETIR LAS LINEAS CON SEQUENCIA DIFERENTE *****
            cadena = "B1"
            i = 1 ' SECUENCIA
            sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
            reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
            Do While Not reg_dtx_deta.EOF
                Select Case Trim(reg_dtx_deta("campo"))
                    Case "1": cadena = cadena & Devuelve_Formato("B1", "1", i)
                    Case "2": cadena = cadena & Devuelve_Formato("B1", "2", reg_blcab("blcab_bultos_tot"))
                    Case "3": cadena = cadena & Completa_Spacios(reg_bldet("bldet_embalaje"), reg_dtx_deta("longitud"))
                    Case "4": cadena = cadena & Completa_Spacios(reg_bldet("bldet_desc_carg"), reg_dtx_deta("longitud"))
                End Select
                reg_dtx_deta.MoveNext
            Loop
            Print #x, cadena
            ' PARA GRABAR LA INFORMACION DE LA DESCRIPCION EN MAS DE 1 LINEA
            reg_dtx_deta.MoveFirst: temp2 = ""
            Do While Not reg_dtx_deta.EOF
                Select Case Trim(reg_dtx_deta("campo"))
                    Case "2": temp2 = temp2 & Devuelve_Formato("B1", "2", "0")
                    Case "3": temp2 = temp2 & Space(reg_dtx_deta("longitud"))
                    Case "4": Exit Do
                End Select
                reg_dtx_deta.MoveNext
            Loop
            If Len(Trim(reg_bldet("bldet_desc_carg"))) > reg_dtx_deta("longitud") Then
                temp = Mid(reg_bldet("bldet_desc_carg"), reg_dtx_deta("longitud") + 1)
                i = 1
                Do While temp <> ""
                    cadena = "B1" & Devuelve_Formato("B1", "1", i) & temp2 & Mid(temp, 1, reg_dtx_deta("longitud"))
                    temp = Mid(temp, reg_dtx_deta("longitud") + 1)
                    Print #x, cadena
                    i = i + 1
                Loop
            End If
            reg_dtx_deta.Close
            '************************* FIN B1 ********************************************
            '************************* B2 *****************************************************************
            '**** POR LA MARCAS Y NUMEROS SE PUEDEN REPETIR LAS LINEAS CON SEQUENCIA DIFERENTE *****
            cadena = "B2"
            sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
            reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
            If Not reg_dtx_deta.EOF Then
                If Len(Trim(reg_bldet("bldet_marcas_num"))) > reg_dtx_deta("longitud") Then
                    temp = Trim(reg_bldet("bldet_marcas_num"))
                    i = 1:
                    reg_dtx_deta.MoveNext 'se lo mueve al siguiente registro que es el las marcas en si (ya que al abrir el cursosr esta el el registro de la secuencia)
                    Do While temp <> "" And i <= 5 ' maximo 5 lineas
                        cadena = "B2" & Devuelve_Formato("B2", "1", i) & Mid(temp, 1, reg_dtx_deta("longitud"))
                        temp = Mid(temp, reg_dtx_deta("longitud") + 1)
                        i = i + 1
                        Print #x, cadena
                    Loop
                Else
                    cadena = cadena & Devuelve_Formato("B2", "1", i) & Completa_Spacios(reg_bldet("bldet_marcas_num"), reg_dtx_deta("longitud"))
                    Print #x, cadena
                End If
            End If
            reg_dtx_deta.Close
            '************************* FIN B2 ********************************************
        End If
        
        
        ' HASTA AQUI SOLO SE ACCESO A LOS DETALLES A BUSCAR CIERTOS CAMPOS PERO DE AQUI EN ADELANTE SE GRABAN LOS DETALLES
        Do While Not reg_bldet.EOF
           '************************* B3 ********************************************
            cadena = "B3"
            If reg_bldet("bldet_contenedor") <> "" Then
                sql = "Select * from " & _
                        " docu_cont_exp " & _
                        " Where cont_linea      ='" & linea & "'" & _
                        "   and cont_buque    ='" & buque & "'" & _
                        "   and cont_viaje    ='" & viaje & "'" & _
                        "   and cont_puer_emb ='" & pto_emb & "'" & _
                        "   and cont_puer_des ='" & pto_descarga & "'" & _
                        "   and cont_contenedor='" & reg_bldet("bldet_contenedor") & "'"
                reg_cont.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                If Not reg_cont.EOF Then
                    cont_sello = reg_cont("Cont_sello1")
                    cont_size = Mid(reg_cont("cont_tam_equip"), 1, 1)
                    cont_loading_serv = Mid(reg_cont("cont_condicion"), 1, 1)
                    cont_discharge_serv = Mid(reg_cont("cont_condicion"), 5, 1)
                    cont_type = Mid(reg_cont("cont_tam_equip"), 4, 1)
                    cont_sign_from = IIf(reg_cont("cont_temp_dur_almac") > 0, "+", "-")
                    cont_temp_from = reg_cont("cont_temp_dur_almac")
                    cont_sign_to = IIf(reg_cont("cont_temp_dur_trans") > 0, "+", "-")
                    cont_temp_to = reg_cont("cont_temp_dur_trans")
                Else
                    MsgBox "El contenedor " & reg_bldet("bldet_contenedor") & " No existe en los datos de contendores", vbInformation
                End If
            Else
                    cont_sello = ""
                    cont_size = ""
                    cont_loading_serv = ""
                    cont_discharge_serv = ""
                    cont_type = ""
                    cont_sign_from = ""
                    cont_temp_from = "0"
                    cont_sign_to = ""
                    cont_temp_to = "0"
            End If
            sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by desde"
            reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
            Do While Not reg_dtx_deta.EOF
                Select Case Trim(reg_dtx_deta("campo"))
                    Case "1": cadena = cadena & Completa_Spacios(reg_bldet("bldet_contenedor"), reg_dtx_deta("longitud"))   ' CONTENEDOR
                    Case "2": cadena = cadena & Completa_Spacios(cont_sello, reg_dtx_deta("longitud"))                      ' SELLO
                    Case "3": cadena = cadena & Completa_Spacios(cont_size, reg_dtx_deta("longitud"))                       ' SIZE DEL CONTENEDOR
                    Case "4": cadena = cadena & Completa_Spacios(cont_loading_serv, reg_dtx_deta("longitud"))               ' loading service
                    Case "5": cadena = cadena & Completa_Spacios(cont_discharge_serv, reg_dtx_deta("longitud"))             ' discharge service
                    Case "6": cadena = cadena & Devuelve_Formato("B3", "6", reg_bldet("bldet_bultos"))                      ' no. of packages
                    Case "10": cadena = cadena & "C"                                                                        ' COC/SOC PENDIENTE POR INCLUIR EN EL MANTENIMIENTO
                    Case "7": cadena = cadena & Completa_Spacios(cont_type, reg_dtx_deta("longitud"))                       ' CONTEINER TYPE
                    Case "8": cadena = cadena & Devuelve_Formato("B3", "8", reg_bldet("bldet_peso"))                        ' WEIGHT
                    Case "9": cadena = cadena & Devuelve_Formato("B3", "9", reg_bldet("bldet_volumen"))                     ' MEASURMENT
                    Case "11": 'no se usan los customs seal
                End Select
                reg_dtx_deta.MoveNext
            Loop
            Print #x, cadena
            reg_dtx_deta.Close
            '************************ FIN B3 ********************************************
            
            '************************* D1 ***********************************************
            If reg_bldet("bldet_DANGER_CARGO") <> "" And reg_bldet("bldet_UN_number") <> "" And _
               reg_bldet("bldet_imdg_page") <> "" Then
                cadena = "D1"
                sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by desde"
                reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                Do While Not reg_dtx_deta.EOF
                    Select Case Trim(reg_dtx_deta("campo"))
                        Case "1": cadena = cadena & Completa_Spacios(reg_bldet("bldet_DANGER_CARGO"), reg_dtx_deta("longitud")) ' peligrosidad KHS
                        Case "2": cadena = cadena & Completa_Spacios(reg_bldet("bldet_UN_number"), reg_dtx_deta("longitud"))    ' UN number KHS
                        Case "3": cadena = cadena & Completa_Spacios(reg_bldet("bldet_imdg_page"), reg_dtx_deta("longitud"))    ' imdg page KHS
                    End Select
                    reg_dtx_deta.MoveNext
                Loop
                Print #x, cadena
                reg_dtx_deta.Close
            End If
            '************************* R1 ***********************************************
            If Trim(cont_temp_from) <> "0" And Trim(cont_temp_to) <> "0" Then
                cadena = "R1"
                sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by desde"
                reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                Do While Not reg_dtx_deta.EOF
                    Select Case Trim(reg_dtx_deta("campo"))
                        Case "1": cadena = cadena & Mid(cont_sign_from, 1, reg_dtx_deta("longitud"))
                        Case "2": cadena = cadena & Devuelve_Formato("R1", "2", cont_temp_from)
                        Case "3": cadena = cadena & Mid(cont_sign_to, 1, reg_dtx_deta("longitud"))
                        Case "4": cadena = cadena & Devuelve_Formato("R1", "4", cont_temp_to)
                        Case "5": cadena = cadena & "C"
                    End Select
                    reg_dtx_deta.MoveNext
                Loop
                Print #x, cadena
                reg_dtx_deta.Close
            End If
            '****************************************************************************
            ' siguiente detalle
            reg_bldet.MoveNext
        Loop
        reg_bldet.Close
        
        reg_blcab.MoveNext
    Loop
    
    Print #x, "ZZ"
            
    MsgBox "Generación de archivo fue realizada con éxito. En el archivo " & Gstr_Ruta_files & buque & "\" & viaje & "\" & Trim(pto_descarga) & "\EXPORT.TXT", vbInformation
     Close #x

    Exit Sub
Errores:
     MsgBox "Ocurrio el error " & Err.Description
     Close
     Exit Sub
     Resume 0

End Sub

Public Sub envia_mail(Archivo As String)
On Error GoTo Error
    If Mail_LogOn = True Then
        Mail_ComposeMessage (Archivo)
    Else
        'Exit Sub
    End If
    'Mail_ComposeMessage (Archivo)
    'LogOff
    Exit Sub
Error:
    MsgBox "Ocurrio el error " & Err.Description
End Sub

Public Sub exporta_datos_back(linea As String, buque As String, viaje As String, pto_emb As String, pto_descarga As String, oper_carga As String)
Dim sql         As String
Dim reg_mani    As New ADODB.Recordset
Dim reg_bls     As New ADODB.Recordset
Dim reg_blcab   As New ADODB.Recordset
Dim reg_bldet   As New ADODB.Recordset
Dim reg_blflete As New ADODB.Recordset
Dim reg_cont    As New ADODB.Recordset
Dim reg_cont_dt As New ADODB.Recordset
Dim REG_viaje   As New ADODB.Recordset
Dim reg_dtx_deta As New ADODB.Recordset
Dim emp_anav    As String
Dim sep         As String
Dim del         As String
Dim x
Dim cadena      As String '* 72
Dim No_bls      As Integer
Dim descrip_marcas As String
Dim descrip_carga As String
Dim viaje_trunc   As String

Dim cont_sello          As String
Dim cont_size           As String
Dim cont_loading_serv   As String
Dim cont_discharge_serv As String
Dim cont_type           As String
Dim cont_sign_from      As String
Dim cont_temp_from      As String
Dim cont_sign_to        As String
Dim cont_temp_to        As String
Dim cont_peso_cy        As Double
Dim cont_vol_cy         As Double
Dim cont_peso_cfs       As Double
Dim cont_vol_cfs        As Double

On Error GoTo Errores

carga_parametros linea
Gstr_Linea = "KHS"

If Me.Lst_bl.SelCount = 0 Then
        If MsgBox("No se ha seleccionado ningun bl, desea generar todo el Puerto.", vbYesNo + vbQuestion) = vbYes Then
           Chk_todos.Value = 1
        Else
            Exit Sub
        End If
'   Else
'        MsgBox "No se ha seleccionado ningun bl", vbQuestion
'        Exit Sub
   End If


    bls = ""
    For i = 0 To Lst_bl.ListCount - 1
        If Lst_bl.Selected(i) Then
           Lst_bl.ListIndex = i
           bls = bls & "'" & Lst_bl.Text & "', "
        End If
    Next i
    bls = Mid(bls, 1, Len(Trim(bls)) - 1) ' se quita la ultima coma

     x = FreeFile
     If Dir(Gstr_Ruta_files & "\" & buque, vbDirectory) = "" Then MkDir (Gstr_Ruta_files & "\" & buque)
     If Dir(Gstr_Ruta_files & "\" & buque & "\" & viaje, vbDirectory) = "" Then MkDir (Gstr_Ruta_files & "\" & buque & "\" & viaje)
     If Dir(Gstr_Ruta_files & "\" & buque & "\" & viaje & "\" & pto_descarga, vbDirectory) = "" Then MkDir (Gstr_Ruta_files & "\" & buque & "\" & viaje & "\" & pto_descarga)
     
     If Dir(Gstr_Ruta_files & "TMP\", vbArchive) <> "" Then Kill Gstr_Ruta_files & "TMP\*.*"
    
     sql = "Select * from " & _
                    " docu_buque_viaje, docu_buque " & _
                    " Where buqv_linea    ='" & linea & "'" & _
                    "   and buqv_buque    ='" & buque & "'" & _
                    "   and buqv_viaje    ='" & viaje & "'" & _
                    "   and buqv_linea    = buqu_linea" & _
                    "   and buqv_buque    = buqu_codigo"
     REG_viaje.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
     
     sql = "Select * from " & _
                    " docu_manifiesto_exp " & _
                    " Where mani_linea    ='" & linea & "'" & _
                    "   and mani_buque    ='" & buque & "'" & _
                    "   and mani_viaje    ='" & viaje & "'" & _
                    "   and mani_puer_emb ='" & pto_emb & "'" & _
                    "   and mani_puer_des ='" & pto_descarga & "'"
     
     reg_mani.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
     
     
     'Open Gstr_Ruta_files & "TMP\export.txt" For Output As #x
     Open Gstr_Ruta_files & buque & "\" & viaje & "\" & pto_descarga & "\EXPORT.TXT" For Output As #x
     '************************* V0 ********************************************
      cadena = "V0"
      sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
      reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
      Do While Not reg_dtx_deta.EOF
        Select Case Trim(reg_dtx_deta("campo"))
            Case "1": cadena = cadena & viaje & Space(reg_dtx_deta("longitud") - Len(viaje))
            Case "2": cadena = cadena & Devuelve_Formato("V0", "2", reg_mani("mani_fecha_zarp_ini")) ' reg_mani("mani_fecha_salida") & Space(reg_dtx_deta("longitud") - Len(reg_mani("mani_fecha_salida")))
            Case "3": cadena = cadena & Devuelve_Formato("V0", "3", "1")                             ' este valor "EXCHANGE RATE" debe ser un parametro o estar en alguna pantalla
        End Select
        reg_dtx_deta.MoveNext
      Loop
      Print #x, cadena
      reg_dtx_deta.Close
            
      sql = "Select * from " & _
                    " docu_blcab_exp " & _
                    " Where blcab_linea      ='" & linea & "'" & _
                    "   and blcab_buque    ='" & buque & "'" & _
                    "   and blcab_viaje    ='" & viaje & "'" & _
                    "   and blcab_puer_emb ='" & pto_emb & "'" & _
                    "   and blcab_puer_des ='" & pto_descarga & "'" & _
                    "   and blcab_emp_resp ='" & oper_carga & "'" & _
                    "   and blcab_cono_emb in (" & bls & ")" & _
            "Order by  blcab_cono_emb"
     reg_blcab.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
     No_bls = 0
     Do While Not reg_blcab.EOF
        No_bls = No_bls + 1
        '************************* A0 ********************************************
        cadena = "A0"
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        Do While Not reg_dtx_deta.EOF
            Select Case Trim(reg_dtx_deta("campo"))
                Case "1":
                        tama_bl = Len(Trim(reg_blcab("blcab_cono_emb")))
                        If tama_bl > reg_dtx_deta("longitud") Then tama_bl = reg_dtx_deta("longitud")
                        cadena = cadena & reg_blcab("blcab_cono_emb") & Space(reg_dtx_deta("longitud") - tama_bl)
                ' no se usa aparentemente 22 enero 2003
                Case "2": cadena = cadena & "" & Space(reg_dtx_deta("longitud") - Len("")) 'service contract number
            End Select
            reg_dtx_deta.MoveNext
        Loop
        Print #x, cadena
        reg_dtx_deta.Close
        '************************* A1 ********************************************
        cadena = "A1"
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        Do While Not reg_dtx_deta.EOF
            Select Case Trim(reg_dtx_deta("campo"))
                Case "1": cadena = cadena & Mid(reg_blcab("blcab_consig"), 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Mid(reg_blcab("blcab_consig"), 1, reg_dtx_deta("longitud"))))
            End Select
            reg_dtx_deta.MoveNext
        Loop
        Print #x, cadena
        reg_dtx_deta.Close
        '************************* A2 ********************************************
        cadena = "A2"
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        If Not reg_dtx_deta.EOF Then
            If Len(Trim(reg_blcab("blcab_direc_consig"))) > reg_dtx_deta("longitud") Then
                temp = Trim(reg_blcab("blcab_direc_consig"))
                i = 1:
                Do While temp <> "" And i <= 5
                    cadena = "A2" & Mid(temp, 1, reg_dtx_deta("longitud"))
                    temp = Mid(temp, reg_dtx_deta("longitud") + 1)
                    Print #x, cadena
                    i = i + 1
                Loop
            Else
                cadena = cadena & Mid(reg_blcab("blcab_direc_consig"), 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Mid(reg_blcab("blcab_direc_consig"), 1, reg_dtx_deta("longitud"))))
                Print #x, cadena
            End If
        End If
        reg_dtx_deta.Close
        '************************* A3 ********************************************
        cadena = "A3"
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        Do While Not reg_dtx_deta.EOF
            Select Case Trim(reg_dtx_deta("campo"))
                Case "1": cadena = cadena & Mid(reg_blcab("blcab_embarcador"), 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Mid(reg_blcab("blcab_embarcador"), 1, reg_dtx_deta("longitud"))))
            End Select
            reg_dtx_deta.MoveNext
        Loop
        Print #x, cadena
        
        reg_dtx_deta.Close
        '************************* A4 ********************************************
        cadena = "A4"
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        If Not reg_dtx_deta.EOF Then
            If Len(Trim(reg_blcab("blcab_direc_embarc"))) > reg_dtx_deta("longitud") Then
                temp = Trim(reg_blcab("blcab_direc_embarc"))
                i = 1
                Do While temp <> "" And i <= 3
                    cadena = "A4" & Mid(temp, 1, reg_dtx_deta("longitud"))
                    temp = Mid(temp, reg_dtx_deta("longitud") + 1)
                    Print #x, cadena:       i = i + 1
                Loop
            Else
                cadena = cadena & Mid(reg_blcab("blcab_direc_embarc"), 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Mid(reg_blcab("blcab_direc_embarc"), 1, reg_dtx_deta("longitud"))))
                Print #x, cadena
            End If
        End If
        reg_dtx_deta.Close
        '************************* A5 ********************************************
        cadena = "A5"
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        Do While Not reg_dtx_deta.EOF
            Select Case Trim(reg_dtx_deta("campo"))
                Case "1": cadena = cadena & Mid(reg_blcab("blcab_notif"), 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Mid(reg_blcab("blcab_notif"), 1, reg_dtx_deta("longitud"))))
            End Select
            reg_dtx_deta.MoveNext
        Loop
        Print #x, cadena
        reg_dtx_deta.Close
        '************************* A6 ********************************************
        cadena = "A6"
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        If Not reg_dtx_deta.EOF Then
            If Len(Trim(reg_blcab("blcab_direc_notif"))) > reg_dtx_deta("longitud") Then
                temp = Trim(reg_blcab("blcab_direc_notif"))
                i = 1:
                Do While temp <> "" And i <= 5 ' maximo 5 lineas
                    cadena = "A6" & Mid(temp, 1, reg_dtx_deta("longitud"))
                    temp = Mid(temp, reg_dtx_deta("longitud") + 1)
                    i = i + 1
                    Print #x, cadena
                Loop
            Else
                cadena = cadena & Mid(reg_blcab("blcab_direc_notif"), 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Mid(reg_blcab("blcab_direc_notif"), 1, reg_dtx_deta("longitud"))))
                Print #x, cadena
            End If
        End If
        reg_dtx_deta.Close
        '************************* A9 ********************************************
        '(no se usa 22 enero 2003)
        '************************* S1 ********************************************
        'totales por contenedores CY y CFS
        sql = "Select convert(char(1),cont_condicion) , sum(bldet_peso), sum(bldet_volumen) from " & _
                        " docu_bldet_exp, docu_cont_exp  " & _
                        " Where bldet_linea      ='" & linea & "'" & _
                        "   and bldet_buque    ='" & buque & "'" & _
                        "   and bldet_viaje    ='" & viaje & "'" & _
                        "   and bldet_puer_emb ='" & pto_emb & "'" & _
                        "   and bldet_puer_des ='" & pto_descarga & "'" & _
                        "   and bldet_emp_resp ='" & oper_carga & "'" & _
                        "   and bldet_cono_emb ='" & reg_blcab("blcab_cono_emb") & "'" & _
                        "   and bldet_linea    = cont_linea " & _
                        "   and bldet_buque    = cont_buque " & _
                        "   and bldet_viaje    = cont_viaje " & _
                        "   and bldet_puer_emb = cont_puer_emb " & _
                        "   and bldet_puer_des = cont_puer_des " & _
                        "   and bldet_contenedor = cont_contenedor" & _
                " Group by  convert(char(1),cont_condicion)"
        reg_cont.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        
        cadena = "S1"
        cont_peso_cy = 0:         cont_vol_cy = 0
        cont_peso_cfs = 0:        cont_vol_cfs = 0
        Do While Not reg_cont.EOF
            If reg_cont(0) = "1" Then
                cont_peso_cy = reg_cont(1)
                cont_vol_cy = reg_cont(2)
            ElseIf reg_cont(0) = "2" Then
                cont_peso_cfs = reg_cont(1)
                cont_vol_cfs = reg_cont(2)
            End If
            reg_cont.MoveNext
        Loop
        cadena = cadena & Devuelve_Formato("S1", "1", cont_vol_cy)
        cadena = cadena & Devuelve_Formato("S1", "2", cont_vol_cfs)
        cadena = cadena & Devuelve_Formato("S1", "3", cont_peso_cy)
        cadena = cadena & Devuelve_Formato("S1", "4", cont_peso_cfs)
        Print #x, cadena
        reg_cont.Close
        '************************* F1 ********************************************
        cadena = "F1"
        sql = "Select * from " & _
                        " docu_blflete_exp " & _
                        " Where flete_linea      ='" & linea & "'" & _
                        "   and flete_buque    ='" & buque & "'" & _
                        "   and flete_viaje    ='" & viaje & "'" & _
                        "   and flete_puer_emb ='" & pto_emb & "'" & _
                        "   and flete_puer_des ='" & pto_descarga & "'" & _
                        "   and flete_emp_resp ='" & oper_carga & "'" & _
                        "   and flete_cono_emb ='" & reg_blcab("blcab_cono_emb") & "'" & _
                " Order by  flete_cono_emb, flete_secuencia "
        reg_blflete.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        
        Do While Not reg_blflete.EOF
            cadena = "F1"
            reg_dtx_deta.MoveFirst
            If Not reg_dtx_deta.EOF Then flete_marcas = IIf(Not IsNull(reg_blflete("flete_marcas")), reg_blflete("flete_marcas"), "")
            Do While Not reg_dtx_deta.EOF
                Select Case Trim(reg_dtx_deta("campo"))
                    Case "1": cadena = cadena & Mid(reg_blflete("flete_cargo"), 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Trim(reg_blflete("flete_cargo")))) ' codigo del cargo
                    Case "2": cadena = cadena & Devuelve_Formato("F1", "2", reg_blflete("flete_valor_unit"))                                                                     ' valor unitario
                    Case "3": cadena = cadena & Mid(reg_blflete("flete_por"), 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(reg_blflete("flete_por")))     ' codigo del nivel de cargo (por)
                    Case "4": cadena = cadena & Mid(reg_blflete("flete_moneda"), 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(reg_blflete("flete_moneda"))) ' codigo de la moneda
                    Case "5": cadena = cadena & Devuelve_Formato("F1", "5", reg_blflete("flete_valor"))                                                                          ' valor total (amount)
                    Case "6": cadena = cadena & Trim(reg_blflete("flete_debtor"))                                                                                                ' debtor prepaid/colect
                    Case "7": cadena = cadena & Devuelve_Formato("F1", "7", reg_blflete("flete_cantidad"))                                                                       ' cantidad
                    Case "8":
                            If Not IsNull(reg_blflete("flete_lugar")) Then
                                cadena = cadena & Mid(reg_blflete("flete_lugar"), 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(reg_blflete("flete_lugar"))) ' lugar de pago
                            Else
                                cadena = cadena & Space(reg_dtx_deta("longitud"))  ' lugar de pago
                            End If
                End Select
                reg_dtx_deta.MoveNext
            Loop
            Print #x, cadena
            reg_blflete.MoveNext
        Loop
        reg_dtx_deta.Close
        '************************* F2 ********************************************
        ' es una observacion del F1
        cadena = "F2"
        sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
        reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        
        'reg_blflete.MoveFirst
        cadena = cadena & Mid(flete_marcas, 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(flete_marcas))    ' marcas del flete
        Print #x, cadena
        reg_dtx_deta.Close
        reg_blflete.Close
        '*************************************************************************
                
        ' **************  SELECT A LOS DETALLES DEL ACTUAL BL   ******************
         sql = "Select * from " & _
                        " docu_bldet_exp  " & _
                        " Where bldet_linea      ='" & linea & "'" & _
                        "   and bldet_buque    ='" & buque & "'" & _
                        "   and bldet_viaje    ='" & viaje & "'" & _
                        "   and bldet_puer_emb ='" & pto_emb & "'" & _
                        "   and bldet_puer_des ='" & pto_descarga & "'" & _
                        "   and bldet_emp_resp ='" & oper_carga & "'" & _
                        "   and bldet_cono_emb ='" & reg_blcab("blcab_cono_emb") & "'" & _
                " Order by  bldet_cono_emb, bldet_secuencia "
         reg_bldet.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
         If Not reg_bldet.EOF Then
            '************************* B1 *****************************************************************
            '**** POR LA DESCRIPCION DE LA CARGA SE PUEDEN REPETIR LAS LINEAS CON SEQUENCIA DIFERENTE *****
            cadena = "B1"
            i = 1 ' SECUENCIA
            sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
            reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
            Do While Not reg_dtx_deta.EOF
                Select Case Trim(reg_dtx_deta("campo"))
                    Case "1": cadena = cadena & Devuelve_Formato("B1", "1", i)
                    Case "2": cadena = cadena & Devuelve_Formato("B1", "2", reg_blcab("blcab_bultos_tot"))
                    Case "3": cadena = cadena & Mid(reg_bldet("bldet_embalaje"), 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(reg_bldet("bldet_embalaje")))
                    Case "4": cadena = cadena & Mid(reg_bldet("bldet_desc_carg"), 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Mid(reg_bldet("bldet_desc_carg"), 1, reg_dtx_deta("longitud"))))
                End Select
                reg_dtx_deta.MoveNext
            Loop
            Print #x, cadena
            ' PARA GRABAR LA INFORMACION DE LA DESCRIPCION EN MAS DE 1 LINEA
            reg_dtx_deta.MoveFirst: temp2 = ""
            Do While Not reg_dtx_deta.EOF
                Select Case Trim(reg_dtx_deta("campo"))
                    Case "2": temp2 = temp2 & Devuelve_Formato("B1", "2", "0")
                    Case "3": temp2 = temp2 & Space(reg_dtx_deta("longitud"))
                    Case "4": Exit Do
                End Select
                reg_dtx_deta.MoveNext
            Loop
            If Len(Trim(reg_bldet("bldet_desc_carg"))) > reg_dtx_deta("longitud") Then
                temp = Trim(reg_bldet("bldet_desc_carg"))
                i = 1
                Do While temp <> ""
                    cadena = "B1" & Devuelve_Formato("B1", "1", i) & temp2 & Mid(temp, 1, reg_dtx_deta("longitud"))
                    temp = Mid(temp, reg_dtx_deta("longitud") + 1)
                    Print #x, cadena
                    i = i + 1
                Loop
            End If
            reg_dtx_deta.Close
            '************************* FIN B1 ********************************************
            '************************* B2 *****************************************************************
            '**** POR LA MARCAS Y NUMEROS SE PUEDEN REPETIR LAS LINEAS CON SEQUENCIA DIFERENTE *****
            cadena = "B2"
            sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by campo"
            reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
            If Not reg_dtx_deta.EOF Then
                If Len(Trim(reg_bldet("bldet_marcas_num"))) > reg_dtx_deta("longitud") Then
                    temp = Trim(reg_bldet("bldet_marcas_num"))
                    i = 1:
                    reg_dtx_deta.MoveNext 'se lo mueve al siguiente registro que es el las marcas en si (ya que al abrir el cursosr esta el el registro de la secuencia)
                    Do While temp <> "" And i <= 5 ' maximo 5 lineas
                        cadena = "B2" & Devuelve_Formato("B2", "1", i) & Mid(temp, 1, reg_dtx_deta("longitud"))
                        temp = Mid(temp, reg_dtx_deta("longitud") + 1)
                        i = i + 1
                        Print #x, cadena
                    Loop
                Else
                    cadena = cadena & Devuelve_Formato("B2", "1", i) & Mid(reg_bldet("bldet_marcas_num"), 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Mid(reg_bldet("bldet_marcas_num"), 1, reg_dtx_deta("longitud"))))
                    Print #x, cadena
                End If
            End If
            reg_dtx_deta.Close
            '************************* FIN B2 ********************************************
        End If
        
        
        ' HASTA AQUI SOLO SE ACCESO A LOS DETALLES A BUSCAR CIERTOS CAMPOS PERO DE AQUI EN ADELANTE SE GRABAN LOS DETALLES
        Do While Not reg_bldet.EOF
           '************************* B3 ********************************************
            cadena = "B3"
            If reg_bldet("bldet_contenedor") <> "" Then
                sql = "Select * from " & _
                        " docu_cont_exp " & _
                        " Where cont_linea      ='" & linea & "'" & _
                        "   and cont_buque    ='" & buque & "'" & _
                        "   and cont_viaje    ='" & viaje & "'" & _
                        "   and cont_puer_emb ='" & pto_emb & "'" & _
                        "   and cont_puer_des ='" & pto_descarga & "'" & _
                        "   and cont_contenedor='" & reg_bldet("bldet_contenedor") & "'"
                reg_cont.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                If Not reg_cont.EOF Then
                    cont_sello = reg_cont("Cont_sello1")
                    cont_size = Mid(reg_cont("cont_tam_equip"), 1, 1)
                    cont_loading_serv = Mid(reg_cont("cont_condicion"), 1, 1)
                    cont_discharge_serv = Mid(reg_cont("cont_condicion"), 5, 1)
                    cont_type = Mid(reg_cont("cont_tam_equip"), 4, 1)
                    cont_sign_from = IIf(reg_cont("cont_temp_dur_almac") > 0, "+", "-")
                    cont_temp_from = reg_cont("cont_temp_dur_almac")
                    cont_sign_to = IIf(reg_cont("cont_temp_dur_trans") > 0, "+", "-")
                    cont_temp_to = reg_cont("cont_temp_dur_trans")
                Else
                    MsgBox "El contenedor " & reg_bldet("bldet_contenedor") & " No existe en los datos de contendores", vbInformation
                End If
            Else
                    cont_sello = ""
                    cont_size = ""
                    cont_loading_serv = ""
                    cont_discharge_serv = ""
                    cont_type = ""
                    cont_sign_from = ""
                    cont_temp_from = "0"
                    cont_sign_to = ""
                    cont_temp_to = "0"
            End If
            sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by desde"
            reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
            Do While Not reg_dtx_deta.EOF
                Select Case Trim(reg_dtx_deta("campo"))
                    Case "1": cadena = cadena & Mid(reg_bldet("bldet_contenedor"), 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Trim(reg_bldet("bldet_contenedor"))))           ' CONTENEDOR
                    Case "2": cadena = cadena & Mid(cont_sello, 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Mid(cont_sello, 1, reg_dtx_deta("longitud"))))                     ' SELLO
                    Case "3": cadena = cadena & Mid(cont_size, 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Mid(cont_size, 1, reg_dtx_deta("longitud"))))                       ' SIZE DEL CONTENEDOR
                    Case "4": cadena = cadena & Mid(cont_loading_serv, 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Mid(cont_loading_serv, 1, reg_dtx_deta("longitud"))))       ' loading service
                    Case "5": cadena = cadena & Mid(cont_discharge_serv, 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Mid(cont_discharge_serv, 1, reg_dtx_deta("longitud"))))   ' discharge service
                    Case "6": cadena = cadena & Devuelve_Formato("B3", "6", reg_bldet("bldet_bultos"))                                                                                                 ' no. of packages
                    Case "10": cadena = cadena & "C"                                                                                                                                                   ' COC/SOC PENDIENTE POR INCLUIR EN EL MANTENIMIENTO
                    Case "7": cadena = cadena & Mid(cont_type, 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Mid(cont_type, 1, reg_dtx_deta("longitud"))))                       ' CONTEINER TYPE
                    Case "8": cadena = cadena & Devuelve_Formato("B3", "8", reg_bldet("bldet_peso"))                                                                                                   ' WEIGHT
                    Case "9": cadena = cadena & Devuelve_Formato("B3", "9", reg_bldet("bldet_volumen"))                                                                                                ' MEASURMENT
                    Case "11": 'no se usan los customs seal
                End Select
                reg_dtx_deta.MoveNext
            Loop
            Print #x, cadena
            reg_dtx_deta.Close
            '************************ FIN B3 ********************************************
            
            '************************* D1 ***********************************************
            If reg_bldet("bldet_DANGER_CARGO") <> "" And reg_bldet("bldet_UN_number") <> "" And _
               reg_bldet("bldet_imdg_page") <> "" Then
                cadena = "D1"
                sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by desde"
                reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                Do While Not reg_dtx_deta.EOF
                    Select Case Trim(reg_dtx_deta("campo"))
                        Case "1": cadena = cadena & Mid(reg_bldet("bldet_DANGER_CARGO"), 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Trim(reg_bldet("bldet_DANGER_CARGO")))) ' peligrosidad KHS
                        Case "2": cadena = cadena & Mid(reg_bldet("bldet_UN_number"), 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Trim(reg_bldet("bldet_UN_number"))))       ' UN number KHS
                        Case "3": cadena = cadena & Mid(reg_bldet("bldet_imdg_page"), 1, reg_dtx_deta("longitud")) & Space(reg_dtx_deta("longitud") - Len(Trim(reg_bldet("bldet_imdg_page"))))       ' imdg page KHS
                    End Select
                    reg_dtx_deta.MoveNext
                Loop
                Print #x, cadena
                reg_dtx_deta.Close
            End If
            '************************* R1 ***********************************************
            If Trim(cont_temp_from) <> "0" And Trim(cont_temp_to) <> "0" Then
                cadena = "R1"
                sql = "Select * From Dtx_detalle Where linea = '" & linea & "' and codigo ='" & cadena & "' order by desde"
                reg_dtx_deta.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                Do While Not reg_dtx_deta.EOF
                    Select Case Trim(reg_dtx_deta("campo"))
                        Case "1": cadena = cadena & Mid(cont_sign_from, 1, reg_dtx_deta("longitud"))
                        Case "2": cadena = cadena & Devuelve_Formato("R1", "2", cont_temp_from)
                        Case "3": cadena = cadena & Mid(cont_sign_to, 1, reg_dtx_deta("longitud"))
                        Case "4": cadena = cadena & Devuelve_Formato("R1", "4", cont_temp_to)
                        Case "5": cadena = cadena & "C"
                    End Select
                    reg_dtx_deta.MoveNext
                Loop
                Print #x, cadena
                reg_dtx_deta.Close
            End If
            '****************************************************************************
            ' siguiente detalle
            reg_bldet.MoveNext
        Loop
        reg_bldet.Close
        
        reg_blcab.MoveNext
    Loop
    
    Print #x, "ZZ"
            
    MsgBox "Generación de archivo fue realizada con éxito. En el archivo " & Gstr_Ruta_files & buque & "\" & viaje & "\" & pto_descarga & "\EXPORT.TXT", vbInformation
     Close #x

    Exit Sub
Errores:
     MsgBox "Ocurrio el error " & Err.Description
     Close
     Exit Sub
     Resume 0

End Sub

Private Function Mail_LogOn() As Boolean
        
    If mapSess.NewSession Then
        MsgBox "Session already established", vbInformation
        Mail_LogOn = True
        Exit Function
    End If
    
    On Error GoTo errLogInFail
    With mapSess
        ' Set DownLoadMail to False to prevent immediate download.
        .DownLoadMail = False
        .LogonUI = True ' Use the underlying email system's logon UI.
        .SignOn ' Signon method.
        Mail_LogOn = True ' If successful, return True
        .NewSession = True ' Set NewSession to True and set0
        bNewSession = .NewSession ' variable flag to true
        mapMess.SessionID = .SessionID ' You must set this before continuing.
    End With
    ' Enabled and disable buttons.
    'ToggleButtonEnabled
    Exit Function
    
errLogInFail:
    
    'Debug.Print Err.Number, Err.Description
    If Err.Number = 32003 Then
        MsgBox "Canceled Login"
        Mail_LogOn = False
    Else
        MsgBox "Ocurrio el error " & Err.Description, vbCritical
    End If
    Exit Function
End Function

Private Sub Mail_ComposeMessage(Archivo As String)
    On Error GoTo ComposeErr
    Dim strMessage  As String
    
    mapMess.Compose
    mapMess.AttachmentIndex = 0
    mapMess.AttachmentPathName = Archivo  '"C:\cardat.zip"
    mapMess.RecipDisplayName = mail_to '"creyes75@hotmail.com"
    mapMess.MsgSubject = mail_subject '"prueba correo"
    mapMess.MsgNoteText = mail_text '"Esto es una prueba"
    mapMess.send True
    
    Exit Sub
ComposeErr:
    MsgBox "Ocurrio el error " & Err.Description, vbCritical
    Resume Next
End Sub

Private Sub Mail_LogOff()
    ' Logoff the MapSessions control.
    With mapSess
        .SignOff ' Close the session.
        .NewSession = False ' Flag for new session.
        bNewSession = .NewSession ' Reset flag.
    End With
    'grdMess.ClearFields ' Clear the grid.
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bNewSession Then Mail_LogOff
End Sub

Private Sub Carga_parametros_MAIL()
    On Error GoTo Errores
    Dim reg         As New ADODB.Recordset
    Dim sql         As String
    
    sql = "Select * from docu_parametro Where para_linea= 'CAE' And para_codigo like 'MAIL_%'"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Do While Not reg.EOF
        Select Case reg("para_codigo")
            Case "MAIL_TO":      mail_to = reg("para_valor")
            Case "MAIL_SUBJECT": mail_subject = reg("para_valor")
            Case "MAIL_TEXT":    mail_text = reg("para_valor")
        End Select
        reg.MoveNext
    Loop
    reg.Close
    
    Exit Sub
Errores:
    MsgBox "Ocurrio el error " & Err.Description
End Sub

