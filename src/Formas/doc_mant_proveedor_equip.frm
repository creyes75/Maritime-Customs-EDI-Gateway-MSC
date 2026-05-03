VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form doc_mant_prov_equip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Proveedor Equipamiento"
   ClientHeight    =   3870
   ClientLeft      =   3360
   ClientTop       =   4110
   ClientWidth     =   6480
   Icon            =   "doc_mant_proveedor_equip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   6480
   Begin VB.PictureBox CReport 
      Height          =   480
      Left            =   105
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   19
      Top             =   3480
      Width           =   1200
   End
   Begin TabDlg.SSTab Tab_dato 
      Height          =   3285
      Left            =   75
      TabIndex        =   4
      Top             =   60
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5794
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Listado"
      TabPicture(0)   =   "doc_mant_proveedor_equip.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "RDC_grid_dato"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Grid_dato"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Detalles"
      TabPicture(1)   =   "doc_mant_proveedor_equip.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_tipo_rol"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "RDC_dato"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   " Buscar "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   705
         Left            =   120
         TabIndex        =   14
         Top             =   2505
         Width           =   6120
         Begin VB.OptionButton Opt_codigo 
            Caption         =   "Por Código"
            Height          =   225
            Left            =   195
            TabIndex        =   18
            Top             =   315
            Width           =   1140
         End
         Begin VB.OptionButton Opt_desc 
            Caption         =   "Por Descripción"
            Height          =   225
            Left            =   1395
            TabIndex        =   17
            Top             =   315
            Width           =   1455
         End
         Begin VB.TextBox Txt_crite 
            Height          =   285
            Left            =   3030
            TabIndex        =   16
            Top             =   270
            Width           =   1965
         End
         Begin VB.CommandButton Cmd_buscar 
            Height          =   465
            Left            =   5085
            Picture         =   "doc_mant_proveedor_equip.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   195
            Width           =   465
         End
      End
      Begin Threed.SSFrame fra_tipo_rol 
         Height          =   2085
         Left            =   -74760
         TabIndex        =   6
         Top             =   480
         Width           =   5805
         _Version        =   65536
         _ExtentX        =   10239
         _ExtentY        =   3678
         _StockProps     =   14
         Caption         =   " Datos de Proveedor Equipamiento "
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
         Begin VB.TextBox txt_desc 
            Height          =   285
            Left            =   1680
            MaxLength       =   60
            TabIndex        =   8
            Top             =   1200
            Width           =   3855
         End
         Begin VB.TextBox txt_codigo 
            Height          =   285
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   7
            Top             =   840
            Width           =   855
         End
         Begin MSDataListLib.DataCombo Cmb_linea 
            Height          =   315
            Left            =   1680
            TabIndex        =   5
            Top             =   360
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
            Left            =   3120
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
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
            Left            =   375
            TabIndex        =   13
            Top             =   435
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   225
            Index           =   3
            Left            =   375
            TabIndex        =   10
            Top             =   1215
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   0
            Left            =   375
            TabIndex        =   9
            Top             =   840
            Width           =   720
         End
      End
      Begin MSDataGridLib.DataGrid Grid_dato 
         Height          =   2115
         Left            =   120
         TabIndex        =   12
         Top             =   390
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3731
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc RDC_grid_dato 
         Height          =   330
         Left            =   1800
         Top             =   2295
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSAdodcLib.Adodc RDC_dato 
         Height          =   330
         Left            =   -74760
         Top             =   2820
         Width           =   5895
         _ExtentX        =   10398
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
   End
   Begin Threed.SSCommand cmd_agregar 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3450
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Agregar"
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
   Begin Threed.SSCommand cmd_modificar 
      Height          =   375
      Left            =   2820
      TabIndex        =   1
      Top             =   3450
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Modificar"
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
      Enabled         =   0   'False
      Font3D          =   3
   End
   Begin Threed.SSCommand cmd_eliminar 
      Height          =   375
      Left            =   4005
      TabIndex        =   2
      Top             =   3450
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Eliminar"
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
      Enabled         =   0   'False
      Font3D          =   3
   End
   Begin Threed.SSCommand cmd_salir 
      Height          =   375
      Left            =   5190
      TabIndex        =   3
      Top             =   3450
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
   Begin Threed.SSCommand cmd_imprimir 
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   3450
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Imprimir"
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
End
Attribute VB_Name = "doc_mant_prov_equip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lstr_opc As String * 1

Public Sub habilita(ByVal opc As Boolean)
   If opc Then
      If Lstr_opc = "M" Then
         Cmb_linea.Enabled = False
         txt_codigo.Enabled = False
         'txt_desc.Enabled = False
      Else
         Cmb_linea.Enabled = opc
         txt_codigo.Enabled = opc
         txt_desc.Enabled = opc
'         Txt_nacio.Enabled = opc
      End If
   Else
      Cmb_linea.Enabled = True
      txt_codigo.Enabled = True
      txt_desc.Enabled = True
'      Txt_nacio.Enabled = True
   End If
   'txt_desc.Enabled = opc
   'Txt_part.Enabled = opc
   'Txt_dias_vaca.Enabled = opc
   'Chk_apli.Enabled = opc
End Sub
Private Sub limpia()
   Cmb_linea.BoundText = ""
   txt_codigo.Text = ""
   txt_desc.Text = ""
End Sub
Private Function valida_Datos() As Boolean
On Error GoTo ERROR_VALIDA
'Dim reg As New ADODB.Recordset

valida_Datos = True
 
 If Not Cmb_linea.MatchedWithList Then
      valida_Datos = False
      MsgBox "Ingrese la Línea", vbInformation, App.Title
      If Cmb_linea.Enabled Then Cmb_linea.SetFocus
      Exit Function
 ElseIf Trim$(txt_codigo.Text) = "" Then
      MsgBox "Ingrese el Código del Buque", vbInformation, App.Title
      valida_Datos = False
      If txt_codigo.Enabled Then txt_codigo.SetFocus
      Exit Function
 ElseIf Trim$(txt_desc.Text) = "" Then
      MsgBox "Ingrese la Descripcion del Buque", vbInformation, App.Title
      txt_desc.SetFocus
      valida_Datos = False
      If txt_desc.Enabled Then txt_desc.SetFocus
      Exit Function
 End If
 
   
   Exit Function
ERROR_VALIDA:
    MsgBox "Error :" & Err.Description, vbInformation, App.Title
    Exit Function
    Resume 0
End Function

Private Sub Graba_trol()
   Dim sql          As String
   Dim st_aplica    As String * 1
   Dim reg          As New ADODB.Recordset
   Dim reg_max      As New ADODB.Recordset
   Dim maximo       As Integer
   
   On Error GoTo ERROR_GRABACION
   
   Gcon_docu.BeginTrans
   
   If Lstr_opc = "A" Then
     sql = "SELECT prov_codigo FROM docu_proveedor_equip WHERE prov_linea ='" & Cmb_linea.BoundText & "' and  prov_codigo = '" & txt_codigo.Text & "'"
     reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
     If reg.EOF Then
        sql = "INSERT INTO docu_proveedor_equip (prov_linea, prov_codigo, prov_descripcion, prov_status)" & _
             " VALUES ('" & Cmb_linea.BoundText & "', '" & Trim$(txt_codigo.Text) & "','" & Trim$(txt_desc.Text) & "', 'A')"
     Else
        MsgBox "Este Proveedor de Equipamiento ya existe, entre por la opcion modificación", vbExclamation, App.Title
        Exit Sub
     End If
   ElseIf Lstr_opc = "M" Then
      sql = "UPDATE docu_proveedor_equip SET prov_descripcion = '" & Trim$(txt_desc.Text) & "' " & _
                               " WHERE prov_linea  = '" & Cmb_linea.BoundText & "'" & _
                               "   and prov_codigo = '" & Trim$(txt_codigo.Text) & "'"
   ElseIf Lstr_opc = "E" Then
       sql = "UPDATE docu_proveedor_equip set prov_status = 'E' " & _
             " WHERE prov_linea  = '" & Cmb_linea.BoundText & "'" & _
             "   and prov_codigo = '" & Trim$(txt_codigo.Text) & "'"
   End If
   Gcon_docu.Execute sql
   MsgBox "Transacción Realizada con éxito", vbInformation, App.Title
   limpia
   If Lstr_opc <> "A" Then
     habilita False
   End If
   Gcon_docu.CommitTrans
   Exit Sub
   
ERROR_GRABACION:
   MsgBox "Ocurrio el Error " & Err.Description, vbInformation, App.Title
   Gcon_docu.RollbackTrans
   Exit Sub
   Resume 0
End Sub

Private Function Nuevo_Tipo_Rol() As String
Dim sql     As String
Dim reg     As New ADODB.Recordset
Dim valor   As String

sql = "Select max(trol_codigo) from rhutipo_rol"
'Set REG = GCON_DOCU.OpenRecordset(sql)
reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
If reg.EOF Then
   valor = "01"
Else
   If IsNull(reg(0)) Or IsEmpty(reg(0)) Then
      valor = "01"
   Else
      valor = Format$(CInt(reg(0)) + 1, "00")
   End If
End If
Nuevo_Tipo_Rol = valor
End Function

Private Sub cmd_agregar_Click()
fra_tipo_rol.Enabled = True
Select Case cmd_agregar.Caption
    Case "&Agregar"
   'modificado darwin llenar combo de linea
   carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
'   data_linea.ConnectionString = Gstr_conex_docu
'   data_linea.RecordSource = "Select * from docu_linea_nav where linea_status ='A'"
'   data_linea.Refresh
'   Cmb_linea.BoundColumn = "linea_codigo"
'   Cmb_linea.ListField = "linea_descripcion"
'   Set Cmb_linea.RowSource = data_linea
   Cmb_linea.SetFocus
'el combo se ha llenado ok
      Lstr_opc = "A"
      'cmd_agregar.ToolTipText = "Grabar"
      cmd_agregar.Caption = "&Grabar"
      cmd_salir.Caption = "&Cancelar"
      Cmd_modificar.Enabled = False
      cmd_eliminar.Enabled = False
      RDC_dato.Caption = ""
      RDC_dato.Enabled = False
      habilita True
      limpia
      'txt_codigo.SetFocus
      txt_codigo.Text = "" 'Nuevo_Tipo_Rol
      Tab_dato.Tab = 1
      Tab_dato.TabEnabled(0) = False
  Case "&Grabar"
      If Not valida_Datos Then Exit Sub
      Graba_trol
      If Lstr_opc = "M" Or Lstr_opc = "E" Then
         Lstr_opc = "C"
         cmd_salir_Click
      Else
         txt_codigo.Text = "" 'Nuevo_Tipo_Rol
      End If
      If txt_codigo.Enabled Then txt_codigo.SetFocus
      RDC_dato.Caption = ""
      If Lstr_opc = "M" Then
         Tab_dato.Tab = 0
         Tab_dato.TabEnabled(0) = True
      End If
      RDC_grid_dato.Refresh:    DoEvents
      Grid_dato.Refresh:        DoEvents
      'Lstr_opc = "C"
      'Tab_dato.TabEnabled(0) = True
     ' Tab_dato.Tab = 0
    '  cmd_salir.Caption = "&Salir"
        limpia
        habilita False
        cmd_agregar.Caption = "&Agregar"
        cmd_salir.Caption = "&Salir"
        Cmd_modificar.Enabled = False
        cmd_eliminar.Enabled = False
        RDC_dato.Enabled = False
        RDC_dato.Caption = ""
        Lstr_opc = ""
        Tab_dato.TabEnabled(0) = True
        Tab_dato.Tab = 0
        fra_tipo_rol.Enabled = False
    End Select
    'fra_tipo_rol.Enabled = False
End Sub

Private Sub Cmd_buscar_Click()
    If Opt_codigo.Value Then
        RDC_grid_dato.Recordset.Find "codigo like '" & Txt_crite.Text & "*'"
    Else
        RDC_grid_dato.Recordset.Find "descripcion like '" & Txt_crite.Text & "*'"
    End If
End Sub

Private Sub cmd_eliminar_Click()
    Lstr_opc = "E"
    'cmd_agregar.Caption = "&Grabar"
    cmd_salir.Caption = "&Cancelar"
    'cmd_eliminar.Enabled = False
    'cmd_modificar.Enabled = False
    RDC_dato.Enabled = False
    'txt_codigo.Enabled = False
    Tab_dato.Tab = 1
    Tab_dato.TabEnabled(0) = False
    If MsgBox("¿Está seguro que desea Eliminar este registro?", vbInformation + vbYesNo, App.Title) = vbYes Then
        'habilita False
        Graba_trol
        Tab_dato.TabEnabled(0) = True
        'Form_Load
        RDC_grid_dato.Refresh
        Set Grid_dato.DataSource = RDC_grid_dato
        Grid_dato.Refresh
        Cmd_modificar.Enabled = False
        cmd_eliminar.Enabled = False
        'Tab_dato.Tab = 0
        limpia
        habilita False
        cmd_agregar.Caption = "&Agregar"
        cmd_salir.Caption = "&Salir"
        Cmd_modificar.Enabled = False
        cmd_eliminar.Enabled = False
        RDC_dato.Enabled = False
        RDC_dato.Caption = ""
        Lstr_opc = ""
        Tab_dato.TabEnabled(0) = True
        Tab_dato.Tab = 0
    End If
End Sub

Private Sub Cmd_Imprimir_Click()
On Error GoTo Errores
'    CReport.ReportFileName = Gstr_ruta_reportes & "cont026.rpt"
'    CReport.Connect = Gstr_conex_cont
'    CReport.SelectionFormula = "{cont_port.port_status}='A'"
'    CReport.WindowTitle = App.Title
    'CReport.Action = 1
    Call Excel_repo_mant("prov_equipo")
    Exit Sub
Errores:
    MsgBox "Ocurrio el error " & Err.Description, vbInformation, App.Title
End Sub

Private Sub cmd_modificar_Click()
    Lstr_opc = "M"
    cmd_agregar.Caption = "&Grabar"
    cmd_salir.Caption = "&Cancelar"
    Cmd_modificar.Enabled = False
    cmd_eliminar.Enabled = False
    RDC_dato.Enabled = False
    habilita True
    fra_tipo_rol.Enabled = True
    Tab_dato.Tab = 1
    Tab_dato.TabEnabled(0) = False
End Sub

Private Sub cmd_salir_Click()
Select Case cmd_salir.Caption
  Case "&Salir"
       Unload Me
  Case "&Cancelar"
       If Lstr_opc <> "C" Then resp = MsgBox("Está seguro que desea cancelar la operación", vbInformation + vbYesNo, App.Title)
       If resp = vbNo Then Exit Sub
       limpia
       habilita False
       cmd_agregar.Caption = "&Agregar"
       cmd_salir.Caption = "&Salir"
       Cmd_modificar.Enabled = False
       cmd_eliminar.Enabled = False
       RDC_dato.Enabled = False
       RDC_dato.Caption = ""
       Lstr_opc = ""
       fra_tipo_rol.Enabled = False
       Tab_dato.TabEnabled(0) = True
       Tab_dato.Tab = 0
End Select
End Sub

Private Sub Form_Load()
   Dim res As Boolean
   Me.Height = 4245
   Me.Width = 6630
   Centra_Pantalla Me
   Opt_codigo.Value = True
   
   carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"

   RDC_dato.ConnectionString = Gstr_conex_docu
   RDC_grid_dato.ConnectionString = Gstr_conex_docu
   RDC_grid_dato.RecordSource = "Select prov_linea as linea, prov_codigo as codigo, prov_descripcion as descripcion, prov_status  as Status from docu_proveedor_equip ORDER BY prov_descripcion"
   RDC_grid_dato.Refresh
   Set Grid_dato.DataSource = RDC_grid_dato
   
   Grid_dato.Refresh
   fra_tipo_rol.Enabled = False
   
End Sub

Private Sub Grid_dato_DblClick()
   If Not RDC_grid_dato.Recordset.EOF Then
      Cmb_linea.BoundText = RDC_grid_dato.Recordset("LINEA")
      txt_codigo.Text = RDC_grid_dato.Recordset("codigo")
      If RDC_grid_dato.Recordset("status") = "E" Then
         If MsgBox("Este Registro esta con STATUS Eliminado, desea Activarlo?", vbQuestion + vbYesNo) = vbYes Then
            sql = "Update docu_proveedor_equip set prov_status='A' " & _
                  "Where prov_linea  ='" & RDC_grid_dato.Recordset("linea") & "'" & _
                  "  And prov_codigo ='" & RDC_grid_dato.Recordset("codigo") & "'"
            Gcon_docu.Execute sql
         End If
      End If
      Call txt_codigo_LostFocus
      Tab_dato.Tab = 1
   End If
End Sub

Private Sub Grid_dato_HeadClick(ByVal ColIndex As Integer)
Me.MousePointer = 11
    RDC_grid_dato.Recordset.Sort = Grid_dato.Columns(ColIndex).DataField
    Grid_dato.Refresh
Me.MousePointer = 0
End Sub

Private Sub rdc_dato_Reposition()
If Not RDC_dato.Recordset.EOF Then
   Cmb_linea.BoundText = Trim$(RDC_dato.Recordset.Fields("prov_linea"))
   txt_codigo.Text = Trim$(RDC_dato.Recordset.Fields("prov_codigo"))
   txt_desc.Text = Trim$(RDC_dato.Recordset.Fields("prov_descripcion"))
   RDC_dato.Caption = Trim(Str(RDC_dato.Recordset.AbsolutePosition)) & " De " & Str(RDC_dato.Recordset.RecordCount) & " Registros Afectados"
End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_codigo_LostFocus()
Dim reg     As New ADODB.Recordset
Dim sql     As String
Dim i       As Integer
If Trim$(txt_codigo.Text) = "" Then
   'limpia
   Cmd_modificar.Enabled = False
   cmd_eliminar.Enabled = False
   RDC_dato.Enabled = False
   Exit Sub
ElseIf Lstr_opc = "A" Or Lstr_opc = "M" Or Lstr_opc = "E" Then
   Cmd_modificar.Enabled = False
   cmd_eliminar.Enabled = False
   RDC_dato.Enabled = False
   Exit Sub
End If

sql = "Select * FROM docu_proveedor_equip WHERE prov_linea ='" & Cmb_linea.BoundText & "' and prov_codigo like '" & txt_codigo.Text & "'"
      
reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
'Set REG = GCON_DOCU.OpenRecordset(sql)
If Not reg.EOF Then
   RDC_dato.RecordSource = sql
   RDC_dato.Refresh
   rdc_dato_Reposition
   
   Cmd_modificar.Enabled = True
   cmd_eliminar.Enabled = True
   RDC_dato.Enabled = True
Else
   limpia
   RDC_dato.Caption = "0 Registros Afectados"
   If txt_codigo.Enabled Then txt_codigo.SetFocus
   RDC_dato.Enabled = False
   Cmd_modificar.Enabled = False
   cmd_eliminar.Enabled = False
End If
End Sub

Private Sub txt_desc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_desc_LostFocus()
Dim reg     As New ADODB.Recordset
Dim sql     As String
Dim i       As Integer
If Trim$(txt_desc.Text) = "" Then
   'limpia
   Cmd_modificar.Enabled = False
   cmd_eliminar.Enabled = False
   RDC_dato.Enabled = False
   Exit Sub
ElseIf Lstr_opc = "A" Or Lstr_opc = "M" Or Lstr_opc = "E" Then
   Cmd_modificar.Enabled = False
   cmd_eliminar.Enabled = False
   RDC_dato.Enabled = False
   Exit Sub
End If

sql = "Select * FROM docu_proveedor_equip WHERE prov_linea ='" & Cmb_linea.BoundText & "' and prov_descripcion like '" & txt_desc.Text & "'"
reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
'Set REG = GCON_DOCU.OpenRecordset(sql)
If Not reg.EOF Then
   RDC_dato.RecordSource = sql
   RDC_dato.Refresh
   rdc_dato_Reposition
   
   Cmd_modificar.Enabled = True
   cmd_eliminar.Enabled = True
   RDC_dato.Enabled = True
Else
   limpia
   RDC_dato.Caption = "0 Registros Afectados"
   If txt_codigo.Enabled Then txt_codigo.SetFocus
   RDC_dato.Enabled = False
   Cmd_modificar.Enabled = False
   cmd_eliminar.Enabled = False
End If

End Sub
