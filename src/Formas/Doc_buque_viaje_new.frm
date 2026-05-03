VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "GTDATE32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form doc_buque_viaje 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Buque/Viaje"
   ClientHeight    =   3990
   ClientLeft      =   3330
   ClientTop       =   4110
   ClientWidth     =   6525
   Icon            =   "Doc_buque_viaje_new.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3990
   ScaleWidth      =   6525
   Begin TabDlg.SSTab Tab_dato 
      Height          =   3360
      Left            =   75
      TabIndex        =   4
      Top             =   60
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5927
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Listado"
      TabPicture(0)   =   "Doc_buque_viaje_new.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Grid_dato"
      Tab(0).Control(1)=   "RDC_grid_dato"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalles"
      TabPicture(1)   =   "Doc_buque_viaje_new.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fra_tipo_rol"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "RDC_dato"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin MSDataGridLib.DataGrid Grid_dato 
         Height          =   2685
         Left            =   -74880
         TabIndex        =   12
         Top             =   480
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4736
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
         Left            =   -73200
         Top             =   2955
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
         Left            =   240
         Top             =   2895
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
      Begin Threed.SSFrame fra_tipo_rol 
         Height          =   2460
         Left            =   315
         TabIndex        =   5
         Top             =   345
         Width           =   5805
         _Version        =   65536
         _ExtentX        =   10239
         _ExtentY        =   4339
         _StockProps     =   14
         Caption         =   " Datos de buque / Viaje "
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
         Begin VB.TextBox txt_vessel 
            Height          =   375
            Left            =   4800
            MaxLength       =   5
            TabIndex        =   16
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Frame Frame1 
            Caption         =   " Direccion "
            ForeColor       =   &H00800000&
            Height          =   540
            Left            =   3390
            TabIndex        =   19
            Top             =   1395
            Width           =   1725
            Begin VB.OptionButton Opt_R 
               Caption         =   "R"
               Height          =   210
               Left            =   915
               TabIndex        =   10
               Top             =   255
               Width           =   555
            End
            Begin VB.OptionButton Opt_A 
               Caption         =   "A"
               Height          =   210
               Left            =   135
               TabIndex        =   20
               Top             =   255
               Width           =   555
            End
         End
         Begin VB.TextBox Txt_voyage 
            Height          =   285
            Left            =   3390
            MaxLength       =   6
            TabIndex        =   17
            Top             =   1065
            Width           =   1230
         End
         Begin VB.TextBox txt_desc 
            Height          =   285
            Left            =   1290
            MaxLength       =   60
            TabIndex        =   21
            Top             =   2070
            Width           =   4005
         End
         Begin GTMaskDate.GTMaskDate GTD_date 
            Height          =   300
            Left            =   285
            TabIndex        =   18
            Top             =   1635
            Width           =   1545
            _Version        =   65537
            _ExtentX        =   2725
            _ExtentY        =   529
            _StockProps     =   77
            BackColor       =   -2147483643
            KeepFocusOnError=   -1  'True
            BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalDayCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo Cmb_vessel 
            Height          =   315
            Left            =   240
            TabIndex        =   15
            Top             =   1080
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_vessel 
            Height          =   330
            Left            =   1920
            Top             =   1035
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
         Begin MSDataListLib.DataCombo cmb_linea 
            Height          =   315
            Left            =   840
            TabIndex        =   13
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
            Left            =   2625
            Top             =   360
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Línea :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   225
            TabIndex        =   14
            Top             =   420
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Arribo :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   195
            TabIndex        =   9
            Top             =   1425
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Viaje :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   3435
            TabIndex        =   8
            Top             =   840
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comentarios :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   165
            TabIndex        =   7
            Top             =   2085
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Buque :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   6
            Top             =   840
            Width           =   555
         End
      End
   End
   Begin Threed.SSCommand cmd_agregar 
      Height          =   375
      Left            =   1635
      TabIndex        =   0
      Top             =   3510
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
      Top             =   3510
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
      Top             =   3510
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
      Top             =   3510
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
      Left            =   450
      TabIndex        =   11
      Top             =   3510
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
   Begin VB.PictureBox CReport 
      Height          =   480
      Left            =   90
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   22
      Top             =   3540
      Width           =   1200
   End
End
Attribute VB_Name = "doc_buque_viaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lstr_opc As String * 1
Dim bandera As Boolean

Public Sub habilita(ByVal opc As Boolean)
   If opc Then
      If Lstr_opc = "M" Then
         Cmb_linea.Enabled = False
         txt_vessel.Enabled = False
         Cmb_vessel.Enabled = False
         Txt_voyage.Enabled = False
         'txt_desc.Enabled = False
      Else
         Cmb_linea.Enabled = opc
         txt_vessel.Enabled = opc
         Cmb_vessel.Enabled = opc
         Txt_voyage.Enabled = opc
         GTD_date.Enabled = opc
         txt_desc.Enabled = opc
      End If
   Else
      Cmb_linea.Enabled = True
      txt_vessel.Enabled = True
      Cmb_vessel.Enabled = True
      Txt_voyage.Enabled = True
      GTD_date.Enabled = True
      txt_desc.Enabled = True
   End If
   'txt_desc.Enabled = opc
   'Txt_part.Enabled = opc
   'Txt_dias_vaca.Enabled = opc
   'Chk_apli.Enabled = opc
End Sub
Private Sub limpia()
   txt_vessel.Text = ""
   Txt_voyage.Text = ""
   GTD_date.Text = ""
   txt_desc.Text = ""
   Cmb_linea.Text = ""
   Cmb_vessel.Text = ""
   'Txt_part.Text = ""
   'Txt_dias_vaca.Text = ""
   'Chk_apli.Value = 0
End Sub
'Private Function Valida_Datos_old() As Boolean
'On Error GoTo ERROR_VALIDA
'Dim reg As New ADODB.Recordset

'valida_Datos = True
   'If Trim$(txt_vessel.Text) = "" Then
 '  If cmb_linea.BoundText = "" Then
  '    MsgBox "Ingrese el Valor de Linea", vbInformation, App.Title
   '   valida_Datos = False
      'If txt_vessel.Enabled Then txt_vessel.SetFocus
    '  If cmb_linea.Enabled Then cmb_linea.SetFocus
     ' Exit Function
   'End If
   
   'If Cmb_vessel.BoundText = "" Then
    '  MsgBox "Ingrese el Valor de Vessel", vbInformation, App.Title
     ' valida_Datos = False
      'If txt_vessel.Enabled Then txt_vessel.SetFocus
      'If Cmb_vessel.Enabled Then Cmb_vessel.SetFocus
      'Exit Function
   'End If
  ' If Trim$(Txt_voyage.Text) = "" Then
   '   MsgBox "Ingrese el Código de Voyage", vbInformation, App.Title
    '  valida_Datos = False
     ' If Txt_voyage.Enabled Then Txt_voyage.SetFocus
      'Exit Function
   'End If
 '  If Trim$(GTD_date.Text) = "" Then
  '    MsgBox "Ingrese el Fecha ", vbInformation, App.Title
   '   valida_Datos = False
    '  If GTD_date.Enabled Then GTD_date.SetFocus
     ' Exit Function
  ' End If

'   If Trim$(txt_desc.Text) = "" Then
'      MsgBox "Ingrese algún comentario ", vbInformation, App.Title
'      Valida_Datos = False
'      If txt_desc.Enabled Then txt_desc.SetFocus
'      Exit Function
'   End If
      
'   Exit Function
'ERROR_VALIDA:
'MsgBox "Error :" & Err.Description, vbInformation, App.Title
'Exit Function
'Resume 0

'End Function
Private Sub valida_Datos()
'On Error GoTo ERROR_VALIDA
'Dim reg As New ADODB.Recordset

'valida_Datos = True
   'If Trim$(txt_vessel.Text) = "" Then
   If Cmb_linea.BoundText = "" Then
      MsgBox "Ingrese la Linea", vbInformation, App.Title
      bandera = False
      'If txt_vessel.Enabled Then txt_vessel.SetFocus
      Cmb_linea.SetFocus
   ElseIf Cmb_vessel.BoundText = "" Then
      MsgBox "Ingrese el Buque", vbInformation, App.Title
      bandera = False
      Cmb_vessel.SetFocus
      'If txt_vessel.Enabled Then txt_vessel.SetFocus
   ElseIf Txt_voyage.Text = "" Then
      MsgBox "Ingrese el Número de Viaje", vbInformation, App.Title
      bandera = False
      Txt_voyage.SetFocus
   ElseIf Len(Trim(Txt_voyage.Text)) > 5 Then
      If MsgBox("El Número de Viaje tiene una longitud de " & Len(Txt_voyage.Text) & " caracteres, cuando el tamaño normal es de 5 caracteres, desea continuar?", vbYesNo + vbInformation, App.Title) = vbNo Then
            bandera = False
            Txt_voyage.SetFocus
       Else
            bandera = True
       End If
 ElseIf GTD_date = "  /  /" Or GTD_date = "" Then
     MsgBox "Ingrese la Fecha", vbInformation, App.Title
      bandera = False
      GTD_date.SetFocus
 Else
      bandera = True
   
   'If Trim$(GTD_date.Text) = "" Then
    '  MsgBox "Ingrese el Fecha ", vbInformation, App.Title
     ' valida_Datos = False
     ' If GTD_date.Enabled Then GTD_date.SetFocus
     ' Exit Sub
   'End If

'   If Trim$(txt_desc.Text) = "" Then
'      MsgBox "Ingrese algún comentario ", vbInformation, App.Title
'      Valida_Datos = False
'      If txt_desc.Enabled Then txt_desc.SetFocus
'      Exit Function
'   End If
      
 End If
   'Exit Sub
'ERROR_VALIDA:
'MsgBox "Error :" & Err.Description, vbInformation, App.Title
'Exit Sub
'Resume 0

End Sub


Private Sub Graba_trol()
   Dim sql          As String
   Dim st_aplica    As String * 1
   Dim reg          As New ADODB.Recordset
   Dim reg_max      As New ADODB.Recordset
   Dim maximo       As Integer
   
   On Error GoTo ERROR_GRABACION
   
   Gcon_docu.BeginTrans
   If Opt_A.Value Then Direction = "A" Else Direction = "R"
   
   If Lstr_opc = "A" Then
     sql = "Select * from docu_buque_viaje Where buqv_linea = '" & Cmb_linea.BoundText & "' and buqv_buque = '" & Cmb_vessel.BoundText & "' and buqv_viaje = '" & Txt_voyage.Text & "'"
     reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
     If reg.EOF Then
        sql = "INSERT INTO docu_buque_viaje (buqv_linea, buqv_buque, buqv_viaje, buqv_fecha_arribo, buqv_direccion, buqv_observacion, buqv_status, buqv_fecha_ing, buqv_user_ing)" & _
             " VALUES ('" & Trim$(Cmb_linea.BoundText) & "','" & Trim$(Cmb_vessel.BoundText) & "','" & Trim$(Txt_voyage.Text) & "','" & Format(GTD_date.Text, GSTR_FORMATO_FECHA) & "','" & Direction & "','" & Trim$(txt_desc.Text) & "', 'A', '" & Format$(Date, GSTR_FORMATO_FECHA) & "','" & Gstr_user & "')"
     Else
        MsgBox "Este Buque/Viaje ya existe, entre por la opción modificación", vbExclamation, App.Title
        Exit Sub
     End If
   ElseIf Lstr_opc = "M" Then
      sql = "UPDATE docu_buque_viaje SET buqv_fecha_arribo = '" & Format(GTD_date.Text, GSTR_FORMATO_FECHA) & "', " & _
                                          "buqv_direccion    = '" & Direction & "'," & _
                                          "buqv_observacion  = '" & Trim$(txt_desc.Text) & "'," & _
                                          "buqv_status       = 'A' " & _
                                    "WHERE buqv_linea   = '" & Trim$(Cmb_linea.BoundText) & "'" & _
                                    "  and buqv_buque   = '" & Trim$(Cmb_vessel.BoundText) & "'" & _
                                    "  and buqv_viaje   = '" & Trim$(Txt_voyage.Text) & "'"
   ElseIf Lstr_opc = "E" Then
'      sql = "Select trcg_cargo from rhutrol_cargo where trcg_tiporol='" & txt_codigo.Text & "' and trcg_st_tupla = 'A'"
'      Set reg = GCON_DOCU.OpenRecordset(sql, rdOpenKeyset)
'      If Not reg.EOF Then
'         MsgBox "Este Tipo de Rol tiene Cargos asociados, NO es posible Eliminarlo", vbInformation, App.Title
'         GCON_DOCU.RollbackTrans
'         Exit Sub
'      End If
        sql = "UPDATE docu_buque_viaje SET " & _
                                            "buqv_status       = 'E' " & _
                                      "WHERE buqv_linea   = '" & Trim$(Cmb_linea.BoundText) & "'" & _
                                      "  and buqv_buque   = '" & Trim$(Cmb_vessel.BoundText) & "'" & _
                                      "  and buqv_viaje   = '" & Trim$(Txt_voyage.Text) & "'"
                                                 
   End If
   If Lstr_opc = "E" Then
      resp = MsgBox("Esta seguro que desea eliminar este Buque/Viaje", vbQuestion + vbYesNo, App.Title)
   Else
      resp = vbYes
   End If
   If resp = vbYes Then
      Gcon_docu.Execute sql
      MsgBox "Transacción Realizada con éxito", vbInformation, App.Title
   Else
      Exit Sub
      'MsgBox "Transacción cancelada "
   End If
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

Private Sub cmb_linea_Change()
If Cmb_linea.MatchedWithList Then
   data_vessel.ConnectionString = Gstr_conex_docu
   data_vessel.RecordSource = "Select * from docu_buque where buqu_linea = '" & Cmb_linea.BoundText & "' and buqu_status ='A' Order by Buqu_descripcion"
   data_vessel.Refresh
   Cmb_vessel.BoundColumn = "buqu_codigo"
   Cmb_vessel.ListField = "buqu_descripcion"
   Set Cmb_vessel.RowSource = data_vessel
End If
End Sub

Public Sub cmd_agregar_Click()
Select Case cmd_agregar.Caption
  Case "&Agregar"
      fra_tipo_rol.Enabled = True
      Cmb_linea.SetFocus
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
      'Cmb_vessel.SetFocus
      'cmb_vessel.lis = "" 'Nuevo_Tipo_Rol
      Txt_voyage.Text = ""
      GTD_date.Text = ""
      txt_desc.Text = ""
      Tab_dato.Tab = 1
      Tab_dato.TabEnabled(0) = False
  Case "&Grabar"
      valida_Datos
      If bandera = True Then
      Graba_trol
      If Lstr_opc = "M" Or Lstr_opc = "E" Then
         Lstr_opc = "C"
         cmd_salir_Click
      Else
         txt_vessel.Text = "" 'Nuevo_Tipo_Rol
      End If
      'If txt_vessel.Enabled Then txt_vessel.SetFocus
      If Cmb_vessel.Enabled Then Cmb_vessel.SetFocus
      RDC_dato.Caption = ""
      If Lstr_opc = "M" Then
         Tab_dato.Tab = 0
         Tab_dato.TabEnabled(0) = True
      End If
      'RDC_grid_dato.Refresh
      'Grid_dato.Refresh
      'Lstr_opc = "C"
      fra_tipo_rol.Enabled = False
      'Tab_dato.TabEnabled(0) = True
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
       RDC_grid_dato.Refresh: DoEvents
       Grid_dato.Refresh:     DoEvents
      End If
      
End Select
End Sub

Public Sub cmd_eliminar_Click()
    Lstr_opc = "E"
    'cmd_agregar.Caption = "&Grabar"
    'cmd_salir.Caption = "&Cancelar"
    'cmd_eliminar.Enabled = False
    'cmd_modificar.Enabled = False
    'RDC_dato.Enabled = False
    Graba_trol
    habilita True
   ' Cmb_vessel.Enabled = False
    'Txt_voyage.Enabled = False
    Tab_dato.Tab = 0
       RDC_grid_dato.Refresh: DoEvents
       Grid_dato.Refresh:     DoEvents
       Cmd_modificar.Enabled = False
       cmd_eliminar.Enabled = False
   ' Tab_dato.TabEnabled(0) = False
End Sub

Private Sub Cmd_Imprimir_Click()
On Error GoTo Errores
'    CReport.ReportFileName = Gstr_ruta_reportes & "cont030.rpt"
'    CReport.Connect = Gstr_conex_docu
'    CReport.SelectionFormula = "{docu_buque_viaje.vess_status}='A'"
'    CReport.WindowTitle = App.Title
'    CReport.Action = 1
    Exit Sub
Errores:
    MsgBox "Ocurrio el error " & Err.Description, vbInformation, App.Title

End Sub

Public Sub cmd_modificar_Click()
    fra_tipo_rol.Enabled = True
    Lstr_opc = "M"
    cmd_agregar.Caption = "&Grabar"
    cmd_salir.Caption = "&Cancelar"
    Cmd_modificar.Enabled = False
    cmd_eliminar.Enabled = False
    RDC_dato.Enabled = False
    habilita True
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
       Tab_dato.TabEnabled(0) = True
       Tab_dato.Tab = 0
       RDC_grid_dato.Refresh: DoEvents
       Grid_dato.Refresh:     DoEvents
       fra_tipo_rol.Enabled = False
End Select
End Sub

Private Sub Form_Load()
   Dim res As Boolean
   Me.Height = 4365
   Me.Width = 6630
   Centra_Pantalla Me
   
   'RDC_dato.ConnectionString = Gstr_conex_docu
   
   RDC_grid_dato.ConnectionString = Gstr_conex_docu
   RDC_grid_dato.RecordSource = "Select buqv_linea as Linea, buqv_buque as Vessel, buqv_viaje as Voyage, buqv_direccion as Direction, buqv_fecha_arribo as SailArrivDate, buqv_observacion as Comments from docu_buque_viaje ORDER BY  Linea, Vessel, Voyage"
   RDC_grid_dato.Refresh
   Set Grid_dato.DataSource = RDC_grid_dato
   Grid_dato.Refresh
   
   data_linea.ConnectionString = Gstr_conex_docu
   data_linea.RecordSource = "Select * from docu_linea_nav where linea_status ='A' Order by linea_descripcion"
   data_linea.Refresh
   Cmb_linea.BoundColumn = "linea_codigo"
   Cmb_linea.ListField = "linea_descripcion"
   Set Cmb_linea.RowSource = data_linea
   fra_tipo_rol.Enabled = False
   Opt_A = True
End Sub

Private Sub Grid_dato_DblClick()
   If Not RDC_grid_dato.Recordset.EOF Then
      Cmb_linea.BoundText = Trim$(RDC_grid_dato.Recordset.Fields("linea"))
      Cmb_vessel.BoundText = Trim$(RDC_grid_dato.Recordset.Fields("vessel"))
      txt_vessel.Text = Trim$(RDC_grid_dato.Recordset.Fields("Vessel"))
      Txt_voyage.Text = Trim$(RDC_grid_dato.Recordset.Fields("Voyage"))
      txt_desc.Text = Trim$(IIf(Not IsNull(RDC_grid_dato.Recordset.Fields("Comments")), RDC_grid_dato.Recordset.Fields("Comments"), ""))
      If Not IsNull(RDC_grid_dato.Recordset.Fields("Direction")) Then
        If Trim$(RDC_grid_dato.Recordset.Fields("Direction")) = "A" Then
           Opt_A.Value = True
        ElseIf Trim$(RDC_grid_dato.Recordset.Fields("Direction")) = "R" Then
           Opt_R.Value = True
        End If
      End If
      GTD_date.Text = IIf(IsDate((RDC_grid_dato.Recordset.Fields("SailArrivDate"))), (RDC_grid_dato.Recordset.Fields("SailArrivDate")), "")
      Cmd_modificar.Enabled = True
      cmd_eliminar.Enabled = True
      'Call txt_codigo_LostFocus
      Tab_dato.Tab = 1
   End If
End Sub

Private Sub Grid_dato_HeadClick(ByVal ColIndex As Integer)
Me.MousePointer = 11
    RDC_grid_dato.Recordset.Sort = Grid_dato.Columns(ColIndex).DataField
    Grid_dato.Refresh
'If ColIndex = 1 Then
'   RDC_grid_dato.RecordSource = "Select buqv_linea as Linea, buqv_buque as Vessel, buqv_viaje as Voyage, buqv_direccion as Direction, buqv_fecha_arribo as SailArrivDate, buqv_observacion as Comments from docu_buque_viaje Where buqv_status = 'A' ORDER BY buqv_linea"
'ElseIf ColIndex = 1 Then
'   RDC_grid_dato.RecordSource = "Select buqv_linea as Linea, buqv_buque as Vessel, buqv_viaje as Voyage, buqv_direccion as Direction, buqv_fecha_arribo as SailArrivDate, buqv_observacion as Comments from docu_buque_viaje Where buqv_status = 'A' ORDER BY buqv_buque"
'ElseIf ColIndex = 2 Then
'   RDC_grid_dato.RecordSource = "Select buqv_linea as Linea, buqv_buque as Vessel, buqv_viaje as Voyage, buqv_direccion as Direction, buqv_fecha_arribo as SailArrivDate, buqv_observacion as Comments from docu_buque_viaje Where buqv_status = 'A' ORDER BY buqv_viaje"
'ElseIf ColIndex = 3 Then
'   RDC_grid_dato.RecordSource = "Select buqv_linea as Linea, buqv_buque as Vessel, buqv_viaje as Voyage, buqv_direccion as Direction, buqv_fecha_arribo as SailArrivDate, buqv_observacion as Comments from docu_buque_viaje Where buqv_status = 'A' ORDER BY buqv_direccion"
'ElseIf ColIndex = 4 Then
'   RDC_grid_dato.RecordSource = "Select buqv_linea as Linea, buqv_buque as Vessel, buqv_viaje as Voyage, buqv_direccion as Direction, buqv_fecha_arribo as SailArrivDate, buqv_observacion as Comments from docu_buque_viaje Where buqv_status = 'A' ORDER BY buqv_fecha_arribo"
'End If
'RDC_grid_dato.Refresh
'Grid_dato.Refresh
Me.MousePointer = 0
End Sub

Private Sub rdc_dato_Reposition()
If Not RDC_dato.Recordset.EOF Then
      Cmb_linea.BoundText = Trim$(RDC_dato.Recordset.Fields("linea"))
      Cmb_vessel.BoundText = Trim$(RDC_dato.Recordset.Fields("vessel"))
      txt_vessel.Text = Trim$(RDC_dato.Recordset.Fields("Vessel"))
      Txt_voyage.Text = Trim$(RDC_dato.Recordset.Fields("Voyage"))
      txt_desc.Text = Trim$(IIf(Not IsNull(RDC_dato.Recordset.Fields("Comments")), RDC_dato.Recordset.Fields("Comments"), ""))
      If Not IsNull(RDC_dato.Recordset.Fields("Direction")) Then
        If Trim$(RDC_dato.Recordset.Fields("Direction")) = "A" Then
           Opt_A.Value = True
        ElseIf Trim$(RDC_dato.Recordset.Fields("Direction")) = "R" Then
           Opt_R.Value = True
        End If
      End If
      GTD_date.Text = IIf(IsDate((RDC_dato.Recordset.Fields("SailArrivDate"))), (RDC_dato.Recordset.Fields("SailArrivDate")), "")
      RDC_dato.Caption = Trim(Str(RDC_dato.Recordset.AbsolutePosition)) & " De " & Str(RDC_dato.Recordset.RecordCount) & " Registros Afectados"
End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txt_codigo_LostFocus()
Dim reg     As New ADODB.Recordset
Dim sql     As String
Dim i       As Integer
If Trim$(txt_codigo.Text) = "" Then
   limpia
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

sql = "Select * FROM cont_Ownership WHERE Owner_codigo like '" & txt_codigo.Text & "'" & _
      " and Owner_status = 'A'"
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
'Dim REG     As new ADODB.recordset
'Dim sql     As String
'Dim i       As Integer
'If Trim$(txt_desc.Text) = "" Then
'   limpia
'   cmd_modificar.Enabled = False
'   cmd_eliminar.Enabled = False
'   RDC_dato.Enabled = False
'   Exit Sub
'ElseIf Lstr_opc = "A" Or Lstr_opc = "M" Or Lstr_opc = "E" Then
'   cmd_modificar.Enabled = False
'   cmd_eliminar.Enabled = False
'   RDC_dato.Enabled = False
'   Exit Sub
'End If
'
'sql = "Select * FROM cont_Ownership WHERE Owner_descripcion like '" & txt_desc.Text & "'" & _
'      " and Owner_status = 'A'"
'Set REG = GCON_DOCU.OpenRecordset(sql)
'If Not REG.EOF Then
'   rdc_dato.recordsource = sql
'   RDC_dato.Refresh
'   rdc_dato_Reposition
'
'   cmd_modificar.Enabled = True
'   cmd_eliminar.Enabled = True
'   RDC_dato.Enabled = True
'Else
'   limpia
'   RDC_dato.Caption = "0 Registros Afectados"
'   If txt_codigo.Enabled Then txt_codigo.SetFocus
'   RDC_dato.Enabled = False
'   cmd_modificar.Enabled = False
'   cmd_eliminar.Enabled = False
'End If

End Sub

Public Sub Txt_voyage_LostFocus()
Dim reg     As New ADODB.Recordset
Dim sql     As String
Dim i       As Integer
If Not Cmb_linea.MatchedWithList Or Not Cmb_vessel.MatchedWithList Or _
   Trim(Txt_voyage.Text) = "" Then
   limpia
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


sql = "Select buqv_linea as Linea, buqv_buque as Vessel, buqv_viaje as Voyage, buqv_direccion as Direction, buqv_fecha_arribo as SailArrivDate, buqv_observacion as Comments " & _
      " from docu_buque_viaje " & _
      " Where buqv_linea ='" & Cmb_linea.BoundText & "' " & _
      "   And buqv_buque ='" & Cmb_vessel.BoundText & "'" & _
      "   And buqv_viaje ='" & Txt_voyage.Text & "'" & _
      "   And buqv_status = 'A' ORDER BY  Linea, Vessel, Voyage"
reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
If Not reg.EOF Then
   RDC_dato.ConnectionString = Gstr_conex_docu
   RDC_dato.RecordSource = sql
   RDC_dato.Refresh
   rdc_dato_Reposition
   
   Cmd_modificar.Enabled = True
   cmd_eliminar.Enabled = True
   RDC_dato.Enabled = True
Else
   limpia
   RDC_dato.Caption = "0 Registros Afectados"
   If Txt_voyage.Enabled Then Txt_voyage.SetFocus
   RDC_dato.Enabled = False
   Cmd_modificar.Enabled = False
   cmd_eliminar.Enabled = False
End If

End Sub
