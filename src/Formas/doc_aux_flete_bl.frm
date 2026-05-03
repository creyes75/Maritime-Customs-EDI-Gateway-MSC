VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{667E8C40-F9B5-11CF-90AB-444553540000}#1.0#0"; "GTNUM32.OCX"
Begin VB.Form doc_aux_flete_bl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fletes"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "doc_aux_flete_bl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fra_flete 
      Height          =   3420
      Left            =   60
      TabIndex        =   0
      Top             =   -45
      Width           =   4260
      Begin VB.Frame Fra_cantidad 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   1815
         TabIndex        =   33
         Top             =   2865
         Width           =   1815
         Begin GTMaskNum.GTMaskNum GTM_cant 
            Height          =   315
            Left            =   795
            TabIndex        =   34
            Top             =   0
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            TabOnEnter      =   -1  'True
            BeginProperty CalcDispFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalcBtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalcMemFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskShowThousandSep=   -1  'True
            MaskType        =   0
            DataType        =   4
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
         Begin VB.Label Lbl_cant 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   0
            Width           =   630
         End
      End
      Begin VB.Frame Fra_valor_unit 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1800
         TabIndex        =   30
         Top             =   2505
         Width           =   2085
         Begin GTMaskNum.GTMaskNum GTM_valor_unit 
            Height          =   315
            Left            =   960
            TabIndex        =   31
            Top             =   0
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            TabOnEnter      =   -1  'True
            BeginProperty CalcDispFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalcBtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalcMemFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskShowThousandSep=   -1  'True
            MaskType        =   0
            DataType        =   4
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
         Begin VB.Label Lbl_valor_unit 
            AutoSize        =   -1  'True
            Caption         =   "Valor Unit"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   0
            Width           =   690
         End
      End
      Begin VB.Frame Fra_valor 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   30
         TabIndex        =   27
         Top             =   1020
         Width           =   2445
         Begin GTMaskNum.GTMaskNum GTM_flete 
            Height          =   315
            Left            =   960
            TabIndex        =   28
            Top             =   0
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            TabOnEnter      =   -1  'True
            BeginProperty CalcDispFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalcBtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CalcMemFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskShowThousandSep=   -1  'True
            MaskType        =   0
            DataType        =   4
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
         Begin VB.Label Lbl_valor 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   0
            Width           =   360
         End
      End
      Begin VB.Frame Fra_nivel_cargo 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   345
         Left            =   30
         TabIndex        =   24
         Top             =   2940
         Width           =   4080
         Begin MSDataListLib.DataCombo Cmb_nivel_cargo 
            Height          =   315
            Left            =   960
            TabIndex        =   25
            Top             =   -30
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_nivel_cargo 
            Height          =   330
            Left            =   2805
            Top             =   -15
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
            Caption         =   "Nivel Cargo"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   105
            TabIndex        =   26
            Top             =   0
            Width           =   825
         End
      End
      Begin VB.CheckBox Chk_impr 
         Caption         =   "Imprimir en Bl"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   165
         TabIndex        =   18
         Top             =   2580
         Width           =   1335
      End
      Begin VB.Frame Fra_lugar 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   30
         TabIndex        =   14
         Top             =   2190
         Width           =   2475
         Begin VB.CommandButton Cmd_hlp_pto_descarga 
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
            Left            =   1980
            Picture         =   "doc_aux_flete_bl.frx":57E2
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Buscar"
            Top             =   0
            Width           =   300
         End
         Begin VB.TextBox Txt_pto 
            Height          =   285
            Left            =   960
            MaxLength       =   5
            TabIndex        =   15
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Lugar Pago"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   45
            Width           =   825
         End
      End
      Begin VB.TextBox Txt_marcas 
         Height          =   285
         Left            =   990
         MaxLength       =   80
         TabIndex        =   12
         Top             =   1830
         Width           =   3015
      End
      Begin VB.Frame Fra_moneda 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   345
         Left            =   30
         TabIndex        =   9
         Top             =   1440
         Width           =   4080
         Begin MSDataListLib.DataCombo Cmb_moneda 
            Height          =   315
            Left            =   960
            TabIndex        =   10
            Top             =   0
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_moneda 
            Height          =   330
            Left            =   2805
            Top             =   -15
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   0
            Width           =   585
         End
      End
      Begin VB.ComboBox Cmb_signo 
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1050
         Width           =   645
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   345
         Left            =   30
         TabIndex        =   4
         Top             =   630
         Width           =   4185
         Begin MSDataListLib.DataCombo Cmb_debtor 
            Height          =   315
            Left            =   975
            TabIndex        =   5
            Top             =   -15
            Width           =   2745
            _ExtentX        =   4842
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_debtor 
            Height          =   330
            Left            =   2775
            Top             =   -15
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
            Caption         =   "Debtor"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   105
            TabIndex        =   6
            Top             =   30
            Width           =   480
         End
      End
      Begin VB.Frame Frame24 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   345
         Left            =   30
         TabIndex        =   1
         Top             =   225
         Width           =   4170
         Begin MSDataListLib.DataCombo Cmb_cargo 
            Height          =   315
            Left            =   975
            TabIndex        =   2
            Top             =   -15
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_cargo 
            Height          =   330
            Left            =   2775
            Top             =   -15
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
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Cargo"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   105
            TabIndex        =   3
            Top             =   30
            Width           =   420
         End
      End
      Begin VB.Label Lbl_marcas 
         AutoSize        =   -1  'True
         Caption         =   "Marcas"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1860
         Width           =   525
      End
      Begin VB.Label Lbl_signo 
         AutoSize        =   -1  'True
         Caption         =   "Signo"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2580
         TabIndex        =   8
         Top             =   1095
         Width           =   405
      End
   End
   Begin Threed.SSCommand btn_sgte 
      Height          =   495
      Left            =   3255
      TabIndex        =   19
      ToolTipText     =   "Aceptar"
      Top             =   3000
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   78
      Picture         =   "doc_aux_flete_bl.frx":5D14
   End
   Begin Threed.SSCommand btn_elim 
      Height          =   495
      Left            =   3795
      TabIndex        =   20
      ToolTipText     =   "Salir"
      Top             =   3000
      Width           =   495
      _Version        =   65536
      _ExtentX        =   868
      _ExtentY        =   868
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "doc_aux_flete_bl.frx":6166
   End
   Begin VB.Frame Fra_linea 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   540
      Left            =   150
      TabIndex        =   21
      Top             =   2955
      Visible         =   0   'False
      Width           =   1500
      Begin MSDataListLib.DataCombo Cmb_linea 
         Height          =   315
         Left            =   135
         TabIndex        =   22
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc data_linea 
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
      Begin VB.Label Label6 
         Caption         =   "Linea"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   105
         TabIndex        =   23
         Top             =   0
         Width           =   825
      End
   End
End
Attribute VB_Name = "doc_aux_flete_bl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_elim_Click()
Unload Me
End Sub

Private Sub btn_sgte_Click()
        
    Gtyp_flete.cargo = Cmb_cargo.Text
    Gtyp_flete.cargo_cod = Cmb_cargo.BoundText
    Gtyp_flete.debtor = Cmb_debtor.Text
    Gtyp_flete.deptor_cod = Cmb_debtor.BoundText
    Gtyp_flete.imp_bl = Chk_impr.Value
    Gtyp_flete.lugar = Txt_pto.Text
    Gtyp_flete.marcas = Txt_marcas.Text
    Gtyp_flete.MONEDA = Cmb_moneda.BoundText
    Gtyp_flete.valor = GTM_flete.Value
    Gtyp_flete.SIGNO = Cmb_signo.Text
    If Cmb_linea.BoundText = "KHS" Then
        Gtyp_flete.valor_unit = GTM_valor_unit.Value
        Gtyp_flete.cantidad = GTM_cant.Value
        Gtyp_flete.nivel_cargo = Cmb_nivel_cargo.BoundText
    End If
    Gtyp_flete.reg_cancel = False
    Unload Me
End Sub

Private Sub cmb_linea_Change()
     carga_datacombo data_moneda, Cmb_moneda, "Select * from docu_moneda  Where mone_linea        = '" & Cmb_linea.BoundText & "' and mone_status      ='A'", "mone_codigo", "mone_descripcion"
     carga_datacombo data_cargo, Cmb_cargo, "Select * from docu_cargo     Where cargo_linea       = '" & Cmb_linea.BoundText & "' and cargo_status     ='A'", "cargo_codigo", "cargo_descripcion"
     carga_datacombo data_debtor, Cmb_debtor, "Select * from docu_debtor  Where debtor_linea      = '" & Cmb_linea.BoundText & "' and debtor_status    ='A'", "debtor_codigo", "debtor_descripcion"
     carga_datacombo data_nivel_cargo, Cmb_nivel_cargo, "Select * from docu_Nivel_cargo  Where nivel_cargo_linea      = '" & Cmb_linea.BoundText & "' and nivel_cargo_status    ='A'", "nivel_cargo_codigo", "nivel_cargo_descripcion"
     If Cmb_linea.BoundText = "KHS" Then
        Cmb_signo.Visible = False:      Lbl_signo.Visible = False
        Fra_valor_unit.Visible = True:  Fra_valor_unit.Left = 30:   Fra_valor_unit.Top = 1080:
        Fra_cantidad.Visible = True:        Fra_cantidad.Left = 2150:       Fra_cantidad.Top = 1095
        Fra_nivel_cargo.Visible = True: Fra_nivel_cargo.Top = 1440
        Fra_moneda.Visible = True: Fra_moneda.Top = 1830
        Fra_valor.Visible = True:  Fra_valor.Left = 30:   Fra_valor.Top = 2210:
        Lbl_marcas.Visible = True: Lbl_marcas.Left = 120: Lbl_marcas.Top = 2580
        Txt_marcas.Visible = True: Txt_marcas.Left = 990: Txt_marcas.Top = 2580
        Fra_lugar.Visible = True: Fra_lugar.Left = 30: Fra_lugar.Top = 2940
        Chk_impr.Visible = True: Chk_impr.Left = 2450: Chk_impr.Top = 2940
        Fra_flete.Height = 3420
        btn_sgte.Top = 3420
        btn_elim.Top = 3420
        Me.Height = 4320
        
     Else
        Fra_valor.Top = 1080:  Cmb_signo.Visible = True: Lbl_signo.Visible = True
        Fra_moneda.Top = 1440
        Lbl_marcas.Top = 1860: Txt_marcas.Top = 1830
        Fra_lugar.Top = 2190
        Chk_impr.Top = 2595
        Fra_valor_unit.Visible = False
        Fra_nivel_cargo.Visible = False
        Fra_cantidad.Visible = False
        Me.Height = 3945
        
        Fra_flete.Height = 2910
        btn_sgte.Top = 3000
        btn_elim.Top = 3000
     End If
     
End Sub

Private Sub Cmd_hlp_pto_descarga_Click()
   Txt_pto.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto")
End Sub

Public Sub Form_Load()
    carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
    Cmb_signo.AddItem "+"
    Cmb_signo.AddItem "-"
    Gtyp_flete.cargo = ""
    Gtyp_flete.cargo_cod = ""
    Gtyp_flete.debtor = ""
    Gtyp_flete.deptor_cod = ""
    Gtyp_flete.imp_bl = False
    Gtyp_flete.lugar = ""
    Gtyp_flete.marcas = ""
    Gtyp_flete.MONEDA = ""
    Gtyp_flete.valor = 0
    Gtyp_flete.SIGNO = ""
    Gtyp_flete.reg_cancel = True
    
End Sub

