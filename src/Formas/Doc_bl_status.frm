VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "GTDATE32.OCX"
Object = "{667E8C40-F9B5-11CF-90AB-444553540000}#1.0#0"; "GTNUM32.OCX"
Begin VB.Form Doc_bl_status 
   Caption         =   "Bl Status"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13245
   Icon            =   "Doc_bl_status.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   13245
   Begin VB.Frame Fra_vessel 
      Height          =   6615
      Left            =   90
      TabIndex        =   1
      Top             =   0
      Width           =   3210
      Begin VB.CommandButton Cmd_buscar 
         Height          =   465
         Left            =   2640
         Picture         =   "Doc_bl_status.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Consultar"
         Top             =   120
         Width           =   465
      End
      Begin VB.CommandButton Cmd_config 
         Caption         =   ">>"
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
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Expande opciones de Filtro"
         Top             =   240
         Width           =   300
      End
      Begin ComctlLib.TreeView TrV_buque 
         Height          =   5280
         Left            =   120
         TabIndex        =   2
         Top             =   1230
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   9313
         _Version        =   327682
         Indentation     =   443
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin VB.Frame Fra_linea 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   660
         Left            =   30
         TabIndex        =   3
         Top             =   585
         Width           =   3060
         Begin MSDataListLib.DataCombo Cmb_linea 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc RDC_linea 
            Height          =   330
            Left            =   1425
            Top             =   75
            Visible         =   0   'False
            Width           =   1470
            _ExtentX        =   2593
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
            Left            =   105
            TabIndex        =   5
            Top             =   0
            Width           =   825
         End
      End
      Begin VB.Frame Fra_opciones 
         Height          =   1575
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
         Begin VB.TextBox Txt_nombre 
            Height          =   285
            Left            =   120
            MaxLength       =   80
            TabIndex        =   19
            ToolTipText     =   "Nombre del buque a buscar (No es necesario ingresar '*')"
            Top             =   1200
            Width           =   2160
         End
         Begin VB.OptionButton Opt_nombre 
            Caption         =   "Nombre contenga"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   "Buques que su nombre inicie con"
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton Opt_rango 
            Caption         =   "Rango Fechas"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            ToolTipText     =   "Buques arribados en este rango de fechas"
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton Opt_dias 
            Caption         =   "No. Días"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   "Todos los Buques arribados desde hace N dias"
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Opt_todos 
            Caption         =   "Todos Buque/Viaje"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   "Todos los buques existentes"
            Top             =   240
            Width           =   1815
         End
         Begin GTMaskNum.GTMaskNum GTM_dias 
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   1200
            Width           =   930
            _Version        =   65536
            _ExtentX        =   1640
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
         Begin GTMaskDate.GTMaskDate Msk_fecha_desde 
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   1200
            Width           =   1215
            _Version        =   65537
            _ExtentX        =   2143
            _ExtentY        =   503
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
            MaskCentury     =   2
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
         Begin GTMaskDate.GTMaskDate Msk_fecha_hasta 
            Height          =   285
            Left            =   1440
            TabIndex        =   18
            Top             =   1200
            Width           =   1215
            _Version        =   65537
            _ExtentX        =   2143
            _ExtentY        =   503
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
            MaskCentury     =   2
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
      End
      Begin VB.Label lbl_opcion 
         Caption         =   "Label1"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   2055
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   2640
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   6
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Doc_bl_status.frx":2004
               Key             =   "ViajeAbierto"
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Doc_bl_status.frx":231E
               Key             =   "Viaje"
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Doc_bl_status.frx":2638
               Key             =   "Buque"
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Doc_bl_status.frx":2952
               Key             =   "BuqueAbierto"
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Doc_bl_status.frx":2C6C
               Key             =   "PM"
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Doc_bl_status.frx":2F86
               Key             =   "ini"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Fra_bls 
      Height          =   6615
      Left            =   3375
      TabIndex        =   0
      Top             =   0
      Width           =   9780
      Begin MSDataGridLib.DataGrid Grid_CONT 
         Height          =   1425
         Left            =   165
         TabIndex        =   9
         Top             =   5040
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   2514
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
      Begin MSDataGridLib.DataGrid Grid_bl 
         Height          =   2400
         Left            =   150
         TabIndex        =   6
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4233
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
      Begin MSAdodcLib.Adodc RDC_grid_bl 
         Height          =   330
         Left            =   5085
         Top             =   2565
         Visible         =   0   'False
         Width           =   5265
         _ExtentX        =   9287
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
      Begin MSDataGridLib.DataGrid Grid_DET 
         Height          =   2000
         Left            =   150
         TabIndex        =   7
         Top             =   2970
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   3519
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
      Begin MSAdodcLib.Adodc RDC_grid_det 
         Height          =   330
         Left            =   5055
         Top             =   4665
         Visible         =   0   'False
         Width           =   5265
         _ExtentX        =   9287
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
      Begin MSAdodcLib.Adodc RDC_grid_cont 
         Height          =   330
         Left            =   5160
         Top             =   6150
         Visible         =   0   'False
         Width           =   5265
         _ExtentX        =   9287
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
      Begin VB.Label Lbl_bls 
         Caption         =   "No. Bls"
         Height          =   255
         Left            =   195
         TabIndex        =   8
         Top             =   2730
         Width           =   2340
      End
   End
End
Attribute VB_Name = "Doc_bl_status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Trv_buque_cod   As String
Public Trv_buque_desc  As String
Public Trv_viaje       As String
Public Trv_puer_emb    As String
Public Trv_ImpExp      As String
Dim Trv_key         As String
'Dim GRID_BL_FILA    As Variant

Private Sub Cmd_buscar_Click()
    Dim sql As String
    Dim reg As New ADODB.Recordset
    Dim REG_viaje As New ADODB.Recordset
    Dim REG_pto As New ADODB.Recordset
    Dim nodo_AC      As Node
    
    If Not Cmb_linea.MatchedWithList Then
        MsgBox "Seleecione una línea de la lista ", vbInformation
        Exit Sub
    End If
    
    carga_parametros Cmb_linea.BoundText
    Screen.MousePointer = 11
    TrV_buque.Nodes.Clear
    '***************************************************************************
    '******************* se evalua los criterios de filtro *********************
    '***************************************************************************
    sql = "Select * from docu_buque " & _
              " Where buqu_linea  = '" & Cmb_linea.BoundText & "'" & _
              "   And buqu_status = 'A'"

    If Opt_todos.Value Then ' obtener todos los buques
        sql = sql & " ORDER BY buqu_descripcion"
    End If
    If Opt_dias.Value Then ' obtener solo los buques cuyo atraque sea dentro de los N dias anteriores
        If GTM_dias.Value <= 0 Then
            MsgBox "Ingrese un número válido de días", vbInformation
            Exit Sub
        End If
              sql = sql & "   And buqu_codigo in (Select buqv_buque from docu_buque_viaje " & _
                                      " where buqv_linea = '" & Cmb_linea.BoundText & "' " & _
                                      "   and buqv_fecha_arribo >='" & Format(DateAdd("d", GTM_dias.Value * -1, Now), GSTR_FORMATO_FECHA) & "')" & _
              " ORDER BY buqu_descripcion"
        lbl_opcion.Caption = "Filtro por No. Días: " & GTM_dias.Value & " días"
    End If
    If Opt_rango.Value Then ' rango de fechas
        If Not IsDate(Msk_fecha_desde.DateValue) Or Not IsDate(Msk_fecha_hasta.DateValue) Then
            MsgBox "Ingrese un rango de fechas válidas ", vbInformation: Exit Sub
        End If
        If Msk_fecha_desde.DateValue > Msk_fecha_hasta.DateValue Then
            MsgBox "La fecha desde debe ser menor a la fecha hasta", vbInformation: Exit Sub
        End If
        sql = sql & _
              "   And buqu_codigo in (Select buqv_buque from docu_buque_viaje " & _
                                      " where buqv_linea = '" & Cmb_linea.BoundText & "' " & _
                                      "   and buqv_fecha_arribo between '" & Format(Msk_fecha_desde.DateValue, GSTR_FORMATO_FECHA) & "' and '" & Format(Msk_fecha_hasta.DateValue, GSTR_FORMATO_FECHA) & "')" & _
              " ORDER BY buqu_descripcion"
         lbl_opcion.Caption = "Filtro por rango de fechas; de :" & Msk_fecha_desde.DateValue & " a:" & Msk_fecha_hasta.DateValue
    End If
    If Opt_nombre.Value Then
        If Trim(Txt_nombre.Text) = "" Then
            MsgBox "Ingrese un nombre válido", vbInformation: Exit Sub
        End If
        sql = sql & _
              "   And buqu_descripcion like '%" & Trim(Txt_nombre.Text) & "%'" & _
              " ORDER BY buqu_descripcion"
        lbl_opcion.Caption = "Filtro por nombre, *" & Txt_nombre.Text & "*"
    End If
    '***************************************************************************
    
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    
    Set nodo = TrV_buque.Nodes.Add(, , "INI", Cmb_linea.Text, "ini")
    Do While Not reg.EOF
        
        Set nodo = TrV_buque.Nodes.Add("INI", 4, "BQ" & reg("buqu_codigo"), reg("buqu_descripcion"), "Buque", "BuqueAbierto")
        sql = "Select * from docu_buque_viaje " & _
          " Where buqv_linea  = '" & Cmb_linea.BoundText & "'" & _
          "   And buqv_buque  = '" & reg("buqu_codigo") & "'" & _
          "   And buqv_status = 'A'"
          ' se evaluan los filtros de fechas y numero de dias
          If Opt_dias.Value Then sql = sql & "   and buqv_fecha_arribo >='" & Format(DateAdd("d", Me.GTM_dias.Value * -1, Now), GSTR_FORMATO_FECHA) & "'"
          If Opt_rango.Value Then sql = sql & "  and buqv_fecha_arribo between '" & Format(Msk_fecha_desde.DateValue, GSTR_FORMATO_FECHA) & "' and '" & Format(Msk_fecha_hasta.DateValue, GSTR_FORMATO_FECHA) & "'"
        
        REG_viaje.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        Do While Not REG_viaje.EOF
            Set nodo = TrV_buque.Nodes.Add("BQ" & reg("buqu_codigo"), 4, "VI" & reg("buqu_codigo") & "|" & REG_viaje("buqv_viaje"), REG_viaje("buqv_viaje"), "Viaje", "ViajeAbierto")
            ' se agregan los nodos de Importacion / Exportacion
            Set nodo = TrV_buque.Nodes.Add("VI" & reg("buqu_codigo") & "|" & REG_viaje("buqv_viaje"), 4, "IE" & reg("buqu_codigo") & "|" & REG_viaje("buqv_viaje") & "|" & "IMP", "Importación", "Viaje", "ViajeAbierto")
            Set nodo = TrV_buque.Nodes.Add("VI" & reg("buqu_codigo") & "|" & REG_viaje("buqv_viaje"), 4, "IE" & reg("buqu_codigo") & "|" & REG_viaje("buqv_viaje") & "|" & "EXP", "Exportación", "Viaje", "ViajeAbierto")
            
            sql = "Select distinct mani_puer_emb, puer_descripcion from docu_manifiesto, docu_puerto " & _
                " Where mani_linea   = '" & Cmb_linea.BoundText & "'" & _
                "   And mani_buque   = '" & reg("buqu_codigo") & "'" & _
                "   And mani_viaje   = '" & REG_viaje("buqv_viaje") & "'" & _
                "   And mani_linea   = puer_linea " & _
                "   And mani_puer_emb= puer_codigo " & _
                "   ORDER BY puer_descripcion"
                '"   And mani_status = 'A'"
        
            REG_pto.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
            Do While Not REG_pto.EOF
                Set nodo = TrV_buque.Nodes.Add("IE" & reg("buqu_codigo") & "|" & REG_viaje("buqv_viaje") & "|" & "IMP", 4, "PT" & reg("buqu_codigo") & "|" & REG_viaje("buqv_viaje") & "|" & REG_pto("mani_puer_emb") & "IMP", REG_pto("puer_descripcion"), "Viaje", "ViajeAbierto")
                REG_pto.MoveNext
            Loop
            REG_pto.Close
            
            sql = "Select distinct mani_puer_des, puer_descripcion from docu_manifiesto_exp, docu_puerto " & _
                " Where mani_linea   = '" & Cmb_linea.BoundText & "'" & _
                "   And mani_buque   = '" & reg("buqu_codigo") & "'" & _
                "   And mani_viaje   = '" & REG_viaje("buqv_viaje") & "'" & _
                "   And mani_linea   = puer_linea " & _
                "   And mani_puer_des= puer_codigo " & _
                "   ORDER BY puer_descripcion"
                '"   And mani_status = 'A'"
        
            REG_pto.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
            Do While Not REG_pto.EOF
                Set nodo = TrV_buque.Nodes.Add("IE" & reg("buqu_codigo") & "|" & REG_viaje("buqv_viaje") & "|" & "EXP", 4, "PT" & reg("buqu_codigo") & "|" & REG_viaje("buqv_viaje") & "|" & REG_pto("mani_puer_des") & "EXP", REG_pto("puer_descripcion"), "Viaje", "ViajeAbierto")
                REG_pto.MoveNext
            Loop
            REG_pto.Close
            
            REG_viaje.MoveNext
        Loop
        REG_viaje.Close
        reg.MoveNext
    Loop
    reg.Close
    Screen.MousePointer = 0
End Sub


Private Sub Cmd_config_Click()
If Fra_opciones.Visible Then
    Fra_opciones.Visible = False
    Me.Fra_linea.Top = 585
    Me.TrV_buque.Top = 1230
    Me.TrV_buque.Height = 5280
    Cmd_config.Caption = ">>": Cmd_config.ToolTipText = "Expander opciones de Filtro"
Else
    Fra_opciones.Visible = True
    Me.Fra_linea.Top = 2145
    Me.TrV_buque.Top = 2790
    Me.TrV_buque.Height = 3720
    Cmd_config.Caption = "<<": Cmd_config.ToolTipText = "Contraer opciones de Filtro"
End If
End Sub

Private Sub Form_Load()
    carga_datacombo RDC_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
    Me.Top = 0
    Me.Left = 0
    Me.Width = 13365
    Me.Height = 7095
    Opt_todos.Value = True
End Sub

Private Sub Form_Resize()
If Me.Width - 3600 > 0 Then Fra_bls.Width = Me.Width - 3600
If (Me.Height - 520) > 0 Then Fra_bls.Height = (Me.Height - 520)
If Me.Width - 3900 > 0 Then Grid_bl.Width = Me.Width - 3900
If (Me.Height - 1095) / 3 > 0 Then Grid_bl.Height = (Me.Height - 1095) / 3


If Me.Width - 3900 > 0 Then Grid_DET.Width = Me.Width - 3900
If (Me.Height - 1095) / 3 > 0 Then Grid_DET.Height = (Me.Height - 1095) / 3
If 500 + (Me.Height - 1095) / 3 > 0 Then Grid_DET.Top = 500 + (Me.Height - 1095) / 3

If 500 + (Me.Height - 1495) / 3 > 0 Then Lbl_bls.Top = Grid_bl.Height + 250 '500 + (Me.Height - 1495) / 3

If Me.Width - 3900 > 0 Then Grid_CONT.Width = Me.Width - 3900
'If (Me.Height - 1095) / 3 > 0 Then Grid_CONT.Height = (Me.Height - 1095) / 4
If (Me.Height - 1095) / 3 > 0 Then Grid_CONT.Height = (Me.Height - 1095) - (Grid_DET.Height + Grid_bl.Height + 300)
If (Me.Height - 1095) / 3 > 0 Then Grid_CONT.Top = Grid_DET.Height + Grid_bl.Height + 800 '+ (Me.Height - 1095) / 3



End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo salir
Unload doc_menu_aux
salir:
Exit Sub
End Sub

Public Sub Grid_bl_DblClick()
Dim frm_bl_mod     As New doc_bl
Dim frm_bl_mod_exp As New doc_expo_bl
'Gstr_opc_bl = "M"
If Not RDC_grid_bl.Recordset.EOF Then
    If Trv_ImpExp = "I" Then    'Importación
        frm_bl_mod.Show
        frm_bl_mod.Gstr_opc_bl = "M"
        Call frm_bl_mod.carga_bl(Cmb_linea.BoundText, Trv_buque_cod, Trv_viaje, RDC_grid_bl.Recordset("Pto_embarque"), RDC_grid_bl.Recordset("Pto_descarga"), RDC_grid_bl.Recordset("bl"))
        frm_bl_mod.Caption = "Bl Impo: " & RDC_grid_bl.Recordset("bl")
    ElseIf Trv_ImpExp = "E" Then ' Exportación
        frm_bl_mod_exp.Show
        frm_bl_mod_exp.Gstr_opc_bl = "M"
        Call frm_bl_mod_exp.carga_bl(Cmb_linea.BoundText, Trv_buque_cod, Trv_viaje, RDC_grid_bl.Recordset("Pto_embarque"), RDC_grid_bl.Recordset("Pto_descarga"), RDC_grid_bl.Recordset("bl"))
        frm_bl_mod_exp.Caption = "Bl Expo: " & RDC_grid_bl.Recordset("bl")
    End If
End If
End Sub

Private Sub Grid_bl_HeadClick(ByVal ColIndex As Integer)
    Me.MousePointer = 11
    Me.RDC_grid_bl.Recordset.Sort = Grid_bl.Columns(ColIndex).DataField
    Grid_bl.Refresh
    Me.MousePointer = 0
End Sub

Private Sub Grid_bl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
   If Grid_bl.Col = -1 Then
    DoEvents
        If Trv_ImpExp = "I" Then
            PopupMenu doc_menu_aux.mnu_bl
        Else
            PopupMenu doc_menu_aux.mnu_bl_export
        End If
   End If
End If

End Sub

Private Sub Grid_bl_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo Errores:
   If Not RDC_grid_bl.Recordset.EOF Then 'And GRID_BL_FILA <> LastRow Then
        If Trv_ImpExp = "I" Then
            sql = "Select bldet_tipo_carga AS Tipo_Carga,   bldet_peso as Peso, bldet_bultos as Bultos," & _
                "       bldet_embalaje     as Embalaje,   bldet_volumen as Volumen, " & _
                "       bldet_contenedor   as Contenedor " & _
                " from docu_bldet " & _
                " Where bldet_linea       ='" & Cmb_linea.BoundText & "'" & _
                "   and bldet_buque       ='" & Trim(Trv_buque_cod) & "'" & _
                "   and bldet_viaje       ='" & Trim(Trv_viaje) & "'" & _
                "   and bldet_puer_emb    ='" & Trim$(Trv_puer_emb) & "'" & _
                "   and bldet_puer_des    ='" & RDC_grid_bl.Recordset("Pto_descarga") & "'" & _
                "   and bldet_cono_emb    ='" & RDC_grid_bl.Recordset("bl") & "'"
            
            RDC_grid_det.ConnectionString = Gstr_conex_docu
            RDC_grid_det.RecordSource = sql
            RDC_grid_det.Refresh
            Set Grid_DET.DataSource = RDC_grid_det
            Grid_DET.Refresh
            
            ' contenedores
            sql = "Select cont_contenedor AS Contenedor,   cont_tam_equip as Tipo_cont, cont_condicion as Condicion," & _
                "       cont_tara     as Tara,   cont_peso as Peso, " & _
                "       cont_sello1   as Sello1, cont_indic_llen_vac as Indic_lleno_vac " & _
                " from docu_cont " & _
                " Where cont_linea       ='" & Cmb_linea.BoundText & "'" & _
                "   and cont_buque       ='" & Trim(Trv_buque_cod) & "'" & _
                "   and cont_viaje       ='" & Trim(Trv_viaje) & "'" & _
                "   and cont_puer_emb    ='" & Trim$(Trv_puer_emb) & "'" & _
                "   and cont_puer_des    ='" & RDC_grid_bl.Recordset("Pto_descarga") & "'" & _
                "   and cont_contenedor  IN (Select bldet_contenedor " & _
                "             from docu_bldet " & _
                "             Where bldet_linea       ='" & Cmb_linea.BoundText & "'" & _
                "               and bldet_buque       ='" & Trim(Trv_buque_cod) & "'" & _
                "               and bldet_viaje       ='" & Trim(Trv_viaje) & "'" & _
                "               and bldet_puer_emb    ='" & Trim$(Trv_puer_emb) & "'" & _
                "               and bldet_puer_des    ='" & RDC_grid_bl.Recordset("Pto_descarga") & "'" & _
                "               and bldet_cono_emb    ='" & RDC_grid_bl.Recordset("bl") & "')"
            
            RDC_grid_cont.ConnectionString = Gstr_conex_docu
            RDC_grid_cont.RecordSource = sql
            RDC_grid_cont.Refresh
            Set Grid_CONT.DataSource = RDC_grid_cont
            Grid_CONT.Refresh
        Else    'Exportacion
            sql = "Select bldet_tipo_carga AS Tipo_Carga,   bldet_peso as Peso, bldet_bultos as Bultos," & _
                "       bldet_embalaje     as Embalaje,   bldet_volumen as Volumen, " & _
                "       bldet_contenedor   as Contenedor, bldet_tam_equi as Tamano_equip," & _
                "       bldet_cond_cont as Condic_Cont,   bldet_sello  as Sello  " & _
                " from docu_bldet_exp " & _
                " Where bldet_linea       ='" & Cmb_linea.BoundText & "'" & _
                "   and bldet_buque       ='" & Trim(Trv_buque_cod) & "'" & _
                "   and bldet_viaje       ='" & Trim(Trv_viaje) & "'" & _
                "   and bldet_puer_emb    ='" & RDC_grid_bl.Recordset("Pto_embarque") & "'" & _
                "   and bldet_puer_des    ='" & RDC_grid_bl.Recordset("Pto_descarga") & "'" & _
                "   and bldet_cono_emb    ='" & RDC_grid_bl.Recordset("bl") & "'"
            
            RDC_grid_det.ConnectionString = Gstr_conex_docu
            RDC_grid_det.RecordSource = sql
            RDC_grid_det.Refresh
            Set Grid_DET.DataSource = RDC_grid_det
            Grid_DET.Refresh
            
            ' contenedores
            sql = "Select cont_contenedor AS Contenedor,   cont_tam_equip as Tipo_cont, cont_condicion as Condicion," & _
                "       cont_tara     as Tara,   cont_peso as Peso, " & _
                "       cont_sello1   as Sello1, cont_indic_llen_vac as Indic_lleno_vac " & _
                " from docu_cont_exp " & _
                " Where cont_linea       ='" & Cmb_linea.BoundText & "'" & _
                "   and cont_buque       ='" & Trim(Trv_buque_cod) & "'" & _
                "   and cont_viaje       ='" & Trim(Trv_viaje) & "'" & _
                "   and cont_puer_emb    ='" & RDC_grid_bl.Recordset("Pto_embarque") & "'" & _
                "   and cont_puer_des    ='" & RDC_grid_bl.Recordset("Pto_descarga") & "'" & _
                "   and cont_contenedor  IN (Select bldet_contenedor " & _
                "             from docu_bldet_exp " & _
                "             Where bldet_linea       ='" & Cmb_linea.BoundText & "'" & _
                "               and bldet_buque       ='" & Trim(Trv_buque_cod) & "'" & _
                "               and bldet_viaje       ='" & Trim(Trv_viaje) & "'" & _
                "               and bldet_puer_emb    ='" & RDC_grid_bl.Recordset("Pto_embarque") & "'" & _
                "               and bldet_puer_des    ='" & RDC_grid_bl.Recordset("Pto_descarga") & "'" & _
                "               and bldet_cono_emb    ='" & RDC_grid_bl.Recordset("bl") & "')"
            
            RDC_grid_cont.ConnectionString = Gstr_conex_docu
            RDC_grid_cont.RecordSource = sql
            RDC_grid_cont.Refresh
            Set Grid_CONT.DataSource = RDC_grid_cont
            Grid_CONT.Refresh
'            Set Grid_CONT.DataSource = Nothing
'            Grid_CONT.Refresh
        End If
   End If
Exit Sub
Errores:
    MsgBox "Ocurrio el Error " & Err.Description
    Exit Sub
    Resume 0
End Sub

Private Sub Grid_CONT_DblClick()
Dim frm_cont_mod     As New doc_contenedor
Dim frm_cont_mod_exp As New doc_expo_contenedor
If Not RDC_grid_cont.Recordset.EOF Then
    If Trv_ImpExp = "I" Then
        frm_cont_mod.Show
        frm_cont_mod.Gbol_CONT_desde_menu = True
        frm_cont_mod.Lstr_opc = "M"
        Call frm_cont_mod.carga_contenedor(Cmb_linea.BoundText, Trv_buque_cod, Trv_viaje, RDC_grid_bl.Recordset("Pto_embarque"), RDC_grid_bl.Recordset("Pto_descarga"), RDC_grid_cont.Recordset("contenedor"))
        frm_cont_mod.Caption = "Contenedor. Impo: " & RDC_grid_cont.Recordset("contenedor")
    Else
        frm_cont_mod_exp.Show
        frm_cont_mod_exp.Gbol_CONT_desde_menu = True
        frm_cont_mod_exp.Lstr_opc = "M"
        Call frm_cont_mod_exp.carga_contenedor(Cmb_linea.BoundText, Trv_buque_cod, Trv_viaje, RDC_grid_bl.Recordset("Pto_embarque"), RDC_grid_bl.Recordset("Pto_descarga"), RDC_grid_cont.Recordset("contenedor"))
        frm_cont_mod_exp.Caption = "Contenedor. Expo: " & RDC_grid_cont.Recordset("contenedor")
    End If
End If
End Sub

Private Sub Grid_CONT_HeadClick(ByVal ColIndex As Integer)
    Me.MousePointer = 11
    Me.RDC_grid_cont.Recordset.Sort = Grid_CONT.Columns(ColIndex).DataField
    Grid_CONT.Refresh
    Me.MousePointer = 0
End Sub

Private Sub Grid_DET_HeadClick(ByVal ColIndex As Integer)
    Me.MousePointer = 11
    Me.RDC_grid_det.Recordset.Sort = Grid_DET.Columns(ColIndex).DataField
    Grid_DET.Refresh
    Me.MousePointer = 0
End Sub

Private Sub Opt_dias_Click()
    Msk_fecha_desde.Visible = False
    Msk_fecha_hasta.Visible = False
    GTM_dias.Visible = True
    Txt_nombre.Visible = False
    lbl_opcion.Caption = " Filtro por No. Días "
End Sub

Private Sub Opt_nombre_Click()
Msk_fecha_desde.Visible = False
Msk_fecha_hasta.Visible = False
GTM_dias.Visible = False
Txt_nombre.Visible = True
lbl_opcion.Caption = " Filtro por nombre "
End Sub

Private Sub Opt_rango_Click()
Msk_fecha_desde.Visible = True
Msk_fecha_hasta.Visible = True
GTM_dias.Visible = False
Txt_nombre.Visible = False
lbl_opcion.Caption = " Filtro por rango de fechas "
End Sub

Private Sub Opt_todos_Click()
Msk_fecha_desde.Visible = False
Msk_fecha_hasta.Visible = False
GTM_dias.Visible = False
Txt_nombre.Visible = False
lbl_opcion.Caption = " Filtro por todos los Buques "
End Sub


Private Sub TrV_buque_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = 2 Then
    Select Case Trv_key
        Case "BQ": PopupMenu doc_menu_aux.mnu_buque
        Case "VI": PopupMenu doc_menu_aux.mnu_viaje
        Case "PT": If Trv_ImpExp = "I" Then PopupMenu doc_menu_aux.mnu_mani Else PopupMenu doc_menu_aux.mnu_mani_expo
        'Case "IE": If Trv_ImpExp = "E" Then PopupMenu doc_menu_aux.mnu_export
    End Select
End If

End Sub

Private Sub TrV_buque_NodeClick(ByVal Node As ComctlLib.Node)
    Dim sql As String
    Dim reg As New ADODB.Recordset

    Trv_key = Mid$(Node.Key, 1, 2)
    Select Case Mid$(Node.Key, 1, 2)
        Case "BQ"
            Trv_buque_cod = Mid(Node.Key, 3)
            Trv_buque_desc = Trim(Node.Text)
        Case "VI"
            Trv_buque_cod = Mid(Node.Parent.Key, 3)
            Trv_buque_desc = Trim(Node.Parent.Text)
            Trv_viaje = Mid(Node.Key, InStr(1, Node.Key, "|") + 1)
        Case "PT":
            'Trv_ImpExp = "I"
            Trv_buque_cod = Mid(Node.Parent.Parent.Parent.Key, 3)
            Trv_buque_desc = Trim(Node.Parent.Parent.Parent.Text)
            Trv_viaje = Mid(Node.Parent.Parent.Key, InStr(1, Node.Parent.Parent.Key, "|") + 1)
            Trv_puer_emb = Right(Mid(Node.Key, InStr(3 + Len(Trim(Trv_buque_cod)) + Len(Trim(Trv_viaje)), Node.Key, "|") + 1), 8)
            Trv_ImpExp = IIf(Right(Trv_puer_emb, 3) = "IMP", "I", "E")
            Trv_puer_emb = Mid(Trv_puer_emb, 1, Len(Trv_puer_emb) - 3)
            
            If Trv_ImpExp = "I" Then
                sql = "Select blcab_cono_emb       as Bl,         blcab_puer_emb as Pto_embarque, " & _
                       "       blcab_puer_origen    as Pto_origen, blcab_puer_des as Pto_descarga, " & _
                       "       blcab_pto_dest_final as Pto_final,  blcab_consig   as Consignatario, " & _
                       "       blcab_embarcador     as Embarcador " & _
                       " from docu_blcab " & _
                       " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
                       "   and blcab_buque       ='" & Trv_buque_cod & "'" & _
                       "   and blcab_viaje       ='" & Trv_viaje & "'" & _
                       "   and blcab_puer_emb    ='" & Trv_puer_emb & "'"
            Else
                sql = "Select blcab_cono_emb       as Bl,         blcab_puer_emb as Pto_embarque, " & _
                       "       blcab_puer_origen    as Pto_origen, blcab_puer_des as Pto_descarga, " & _
                       "       blcab_pto_dest_final as Pto_final,  blcab_consig   as Consignatario, " & _
                       "       blcab_embarcador     as Embarcador " & _
                       " from docu_blcab_exp " & _
                       " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
                       "   and blcab_buque       ='" & Trv_buque_cod & "'" & _
                       "   and blcab_viaje       ='" & Trv_viaje & "'" & _
                       "   and blcab_puer_des    ='" & Trv_puer_emb & "'"
            End If
            RDC_grid_bl.ConnectionString = Gstr_conex_docu
            RDC_grid_bl.RecordSource = sql
            RDC_grid_bl.Refresh
            Set Grid_bl.DataSource = RDC_grid_bl
            Grid_bl.Refresh
            Lbl_bls.Caption = "No. Bls = " & RDC_grid_bl.Recordset.RecordCount
        Case "IE":
            If Right(Node.Key, 3) = "EXP" Then
                Trv_ImpExp = "E"
'                Trv_buque_cod = Mid(Node.Parent.Parent.Key, 3)
'                Trv_buque_desc = Trim(Node.Parent.Parent.Text)
'                Trv_viaje = Mid(Node.Parent.Key, InStr(1, Node.Parent.Key, "|") + 1)
'
'                sql = "Select blcab_cono_emb       as Bl,         blcab_puer_emb as Pto_embarque, " & _
'                       "       blcab_puer_origen    as Pto_origen, blcab_puer_des as Pto_descarga, " & _
'                       "       blcab_pto_dest_final as Pto_final,  blcab_consig   as Consignatario, " & _
'                       "       blcab_embarcador     as Embarcador " & _
'                       " from docu_blcab_exp " & _
'                       " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
'                       "   and blcab_buque       ='" & Trv_buque_cod & "'" & _
'                       "   and blcab_viaje       ='" & Trv_viaje & "'"
'
'                RDC_grid_bl.ConnectionString = Gstr_conex_docu
'                RDC_grid_bl.RecordSource = sql
'                RDC_grid_bl.Refresh
'                Set Grid_bl.DataSource = RDC_grid_bl
'                Grid_bl.Refresh
'                Lbl_bls.Caption = "No. Bls = " & RDC_grid_bl.Recordset.RecordCount
            Else
                Trv_ImpExp = "I"
            End If
    End Select
End Sub
