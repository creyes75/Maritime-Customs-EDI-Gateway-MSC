VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "GTDATE32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form doc_manifiesto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de manifiesto"
   ClientHeight    =   5040
   ClientLeft      =   3330
   ClientTop       =   4110
   ClientWidth     =   8715
   Icon            =   "Doc_manifiesto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5040
   ScaleWidth      =   8715
   Begin TabDlg.SSTab Tab_dato 
      Height          =   4425
      Left            =   75
      TabIndex        =   23
      Top             =   60
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   7805
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Listado"
      TabPicture(0)   =   "Doc_manifiesto.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "RDC_grid_dato"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Grid_dato"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Detalles"
      TabPicture(1)   =   "Doc_manifiesto.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_tipo_rol"
      Tab(1).Control(1)=   "RDC_dato"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
         Caption         =   " Filtrar por "
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
         Height          =   795
         Left            =   135
         TabIndex        =   54
         Top             =   3525
         Width           =   8295
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   5580
            TabIndex        =   58
            Top             =   135
            Width           =   1590
            Begin VB.CommandButton Cmd_busc_hlp_pto_emb 
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
               Picture         =   "Doc_manifiesto.frx":0342
               Style           =   1  'Graphical
               TabIndex        =   60
               ToolTipText     =   "Buscar"
               Top             =   255
               Width           =   300
            End
            Begin VB.TextBox Txt_busc_pto_emb 
               Height          =   285
               Left            =   120
               MaxLength       =   5
               TabIndex        =   59
               Top             =   270
               Width           =   975
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Puerto Embarque"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   61
               Top             =   0
               Width           =   1230
            End
         End
         Begin VB.CommandButton Cmd_buscar 
            Height          =   465
            Left            =   7605
            Picture         =   "Doc_manifiesto.frx":0874
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Filtrar"
            Top             =   180
            Width           =   465
         End
         Begin MSDataListLib.DataCombo cmb_busc_linea 
            Height          =   315
            Left            =   90
            TabIndex        =   56
            Top             =   390
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_busc_linea 
            Height          =   330
            Left            =   225
            Top             =   405
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
         Begin MSDataListLib.DataCombo cmb_busc_buque 
            Height          =   315
            Left            =   1530
            TabIndex        =   62
            Top             =   390
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_busc_buque 
            Height          =   330
            Left            =   2685
            Top             =   375
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
         Begin MSDataListLib.DataCombo cmb_busc_viaje 
            Height          =   315
            Left            =   3960
            TabIndex        =   63
            Top             =   375
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_busc_viaje 
            Height          =   330
            Left            =   4320
            Top             =   390
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
            Caption         =   "Buque :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   1515
            TabIndex        =   65
            Top             =   165
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Viaje :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   3960
            TabIndex        =   64
            Top             =   165
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Línea :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   135
            TabIndex        =   57
            Top             =   165
            Width           =   510
         End
      End
      Begin MSDataGridLib.DataGrid Grid_dato 
         Height          =   3045
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   5371
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
         Left            =   7050
         Top             =   3270
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
         Left            =   -74805
         Top             =   3900
         Width           =   6735
         _ExtentX        =   11880
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
         Height          =   3465
         Left            =   -74865
         TabIndex        =   24
         Top             =   390
         Width           =   8250
         _Version        =   65536
         _ExtentX        =   14552
         _ExtentY        =   6112
         _StockProps     =   14
         Caption         =   " Datos de Manifiesto "
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
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   615
            Left            =   3495
            TabIndex        =   51
            Top             =   1485
            Width           =   3870
            Begin VB.TextBox Txt_capitan 
               Height          =   285
               Left            =   150
               MaxLength       =   60
               TabIndex        =   52
               Top             =   270
               Width           =   3525
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Capitán :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   13
               Left            =   180
               TabIndex        =   53
               Top             =   30
               Width           =   630
            End
         End
         Begin VB.Frame Fra_fecha_emb 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   585
            Left            =   3345
            TabIndex        =   43
            Top             =   2820
            Width           =   1635
            Begin GTMaskDate.GTMaskDate Msk_fecha_zarpe 
               Height          =   285
               Left            =   30
               TabIndex        =   17
               Top             =   255
               Width           =   1275
               _Version        =   65537
               _ExtentX        =   2249
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
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               Caption         =   "Fech. Zarpe Pto. Ini."
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   30
               TabIndex        =   44
               Top             =   30
               Width           =   1455
            End
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   255
            TabIndex        =   41
            Top             =   2820
            Width           =   1590
            Begin VB.CommandButton Cmd_hlp_pto_inicio 
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
               Picture         =   "Doc_manifiesto.frx":0B7E
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Buscar"
               Top             =   240
               Width           =   300
            End
            Begin VB.TextBox Txt_pto_inicio 
               Height          =   285
               Left            =   105
               MaxLength       =   5
               TabIndex        =   15
               Top             =   255
               Width           =   975
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Puerto Inicio Viaje"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   42
               Top             =   15
               Width           =   1275
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   " Orden "
            ForeColor       =   &H00800000&
            Height          =   645
            Left            =   300
            TabIndex        =   38
            Top             =   1515
            Width           =   3195
            Begin VB.TextBox txt_no_orden 
               Height          =   285
               Left            =   1875
               MaxLength       =   6
               TabIndex        =   12
               Top             =   225
               Width           =   960
            End
            Begin VB.TextBox txt_año_orden 
               Height          =   285
               Left            =   585
               MaxLength       =   4
               TabIndex        =   11
               Top             =   210
               Width           =   720
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "No.:"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   10
               Left            =   1500
               TabIndex        =   40
               Top             =   270
               Width           =   300
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Año :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   9
               Left            =   90
               TabIndex        =   39
               Top             =   240
               Width           =   375
            End
         End
         Begin MSDataListLib.DataCombo Cmb_naviera 
            Height          =   315
            Left            =   3600
            TabIndex        =   14
            Top             =   2475
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_naviera 
            Height          =   330
            Left            =   5685
            Top             =   2505
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
         Begin MSDataListLib.DataCombo Cmb_linea_nav 
            Height          =   315
            Left            =   330
            TabIndex        =   13
            Top             =   2475
            Width           =   2880
            _ExtentX        =   5080
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_linea_nav 
            Height          =   330
            Left            =   2160
            Top             =   2475
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
         Begin MSMask.MaskEdBox Msk_manifiesto 
            Height          =   315
            Left            =   3660
            TabIndex        =   49
            Top             =   1110
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            Mask            =   "###-##-##-######-#"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Txt_tipo_mani 
            Height          =   285
            Left            =   3615
            MaxLength       =   2
            TabIndex        =   8
            Top             =   1725
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox Txt_digito_v 
            Height          =   285
            Left            =   5550
            MaxLength       =   1
            TabIndex        =   10
            Top             =   1725
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox Txt_no_mani 
            Height          =   285
            Left            =   4395
            MaxLength       =   6
            TabIndex        =   9
            Top             =   1725
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.TextBox Txt_año_carga 
            Height          =   285
            Left            =   2520
            MaxLength       =   4
            TabIndex        =   7
            Top             =   1725
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Frame Fra_pto_emb 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   180
            TabIndex        =   32
            Top             =   900
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
               Picture         =   "Doc_manifiesto.frx":10B0
               Style           =   1  'Graphical
               TabIndex        =   4
               ToolTipText     =   "Buscar"
               Top             =   255
               Width           =   300
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Puerto Embarque"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   33
               Top             =   0
               Width           =   1230
            End
         End
         Begin VB.Frame Fra_pto_descarga 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   1740
            TabIndex        =   30
            Top             =   900
            Width           =   1590
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
               Left            =   1140
               Picture         =   "Doc_manifiesto.frx":15E2
               Style           =   1  'Graphical
               TabIndex        =   6
               ToolTipText     =   "Buscar"
               Top             =   255
               Width           =   300
            End
            Begin VB.TextBox Txt_pto_descarga 
               Height          =   285
               Left            =   120
               MaxLength       =   5
               TabIndex        =   5
               Top             =   270
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Puerto Descarga"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   31
               Top             =   15
               Width           =   1200
            End
         End
         Begin VB.TextBox txt_buque 
            Height          =   285
            Left            =   3300
            MaxLength       =   5
            TabIndex        =   25
            Top             =   225
            Visible         =   0   'False
            Width           =   840
         End
         Begin MSDataListLib.DataCombo Cmb_buque 
            Height          =   315
            Left            =   2400
            TabIndex        =   1
            Top             =   525
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_buque 
            Height          =   330
            Left            =   4005
            Top             =   510
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
            Left            =   285
            TabIndex        =   0
            Top             =   525
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_linea 
            Height          =   330
            Left            =   1140
            Top             =   525
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
         Begin MSDataListLib.DataCombo Cmb_viaje 
            Height          =   315
            Left            =   5580
            TabIndex        =   2
            Top             =   510
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_viaje 
            Height          =   330
            Left            =   5940
            Top             =   525
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
         Begin MSDataListLib.DataCombo cmb_adua_ing 
            Height          =   315
            Left            =   300
            TabIndex        =   47
            Top             =   1725
            Visible         =   0   'False
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_adua_ing 
            Height          =   330
            Left            =   1260
            Top             =   1740
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
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Ag. Naviera de Carga"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   360
            TabIndex        =   46
            Top             =   2235
            Width           =   1530
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Ag. Oper. de Carga (Linea)"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   3615
            TabIndex        =   45
            Top             =   2265
            Width           =   1890
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "No. Manifiesto"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   3660
            TabIndex        =   50
            Top             =   900
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo manif"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   3570
            TabIndex        =   48
            Top             =   1500
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Dígito Ver."
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   5505
            TabIndex        =   37
            Top             =   1515
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "No. Manifiesto"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   4365
            TabIndex        =   36
            Top             =   1515
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Año de Carga:"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   2475
            TabIndex        =   35
            Top             =   1515
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Aduana Ingreso :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   330
            TabIndex        =   34
            Top             =   1515
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Línea :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   285
            TabIndex        =   29
            Top             =   285
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Viaje :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   5580
            TabIndex        =   27
            Top             =   315
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Buque :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   2385
            TabIndex        =   26
            Top             =   300
            Width           =   555
         End
      End
   End
   Begin Threed.SSCommand cmd_agregar 
      Height          =   375
      Left            =   3915
      TabIndex        =   18
      Top             =   4560
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
      Left            =   5100
      TabIndex        =   19
      Top             =   4560
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
      Left            =   6285
      TabIndex        =   20
      Top             =   4560
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
      Left            =   7470
      TabIndex        =   21
      Top             =   4560
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
      Left            =   2730
      TabIndex        =   22
      Top             =   4560
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
      Left            =   2370
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   66
      Top             =   4590
      Width           =   1200
   End
End
Attribute VB_Name = "doc_manifiesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lstr_opc As String * 1 ' para saber la transaccion que se esta ejecutando Ingreso, modificacion, eliminacion
'Dim Lstr_tipo As String * 1 ' para saber si se trata de Importacion o exportacion

Public Sub habilita(ByVal opc As Boolean)
   If opc Then
      If Lstr_opc = "M" Then
         Cmb_linea.Enabled = False
         txt_buque.Enabled = False
         Cmb_buque.Enabled = False
         Cmb_viaje.Enabled = False
         Txt_pto_emb.Enabled = False
         Txt_pto_descarga.Enabled = False
         Cmd_hlp_pto_emb.Enabled = False
         Cmd_hlp_pto_descarga.Enabled = False
      Else
         Cmb_linea.Enabled = opc
         txt_buque.Enabled = opc
         Cmb_buque.Enabled = opc
         Cmb_viaje.Enabled = opc
         Txt_pto_emb.Enabled = opc
         Txt_pto_descarga.Enabled = opc
         Cmd_hlp_pto_emb.Enabled = opc
         Cmd_hlp_pto_descarga.Enabled = opc
         
         cmb_adua_ing.Enabled = opc
         Txt_año_carga.Enabled = opc
         Txt_tipo_mani.Enabled = opc
         'Opt_tipo_mani1.Enabled = opc
         'Opt_tipo_mani2.Enabled = opc
         Txt_no_mani.Enabled = opc
         Txt_digito_v.Enabled = opc
         Cmb_linea_nav.Enabled = opc
         Cmb_naviera.Enabled = opc
         txt_año_orden.Enabled = opc
         txt_no_orden.Enabled = opc
         Txt_capitan.Enabled = opc
         Txt_pto_inicio.Enabled = opc
         Cmb_linea_nav.Enabled = opc
         Cmb_naviera.Enabled = opc
         Msk_fecha_zarpe.Enabled = opc
      End If
   Else
      Cmb_linea.Enabled = True
      txt_buque.Enabled = True
      Cmb_buque.Enabled = True
      Cmb_viaje.Enabled = True
      Txt_pto_emb.Enabled = True
      Txt_pto_descarga.Enabled = True
      Cmd_hlp_pto_emb.Enabled = True
      Cmd_hlp_pto_descarga.Enabled = True
      
      cmb_adua_ing.Enabled = True
         Txt_año_carga.Enabled = True
         Txt_tipo_mani.Enabled = True
         'Opt_tipo_mani1.Enabled = True
         'Opt_tipo_mani2.Enabled = True
         Txt_no_mani.Enabled = True
         Txt_digito_v.Enabled = True
         txt_año_orden.Enabled = True
         txt_no_orden.Enabled = True
         Txt_capitan.Enabled = True
         Txt_pto_inicio.Enabled = True
         Cmb_linea_nav.Enabled = True
         Cmb_naviera.Enabled = True
         Msk_fecha_zarpe.Enabled = True
   End If
   
End Sub

Private Sub limpia()
   Cmb_linea.BoundText = ""
   txt_buque.Text = ""
   Cmb_buque.BoundText = ""
   Cmb_viaje.Text = ""
   Txt_pto_emb.Text = ""
   Txt_pto_descarga.Text = ""
      
   cmb_adua_ing.BoundText = ""
   Txt_año_carga.Text = ""
   
   Txt_no_mani.Text = ""
   Txt_digito_v.Text = ""
   txt_año_orden.Text = ""
   txt_no_orden.Text = ""
   Txt_capitan.Text = ""
   Txt_pto_inicio.Text = ""
   Cmb_linea_nav.BoundText = ""
   Cmb_naviera.BoundText = ""
   Msk_fecha_zarpe.Text = ""
   'Txt_part.Text = ""
   'Txt_dias_vaca.Text = ""
   'Chk_apli.Value = 0
End Sub

Private Function valida_Datos() As Boolean
On Error GoTo ERROR_VALIDA
Dim reg As New ADODB.Recordset

valida_Datos = True
   'If Trim$(txt_Buque.Text) = "" Then
   If Cmb_linea.BoundText = "" Then
      MsgBox "Ingrese el Valor de Linea", vbInformation, App.Title
      valida_Datos = False
      'If txt_Buque.Enabled Then txt_Buque.SetFocus
      If Cmb_linea.Enabled Then Cmb_linea.SetFocus
      Exit Function
   End If
   
   If Cmb_buque.BoundText = "" Then
      MsgBox "Ingrese el Valor de Buque", vbInformation, App.Title
      valida_Datos = False
      'If txt_Buque.Enabled Then txt_Buque.SetFocus
      If Cmb_buque.Enabled Then Cmb_buque.SetFocus
      Exit Function
   End If
   If Trim$(Cmb_viaje.BoundText) = "" Then
      MsgBox "Ingrese el Código de Viaje", vbInformation, App.Title
      valida_Datos = False
      If Cmb_viaje.Enabled Then Cmb_viaje.SetFocus
      Exit Function
   End If
   If Txt_pto_emb.Text = "" Then
      MsgBox "Ingrese el Valor de Puerto de Embarque", vbInformation, App.Title
      valida_Datos = False
      If Txt_pto_emb.Enabled Then Txt_pto_emb.SetFocus
      Exit Function
   Else
      If Codigo_equivalente("docu_puerto", Txt_pto_emb.Text, "puer_codigo", "puer_codigo") = "NO" Then
         valida_Datos = False
         If Txt_pto_emb.Enabled Then Txt_pto_emb.SetFocus
         MsgBox "El código de Puerto de Embarque no existe, por favor verifique", vbInformation
         Exit Function
      End If
   End If
   
   If Txt_pto_descarga.Text = "" Then
      MsgBox "Ingrese el Valor de Puerto de Descarga", vbInformation, App.Title
      valida_Datos = False
      If Txt_pto_descarga.Enabled Then Txt_pto_descarga.SetFocus
      Exit Function
   Else
      If Codigo_equivalente("docu_puerto", Txt_pto_descarga.Text, "puer_codigo", "puer_codigo") = "NO" Then
         valida_Datos = False
         If Txt_pto_descarga.Enabled Then Txt_pto_descarga.SetFocus
         MsgBox "El código de Puerto de Descarga no existe, por favor verifique", vbInformation
         Exit Function
      End If
   End If
      
   If Txt_pto_inicio.Text = "" Then
      MsgBox "Ingrese el Valor de Puerto de Inicio", vbInformation, App.Title
      valida_Datos = False
      If Txt_pto_inicio.Enabled Then Txt_pto_inicio.SetFocus
      Exit Function
   Else
      If Codigo_equivalente("docu_puerto", Txt_pto_inicio.Text, "puer_codigo", "puer_codigo") = "NO" Then
         valida_Datos = False
         If Txt_pto_inicio.Enabled Then Txt_pto_inicio.SetFocus
         MsgBox "El código de Puerto de Inicio no existe, por favor verifique", vbInformation
         Exit Function
      End If
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
   
   'If Opt_tipo_mani1.Value Then tipo_manifiesto = "01" Else tipo_manifiesto = "02"
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'TL aumento los dos digitos del año completo en los dos campos de año
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   If Trim$(txt_año_orden.Text) <> "" Then
      txt_año_orden.Text = "20" & Trim$(txt_año_orden.Text)
   End If
   If Len(Txt_año_carga.Text) = 2 Then
      Txt_año_carga.Text = "20" & Trim$(Txt_año_carga.Text)
   End If
'   If Lstr_tipo = "I" Then tabla = "docu_manifiesto" Else tabla = "docu_manifiesto_exp"
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   If Lstr_opc = "A" Then
     sql = "Select * from docu_manifiesto Where mani_linea = '" & Cmb_linea.BoundText & "' and mani_buque = '" & Cmb_buque.BoundText & "' and mani_viaje = '" & Cmb_viaje.BoundText & "' and mani_puer_emb ='" & Txt_pto_emb.Text & "' and mani_puer_des ='" & Txt_pto_descarga.Text & "'"
     reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
     If reg.EOF Then
        sql = "INSERT INTO docu_manifiesto (mani_linea,         mani_buque,         mani_viaje, " & _
                                           "mani_puer_emb,      mani_puer_des,      mani_adua_ing," & _
                                           "mani_ano_car,       mani_tip_mani,      mani_mani_car," & _
                                           "mani_digi_ver,      mani_ano_ord,       mani_num_ord, " & _
                                           "mani_capitan,       mani_emp_prop,      mani_emp_nave, " & _
                                           "mani_puer_ini,      mani_fecha_zarp_ini)" & _
              " VALUES ('" & Trim$(Cmb_linea.BoundText) & "',    '" & Trim$(Cmb_buque.BoundText) & "','" & Trim$(Cmb_viaje.BoundText) & "','" & _
                             Txt_pto_emb.Text & "',              '" & Txt_pto_descarga.Text & "',      '" & cmb_adua_ing.BoundText & "', '" & _
                             Txt_año_carga.Text & "',            '" & Txt_tipo_mani.Text & "',            '" & Txt_no_mani.Text & "', '" & _
                             Txt_digito_v.Text & "',             '" & txt_año_orden.Text & "',         '" & txt_no_orden.Text & "', '" & _
                             Txt_capitan.Text & "',              '" & Cmb_linea_nav.BoundText & "',    '" & Cmb_naviera.BoundText & "' ,'" & _
                             Txt_pto_inicio.Text & "',           '" & Format(Msk_fecha_zarpe.Text, GSTR_FORMATO_FECHA) & "')"
     Else
        MsgBox "Este Manifiesto ya existe, entre por la opción modificación", vbExclamation, App.Title
        Exit Sub
     End If
   ElseIf Lstr_opc = "M" Then
      sql = "UPDATE docu_manifiesto SET mani_adua_ing = '" & cmb_adua_ing.BoundText & "', " & _
                                          "mani_ano_car   = '" & Txt_año_carga.Text & "'," & _
                                          "mani_tip_mani  = '" & Txt_tipo_mani.Text & "', " & _
                                          "mani_mani_car  = '" & Txt_no_mani.Text & "', " & _
                                          "mani_digi_ver  = '" & Txt_digito_v.Text & "', " & _
                                          "mani_ano_ord   = '" & txt_año_orden.Text & "', " & _
                                          "mani_num_ord   = '" & txt_no_orden.Text & "', " & _
                                          "mani_capitan   ='" & Txt_capitan.Text & "', " & _
                                          "mani_emp_prop  ='" & Cmb_linea_nav.BoundText & "', " & _
                                          "mani_emp_nave  ='" & Cmb_naviera.BoundText & "', " & _
                                          "mani_puer_ini  ='" & Txt_pto_inicio.Text & "', " & _
                                          "mani_fecha_zarp_ini='" & Format(Msk_fecha_zarpe.Text, GSTR_FORMATO_FECHA) & "' " & _
                                          "Where mani_linea = '" & Cmb_linea.BoundText & "' and mani_buque = '" & Cmb_buque.BoundText & "' and mani_viaje = '" & Cmb_viaje.BoundText & "' and mani_puer_emb ='" & Txt_pto_emb.Text & "' and mani_puer_des ='" & Txt_pto_descarga.Text & "'"

   ElseIf Lstr_opc = "E" Then
      resp = MsgBox("Esta seguro que desea eliminar este Manifiesto", vbQuestion + vbYesNo, App.Title)
      If resp Then resp = Elimina_manifiesto(Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, Txt_pto_emb.Text, Txt_pto_descarga.Text)
   End If
   If Lstr_opc <> "E" Then
      Gcon_docu.Execute sql
      MsgBox "Transacción Realizada con éxito", vbInformation, App.Title
   Else
      If resp Then MsgBox "Transacción Realizada con éxito", vbInformation, App.Title
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

Private Sub cmb_adua_ing_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmb_busc_buque_Change()
   Dim sql As String
   'If cmb_busc_buque.MatchedWithList Then
        sql = "Select * from docu_buque_viaje " & _
              " Where buqv_linea ='" & cmb_busc_linea.BoundText & "'" & _
              "   And buqv_buque ='" & cmb_busc_buque.BoundText & "'" & _
              "   And buqv_status = 'A'"
        
        carga_datacombo data_busc_viaje, cmb_busc_viaje, sql, "buqv_viaje", "buqv_viaje"
   'Else
   '
   'End If
End Sub

Private Sub cmb_busc_linea_Change()
If cmb_busc_linea.MatchedWithList Then
   carga_datacombo data_busc_buque, cmb_busc_buque, "Select * from docu_buque Where buqu_linea = '" & cmb_busc_linea.BoundText & "' and buqu_status ='A'", "buqu_codigo", "buqu_descripcion"
End If
End Sub


Private Sub cmb_linea_Change()
If Cmb_linea.MatchedWithList Then
   data_buque.ConnectionString = Gstr_conex_docu
   data_buque.RecordSource = "Select * from docu_buque where buqu_linea = '" & Cmb_linea.BoundText & "' and buqu_status ='A'"
   data_buque.Refresh
   Cmb_buque.BoundColumn = "buqu_codigo"
   Cmb_buque.ListField = "buqu_descripcion"
   Set Cmb_buque.RowSource = data_buque
   
   carga_datacombo data_linea_nav, Cmb_linea_nav, "Select * from docu_emp_anav Where emp_anav_linea ='" & Cmb_linea.BoundText & "' and emp_anav_status ='A'", "emp_anav_codigo", "emp_anav_descripcion"
   carga_datacombo data_naviera, Cmb_naviera, "Select * from docu_emp_resp Where emp_resp_linea ='" & Cmb_linea.BoundText & "' and emp_resp_status ='A'", "emp_resp_codigo", "emp_resp_descripcion"
   carga_datacombo data_adua_ing, cmb_adua_ing, "Select * from docu_adua_ing Where adua_ing_linea ='" & Cmb_linea.BoundText & "'", "adua_ing_codigo", "adua_ing_descripcion"
   
   carga_parametros Cmb_linea.BoundText
   If Lstr_opc = "A" Then
      Txt_pto_descarga.Text = Gstr_param.pto_descarga
      cmb_adua_ing.BoundText = Gstr_param.adua_ing
      Txt_tipo_mani.Text = Gstr_param.tipo_manif
      Txt_año_carga.Text = Year(Now)
      'If Gstr_param.tipo_manif = "00" Then Opt_tipo_mani1.Value = True
      Cmb_linea_nav.BoundText = Gstr_param.Linea_Nav
      Cmb_naviera.BoundText = Gstr_param.naviera
   End If
End If
End Sub

Private Sub Cmb_linea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_linea_nav_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub Cmb_naviera_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_buque_Change()
   Dim sql As String
   sql = "Select * from docu_buque_viaje " & _
         " Where buqv_linea ='" & Cmb_linea.BoundText & "'" & _
         "   And buqv_buque ='" & Cmb_buque.BoundText & "'" & _
         "   And buqv_status = 'A'"
   
   carga_datacombo data_viaje, Cmb_viaje, sql, "buqv_viaje", "buqv_viaje"

End Sub

Private Sub Cmb_buque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_viaje_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Public Sub cmd_agregar_Click()
Select Case cmd_agregar.Caption
  Case "&Agregar"
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
      Cmb_buque.SetFocus
      'cmb_Buque.lis = "" 'Nuevo_Tipo_Rol
      Cmb_viaje.BoundText = ""
      Txt_pto_emb.Text = ""
      Txt_pto_inicio.Text = ""
      Txt_capitan.Text = ""
      Txt_pto_descarga.Text = ""
      Msk_fecha_zarpe.Text = ""
      Tab_dato.Tab = 1
      Tab_dato.TabEnabled(0) = False
  Case "&Grabar"
      If Not valida_Datos Then Exit Sub
      Graba_trol
      If Lstr_opc = "M" Or Lstr_opc = "E" Then
         Lstr_opc = "C"
         cmd_salir_Click
      Else
         txt_buque.Text = "" 'Nuevo_Tipo_Rol
      End If
      'If txt_Buque.Enabled Then txt_Buque.SetFocus
      If Cmb_buque.Enabled Then Cmb_buque.SetFocus
      RDC_dato.Caption = ""
      If Lstr_opc = "M" Then
         Tab_dato.Tab = 0
         Tab_dato.TabEnabled(0) = True
      End If
      'RDC_grid_dato.Refresh
      'Grid_dato.Refresh
      'Lstr_opc = "C"
End Select
End Sub

Private Sub Cmd_busc_hlp_pto_emb_Click()
   Txt_busc_pto_emb.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & cmb_busc_linea.BoundText & "' and puer_status ='A' ", "Puerto de Embarque(Búsqueda)")
End Sub

Private Sub Cmd_buscar_Click()
Dim sql As String
sql = ""
If cmb_busc_linea.MatchedWithList Then sql = sql & " linea='" & cmb_busc_linea.BoundText & "' and"
If cmb_busc_buque.MatchedWithList Then sql = sql & " Buque='" & cmb_busc_buque.BoundText & "' and"
If cmb_busc_viaje.MatchedWithList Then sql = sql & " Viaje='" & cmb_busc_viaje.BoundText & "' and"
If Me.Txt_busc_pto_emb.Text <> "" Then sql = sql & " pto_embarque like '" & Txt_busc_pto_emb.Text & "*'"
If Right(sql, 3) = "and" Then sql = Mid(sql, 1, Len(sql) - 3)

RDC_grid_dato.Recordset.Filter = Trim$(sql)

End Sub

Public Sub cmd_eliminar_Click()
    Lstr_opc = "E"
    cmd_agregar.Caption = "&Grabar"
    cmd_salir.Caption = "&Cancelar"
    cmd_eliminar.Enabled = False
    Cmd_modificar.Enabled = False
    RDC_dato.Enabled = False
    habilita True
    Cmb_buque.Enabled = False
    Cmb_viaje.Enabled = False
    Tab_dato.Tab = 1
    Tab_dato.TabEnabled(0) = False
End Sub

Private Sub Cmd_hlp_pto_descarga_Click()
   Txt_pto_descarga.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Descarga")
End Sub

Private Sub Cmd_hlp_pto_emb_Click()
   Txt_pto_emb.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Embarque")
End Sub

Private Sub Cmd_hlp_pto_inicio_Click()
   Txt_pto_inicio.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Inicio de Viaje")
End Sub

Private Sub Cmd_Imprimir_Click()
On Error GoTo Errores
'    CReport.ReportFileName = Gstr_ruta_reportes & "cont030.rpt"
'    CReport.Connect = Gstr_conex_docu
'    CReport.SelectionFormula = "{docu_buque_viaje.vess_status}='A'"
'    CReport.WindowTitle = App.Title
    'CReport.Action = 1
    Exit Sub
Errores:
    MsgBox "Ocurrio el error " & Err.Description, vbInformation, App.Title

End Sub

Public Sub cmd_modificar_Click()
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
End Select
End Sub

Private Sub Form_Load()
   Dim res As Boolean
   Me.Height = 5415
   Me.Width = 8805
   Centra_Pantalla Me
   Lstr_opc = "C"
   
   'RDC_dato.ConnectionString = Gstr_conex_docu
   
   RDC_grid_dato.ConnectionString = Gstr_conex_docu
   RDC_grid_dato.RecordSource = "Select mani_linea as Linea, mani_buque as Buque, mani_viaje as Viaje, " & _
                                "mani_puer_emb as Pto_embarque,     mani_puer_des as Pto_descarga,  mani_adua_ing as Aduana_ing, " & _
                                "mani_ano_car as Anio_carga,        mani_tip_mani as Tipo_manif,    mani_mani_car as Manifiesto, " & _
                                "mani_digi_ver as Dig_verific,      mani_ano_ord as Anio_Orden,     mani_num_ord as Num_Orden, " & _
                                "mani_tot_doc as Tot_docum,         mani_tot_peso as Tot_peso, " & _
                                "mani_tot_bult as Tot_bultos,       mani_tot_vol as Tot_vol,        mani_tot_con as Tot_cont, " & _
                                "mani_emp_prop as Linea_naviera, " & _
                                "mani_emp_nave as Naviera,          mani_capitan as Capitan,        mani_puer_ini as Pto_inicio, " & _
                                "mani_fecha_zarp_ini as Fecha_zarpe " & _
                                "from docu_manifiesto "
   RDC_grid_dato.Refresh
   Set Grid_dato.DataSource = RDC_grid_dato
   Grid_dato.Refresh
   
   data_linea.ConnectionString = Gstr_conex_docu
   data_linea.RecordSource = "Select * from docu_linea_nav where linea_status ='A'"
   data_linea.Refresh
   Cmb_linea.BoundColumn = "linea_codigo"
   Cmb_linea.ListField = "linea_descripcion"
   Set Cmb_linea.RowSource = data_linea
   
   carga_datacombo data_busc_linea, cmb_busc_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_codigo"
   
End Sub

Private Sub Grid_dato_DblClick()
   If Not RDC_grid_dato.Recordset.EOF Then
      Cmb_linea.BoundText = Trim$(RDC_grid_dato.Recordset.Fields("linea"))
      Cmb_buque.BoundText = Trim$(RDC_grid_dato.Recordset.Fields("Buque"))
      txt_buque.Text = Trim$(RDC_grid_dato.Recordset.Fields("Buque"))
      Cmb_viaje.BoundText = Trim$(RDC_grid_dato.Recordset.Fields("Viaje"))
      
      
         Txt_pto_emb.Text = Trim$(RDC_grid_dato.Recordset.Fields("Pto_embarque"))
         Txt_pto_descarga.Text = Trim$(RDC_grid_dato.Recordset.Fields("Pto_descarga"))
         
         cmb_adua_ing.BoundText = IIf(Not IsNull(RDC_grid_dato.Recordset.Fields("Aduana_ing")), RDC_grid_dato.Recordset.Fields("Aduana_ing"), " ")
         
         Txt_año_carga.Text = IIf(Not IsNull((RDC_grid_dato.Recordset.Fields("anio_carga"))), (RDC_grid_dato.Recordset.Fields("anio_carga")), "")
         Txt_año_carga.Text = Right(Txt_año_carga.Text, 2)
         
         Txt_tipo_mani.Text = IIf(Not IsNull(RDC_grid_dato.Recordset.Fields("Tipo_manif")), RDC_grid_dato.Recordset.Fields("Tipo_manif"), "")
         Txt_no_mani.Text = IIf(Not IsNull(RDC_grid_dato.Recordset.Fields("Manifiesto")), RDC_grid_dato.Recordset.Fields("Manifiesto"), "")
         
         Txt_digito_v.Text = IIf(Not IsNull(RDC_grid_dato.Recordset.Fields("Dig_verific")), RDC_grid_dato.Recordset.Fields("Dig_verific"), "")
         
         Msk_manifiesto.Text = cmb_adua_ing.BoundText & Txt_año_carga.Text & Txt_tipo_mani.Text & Txt_no_mani.Text & Txt_digito_v.Text
         
         Cmb_linea_nav.BoundText = IIf(Not IsNull(RDC_grid_dato.Recordset.Fields("Linea_naviera")), RDC_grid_dato.Recordset.Fields("Linea_naviera"), "")
         Cmb_naviera.BoundText = IIf(Not IsNull(RDC_grid_dato.Recordset.Fields("Naviera")), RDC_grid_dato.Recordset.Fields("Naviera"), "")
         
         txt_año_orden.Text = IIf(Not IsNull(RDC_grid_dato.Recordset.Fields("Anio_orden")), RDC_grid_dato.Recordset.Fields("Anio_orden"), "")
         txt_no_orden.Text = IIf(Not IsNull(RDC_grid_dato.Recordset.Fields("Num_orden")), RDC_grid_dato.Recordset.Fields("Num_orden"), "")
              
      Txt_capitan.Text = IIf(Not IsNull(RDC_grid_dato.Recordset.Fields("Capitan")), RDC_grid_dato.Recordset.Fields("Capitan"), "")
      Txt_pto_inicio.Text = IIf(Not IsNull(RDC_grid_dato.Recordset.Fields("Pto_Inicio")), RDC_grid_dato.Recordset.Fields("Pto_Inicio"), "")
      Msk_fecha_zarpe.Text = IIf(Not IsNull(RDC_grid_dato.Recordset.Fields("Fecha_zarpe")), RDC_grid_dato.Recordset.Fields("Fecha_zarpe"), "")
      
      
'      If (RDC_grid_dato.Recordset.Fields("Tipo_manif")) = "01" Then
'         Opt_tipo_mani1.Value = True
'      ElseIf (RDC_grid_dato.Recordset.Fields("Tipo_manif")) = "02" Then
'         Opt_tipo_mani2.Value = True
'      End If
      
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
Me.MousePointer = 0
End Sub

Private Sub rdc_dato_Reposition()
    If Not RDC_dato.Recordset.EOF Then
         Cmb_linea.BoundText = Trim$(RDC_dato.Recordset.Fields("linea"))
         Cmb_buque.BoundText = Trim$(RDC_dato.Recordset.Fields("Buque"))
         txt_buque.Text = Trim$(RDC_dato.Recordset.Fields("Buque"))
         Cmb_viaje.BoundText = Trim$(RDC_dato.Recordset.Fields("Viaje"))
         Txt_pto_emb.Text = Trim$(RDC_dato.Recordset.Fields("Pto_embarque"))
         Txt_pto_descarga.Text = Trim$(RDC_dato.Recordset.Fields("Pto_descarga"))
         
         cmb_adua_ing.BoundText = IIf(Not IsNull(RDC_dato.Recordset.Fields("Aduana_ing")), RDC_dato.Recordset.Fields("Aduana_ing"), " ")
         
         Txt_año_carga.Text = IIf(Not IsNull((RDC_dato.Recordset.Fields("anio_carga"))), (RDC_dato.Recordset.Fields("anio_carga")), "")
         Txt_año_carga.Text = Right(Txt_año_carga.Text, 2)
         
         Txt_tipo_mani.Text = IIf(Not IsNull(RDC_dato.Recordset.Fields("Tipo_manif")), RDC_dato.Recordset.Fields("Tipo_manif"), "")
         Txt_no_mani.Text = IIf(Not IsNull(RDC_dato.Recordset.Fields("Manifiesto")), RDC_dato.Recordset.Fields("Manifiesto"), "")
         
         Txt_digito_v.Text = IIf(Not IsNull(RDC_dato.Recordset.Fields("Dig_verific")), RDC_dato.Recordset.Fields("Dig_verific"), "")
         
         Msk_manifiesto.Text = cmb_adua_ing.BoundText & Txt_año_carga.Text & Txt_tipo_mani.Text & Txt_no_mani.Text & Txt_digito_v.Text
         DoEvents
         
         Cmb_linea_nav.BoundText = IIf(Not IsNull(RDC_dato.Recordset.Fields("Linea_naviera")), RDC_dato.Recordset.Fields("Linea_naviera"), "")
         Cmb_naviera.BoundText = IIf(Not IsNull(RDC_dato.Recordset.Fields("Naviera")), RDC_dato.Recordset.Fields("Naviera"), "")
         
         txt_año_orden.Text = IIf(Not IsNull(RDC_dato.Recordset.Fields("Anio_orden")), RDC_dato.Recordset.Fields("Anio_orden"), "")
         txt_no_orden.Text = IIf(Not IsNull(RDC_dato.Recordset.Fields("Num_orden")), RDC_dato.Recordset.Fields("Num_orden"), "")
         
         Txt_capitan.Text = IIf(Not IsNull(RDC_dato.Recordset.Fields("Capitan")), RDC_dato.Recordset.Fields("Capitan"), "")
         Txt_pto_inicio.Text = IIf(Not IsNull(RDC_dato.Recordset.Fields("Pto_Inicio")), RDC_dato.Recordset.Fields("Pto_Inicio"), "")
         Msk_fecha_zarpe.Text = IIf(Not IsNull(RDC_dato.Recordset.Fields("Fecha_zarpe")), RDC_dato.Recordset.Fields("Fecha_zarpe"), "")
    End If
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

Private Sub Msk_fecha_zarpe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Msk_manifiesto_LostFocus()
If Len(Trim$(Msk_manifiesto.Text)) = 14 Then
   cmb_adua_ing.BoundText = Mid(Msk_manifiesto.Text, 1, 3)
   Me.Txt_año_carga.Text = Mid(Msk_manifiesto.Text, 4, 2)
   Me.Txt_tipo_mani.Text = Mid(Msk_manifiesto.Text, 6, 2)
   Me.Txt_no_mani.Text = Mid(Msk_manifiesto.Text, 8, 6)
   Me.Txt_digito_v.Text = Mid(Msk_manifiesto.Text, 14, 1)
   
Else
    MsgBox "El No. de manifiesto no tiene el tamaño adecuado", vbInformation
End If
End Sub

Private Sub Txt_año_carga_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txt_año_orden_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_capitan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txt_no_orden_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_pto_descarga_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txt_pto_emb_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Public Sub Txt_pto_emb_LostFocus()
    If Txt_pto_emb.Text <> "" Then
        If Codigo_equivalente("docu_puerto", Txt_pto_emb.Text, "puer_codigo", "puer_codigo") = "NO" Then
           If Txt_pto_emb.Enabled Then Txt_pto_emb.SetFocus
           MsgBox "Este codigo no existe, por favor Verifique", vbInformation
           Exit Sub
        End If
        If Lstr_opc = "A" Then Txt_pto_inicio.Text = Txt_pto_emb.Text
    End If
    
    ' esta parte esta solo para que sea llamada desde la pantalla de bl_status con click derecho
    Dim reg     As New ADODB.Recordset
    Dim sql     As String
    Dim i       As Integer
    
    If (Not Cmb_linea.MatchedWithList Or Not Me.Cmb_buque.MatchedWithList Or _
       Not Cmb_viaje.MatchedWithList Or Trim$(Txt_pto_emb.Text) = "") And Lstr_opc = "C" Then
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
    
    sql = "Select mani_linea as Linea, mani_buque as Buque, mani_viaje as Viaje, " & _
          "mani_puer_emb as Pto_embarque,     mani_puer_des as Pto_descarga,  mani_adua_ing as Aduana_ing, " & _
          "mani_ano_car as Anio_carga,        mani_tip_mani as Tipo_manif,    mani_mani_car as Manifiesto, " & _
          "mani_digi_ver as Dig_verific,      mani_ano_ord as Anio_Orden,     mani_num_ord as Num_Orden, " & _
          "mani_tot_doc as Tot_docum,         mani_tot_peso as Tot_peso, " & _
          "mani_tot_bult as Tot_bultos,       mani_tot_vol as Tot_vol,        mani_tot_con as Tot_cont, " & _
          "mani_emp_prop as Linea_naviera,    mani_emp_nave as Naviera,       mani_capitan as Capitan,  " & _
          "mani_puer_ini as Pto_inicio,       mani_fecha_zarp_ini as Fecha_zarpe " & _
          "from docu_manifiesto " & _
          " Where mani_linea    ='" & Cmb_linea.BoundText & "'" & _
          "   And mani_buque    ='" & Cmb_buque.BoundText & "'" & _
          "   And mani_viaje    ='" & Cmb_viaje.BoundText & "'" & _
          "   And mani_puer_emb ='" & Txt_pto_emb.Text & "'" & _
          "   And mani_puer_des ='" & Txt_pto_descarga.Text & "'"
                                    
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
       If Txt_pto_emb.Enabled Then Txt_pto_emb.SetFocus
       RDC_dato.Enabled = False
       Cmd_modificar.Enabled = False
       cmd_eliminar.Enabled = False
    End If
End Sub

Private Sub Txt_pto_inicio_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_tipo_mani_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
