VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{667E8C40-F9B5-11CF-90AB-444553540000}#1.0#0"; "GTNUM32.OCX"
Begin VB.Form doc_contenedor 
   Caption         =   "Mantenimiento de contenedores"
   ClientHeight    =   7125
   ClientLeft      =   3345
   ClientTop       =   4125
   ClientWidth     =   10905
   Icon            =   "Doc_contenedor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   10905
   Begin VB.Frame Frame1 
      Height          =   6660
      Left            =   15
      TabIndex        =   37
      Top             =   30
      Width           =   10800
      Begin VB.Frame Frame13 
         Height          =   1470
         Left            =   105
         TabIndex        =   76
         Top             =   150
         Width           =   10530
         Begin VB.Frame Frame14 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   75
            TabIndex        =   86
            Top             =   825
            Width           =   1500
            Begin MSDataListLib.DataCombo Cmb_pto_emb 
               Height          =   315
               Left            =   135
               TabIndex        =   3
               Top             =   225
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "3"
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin MSAdodcLib.Adodc data_pto_emb 
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
               Caption         =   "Puerto Embarque"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   87
               Top             =   -15
               Width           =   1230
            End
         End
         Begin VB.TextBox Txt_contenedor 
            Height          =   285
            Left            =   4230
            MaxLength       =   11
            TabIndex        =   6
            Top             =   1065
            Width           =   1440
         End
         Begin VB.TextBox txt_vessel 
            Height          =   285
            Left            =   3210
            MaxLength       =   5
            TabIndex        =   79
            Top             =   180
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Frame Fra_pto_descarga 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   2145
            TabIndex        =   77
            Top             =   810
            Width           =   1755
            Begin VB.TextBox Txt_pto_descarga 
               Height          =   285
               Left            =   195
               MaxLength       =   5
               TabIndex        =   4
               Top             =   270
               Width           =   975
            End
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
               Left            =   1215
               Picture         =   "Doc_contenedor.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   5
               ToolTipText     =   "Buscar"
               Top             =   255
               Width           =   300
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Puerto Descarga"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   2
               Left            =   195
               TabIndex        =   78
               Top             =   15
               Width           =   1260
            End
         End
         Begin MSDataListLib.DataCombo Cmb_buque 
            Height          =   315
            Left            =   2310
            TabIndex        =   1
            Top             =   465
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_vessel 
            Height          =   330
            Left            =   3960
            Top             =   450
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
            Left            =   210
            TabIndex        =   0
            Top             =   465
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_linea 
            Height          =   330
            Left            =   1110
            Top             =   450
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
            Left            =   5415
            TabIndex        =   2
            Top             =   465
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
            Left            =   5775
            Top             =   435
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
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "Contenedor"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   4230
            TabIndex        =   83
            Top             =   825
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Línea :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   225
            TabIndex        =   82
            Top             =   210
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Viaje :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   5415
            TabIndex        =   81
            Top             =   210
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Buque :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   2310
            TabIndex        =   80
            Top             =   210
            Width           =   555
         End
      End
      Begin MSDataGridLib.DataGrid Grid_dato 
         Height          =   4845
         Left            =   105
         TabIndex        =   38
         Top             =   1695
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   8546
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
         ColumnCount     =   1
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
         EndProperty
      End
      Begin Threed.SSFrame fra_tipo_rol 
         Height          =   4935
         Left            =   3930
         TabIndex        =   39
         Top             =   1605
         Width           =   6705
         _Version        =   65536
         _ExtentX        =   11827
         _ExtentY        =   8705
         _StockProps     =   14
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
         Begin VB.TextBox Txt_sello1 
            Height          =   285
            Left            =   225
            MaxLength       =   15
            TabIndex        =   19
            Top             =   3015
            Width           =   1455
         End
         Begin VB.Frame Frame37 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   2385
            TabIndex        =   62
            Top             =   840
            Width           =   1695
            Begin GTMaskNum.GTMaskNum GTM_tara 
               Height          =   315
               Left            =   75
               TabIndex        =   11
               Top             =   225
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
            Begin VB.Label Label53 
               AutoSize        =   -1  'True
               Caption         =   "Tara"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   90
               TabIndex        =   63
               Top             =   0
               Width           =   330
            End
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   4680
            TabIndex        =   60
            Top             =   840
            Width           =   1695
            Begin GTMaskNum.GTMaskNum GTM_peso 
               Height          =   315
               Left            =   75
               TabIndex        =   12
               Top             =   225
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
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Peso :"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   90
               TabIndex        =   61
               Top             =   0
               Width           =   450
            End
         End
         Begin VB.TextBox Txt_sello2 
            Height          =   285
            Left            =   1755
            MaxLength       =   15
            TabIndex        =   20
            Top             =   3015
            Width           =   1515
         End
         Begin VB.TextBox Txt_sello3 
            Height          =   285
            Left            =   3375
            MaxLength       =   15
            TabIndex        =   21
            Top             =   3015
            Width           =   1485
         End
         Begin VB.TextBox Txt_sello4 
            Height          =   285
            Left            =   5055
            MaxLength       =   15
            TabIndex        =   22
            Top             =   3015
            Width           =   1455
         End
         Begin VB.Frame Frame2 
            Caption         =   " Temperatura "
            Height          =   1500
            Left            =   180
            TabIndex        =   40
            Top             =   3330
            Width           =   6375
            Begin VB.Frame Frame4 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   540
               Left            =   135
               TabIndex        =   58
               Top             =   240
               Width           =   1200
               Begin GTMaskNum.GTMaskNum GTM_temp_dur_almac 
                  Height          =   315
                  Left            =   75
                  TabIndex        =   23
                  Top             =   225
                  Width           =   990
                  _Version        =   65536
                  _ExtentX        =   1746
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
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  Caption         =   "Dur Almac:"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   90
                  TabIndex        =   59
                  Top             =   0
                  Width           =   780
               End
            End
            Begin VB.Frame Frame5 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   540
               Left            =   1350
               TabIndex        =   56
               Top             =   240
               Width           =   1200
               Begin GTMaskNum.GTMaskNum GTM_temp_dur_trans 
                  Height          =   315
                  Left            =   75
                  TabIndex        =   24
                  Top             =   225
                  Width           =   990
                  _Version        =   65536
                  _ExtentX        =   1746
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
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  Caption         =   "Dur Transp. :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   90
                  TabIndex        =   57
                  Top             =   0
                  Width           =   930
               End
            End
            Begin VB.Frame Frame6 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   540
               Left            =   2595
               TabIndex        =   54
               Top             =   240
               Width           =   1200
               Begin GTMaskNum.GTMaskNum GTM_temp_dur_manip 
                  Height          =   315
                  Left            =   75
                  TabIndex        =   25
                  Top             =   225
                  Width           =   990
                  _Version        =   65536
                  _ExtentX        =   1746
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
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  Caption         =   "Dur Manip. :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   90
                  TabIndex        =   55
                  Top             =   0
                  Width           =   870
               End
            End
            Begin VB.Frame Frame7 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   540
               Left            =   3810
               TabIndex        =   51
               Top             =   240
               Width           =   1200
               Begin GTMaskNum.GTMaskNum GTM_temp_min_almac 
                  Height          =   315
                  Left            =   75
                  TabIndex        =   52
                  Top             =   225
                  Width           =   990
                  _Version        =   65536
                  _ExtentX        =   1746
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
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "Min. Almac. :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   90
                  TabIndex        =   53
                  Top             =   0
                  Width           =   915
               End
            End
            Begin VB.Frame Frame8 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   540
               Left            =   5040
               TabIndex        =   49
               Top             =   240
               Width           =   1200
               Begin GTMaskNum.GTMaskNum GTM_temp_max_almac 
                  Height          =   315
                  Left            =   75
                  TabIndex        =   26
                  Top             =   225
                  Width           =   990
                  _Version        =   65536
                  _ExtentX        =   1746
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
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  Caption         =   "Max. Almac. :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   105
                  TabIndex        =   50
                  Top             =   0
                  Width           =   960
               End
            End
            Begin VB.Frame Frame9 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   540
               Left            =   120
               TabIndex        =   47
               Top             =   840
               Width           =   1200
               Begin GTMaskNum.GTMaskNum GTM_temp_min_trans 
                  Height          =   315
                  Left            =   75
                  TabIndex        =   27
                  Top             =   225
                  Width           =   990
                  _Version        =   65536
                  _ExtentX        =   1746
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
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  Caption         =   "Min. Transp. :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   90
                  TabIndex        =   48
                  Top             =   0
                  Width           =   975
               End
            End
            Begin VB.Frame Frame10 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   540
               Left            =   1350
               TabIndex        =   45
               Top             =   840
               Width           =   1200
               Begin GTMaskNum.GTMaskNum GTM_temp_max_trans 
                  Height          =   315
                  Left            =   75
                  TabIndex        =   28
                  Top             =   225
                  Width           =   990
                  _Version        =   65536
                  _ExtentX        =   1746
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
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  Caption         =   "Max. Transp. :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   105
                  TabIndex        =   46
                  Top             =   0
                  Width           =   1020
               End
            End
            Begin VB.Frame Frame11 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   540
               Left            =   2640
               TabIndex        =   43
               Top             =   840
               Width           =   1200
               Begin GTMaskNum.GTMaskNum GTM_temp_min_manip 
                  Height          =   315
                  Left            =   75
                  TabIndex        =   29
                  Top             =   225
                  Width           =   990
                  _Version        =   65536
                  _ExtentX        =   1746
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
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  Caption         =   "Min. Manip. :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   105
                  TabIndex        =   44
                  Top             =   0
                  Width           =   915
               End
            End
            Begin VB.Frame Frame12 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   540
               Left            =   3870
               TabIndex        =   41
               Top             =   840
               Width           =   1200
               Begin GTMaskNum.GTMaskNum GTM_temp_max_manip 
                  Height          =   315
                  Left            =   75
                  TabIndex        =   30
                  Top             =   225
                  Width           =   990
                  _Version        =   65536
                  _ExtentX        =   1746
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
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "Max. Manip. :"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   105
                  TabIndex        =   42
                  Top             =   0
                  Width           =   960
               End
            End
         End
         Begin MSDataListLib.DataCombo Cmb_reg_serv_transp 
            Height          =   315
            Left            =   4740
            TabIndex        =   15
            Top             =   1755
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_reg_serv_transp 
            Height          =   330
            Left            =   5700
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
         Begin MSDataListLib.DataCombo Cmb_tipo_equip 
            Height          =   315
            Left            =   240
            TabIndex        =   7
            Top             =   435
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_tipo_equip 
            Height          =   330
            Left            =   1245
            Top             =   420
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
         Begin MSDataListLib.DataCombo Cmb_proved_equip 
            Height          =   315
            Left            =   4740
            TabIndex        =   9
            Top             =   435
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_proved_equip 
            Height          =   330
            Left            =   5610
            Top             =   420
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
         Begin MSDataListLib.DataCombo Cmb_estad_equip 
            Height          =   315
            Left            =   285
            TabIndex        =   10
            Top             =   1080
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_estad_equip 
            Height          =   330
            Left            =   1335
            Top             =   1020
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
         Begin MSDataListLib.DataCombo Cmb_indic_llen_vac 
            Height          =   315
            Left            =   225
            TabIndex        =   13
            Top             =   1755
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_indic_llen_vac 
            Height          =   330
            Left            =   1245
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
         Begin MSDataListLib.DataCombo Cmb_cond_contrato 
            Height          =   315
            Left            =   2475
            TabIndex        =   14
            Top             =   1755
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_cond_contrato 
            Height          =   330
            Left            =   3555
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
         Begin MSDataListLib.DataCombo Cmb_natu_carg 
            Height          =   315
            Left            =   210
            TabIndex        =   16
            Top             =   2370
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_natu_carg 
            Height          =   330
            Left            =   1290
            Top             =   2355
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
         Begin MSDataListLib.DataCombo Cmb_tipo_bien 
            Height          =   315
            Left            =   2475
            TabIndex        =   17
            Top             =   2370
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_tipo_bien 
            Height          =   330
            Left            =   3585
            Top             =   2370
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
         Begin MSDataListLib.DataCombo Cmb_tam_Equip 
            Height          =   315
            Left            =   2430
            TabIndex        =   8
            Top             =   435
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_tam_equip 
            Height          =   330
            Left            =   3555
            Top             =   435
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
         Begin MSDataListLib.DataCombo cmb_condicion 
            Height          =   315
            Left            =   4740
            TabIndex        =   18
            Top             =   2370
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_condicion 
            Height          =   330
            Left            =   5610
            Top             =   2370
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
            Caption         =   "Condición de Contenedor :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   4725
            TabIndex        =   85
            Top             =   2130
            Width           =   1890
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tamaño de Equipo:"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   2430
            TabIndex        =   84
            Top             =   195
            Width           =   1395
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Requer. Servicio Transp. :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   4740
            TabIndex        =   75
            Top             =   1530
            Width           =   1860
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sello 1 :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   225
            TabIndex        =   74
            Top             =   2745
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Equipo:"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   225
            TabIndex        =   73
            Top             =   195
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor equipo:"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   4725
            TabIndex        =   72
            Top             =   195
            Width           =   1305
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Estado del Equipo :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   270
            TabIndex        =   71
            Top             =   840
            Width           =   1380
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Indicador Lleno/vacío :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   210
            TabIndex        =   70
            Top             =   1530
            Width           =   1680
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Condición del Contrato :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   2460
            TabIndex        =   69
            Top             =   1530
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Naturaleza Carga :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   195
            TabIndex        =   68
            Top             =   2145
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Bien :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   2460
            TabIndex        =   67
            Top             =   2145
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sello 2 :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   1785
            TabIndex        =   66
            Top             =   2745
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sello 3 :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   3390
            TabIndex        =   65
            Top             =   2745
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sello 4 :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   18
            Left            =   5070
            TabIndex        =   64
            Top             =   2745
            Width           =   570
         End
      End
      Begin MSAdodcLib.Adodc RDC_grid_dato 
         Height          =   330
         Left            =   1395
         Top             =   6240
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
         Left            =   4350
         Top             =   6300
         Visible         =   0   'False
         Width           =   6390
         _ExtentX        =   11271
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
      Left            =   6045
      TabIndex        =   33
      Top             =   6735
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
      Left            =   7230
      TabIndex        =   34
      Top             =   6735
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
      Left            =   8415
      TabIndex        =   35
      Top             =   6735
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
      Left            =   9600
      TabIndex        =   36
      Top             =   6735
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
      Left            =   4860
      TabIndex        =   32
      Top             =   6735
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
      Left            =   1380
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   88
      Top             =   6765
      Width           =   1200
   End
   Begin Threed.SSCommand cmd_consultar 
      Height          =   375
      Left            =   3675
      TabIndex        =   31
      Top             =   6735
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Consultar"
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
Attribute VB_Name = "doc_contenedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Lstr_opc As String
Public forma_bl As Form
Public Gbol_CONT_desde_menu As Boolean

Public Sub habilita(ByVal opc As Boolean)
   If opc Then
      If Lstr_opc = "M" Then
         Cmb_linea.Enabled = False
         txt_vessel.Enabled = False
         Cmb_buque.Enabled = False
         Cmb_viaje.Enabled = False
         Cmb_pto_emb.Enabled = False
         Txt_pto_descarga.Enabled = False
         'Cmd_hlp_pto_emb.Enabled = False
         Cmd_hlp_pto_descarga.Enabled = False
         Txt_contenedor.Enabled = False
      Else
         Cmb_linea.Enabled = opc
         txt_vessel.Enabled = opc
         Cmb_buque.Enabled = opc
         Cmb_viaje.Enabled = opc
         Cmb_pto_emb.Enabled = opc
         Txt_pto_descarga.Enabled = opc
         'Cmd_hlp_pto_emb.Enabled = opc
         Cmd_hlp_pto_descarga.Enabled = opc
         Txt_contenedor.Enabled = opc
         
         Cmb_cond_contrato.Enabled = opc
         cmb_condicion.Enabled = opc
         Cmb_estad_equip.Enabled = opc
         Cmb_indic_llen_vac.Enabled = opc
         Cmb_natu_carg.Enabled = opc
         Cmb_proved_equip.Enabled = opc
         Cmb_reg_serv_transp.Enabled = opc
         Cmb_tipo_bien.Enabled = opc
         Cmb_tipo_equip.Enabled = opc
         Cmb_tam_Equip.Enabled = opc
         Txt_sello1.Enabled = opc
         Txt_sello2.Enabled = opc
         Txt_sello3.Enabled = opc
         Txt_sello4.Enabled = opc
         GTM_peso.Enabled = opc
         GTM_tara.Enabled = opc
         GTM_temp_dur_almac.Enabled = opc
         GTM_temp_dur_manip.Enabled = opc
         GTM_temp_dur_trans.Enabled = opc
         GTM_temp_max_almac.Enabled = opc
         GTM_temp_max_manip.Enabled = opc
         GTM_temp_max_trans.Enabled = opc
         GTM_temp_min_almac.Enabled = opc
         GTM_temp_min_manip.Enabled = opc
         GTM_temp_min_trans.Enabled = opc
      End If
   Else
      Cmb_linea.Enabled = True
      txt_vessel.Enabled = True
      Cmb_buque.Enabled = True
      Cmb_viaje.Enabled = True
      Cmb_pto_emb.Enabled = True
      Txt_pto_descarga.Enabled = True
      'Cmd_hlp_pto_emb.Enabled = True
      Cmd_hlp_pto_descarga.Enabled = True
      Txt_contenedor.Enabled = True
      
      Cmb_cond_contrato.Enabled = True
      cmb_condicion.Enabled = True
      Cmb_estad_equip.Enabled = True
      Cmb_indic_llen_vac.Enabled = True
      Cmb_natu_carg.Enabled = True
      Cmb_proved_equip.Enabled = True
      Cmb_reg_serv_transp.Enabled = True
      Cmb_tipo_bien.Enabled = True
      Cmb_tipo_equip.Enabled = True
      Cmb_tam_Equip.Enabled = True
      Txt_sello1.Enabled = True
      Txt_sello2.Enabled = True
      Txt_sello3.Enabled = True
      Txt_sello4.Enabled = True
      GTM_peso.Enabled = True
      GTM_tara.Enabled = True
      GTM_temp_dur_almac.Enabled = True
      GTM_temp_dur_manip.Enabled = True
      GTM_temp_dur_trans.Enabled = True
      GTM_temp_max_almac.Enabled = True
      GTM_temp_max_manip.Enabled = True
      GTM_temp_max_trans.Enabled = True
      GTM_temp_min_almac.Enabled = True
      GTM_temp_min_manip.Enabled = True
      GTM_temp_min_trans.Enabled = True
   End If
   
End Sub
Private Sub limpia()
   txt_vessel.Text = ""
   Cmb_viaje.Text = ""
   Cmb_pto_emb.Text = ""
   Txt_pto_descarga.Text = ""
      
   Txt_contenedor.Text = ""
      
   Cmb_cond_contrato.BoundText = ""
   cmb_condicion.BoundText = ""
   Cmb_estad_equip.BoundText = ""
   Cmb_indic_llen_vac.BoundText = ""
   Cmb_natu_carg.BoundText = ""
   Cmb_proved_equip.BoundText = ""
   Cmb_reg_serv_transp.BoundText = ""
   Cmb_tipo_bien.BoundText = ""
   Cmb_tipo_equip.BoundText = ""
   Cmb_tam_Equip.BoundText = ""
   Txt_sello1.Text = ""
   Txt_sello2.Text = ""
   Txt_sello3.Text = ""
   Txt_sello4.Text = ""
   GTM_peso.Value = 0
   GTM_tara.Value = 0
   GTM_temp_dur_almac.Value = 0
   GTM_temp_dur_manip.Value = 0
   GTM_temp_dur_trans.Value = 0
   GTM_temp_max_almac.Value = 0
   GTM_temp_max_manip.Value = 0
   GTM_temp_max_trans.Value = 0
   GTM_temp_min_almac.Value = 0
   GTM_temp_min_manip.Value = 0
   GTM_temp_min_trans.Value = 0
   'Tab_datoGTM_temp_min_trans.Value = 0
End Sub

Private Function valida_Datos() As Boolean
On Error GoTo ERROR_VALIDA
Dim reg As New ADODB.Recordset

valida_Datos = True
   'If Trim$(txt_vessel.Text) = "" Then
   If Cmb_linea.BoundText = "" Then
      MsgBox "Ingrese el Valor de Linea", vbInformation, App.Title
      valida_Datos = False
      'If txt_vessel.Enabled Then txt_vessel.SetFocus
      If Cmb_linea.Enabled Then Cmb_linea.SetFocus
      Exit Function
   End If
   
   If Cmb_buque.BoundText = "" Then
      MsgBox "Ingrese el Valor de Vessel", vbInformation, App.Title
      valida_Datos = False
      'If txt_vessel.Enabled Then txt_vessel.SetFocus
      If Cmb_buque.Enabled Then Cmb_buque.SetFocus
      Exit Function
   End If
   If Trim$(Cmb_viaje.BoundText) = "" Then
      MsgBox "Ingrese el Código de Viaje", vbInformation, App.Title
      valida_Datos = False
      If Cmb_viaje.Enabled Then Cmb_viaje.SetFocus
      Exit Function
   End If
   

'   If Trim$(txt_desc.Text) = "" Then
'      MsgBox "Ingrese algún comentario ", vbInformation, App.Title
'      Valida_Datos = False
'      If txt_desc.Enabled Then txt_desc.SetFocus
'      Exit Function
'   End If
      
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
   Dim reg_exis_bl  As New ADODB.Recordset
   Dim maximo       As Integer
   
   On Error GoTo ERROR_GRABACION
   
   'Gcon_docu.BeginTrans
   
   
   If Lstr_opc = "A" Then
     sql = "Select * from docu_cont " & _
           " Where " & _
           " cont_linea = '" & Cmb_linea.BoundText & "'" & _
           " And cont_buque = '" & Cmb_buque.BoundText & "'" & _
           " And cont_viaje = '" & Cmb_viaje.BoundText & "'" & _
           " And cont_puer_emb ='" & Cmb_pto_emb.BoundText & "'" & _
           " And cont_puer_des ='" & Txt_pto_descarga.Text & "'" & _
           " And cont_contenedor='" & Txt_contenedor.Text & "'"
                                
     
     reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
     If reg.EOF Then
        sql = "INSERT INTO docu_cont (cont_linea, cont_buque, cont_viaje, cont_puer_emb, cont_puer_des, cont_contenedor, cont_tam_equip, cont_condicion, cont_tara, cont_tipo_equip, cont_proved_equip, cont_estad_equip, cont_indic_llen_vac, cont_cond_contrato, cont_reg_serv_transp, cont_natu_carg, cont_peso, cont_tipo_bien, cont_sello1, cont_sello2, cont_sello3, cont_sello4, cont_temp_dur_almac, cont_temp_dur_trans, cont_temp_dur_manip, cont_temp_min_almac, cont_temp_max_almac, cont_temp_min_trans, cont_temp_max_trans, cont_temp_min_manip, cont_temp_max_manip)" & _
             " VALUES ('" & Trim$(Cmb_linea.BoundText) & "',  '" & Trim$(Cmb_buque.BoundText) & "', '" & Trim$(Cmb_viaje.BoundText) & "'," & _
                      "'" & Cmb_pto_emb.BoundText & "',            '" & Txt_pto_descarga.Text & "',      '" & Txt_contenedor.Text & "', " & _
                      "'" & Cmb_tam_Equip.BoundText & "',    '" & cmb_condicion.BoundText & "',     " & GTM_tara.Value & ", " & _
                      "'" & Cmb_tipo_equip.BoundText & "',    '" & Cmb_proved_equip.BoundText & "', '" & Cmb_estad_equip.BoundText & "'," & _
                      "'" & Cmb_indic_llen_vac.BoundText & "','" & Cmb_cond_contrato.BoundText & "','" & Cmb_reg_serv_transp.BoundText & "'," & _
                      "'" & Cmb_natu_carg.BoundText & "',      " & GTM_peso.Value & ",              '" & Cmb_tipo_bien.BoundText & "', " & _
                      "'" & Txt_sello1.Text & "',             '" & Txt_sello2.Text & "',            '" & Txt_sello3.Text & "', " & _
                      "'" & Txt_sello4.Text & "',              " & GTM_temp_dur_almac.Value & ",     " & GTM_temp_dur_trans.Value & ", " & _
                      " " & GTM_temp_dur_manip.Value & ",      " & GTM_temp_min_almac.Value & ",     " & GTM_temp_max_almac.Value & ", " & _
                      " " & GTM_temp_min_trans.Value & ",      " & GTM_temp_max_trans.Value & ",     " & GTM_temp_min_manip.Value & ", " & _
                      " " & GTM_temp_max_manip.Value & ")"
     Else
        MsgBox "Este Contenedor ya existe, entre por la opción modificación", vbExclamation, App.Title
        Exit Sub
     End If
   ElseIf Lstr_opc = "M" Then
      sql = "UPDATE docu_cont       SET  cont_cond_contrato = '" & Cmb_cond_contrato.BoundText & "'," & _
                                        "cont_condicion     = '" & cmb_condicion.BoundText & "', " & _
                                        "cont_estad_equip   = '" & Cmb_estad_equip.BoundText & "', " & _
                                        "cont_indic_lleN_vac= '" & Cmb_indic_llen_vac.BoundText & "', " & _
                                        "cont_natu_carg     = '" & Cmb_natu_carg.BoundText & "', " & _
                                        "cont_proved_equip  = '" & Cmb_proved_equip.BoundText & "', " & _
                                        "cont_reg_serv_transp='" & Cmb_reg_serv_transp.BoundText & "', " & _
                                        "cont_tipo_bien     = '" & Cmb_tipo_bien.BoundText & "', " & _
                                        "cont_tipo_equip    = '" & Cmb_tipo_equip.BoundText & "', " & _
                                        "cont_tam_equip     = '" & Cmb_tam_Equip.BoundText & "', "
      sql = sql & "                      cont_sello1        = '" & Txt_sello1.Text & "', " & _
                                        "cont_sello2        = '" & Txt_sello2.Text & "', " & _
                                        "cont_sello3        = '" & Txt_sello3.Text & "', " & _
                                        "cont_sello4        = '" & Txt_sello4.Text & "', " & _
                                        "cont_peso          =  " & CDbl(GTM_peso.Value) & ", " & _
                                        "cont_tara          =  " & CDbl(GTM_tara.Value) & ", " & _
                                        "cont_temp_dur_almac=  " & CDbl(GTM_temp_dur_almac.Value) & ", " & _
                                        "cont_temp_dur_manip=  " & CDbl(GTM_temp_dur_manip.Value) & ", " & _
                                        "cont_temp_dur_trans=  " & CDbl(GTM_temp_dur_trans.Value) & ", " & _
                                        "cont_temp_max_almac=  " & CDbl(GTM_temp_max_almac.Value) & ", " & _
                                        "cont_temp_max_manip=  " & CDbl(GTM_temp_max_manip.Value) & ", " & _
                                        "cont_temp_max_trans=  " & CDbl(GTM_temp_max_trans.Value) & ", " & _
                                        "cont_temp_min_almac=  " & CDbl(GTM_temp_min_almac.Value) & ", " & _
                                        "cont_temp_min_manip=  " & CDbl(GTM_temp_min_manip.Value) & ", " & _
                                        "cont_temp_min_trans=  " & CDbl(GTM_temp_min_trans.Value) & "  "
                                        
      sql = sql & "Where cont_linea = '" & Cmb_linea.BoundText & "'" & _
                  "  And cont_buque = '" & Cmb_buque.BoundText & "'" & _
                  "  And cont_viaje = '" & Cmb_viaje.BoundText & "'" & _
                  "  And cont_puer_emb ='" & Cmb_pto_emb.BoundText & "'" & _
                  "  And cont_puer_des ='" & Txt_pto_descarga.Text & "'" & _
                  "  And cont_contenedor ='" & RDC_grid_dato.Recordset("contenedor") & "'"
                                    

   ElseIf Lstr_opc = "E" Then
        sql = "SELECT * FROM docu_bldet " & _
               "Where bldet_linea      ='" & Cmb_linea.BoundText & "'" & _
               "   and bldet_buque     ='" & Cmb_buque.BoundText & "'" & _
               "   and bldet_viaje     ='" & Cmb_viaje.BoundText & "'" & _
               "   and bldet_puer_emb  ='" & Cmb_pto_emb.BoundText & "'" & _
               "   and bldet_puer_des  ='" & Txt_pto_descarga.Text & "'" & _
               "   and bldet_contenedor='" & RDC_grid_dato.Recordset("contenedor") & "'"
        reg_exis_bl.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        If Not reg_exis_bl.EOF Then
            MsgBox "Este contenedor esta asociado a uno o mas bls por favor verifique antes de eliminarlo", vbInformation
            'Gcon_docu.RollbackTrans
            Exit Sub
        Else
            sql = "delete from docu_cont "
            sql = sql & "Where cont_linea = '" & Cmb_linea.BoundText & "'" & _
                  "  And cont_buque = '" & Cmb_buque.BoundText & "'" & _
                  "  And cont_viaje = '" & Cmb_viaje.BoundText & "'" & _
                  "  And cont_puer_emb ='" & Cmb_pto_emb.BoundText & "'" & _
                  "  And cont_puer_des ='" & Txt_pto_descarga.Text & "'" & _
                  "  And cont_contenedor ='" & RDC_grid_dato.Recordset("contenedor") & "'"
        End If
                                                 
   End If
   If Lstr_opc = "E" Then
      resp = MsgBox("Esta seguro que desea eliminar este Contenedor", vbQuestion + vbYesNo, App.Title)
   Else
      resp = vbYes
   End If
   If resp = vbYes Then
      Gcon_docu.Execute sql, n
      'Gcon_docu.CommitTrans
      MsgBox "Transacción Realizada con éxito. Se afecto " & n & " registro(s)", vbInformation, App.Title
   Else
      Exit Sub
      'MsgBox "Transacción cancelada "
   End If
   
   If Not Gbol_CONT_desde_menu Then forma_bl.GTM_peso.Value = GTM_peso.Value

   limpia
   If Lstr_opc <> "A" Then
     habilita False
   End If
   
   Exit Sub
   
ERROR_GRABACION:
   MsgBox "Ocurrio el Error " & Err.Description, vbInformation, App.Title
   'Gcon_docu.RollbackTrans
   Exit Sub
   Resume 0
End Sub

Private Sub Cmb_buque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub Cmb_cond_contrato_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmb_condicion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_estad_equip_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_indic_llen_vac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmb_linea_Change()
If Cmb_linea.MatchedWithList Then
   carga_datacombo data_vessel, Cmb_buque, "Select * from docu_buque                             Where buqu_linea       ='" & Cmb_linea.BoundText & "' and buqu_status      ='A'", "buqu_codigo", "buqu_descripcion"
   carga_datacombo data_condicion, cmb_condicion, "Select * from docu_cond_contenedor            Where cond_cont_linea  ='" & Cmb_linea.BoundText & "' and cond_cont_status ='A'", "cond_cont_codigo", "cond_cont_descripcion"
   carga_datacombo data_tipo_equip, Cmb_tipo_equip, "Select * from docu_tipo_equip               Where tipo_equip_linea ='" & Cmb_linea.BoundText & "' and tipo_equip_status='A' ", "tipo_equip_codigo", "tipo_equip_descripcion"
   carga_datacombo data_tam_equip, Cmb_tam_Equip, "Select * from docu_tamano_equip               Where tama_equip_linea ='" & Cmb_linea.BoundText & "' and tama_equip_status='A'", "tama_equip_codigo", "tama_equip_descripcion"
   carga_datacombo data_proved_equip, Cmb_proved_equip, "Select * from docu_proveedor_equip      Where prov_linea       ='" & Cmb_linea.BoundText & "' and prov_status      ='A'", "prov_codigo", "prov_descripcion"
   carga_datacombo data_estad_equip, Cmb_estad_equip, "Select * from docu_estado_equipo          Where esta_linea       ='" & Cmb_linea.BoundText & "'                          ", "esta_codigo", "esta_descripcion"
   carga_datacombo data_indic_llen_vac, Cmb_indic_llen_vac, "Select * from docu_indicad_llen_vac Where indic_linea      ='" & Cmb_linea.BoundText & "' and indic_status     ='A'", "indic_codigo", "indic_descripcion"
   carga_datacombo data_cond_contrato, Cmb_cond_contrato, "Select * from docu_cond_contrato      Where cond_linea       ='" & Cmb_linea.BoundText & "' and cond_status      ='A'", "cond_codigo", "cond_descripcion"
   carga_datacombo data_reg_serv_transp, Cmb_reg_serv_transp, "Select * from docu_requer_serv_trans Where requ_serv_linea ='" & Cmb_linea.BoundText & "' and requ_serv_status ='A'", "requ_serv_codigo", "requ_serv_descripcion"
   carga_datacombo data_natu_carg, Cmb_natu_carg, "Select * from docu_natura_carga               Where carg_linea       ='" & Cmb_linea.BoundText & "' and carg_status      ='A'", "carg_codigo", "carg_descripcion"
   
   'carga_datacombo data_tipo_bien, Cmb_tipo_bien, "Select * from docu_tipo_bien Where bien_linea ='" & cmb_linea.BoundText & "'", "bien_codigo", "requ_serv_descripcion"
   

If Lstr_opc = "A" Then
      carga_parametros Cmb_linea.BoundText
      Txt_pto_descarga.Text = Gstr_param.pto_descarga
      Cmb_natu_carg.BoundText = Gstr_param.naturaleza_carga
      cmb_condicion.BoundText = Gstr_param.condicion_cont
      Cmb_proved_equip.BoundText = Gstr_param.proveed_eq
      Cmb_reg_serv_transp.BoundText = Gstr_param.requer_serv
      Cmb_cond_contrato.BoundText = Gstr_param.cond_contrato
      Cmb_estad_equip.BoundText = Gstr_param.estado_equip
      Cmb_indic_llen_vac.BoundText = Gstr_param.indic_llen_vac
      Cmb_tipo_equip.BoundText = Gstr_param.tipo_equipo
End If
End If
End Sub

Private Sub Cmb_buque_Change()
   Dim sql As String
   sql = "Select * from docu_buque_viaje " & _
         " Where buqv_linea ='" & Cmb_linea.BoundText & "'" & _
         "   And buqv_buque ='" & Cmb_buque.BoundText & "'" & _
         "   And buqv_status = 'A'"
   
   carga_datacombo data_viaje, Cmb_viaje, sql, "buqv_viaje", "buqv_viaje"

End Sub

Private Sub Cmb_linea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_natu_carg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_proved_equip_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_pto_emb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_reg_serv_transp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_tam_Equip_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_tipo_bien_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_tipo_equip_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_viaje_Change()
   Dim sql As String
   sql = "Select * from docu_manifiesto " & _
         " Where mani_linea ='" & Cmb_linea.BoundText & "'" & _
         "   And mani_buque ='" & Cmb_buque.BoundText & "'" & _
         "   And mani_viaje ='" & Cmb_viaje.BoundText & "'"
         '"   And buqv_status = 'A'"
   
   carga_datacombo data_pto_emb, Cmb_pto_emb, sql, "mani_puer_emb", "mani_puer_emb"

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
      cmd_consultar.Enabled = False
      RDC_dato.Caption = ""
      RDC_dato.Enabled = False
      habilita True
      limpia
      Cmb_buque.SetFocus
      'cmb_buque.lis = "" 'Nuevo_Tipo_Rol
      Cmb_viaje.BoundText = ""
      Cmb_pto_emb.BoundText = ""
      Txt_contenedor.Text = ""
      Txt_pto_descarga.Text = ""
      
      
      'Tab_dato.Tab = 1
      'Tab_dato.TabEnabled(0) = False
  Case "&Grabar"
      If Not valida_Datos Then Exit Sub
      Graba_trol

      If Lstr_opc = "M" Or Lstr_opc = "E" Then
         Lstr_opc = "C"
         cmd_salir_Click
      Else
         txt_vessel.Text = "" 'Nuevo_Tipo_Rol
      End If
      'If txt_vessel.Enabled Then txt_vessel.SetFocus
      If Cmb_buque.Enabled Then Cmb_buque.SetFocus
      RDC_dato.Caption = ""
      If Lstr_opc = "M" Then
         'Tab_dato.Tab = 0
         'Tab_dato.TabEnabled(0) = True
      End If
      'RDC_grid_dato.Refresh
      'Grid_dato.Refresh
      'Lstr_opc = "C"
      If Not Gbol_CONT_desde_menu Then Unload Me
      
End Select
End Sub

Private Sub cmd_consultar_Click()
    Call busca_contenedor
End Sub

Private Sub cmd_eliminar_Click()
    Lstr_opc = "E"
    cmd_agregar.Caption = "&Grabar"
    cmd_salir.Caption = "&Cancelar"
    cmd_eliminar.Enabled = False
    Cmd_modificar.Enabled = False
    cmd_consultar.Enabled = False
    RDC_dato.Enabled = False
    habilita True
    Cmb_buque.Enabled = False
    Cmb_viaje.Enabled = False
    'Tab_dato.Tab = 1
    'Tab_dato.TabEnabled(0) = False
End Sub

Private Sub Cmd_hlp_pto_descarga_Click()
   Txt_pto_descarga.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Descarga")
End Sub

'Private Sub Cmd_hlp_pto_emb_Click()
'   Txt_pto_emb.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Embarque")
'End Sub

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

Private Sub cmd_modificar_Click()
    Lstr_opc = "M"
    cmd_agregar.Caption = "&Grabar"
    cmd_salir.Caption = "&Cancelar"
    Cmd_modificar.Enabled = False
    cmd_eliminar.Enabled = False
    cmd_consultar.Enabled = False
    RDC_dato.Enabled = False
    habilita True
'    Tab_dato.Tab = 1
'    Tab_dato.TabEnabled(0) = False
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
       cmd_consultar.Enabled = True
       RDC_dato.Enabled = False
       RDC_dato.Caption = ""
       Lstr_opc = ""
       'Tab_dato.TabEnabled(0) = True
       'Tab_dato.Tab = 0
       If RDC_grid_dato.RecordSource <> "" Then
            RDC_grid_dato.Refresh: DoEvents
       End If
       Grid_dato.Refresh:     DoEvents
End Select
End Sub

Private Sub Form_Load()
   Dim res As Boolean
   Me.Height = 7530
   Me.Width = 11025
   Centra_Pantalla Me
   data_linea.ConnectionString = Gstr_conex_docu
   data_linea.RecordSource = "Select * from docu_linea_nav where linea_status ='A'"
   data_linea.Refresh
   Cmb_linea.BoundColumn = "linea_codigo"
   Cmb_linea.ListField = "linea_descripcion"
   Set Cmb_linea.RowSource = data_linea
End Sub

Private Sub Grid_dato_DblClick()
On Error GoTo Errores
Dim reg As New ADODB.Recordset
   If Not RDC_grid_dato.Recordset.EOF Then
        sql = "Select * " & _
              "from docu_cont Where "

           sql = Trim$(sql) & " cont_linea = '" & RDC_grid_dato.Recordset("linea") & "'"
           sql = Trim$(sql) & " and cont_buque = '" & RDC_grid_dato.Recordset("cbuque") & "'"
           sql = Trim$(sql) & " and cont_viaje = '" & RDC_grid_dato.Recordset("viaje") & "'"
           sql = Trim$(sql) & " and cont_puer_emb = '" & RDC_grid_dato.Recordset("pto_emb") & "'"
           sql = Trim$(sql) & " and cont_puer_des = '" & RDC_grid_dato.Recordset("pto_des") & "'"
           sql = Trim$(sql) & " and cont_contenedor = '" & RDC_grid_dato.Recordset("contenedor") & "'"

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'TL Pongo en comentario porque los datos pueden ser tomados desde las cajas o'
        '   desde el grid de consulta masiva                                         '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '      " And cont_buque = '" & Cmb_buque.BoundText & "'" & _
        '      " And cont_viaje = '" & Cmb_viaje.BoundText & "'" & _
        '      " And cont_puer_emb ='" & Txt_pto_emb.Text & "'" & _
        '      " And cont_puer_des ='" & Txt_pto_descarga.Text & "'" & _
        '      " And cont_contenedor ='" & RDC_grid_dato.Recordset("contenedor") & "'"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
              
        reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        
        If Not reg.EOF Then
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Tl aqui se debe aumenta la carga de las variables generales '
            '   del contenedor despues de dar doble click en el grid     '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Cmb_linea.BoundText = RDC_grid_dato.Recordset("linea")
            Cmb_buque.BoundText = RDC_grid_dato.Recordset("cbuque")
            Cmb_viaje.BoundText = RDC_grid_dato.Recordset("viaje")
            Cmb_pto_emb.BoundText = RDC_grid_dato.Recordset("pto_emb")
            Txt_pto_descarga.Text = RDC_grid_dato.Recordset("pto_des")
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Txt_contenedor.Text = reg("cont_contenedor")
            Cmb_cond_contrato.BoundText = reg("cont_cond_contrato")
            cmb_condicion.BoundText = reg("cont_condicion")
            Cmb_estad_equip.BoundText = reg("cont_estad_equip")
            Cmb_indic_llen_vac.BoundText = reg("cont_indic_llen_vac")
            Cmb_natu_carg.BoundText = reg("cont_natu_carg")
            Cmb_proved_equip.BoundText = reg("cont_proved_equip")
            Cmb_reg_serv_transp.BoundText = reg("cont_reg_serv_transp")
            Cmb_tipo_bien.BoundText = reg("cont_tipo_bien")
            Cmb_tipo_equip.BoundText = reg("cont_tipo_equip")
            Cmb_tam_Equip.BoundText = reg("cont_tam_equip")
            Txt_sello1.Text = reg("cont_sello1")
            Txt_sello2.Text = reg("cont_sello2")
            Txt_sello3.Text = reg("cont_sello3")
            Txt_sello4.Text = reg("cont_sello4")
            GTM_peso.Value = CDbl(reg("cont_peso"))
            GTM_tara.Value = CDbl(reg("cont_tara"))
            GTM_temp_dur_almac.Value = CDbl(reg("cont_temp_dur_almac"))
            GTM_temp_dur_manip.Value = CDbl(reg("cont_temp_dur_manip"))
            GTM_temp_dur_trans.Value = CDbl(reg("cont_temp_dur_trans"))
            GTM_temp_max_almac.Value = CDbl(reg("cont_temp_max_almac"))
            GTM_temp_max_manip.Value = CDbl(reg("cont_temp_max_manip"))
            GTM_temp_max_trans.Value = CDbl(reg("cont_temp_max_trans"))
            GTM_temp_min_almac.Value = CDbl(reg("cont_temp_min_almac"))
            GTM_temp_min_manip.Value = CDbl(reg("cont_temp_min_manip"))
            GTM_temp_min_trans.Value = CDbl(reg("cont_temp_min_trans"))
            
            Cmd_modificar.Enabled = True
            cmd_eliminar.Enabled = True
            cmd_consultar.Enabled = False
        End If
   End If
   Exit Sub
Errores:
   MsgBox "Ocurrio el Error " & Err.Description, vbInformation
   Exit Sub
   Resume 0
End Sub

Public Sub busca_contenedor()
Dim sql As String
   RDC_dato.ConnectionString = Gstr_conex_docu
   RDC_grid_dato.ConnectionString = Gstr_conex_docu
   sql = "Select cont_contenedor as Contenedor,buqu_descripcion as Buque, " & _
         "       cont_viaje as Viaje,cont_linea as Linea,cont_buque cbuque," & _
         "       cont_puer_emb as pto_emb,cont_puer_des as pto_des " & _
         " from docu_cont,docu_buque Where "
   If Cmb_linea.MatchedWithList Then
       sql = Trim$(sql) & " cont_linea = '" & Cmb_linea.BoundText & "'"
       If Cmb_buque.MatchedWithList Then
          sql = Trim$(sql) & " And cont_buque = '" & Cmb_buque.BoundText & "'"
       End If
       If Cmb_viaje.MatchedWithList Then
          sql = Trim$(sql) & " And cont_viaje = '" & Cmb_viaje.BoundText & "'"
       End If
       If Cmb_pto_emb.MatchedWithList Then
          sql = Trim$(sql) & " And cont_puer_emb = '" & Cmb_pto_emb.BoundText & "'"
       End If
       If Txt_pto_descarga.Text <> "" Then
          sql = Trim$(sql) & " And cont_puer_des = '" & Txt_pto_descarga.Text & "'"
       End If
       If Txt_contenedor.Text <> "" Then
          sql = Trim$(sql) & " And cont_contenedor = '" & Txt_contenedor.Text & "'"
       End If
   Else
       If Cmb_buque.MatchedWithList Then
          sql = Trim$(sql) & " cont_buque = '" & Cmb_buque.BoundText & "'"
          If Cmb_viaje.MatchedWithList Then
             sql = Trim$(sql) & " And cont_viaje = '" & Cmb_viaje.BoundText & "'"
          End If
          If Cmb_pto_emb.MatchedWithList Then
             sql = Trim$(sql) & " And cont_puer_emb = '" & Cmb_pto_emb.BoundText & "'"
          End If
          If Txt_pto_descarga.Text <> "" Then
             sql = Trim$(sql) & " And cont_puer_des = '" & Txt_pto_descarga.Text & "'"
          End If
          If Txt_contenedor.Text <> "" Then
             sql = Trim$(sql) & " And cont_contenedor = '" & Txt_contenedor.Text & "'"
          End If
       Else
          If Cmb_viaje.MatchedWithList Then
             sql = Trim$(sql) & " cont_viaje = '" & Cmb_viaje.BoundText & "'"
             If Cmb_pto_emb.MatchedWithList Then
                sql = Trim$(sql) & " And cont_puer_emb = '" & Cmb_pto_emb.BoundText & "'"
             End If
             If Txt_pto_descarga.Text <> "" Then
                sql = Trim$(sql) & " And cont_puer_des = '" & Txt_pto_descarga.Text & "'"
             End If
             If Txt_contenedor.Text <> "" Then
                sql = Trim$(sql) & " And cont_contenedor = '" & Txt_contenedor.Text & "'"
             End If
          Else
             If Cmb_pto_emb.MatchedWithList Then
                sql = Trim$(sql) & " cont_puer_emb = '" & Cmb_pto_emb.BoundText & "'"
                If Txt_pto_descarga.Text <> "" Then
                   sql = Trim$(sql) & " And cont_puer_des = '" & Txt_pto_descarga.Text & "'"
                End If
                If Txt_contenedor.Text <> "" Then
                   sql = Trim$(sql) & " And cont_contenedor = '" & Txt_contenedor.Text & "'"
                End If
             Else
                If Txt_pto_descarga.Text <> "" Then
                   sql = Trim$(sql) & " cont_puer_des = '" & Txt_pto_descarga.Text & "'"
                   If Txt_contenedor.Text <> "" Then
                      sql = Trim$(sql) & " And cont_contenedor = '" & Txt_contenedor.Text & "'"
                   End If
                Else
                   If Txt_contenedor.Text <> "" Then
                      sql = Trim$(sql) & " cont_contenedor = '" & Txt_contenedor.Text & "'"
                   Else
                      MsgBox "No ha ingresado ningun parametro para realizar la consulta"
                      Exit Sub
                   End If
                End If
             End If
          End If
       End If
   End If
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'TL aumento en el sql el joing con la tabla de buques
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   sql = Trim$(sql) & " and cont_linea = buqu_linea and cont_buque = buqu_codigo "
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   RDC_grid_dato.RecordSource = sql
   RDC_grid_dato.Refresh
   Set Grid_dato.DataSource = RDC_grid_dato
   Grid_dato.Refresh
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   

End Sub

Private Sub GTM_peso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_tara_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_temp_dur_almac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_temp_dur_manip_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_temp_dur_trans_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_temp_max_almac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_temp_max_manip_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_temp_max_trans_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_temp_min_almac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_temp_min_manip_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_temp_min_trans_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_contenedor_KeyPress(KeyAscii As Integer)

If Len(Trim$(Txt_contenedor.Text)) >= 4 And Not InStr(1, "0123456789" & Chr(32) & Chr(8), Chr(KeyAscii)) <> 0 Then KeyAscii = 0

KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Txt_pto_descarga_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

'Private Sub Txt_pto_emb_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub

Private Sub Txt_sello1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_sello2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_sello3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_sello4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

'procedimiento usado para poder cargar desde afuera de la forma un contenedor
'para posteriormente modificarlo, esto desde la pantalla de bl status
' 3 de octubre 2002 carlos reyes a
Public Sub carga_contenedor(linea As String, buque As String, viaje As String, _
                            pto_emb As String, pto_desc As String, contenedor As String)
On Error GoTo errores_cont

    Cmb_linea.BoundText = linea
    Cmb_buque.BoundText = buque
    Cmb_viaje.BoundText = viaje
    Cmb_pto_emb.BoundText = pto_emb
    Txt_pto_descarga = pto_desc
    Txt_contenedor = contenedor
    Call cmd_consultar_Click
    If Not RDC_grid_dato.Recordset.EOF Then
        RDC_grid_dato.Recordset.MoveFirst
        Grid_dato_DblClick
    End If
    
    Exit Sub
errores_cont:
MsgBox "Ocurrio el error " & Err.Description
Exit Sub
End Sub

