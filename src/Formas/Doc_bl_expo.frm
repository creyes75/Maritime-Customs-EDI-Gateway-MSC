VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "GTDATE32.OCX"
Object = "{667E8C40-F9B5-11CF-90AB-444553540000}#1.0#0"; "GTNUM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form doc_bl_expo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill of lading de Exportación( no se usara mas)"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   10260
   Begin VB.CommandButton Cmd_salir 
      Height          =   465
      Left            =   1875
      Picture         =   "Doc_bl_expo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   0
      Width           =   465
   End
   Begin VB.CommandButton Cmd_cancelar 
      Height          =   465
      Left            =   1410
      Picture         =   "Doc_bl_expo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   0
      Width           =   465
   End
   Begin VB.CommandButton Cmd_grabar 
      Height          =   465
      Left            =   945
      Picture         =   "Doc_bl_expo.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   0
      Width           =   465
   End
   Begin VB.CommandButton Cmd_eliminar 
      Height          =   465
      Left            =   465
      Picture         =   "Doc_bl_expo.frx":0FDE
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   0
      Width           =   465
   End
   Begin VB.CommandButton Cmd_nuevo 
      Height          =   465
      Left            =   0
      Picture         =   "Doc_bl_expo.frx":1420
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   0
      Width           =   465
   End
   Begin TabDlg.SSTab STab_bl 
      Height          =   6000
      Left            =   60
      TabIndex        =   53
      Top             =   525
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10583
      _Version        =   393216
      TabsPerRow      =   5
      TabHeight       =   503
      TabCaption(0)   =   "Datos Generales Bl"
      TabPicture(0)   =   "Doc_bl_expo.frx":1A8A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Fletes"
      TabPicture(1)   =   "Doc_bl_expo.frx":1AA6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Detalle de Bl"
      TabPicture(2)   =   "Doc_bl_expo.frx":1AC2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame29"
      Tab(2).Control(1)=   "Cmd_nuevo_det"
      Tab(2).Control(2)=   "Cmd_eliminar_det"
      Tab(2).Control(3)=   "Cmd_primer_det"
      Tab(2).Control(4)=   "Cmd_anterior_det"
      Tab(2).Control(5)=   "Cmd_siguiente_det"
      Tab(2).Control(6)=   "Cmd_ultimo_det"
      Tab(2).Control(7)=   "Lbl_no_det_act"
      Tab(2).Control(8)=   "Lbl_no_det_tot"
      Tab(2).Control(9)=   "Label47"
      Tab(2).ControlCount=   10
      Begin VB.Frame Frame29 
         Height          =   4830
         Left            =   -74715
         TabIndex        =   104
         Top             =   960
         Width           =   8535
         Begin VB.Frame Frame3 
            Caption         =   " Contenedor "
            Height          =   1680
            Left            =   4050
            TabIndex        =   124
            Top             =   315
            Width           =   4365
            Begin VB.TextBox Txt_sello1 
               Height          =   285
               Left            =   2325
               MaxLength       =   15
               TabIndex        =   42
               Top             =   1110
               Width           =   1455
            End
            Begin VB.Frame Frame40 
               BorderStyle     =   0  'None
               Height          =   510
               Left            =   120
               TabIndex        =   125
               Top             =   315
               Width           =   1620
               Begin VB.TextBox Txt_contenedor 
                  Height          =   285
                  Left            =   120
                  MaxLength       =   11
                  TabIndex        =   39
                  Top             =   195
                  Width           =   1305
               End
               Begin VB.Label Label56 
                  AutoSize        =   -1  'True
                  Caption         =   "Número"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   105
                  TabIndex        =   126
                  Top             =   -30
                  Width           =   555
               End
            End
            Begin MSDataListLib.DataCombo Cmb_tam_Equip 
               Height          =   315
               Left            =   1860
               TabIndex        =   40
               Top             =   480
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
               Left            =   2850
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
            Begin MSDataListLib.DataCombo cmb_condicion 
               Height          =   315
               Left            =   225
               TabIndex        =   41
               Top             =   1065
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
               Left            =   1065
               Top             =   1080
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
               Caption         =   "Sello 1 :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   6
               Left            =   2325
               TabIndex        =   129
               Top             =   840
               Width           =   570
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Condición de Contenedor :"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   13
               Left            =   210
               TabIndex        =   128
               Top             =   840
               Width           =   1890
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Tamaño de Equipo:"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   3
               Left            =   1860
               TabIndex        =   127
               Top             =   255
               Width           =   1395
            End
         End
         Begin VB.Frame Frame32 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   5880
            TabIndex        =   119
            Top             =   4080
            Width           =   1695
            Begin GTMaskNum.GTMaskNum GTM_volumen 
               Height          =   315
               Left            =   75
               TabIndex        =   47
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
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               Caption         =   "Volumen"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   90
               TabIndex        =   120
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame Frame34 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   195
            TabIndex        =   117
            Top             =   870
            Width           =   1695
            Begin GTMaskNum.GTMaskNum GTM_bultos 
               Height          =   315
               Left            =   75
               TabIndex        =   37
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
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               Caption         =   "No. bultos"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   90
               TabIndex        =   118
               Top             =   -15
               Width           =   720
            End
         End
         Begin VB.Frame Frame35 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   120
            TabIndex        =   115
            Top             =   4065
            Width           =   3705
            Begin MSDataListLib.DataCombo Cmb_IMO 
               Height          =   315
               Left            =   105
               TabIndex        =   45
               Top             =   240
               Width           =   3345
               _ExtentX        =   5900
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin MSAdodcLib.Adodc data_IMO 
               Height          =   330
               Left            =   2400
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
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               Caption         =   "IMO"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   116
               Top             =   -30
               Width           =   300
            End
         End
         Begin VB.Frame Frame37 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   4050
            TabIndex        =   113
            Top             =   4050
            Width           =   1695
            Begin GTMaskNum.GTMaskNum GTM_peso 
               Height          =   315
               Left            =   75
               TabIndex        =   46
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
               Caption         =   "Peso"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   90
               TabIndex        =   114
               Top             =   0
               Width           =   360
            End
         End
         Begin VB.Frame Frame38 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   120
            TabIndex        =   111
            Top             =   1515
            Width           =   3705
            Begin MSDataListLib.DataCombo Cmb_embalaje 
               Height          =   315
               Left            =   105
               TabIndex        =   38
               Top             =   225
               Width           =   3360
               _ExtentX        =   5927
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin MSAdodcLib.Adodc data_embalaje 
               Height          =   330
               Left            =   2565
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
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               Caption         =   "Embalaje"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   112
               Top             =   0
               Width           =   645
            End
         End
         Begin VB.Frame Frame39 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   150
            TabIndex        =   109
            Top             =   180
            Width           =   3750
            Begin MSDataListLib.DataCombo Cmb_tipo_carga 
               Height          =   315
               Left            =   105
               TabIndex        =   36
               Top             =   225
               Width           =   3360
               _ExtentX        =   5927
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin MSAdodcLib.Adodc data_tipo_carga 
               Height          =   330
               Left            =   2490
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
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Carga"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   110
               Top             =   0
               Width           =   780
            End
         End
         Begin VB.Frame Frame41 
            BorderStyle     =   0  'None
            Height          =   1575
            Left            =   75
            TabIndex        =   107
            Top             =   2310
            Width           =   3810
            Begin VB.TextBox Txt_marcas 
               Height          =   1260
               Left            =   135
               MultiLine       =   -1  'True
               TabIndex        =   43
               Top             =   195
               Width           =   3585
            End
            Begin VB.Label Label57 
               AutoSize        =   -1  'True
               Caption         =   "Marcas/Números"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   108
               Top             =   -30
               Width           =   1230
            End
         End
         Begin VB.Frame Frame42 
            BorderStyle     =   0  'None
            Height          =   1575
            Left            =   3960
            TabIndex        =   105
            Top             =   2310
            Width           =   3810
            Begin VB.TextBox Txt_desc_carga 
               Height          =   1260
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   44
               Top             =   195
               Width           =   3585
            End
            Begin VB.Label Label58 
               AutoSize        =   -1  'True
               Caption         =   "Descripción carga"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   106
               Top             =   -30
               Width           =   1290
            End
         End
      End
      Begin VB.CommandButton Cmd_nuevo_det 
         Height          =   465
         Left            =   -74760
         Picture         =   "Doc_bl_expo.frx":1ADE
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   510
         Width           =   465
      End
      Begin VB.CommandButton Cmd_eliminar_det 
         Height          =   465
         Left            =   -74295
         Picture         =   "Doc_bl_expo.frx":2148
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   510
         Width           =   465
      End
      Begin VB.CommandButton Cmd_primer_det 
         Height          =   465
         Left            =   -73830
         Picture         =   "Doc_bl_expo.frx":258A
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   510
         Width           =   465
      End
      Begin VB.CommandButton Cmd_anterior_det 
         Height          =   465
         Left            =   -73365
         Picture         =   "Doc_bl_expo.frx":268C
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   510
         Width           =   465
      End
      Begin VB.CommandButton Cmd_siguiente_det 
         Height          =   465
         Left            =   -72900
         Picture         =   "Doc_bl_expo.frx":278E
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   510
         Width           =   465
      End
      Begin VB.CommandButton Cmd_ultimo_det 
         Height          =   465
         Left            =   -72435
         Picture         =   "Doc_bl_expo.frx":2890
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   510
         Width           =   465
      End
      Begin VB.Frame Frame8 
         Caption         =   " Fletes "
         Height          =   3150
         Left            =   -74835
         TabIndex        =   103
         Top             =   465
         Width           =   9750
         Begin VB.CommandButton Cmd_eliminar_flete 
            Height          =   465
            Left            =   9030
            Picture         =   "Doc_bl_expo.frx":2992
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   2625
            Width           =   465
         End
         Begin VB.CommandButton Cmd_nuevo_flete 
            Height          =   465
            Left            =   8565
            Picture         =   "Doc_bl_expo.frx":2DD4
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   2625
            Width           =   465
         End
         Begin MSGrid.Grid Grid_flete 
            Height          =   2250
            Left            =   150
            TabIndex        =   27
            Top             =   300
            Width           =   9330
            _Version        =   65536
            _ExtentX        =   16457
            _ExtentY        =   3969
            _StockProps     =   77
            BackColor       =   16777215
            Cols            =   12
            FixedCols       =   0
         End
      End
      Begin VB.Frame Frame10 
         Height          =   1455
         Left            =   6870
         TabIndex        =   94
         Top             =   345
         Width           =   3045
         Begin VB.Frame Fra_fecha_emb 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   585
            Left            =   1590
            TabIndex        =   101
            Top             =   750
            Width           =   1380
            Begin GTMaskDate.GTMaskDate Msk_fecha_emb 
               Height          =   285
               Left            =   30
               TabIndex        =   12
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
               Caption         =   "Fecha Salida"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   30
               TabIndex        =   102
               Top             =   30
               Width           =   930
            End
         End
         Begin VB.Frame Frame7 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   45
            TabIndex        =   99
            Top             =   795
            Width           =   1500
            Begin VB.TextBox Txt_pto_emb 
               Height          =   285
               Left            =   105
               MaxLength       =   5
               TabIndex        =   11
               Top             =   225
               Width           =   975
            End
            Begin MSAdodcLib.Adodc data_pto_emb 
               Height          =   330
               Left            =   480
               Top             =   210
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
               TabIndex        =   100
               Top             =   0
               Width           =   1230
            End
         End
         Begin VB.Frame Fra_pto_destino 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   1470
            TabIndex        =   97
            Top             =   210
            Width           =   1560
            Begin VB.TextBox Txt_pto_destino 
               Height          =   285
               Left            =   120
               MaxLength       =   5
               TabIndex        =   9
               Top             =   255
               Width           =   975
            End
            Begin VB.CommandButton Cmd_hlp_pto_destino 
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
               Picture         =   "Doc_bl_expo.frx":343E
               Style           =   1  'Graphical
               TabIndex        =   10
               ToolTipText     =   "Buscar"
               Top             =   255
               Width           =   300
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               Caption         =   "Puerto Destino final"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   98
               Top             =   15
               Width           =   1380
            End
         End
         Begin VB.Frame Fra_pto_origen 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   30
            TabIndex        =   95
            Top             =   210
            Width           =   1560
            Begin VB.CommandButton Cmd_hlp_pto_origen 
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
               Picture         =   "Doc_bl_expo.frx":3970
               Style           =   1  'Graphical
               TabIndex        =   8
               ToolTipText     =   "Buscar"
               Top             =   255
               Width           =   300
            End
            Begin VB.TextBox Txt_pto_origen 
               Height          =   285
               Left            =   105
               MaxLength       =   5
               TabIndex        =   7
               Top             =   270
               Width           =   975
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "Puerto Origen"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   96
               Top             =   15
               Width           =   975
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1440
         Left            =   150
         TabIndex        =   54
         Top             =   345
         Width           =   6645
         Begin VB.Frame Fra_pto_descarga 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   105
            TabIndex        =   63
            Top             =   735
            Width           =   1590
            Begin VB.TextBox Txt_pto_descarga 
               Height          =   285
               Left            =   120
               MaxLength       =   5
               TabIndex        =   3
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
               Left            =   1140
               Picture         =   "Doc_bl_expo.frx":3EA2
               Style           =   1  'Graphical
               TabIndex        =   4
               ToolTipText     =   "Buscar"
               Top             =   255
               Width           =   300
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Puerto Descarga"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   64
               Top             =   15
               Width           =   1200
            End
         End
         Begin VB.Frame Fra_linea 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   75
            TabIndex        =   61
            Top             =   195
            Width           =   1500
            Begin MSDataListLib.DataCombo Cmb_linea 
               Height          =   315
               Left            =   135
               TabIndex        =   0
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
            Begin VB.Label Label5 
               Caption         =   "Linea"
               ForeColor       =   &H00800000&
               Height          =   210
               Left            =   105
               TabIndex        =   62
               Top             =   0
               Width           =   825
            End
         End
         Begin VB.Frame Fra_bl 
            BorderStyle     =   0  'None
            Height          =   510
            Left            =   1650
            TabIndex        =   59
            Top             =   825
            Width           =   2535
            Begin VB.TextBox Txt_bl 
               Height          =   285
               Left            =   120
               MaxLength       =   25
               TabIndex        =   5
               Top             =   180
               Width           =   2385
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Bill of lading"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   60
               Top             =   -30
               Width           =   840
            End
         End
         Begin VB.Frame Fra_buque 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   1605
            TabIndex        =   57
            Top             =   210
            Width           =   2640
            Begin MSDataListLib.DataCombo Cmb_buque 
               Height          =   315
               Left            =   135
               TabIndex        =   1
               Top             =   240
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
               Left            =   1500
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
               TabIndex        =   58
               Top             =   0
               Width           =   465
            End
         End
         Begin VB.Frame Fra_viaje 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   4170
            TabIndex        =   55
            Top             =   240
            Width           =   2400
            Begin MSDataListLib.DataCombo Cmb_viaje 
               Height          =   315
               Left            =   120
               TabIndex        =   2
               Top             =   240
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin MSAdodcLib.Adodc data_viaje 
               Height          =   330
               Left            =   1155
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
               TabIndex        =   56
               Top             =   0
               Width           =   345
            End
         End
         Begin MSDataListLib.DataCombo Cmb_oper_carga 
            Height          =   315
            Left            =   4305
            TabIndex        =   6
            Top             =   990
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_oper_carga 
            Height          =   330
            Left            =   5565
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
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Ag. Oper. de Carga (Agencia)"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   4320
            TabIndex        =   90
            Top             =   780
            Width           =   2085
         End
      End
      Begin VB.Frame Frame9 
         Height          =   4020
         Left            =   135
         TabIndex        =   65
         Top             =   1740
         Width           =   9795
         Begin VB.Frame Fra_embarcador 
            Caption         =   " Embarcador "
            Height          =   1305
            Left            =   120
            TabIndex        =   91
            Top             =   2100
            Width           =   4725
            Begin VB.TextBox Txt_direc_embarcador 
               Height          =   285
               Left            =   135
               MaxLength       =   80
               TabIndex        =   26
               Top             =   930
               Width           =   4470
            End
            Begin VB.TextBox Txt_nombre_embarcador 
               Height          =   285
               Left            =   135
               MaxLength       =   80
               TabIndex        =   25
               Top             =   405
               Width           =   3360
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Dirección"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   150
               TabIndex        =   93
               Top             =   720
               Width           =   675
            End
            Begin VB.Label Label23 
               Caption         =   "Nombre"
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   135
               TabIndex        =   92
               Top             =   210
               Width           =   1260
            End
         End
         Begin VB.Frame Fra_consignatario 
            Caption         =   " Exportador "
            Height          =   1965
            Left            =   120
            TabIndex        =   71
            Top             =   120
            Width           =   4740
            Begin VB.CommandButton Cmd_hlp_consignatario 
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
               Left            =   3585
               Picture         =   "Doc_bl_expo.frx":43D4
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Buscar"
               Top             =   450
               Width           =   300
            End
            Begin VB.TextBox Txt_id_consignatario 
               Height          =   285
               Left            =   2025
               MaxLength       =   15
               TabIndex        =   14
               Top             =   465
               Width           =   1500
            End
            Begin VB.TextBox Txt_nombre_consignatario 
               Height          =   285
               Left            =   135
               MaxLength       =   80
               TabIndex        =   16
               Top             =   1035
               Width           =   3360
            End
            Begin VB.TextBox Txt_direc_consignatario 
               Height          =   285
               Left            =   135
               MaxLength       =   80
               TabIndex        =   19
               Top             =   1575
               Width           =   4470
            End
            Begin VB.CommandButton Cmd_hlp_nac_consignatario 
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
               Left            =   4260
               Picture         =   "Doc_bl_expo.frx":4906
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Buscar"
               Top             =   1020
               Width           =   300
            End
            Begin VB.TextBox Txt_nac_consignatario 
               Height          =   285
               Left            =   3570
               MaxLength       =   5
               TabIndex        =   17
               Top             =   1035
               Width           =   615
            End
            Begin MSAdodcLib.Adodc data_tipo_id_consignatario 
               Height          =   330
               Left            =   3870
               Top             =   195
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
            Begin MSDataListLib.DataCombo Cmb_tipo_id_consignatario 
               Height          =   315
               Left            =   135
               TabIndex        =   13
               Top             =   435
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Identificación"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   2025
               TabIndex        =   76
               Top             =   210
               Width           =   945
            End
            Begin VB.Label Label4 
               Caption         =   "Nombre"
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   135
               TabIndex        =   75
               Top             =   780
               Width           =   1260
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Dirección"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   150
               TabIndex        =   74
               Top             =   1350
               Width           =   675
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Nacionalidad"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   3585
               TabIndex        =   73
               Top             =   780
               Width           =   930
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Identif."
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   72
               Top             =   195
               Width           =   1065
            End
         End
         Begin VB.Frame Fra_notificador 
            Caption         =   " Notificador "
            Height          =   1965
            Left            =   4935
            TabIndex        =   66
            Top             =   120
            Width           =   4725
            Begin VB.TextBox Txt_direc_notificador 
               Height          =   285
               Left            =   150
               MaxLength       =   80
               TabIndex        =   24
               Top             =   1575
               Width           =   4500
            End
            Begin VB.TextBox Txt_nombre_notificador 
               Height          =   285
               Left            =   135
               MaxLength       =   80
               TabIndex        =   23
               Top             =   1035
               Width           =   3360
            End
            Begin VB.TextBox Txt_id_notificador 
               Height          =   285
               Left            =   2025
               MaxLength       =   15
               TabIndex        =   21
               Top             =   465
               Width           =   1500
            End
            Begin VB.CommandButton Cmd_hlp_notificador 
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
               Left            =   3585
               Picture         =   "Doc_bl_expo.frx":4E38
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "Buscar"
               Top             =   450
               Width           =   300
            End
            Begin MSAdodcLib.Adodc data_tipo_id_notificador 
               Height          =   330
               Left            =   3870
               Top             =   195
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
            Begin MSDataListLib.DataCombo Cmb_tipo_id_notificador 
               Height          =   315
               Left            =   135
               TabIndex        =   20
               Top             =   435
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Identif."
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   70
               Top             =   195
               Width           =   1065
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Dirección"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   150
               TabIndex        =   69
               Top             =   1350
               Width           =   675
            End
            Begin VB.Label Label21 
               Caption         =   "Nombre"
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   135
               TabIndex        =   68
               Top             =   780
               Width           =   1260
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Identificación"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   2025
               TabIndex        =   67
               Top             =   210
               Width           =   945
            End
         End
         Begin VB.Frame Frame20 
            Height          =   570
            Left            =   135
            TabIndex        =   77
            Top             =   3345
            Width           =   9525
            Begin VB.Frame Frame13 
               BorderStyle     =   0  'None
               Height          =   360
               Left            =   60
               TabIndex        =   78
               Top             =   165
               Width           =   2370
               Begin VB.TextBox Txt_peso_tot 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000016&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1110
                  Locked          =   -1  'True
                  MaxLength       =   80
                  TabIndex        =   79
                  Top             =   15
                  Width           =   1155
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  Caption         =   "Peso Total"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   105
                  TabIndex        =   80
                  Top             =   75
                  Width           =   765
               End
            End
            Begin VB.Frame Frame15 
               BorderStyle     =   0  'None
               Height          =   360
               Left            =   7200
               TabIndex        =   87
               Top             =   165
               Width           =   2235
               Begin VB.TextBox Txt_cont_tot 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000016&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1020
                  Locked          =   -1  'True
                  MaxLength       =   80
                  TabIndex        =   88
                  Top             =   15
                  Width           =   1155
               End
               Begin VB.Label Label26 
                  AutoSize        =   -1  'True
                  Caption         =   "No.Contens."
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   15
                  TabIndex        =   89
                  Top             =   75
                  Width           =   885
               End
            End
            Begin VB.Frame Frame16 
               BorderStyle     =   0  'None
               Height          =   360
               Left            =   2475
               TabIndex        =   84
               Top             =   165
               Width           =   2370
               Begin VB.TextBox Txt_volumen_tot 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000016&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1050
                  Locked          =   -1  'True
                  MaxLength       =   80
                  TabIndex        =   85
                  Top             =   15
                  Width           =   1155
               End
               Begin VB.Label Label17 
                  Caption         =   "Volum Total"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   75
                  TabIndex        =   86
                  Top             =   75
                  Width           =   1020
               End
            End
            Begin VB.Frame Frame14 
               BorderStyle     =   0  'None
               Height          =   360
               Left            =   4860
               TabIndex        =   81
               Top             =   165
               Width           =   2235
               Begin VB.TextBox Txt_bultos_tot 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000016&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1020
                  Locked          =   -1  'True
                  MaxLength       =   80
                  TabIndex        =   82
                  Top             =   15
                  Width           =   1155
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  Caption         =   "BultosTotal"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   15
                  TabIndex        =   83
                  Top             =   75
                  Width           =   795
               End
            End
         End
      End
      Begin VB.Label Lbl_no_det_act 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   -71355
         TabIndex        =   123
         Top             =   645
         Width           =   90
      End
      Begin VB.Label Lbl_no_det_tot 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   -70950
         TabIndex        =   122
         Top             =   645
         Width           =   90
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "de"
         Height          =   195
         Left            =   -71205
         TabIndex        =   121
         Top             =   645
         Width           =   180
      End
   End
End
Attribute VB_Name = "doc_bl_expo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Gstr_opc_bl As String
Dim mani_bl_det()   As struct_blimpodt  ' para el manejo de los detalles de los bls.

Private Sub Cmb_Buque_Change()
   Dim sql As String
   sql = "Select * from docu_buque_viaje " & _
         " Where buqv_linea ='" & cmb_linea.BoundText & "'" & _
         "   And buqv_buque ='" & Cmb_buque.BoundText & "'" & _
         "   And buqv_status = 'A'"
   carga_datacombo data_viaje, Cmb_viaje, sql, "buqv_viaje", "buqv_viaje"
End Sub

Private Sub Cmb_Buque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmb_condicion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmb_condicion_LostFocus()
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).cond_cont = cmb_condicion.BoundText
End Sub

Private Sub Cmb_embalaje_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_embalaje_LostFocus()
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).embalaje = Cmb_embalaje.BoundText
End Sub

Private Sub Cmb_IMO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_IMO_LostFocus()
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).IMO = Cmb_IMO.BoundText
End Sub

Private Sub cmb_linea_Change()
    If Not cmb_linea.MatchedWithList Then Exit Sub
    carga_datacombo data_buque, Cmb_buque, "Select * from docu_buque                                Where buqu_linea        = '" & cmb_linea.BoundText & "' and buqu_status      ='A'", "buqu_codigo", "buqu_descripcion"
    carga_datacombo data_tipo_carga, Cmb_tipo_carga, "Select * from docu_tipo_carga                 Where tipo_carga_linea  = '" & cmb_linea.BoundText & "' and tipo_carga_status='A'", "tipo_carga_codigo", "tipo_carga_descripcion"
    carga_datacombo data_embalaje, Cmb_embalaje, "Select * from docu_embalaje                       Where emba_linea        = '" & cmb_linea.BoundText & "' and emba_status      ='A'", "emba_codigo", "emba_descripcion"
    carga_datacombo data_IMO, Cmb_IMO, "Select * from docu_imo                                      Where imo_linea         = '" & cmb_linea.BoundText & "' and imo_status       ='A'", "imo_codigo", "imo_descripcion"
    carga_datacombo data_oper_carga, Cmb_oper_carga, "Select * from docu_emp_resp                   Where emp_resp_linea    = '" & cmb_linea.BoundText & "' and emp_resp_status  ='A'", "emp_resp_codigo", "emp_resp_descripcion"
    
    carga_datacombo data_condicion, cmb_condicion, "Select * from docu_cond_contenedor            Where cond_cont_linea  ='" & cmb_linea.BoundText & "' and cond_cont_status ='A'", "cond_cont_codigo", "cond_cont_descripcion"
    carga_datacombo data_tam_equip, Cmb_tam_Equip, "Select * from docu_tamano_equip               Where tama_equip_linea ='" & cmb_linea.BoundText & "' and tama_equip_status='A'", "tama_equip_codigo", "tama_equip_descripcion"

    carga_parametros cmb_linea.BoundText
    
    Me.Txt_pto_origen.Text = Gstr_param.pto_descarga
    Me.Txt_pto_emb.Text = Gstr_param.pto_descarga
    Cmb_oper_carga.BoundText = Gstr_param.naviera
    
    Me.Cmb_tipo_id_consignatario.BoundText = Gstr_param.tipo_id
    Me.Cmb_tipo_id_notificador.BoundText = Gstr_param.tipo_id
    Me.Txt_nac_consignatario.Text = Gstr_param.nacionalidad
End Sub

Private Sub Cmb_linea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_oper_carga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_tam_Equip_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_tam_Equip_LostFocus()
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).tamano_equip = Cmb_tam_Equip.BoundText
End Sub

Private Sub Cmb_tipo_carga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_tipo_carga_LostFocus()
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).tipo_carga = Cmb_tipo_carga.BoundText
End Sub

Private Sub Cmb_tipo_id_consignatario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_tipo_id_notificador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_viaje_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmd_anterior_det_Click()
If CDbl(Lbl_no_det_act.Caption) > 1 Then
    Lbl_no_det_act.Caption = Str(CDbl(Lbl_no_det_act.Caption) - 1)
    Call muestra_detalle
End If
End Sub

Public Sub cmd_eliminar_Click()
If MsgBox("Esta seguro que desea eliminar el bl " & Txt_bl.Text, vbYesNo + vbQuestion) = vbYes Then
    If Elimina_bl_EXP(cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, Txt_pto_emb.Text, Txt_pto_descarga.Text, Cmb_oper_carga.BoundText, Txt_bl.Text) Then
        inicia_controles
        limpia_detalle
     End If
End If
End Sub

Private Sub Cmd_eliminar_det_Click()
If CDbl(Lbl_no_det_act.Caption) > 0 Then
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).elim = "E"
    If CDbl(Lbl_no_det_tot.Caption) > 1 Then
       If Lbl_no_det_act.Caption = Lbl_no_det_tot.Caption Then
          Lbl_no_det_act.Caption = Str(CDbl(Lbl_no_det_act.Caption) - 1)
       End If
       Lbl_no_det_tot.Caption = Str(CDbl(Lbl_no_det_tot.Caption) - 1)
       Call muestra_detalle
    Else
       Lbl_no_det_tot.Caption = "0"
       Lbl_no_det_act.Caption = "0"
    End If
End If
End Sub

Private Sub Cmd_eliminar_flete_Click()
    If Grid_flete.Rows <= 2 Or Grid_flete.Row = Grid_flete.Rows - 1 Then Exit Sub
    
    For i = Grid_flete.Row To Grid_flete.Rows - 2
        Grid_flete.Row = i + 1
        Grid_flete.Col = 0: cargo_cod = Grid_flete.Text
        Grid_flete.Col = 1: cargo = Grid_flete.Text
        Grid_flete.Col = 2: debtor_cod = Grid_flete.Text
        Grid_flete.Col = 3: debtor = Grid_flete.Text
        Grid_flete.Col = 4: valor = Grid_flete.Text
        Grid_flete.Col = 5: SIGNO = Grid_flete.Text
        Grid_flete.Col = 6: MONEDA = Grid_flete.Text
        Grid_flete.Col = 7: marcas = Grid_flete.Text
        Grid_flete.Col = 8: lugar = Grid_flete.Text
        Grid_flete.Col = 9: imp_bl = Grid_flete.Text
        
        Grid_flete.Row = i
        Grid_flete.Col = 0: Grid_flete.Text = cargo_cod
        Grid_flete.Col = 1: Grid_flete.Text = cargo
        Grid_flete.Col = 2: Grid_flete.Text = deptor_cod
        Grid_flete.Col = 3: Grid_flete.Text = debtor
        Grid_flete.Col = 4: Grid_flete.Text = valor
        Grid_flete.Col = 5: Grid_flete.Text = SIGNO
        Grid_flete.Col = 6: Grid_flete.Text = MONEDA
        Grid_flete.Col = 7: Grid_flete.Text = marcas
        Grid_flete.Col = 8: Grid_flete.Text = lugar
        Grid_flete.Col = 9: Grid_flete.Text = imp_bl
    Next i
    Grid_flete.Rows = Grid_flete.Rows - 1
    
End Sub

Private Sub Cmd_grabar_Click()
On Error GoTo Errores

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'TL Probar las validaciones de ingreso antes de la grabacion'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If valida_Datos = False Then
   Exit Sub
End If
Gcon_docu.BeginTrans
If Gstr_opc_bl = "A" Then
        sql = "Insert into docu_blcab_exp (blcab_linea,     blcab_buque,    blcab_viaje, " & _
                                     " blcab_puer_emb,  blcab_puer_des, blcab_emp_resp, " & _
                                     " blcab_cono_emb,   " & _
                                     " blcab_peso_tot, blcab_bultos_tot, " & _
                                     " blcab_tipo_ide_cons, blcab_ide_cons,     blcab_consig, " & _
                                     " blcab_direc_consig,  blcab_nacio_consig, blcab_tipo_ide_notif," & _
                                     " blcab_ide_notif,     blcab_notif,        blcab_direc_notif, " & _
                                     " blcab_embarcador,    blcab_direc_embarc, blcab_puer_origen," & _
                                     " blcab_no_cont,      blcab_volumen_tot," & _
                                     " blcab_pto_dest_final," & _
                                     " blcab_fecha_emb,     blcab_user_ing,     blcab_fecha_ing, " & _
                                     " blcab_hora_ing)"
         sql = sql & "VALUES ('" & cmb_linea.BoundText & "',    '" & Cmb_buque.BoundText & "',       '" & Cmb_viaje.BoundText & "', " & _
                        "'" & Txt_pto_emb.Text & "',       '" & Txt_pto_descarga.Text & "',     '" & Cmb_oper_carga.BoundText & "', " & _
                        "'" & Txt_bl.Text & "'," & _
                        " " & CDbl(Txt_peso_tot.Text) & ",     " & CDbl(Txt_bultos_tot.Text) & ", " & _
                        "'" & Cmb_tipo_id_consignatario.BoundText & "', '" & Txt_id_consignatario.Text & "',  '" & Txt_nombre_consignatario.Text & "', " & _
                        "'" & Txt_direc_consignatario.Text & "',        '" & Txt_nac_consignatario.Text & "', '" & Cmb_tipo_id_notificador.BoundText & "', " & _
                        "'" & Txt_id_notificador.Text & "',             '" & Txt_nombre_notificador.Text & "','" & Txt_direc_notificador.Text & "', " & _
                        "'" & Txt_nombre_embarcador.Text & "',          '" & Txt_direc_embarcador.Text & "',  '" & Txt_pto_origen.Text & "', "
        sql = sql & "     " & _
                        " " & CDbl(Txt_cont_tot.Text) & ",      " & CDbl(Txt_volumen_tot.Text) & ", " & _
                        "'" & Txt_pto_destino.Text & "', " & _
                        "'" & Msk_fecha_emb.Text & "',                  '" & Gstr_user & "',                  '" & Format(Now, "dd/mm/yyyy") & "', " & _
                        "'" & Format(Time, "hh:mm:ss") & "')"

Else
        sql = "UPDATE docu_blcab_exp SET " & _
                                    "blcab_peso_tot   = " & CDbl(Txt_peso_tot.Text) & ", " & _
                                    "blcab_bultos_tot = " & CDbl(Txt_bultos_tot.Text) & ", " & _
                                    "blcab_tipo_ide_cons='" & Trim$(Cmb_tipo_id_consignatario.BoundText) & "', " & _
                                    "blcab_ide_cons     ='" & Trim$(Txt_id_consignatario.Text) & "'," & _
                                    "blcab_consig       ='" & Trim$(Txt_nombre_consignatario.Text) & "', " & _
                                    "blcab_direc_consig ='" & Trim$(Txt_direc_consignatario.Text) & "', " & _
                                    "blcab_nacio_consig ='" & Trim$(Txt_nac_consignatario.Text) & "', " & _
                                    "blcab_tipo_ide_notif='" & Trim$(Cmb_tipo_id_notificador.BoundText) & "', " & _
                                    "blcab_ide_notif     ='" & Trim$(Txt_id_notificador.Text) & "', " & _
                                    "blcab_notif         ='" & Trim$(Txt_nombre_notificador.Text) & "', " & _
                                    "blcab_direc_notif   ='" & Trim$(Txt_direc_notificador.Text) & "', " & _
                                    "blcab_embarcador    ='" & Trim$(Txt_nombre_embarcador.Text) & "', " & _
                                    "blcab_direc_embarc  ='" & Trim$(Txt_direc_embarcador.Text) & "', " & _
                                    "blcab_puer_origen   ='" & Txt_pto_origen.Text & "', " & _
                                    "blcab_no_cont       = " & CDbl(Txt_cont_tot.Text) & ", " & _
                                    "blcab_volumen_tot   = " & CDbl(Txt_volumen_tot.Text) & "," & _
                                    "blcab_pto_dest_final='" & Txt_pto_destino.Text & "', " & _
                                    "blcab_fecha_emb     ='" & Msk_fecha_emb.Text & "', " & _
                                    "blcab_user_mod      ='" & Gstr_user & "', " & _
                                    "blcab_fecha_mod     ='" & Format(Now, "dd/mm/yyyy") & "', " & _
                                    "blcab_hora_mod      ='" & Format(Time, "hh:mm:ss") & "'"
                                    
        sql = sql & "  Where blcab_linea       ='" & cmb_linea.BoundText & "'" & _
                        "   and blcab_buque    ='" & Cmb_buque.BoundText & "'" & _
                        "   and blcab_viaje    ='" & Cmb_viaje.BoundText & "'" & _
                        "   and blcab_puer_emb ='" & Txt_pto_emb.Text & "'" & _
                        "   and blcab_puer_des ='" & Txt_pto_descarga.Text & "'" & _
                        "   and blcab_emp_resp ='" & Cmb_oper_carga.BoundText & "'" & _
                        "   and blcab_cono_emb ='" & Txt_bl.Text & "'"
      Gcon_docu.Execute sql
      
      sql = "Delete from docu_bldet_exp " & _
              "Where bldet_linea      ='" & cmb_linea.BoundText & "'" & _
               "   and bldet_buque    ='" & Cmb_buque.BoundText & "'" & _
               "   and bldet_viaje    ='" & Cmb_viaje.BoundText & "'" & _
               "   and bldet_puer_emb ='" & Txt_pto_emb.Text & "'" & _
               "   and bldet_puer_des ='" & Txt_pto_descarga.Text & "'" & _
               "   and bldet_emp_resp ='" & Cmb_oper_carga.BoundText & "'" & _
               "   and bldet_cono_emb ='" & Txt_bl.Text & "'"
      
End If
Gcon_docu.Execute sql
If CDbl(Lbl_no_det_tot.Caption) = 0 Then
    If MsgBox("No existe ningún detalle de carga ingresado, desea grabar solo los datos del Bill of lading ?", vbQuestion + vbYesNo) = vbNo Then
        Gcon_docu.RollbackTrans
        Exit Sub
    End If
Else
    For i = 1 To UBound(mani_bl_det)
        If mani_bl_det(i).elim <> "E" Then
            'sql = "insert into docu_bldet_exp (bldet_linea, bldet_buque, bldet_viaje," & _
                  " bldet_puer_emb,     bldet_puer_des,     bldet_emp_resp, " & _
                  " bldet_cono_emb,     bldet_secuencia,    bldet_tipo_carga, " & _
                  " bldet_peso,         bldet_bultos,       bldet_embalaje, " & _
                  " bldet_marcas_num,   bldet_desc_carg,    bldet_volumen," & _
                  " bldet_contenedor,   bldet_just_cambio,  bldet_imo, " & _
                  " bldet_pelig_onu,    bldet_sello,        bldet_tam_equi, " & _
                  " bldet_tip_equi,     bldet_cond_cont)"
                  
          sql = "insert into docu_bldet_exp " & _
            " VALUES ('" & cmb_linea.BoundText & "',    '" & Cmb_buque.BoundText & "',      '" & Cmb_viaje.BoundText & "', " & _
                  "'" & Txt_pto_emb.Text & "',       '" & Txt_pto_descarga.Text & "',     '" & Cmb_oper_carga.BoundText & "', " & _
                  "'" & Txt_bl.Text & "',             " & i & ",                          '" & mani_bl_det(i).tipo_carga & "', " & _
                  " " & mani_bl_det(i).peso & ",              " & mani_bl_det(i).no_bultos & ",          '" & Trim(mani_bl_det(i).embalaje) & "', " & _
                  "'" & Trim(mani_bl_det(i).marca_num) & "', '" & Trim(mani_bl_det(i).desc_carga) & "',   " & mani_bl_det(i).volumen & ", " & _
                  "'" & Trim$(mani_bl_det(i).CONTEINER) & "','" & Trim(mani_bl_det(i).IMO) & "'," & _
                  "'" & Trim(mani_bl_det(i).Sello) & "',       '" & Trim(mani_bl_det(i).tamano_equip) & "'," & _
                  "'" & Trim(mani_bl_det(i).cond_cont) & "')"
                  
            Gcon_docu.Execute sql
        End If
    Next i
End If
'Call actualiza_totales_manifiesto(Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, TXT_pto_emb.Text, Txt_pto_descarga.Text, Cmb_oper_carga.BoundText)
' ------ seccion de fletes ------
 sql = "Delete from docu_blflete_exp " & _
              "Where flete_linea      ='" & cmb_linea.BoundText & "'" & _
               "   and flete_buque    ='" & Cmb_buque.BoundText & "'" & _
               "   and flete_viaje    ='" & Cmb_viaje.BoundText & "'" & _
               "   and flete_puer_emb ='" & Txt_pto_emb.Text & "'" & _
               "   and flete_puer_des ='" & Txt_pto_descarga.Text & "'" & _
               "   and flete_emp_resp ='" & Cmb_oper_carga.BoundText & "'" & _
               "   and flete_cono_emb ='" & Txt_bl.Text & "'"
Gcon_docu.Execute sql

Call graba_fletes

Gcon_docu.CommitTrans
MsgBox "Grabación OK.", vbInformation
If Gstr_opc_bl = "A" Then Unload Me

Exit Sub
Errores:
    Gcon_docu.RollbackTrans
    MsgBox "ERROR " & Err.Description
    Exit Sub
    Resume 0
End Sub

Private Sub Cmd_hlp_consignatario_Click()
   Txt_id_consignatario.Text = Ayuda("docu_consignatario", "consig_codigo", "consig_nombre", " consig_tipo_id = '" & Cmb_tipo_id_consignatario.BoundText & "' and consig_status ='A' ", "Consignatario")
   Txt_id_consignatario.SetFocus
End Sub


Private Sub Cmd_hlp_nac_consignatario_Click()
    Txt_nac_consignatario.Text = Ayuda("docu_nacionalidad", "naci_codigo", "naci_descripcion", " naci_status ='A' ", "Nacionalidad")
End Sub

Private Sub Cmd_hlp_notificador_Click()
    Txt_id_notificador.Text = Ayuda("docu_consignatario", "consig_codigo", "consig_nombre", " consig_tipo_id = '" & Cmb_tipo_id_notificador.BoundText & "' and consig_status ='A' ", "Notificador")
    Txt_id_notificador.SetFocus
End Sub

Private Sub Cmd_hlp_pto_descarga_Click()
   Txt_pto_descarga.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Descarga")
End Sub

Private Sub Cmd_hlp_pto_destino_Click()
    Txt_pto_destino.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Destino")
End Sub

Private Sub Cmd_hlp_pto_origen_Click()
   Txt_pto_origen.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Origen")
End Sub

Private Sub Cmd_nuevo_Click()
    Gstr_opc_bl = "A"
    Call inicia_controles
    habilita_detalle False
End Sub

Private Sub Cmd_nuevo_det_Click()
    Call limpia_detalle
    If Lbl_no_det_tot.Caption = "0" Then
        ReDim Preserve mani_bl_det(1 To 1)
        inicia_arreglo_detalle 1
        Lbl_no_det_tot.Caption = "1"
        Lbl_no_det_act.Caption = "1"
        habilita_detalle True
    Else
        ReDim Preserve mani_bl_det(1 To CDbl(Lbl_no_det_tot.Caption) + 1)
        inicia_arreglo_detalle CDbl(Lbl_no_det_tot.Caption) + 1
        Lbl_no_det_tot.Caption = Str(CDbl(Lbl_no_det_tot.Caption) + 1)
        Lbl_no_det_act.Caption = Str(CDbl(Lbl_no_det_tot.Caption))
    End If
    
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).SECUENCIA = CDbl(Lbl_no_det_tot.Caption)
    Cmb_tipo_carga.BoundText = Gstr_param.tipo_carga
    Call Cmb_tipo_carga_LostFocus
End Sub

Private Sub Cmd_nuevo_flete_Click()
    Load doc_aux_flete_bl
    'doc_aux_flete_bl.Form_Load
    doc_aux_flete_bl.cmb_linea.BoundText = cmb_linea.BoundText
    doc_aux_flete_bl.Show 1
    If Gtyp_flete.reg_cancel = False Then
        Grid_flete.Row = Grid_flete.Rows - 1
        Grid_flete.Col = 0: Grid_flete.Text = Gtyp_flete.cargo_cod
        Grid_flete.Col = 1: Grid_flete.Text = Gtyp_flete.cargo
        Grid_flete.Col = 2: Grid_flete.Text = Gtyp_flete.deptor_cod
        Grid_flete.Col = 3: Grid_flete.Text = Gtyp_flete.debtor
        Grid_flete.Col = 4: Grid_flete.Text = Format(Gtyp_flete.valor, "###,###,##0.00")
        Grid_flete.Col = 5: Grid_flete.Text = Gtyp_flete.SIGNO
        Grid_flete.Col = 6: Grid_flete.Text = Gtyp_flete.MONEDA
        Grid_flete.Col = 7: Grid_flete.Text = Gtyp_flete.marcas
        Grid_flete.Col = 8: Grid_flete.Text = Gtyp_flete.lugar
        Grid_flete.Col = 9: Grid_flete.Text = IIf(Gtyp_flete.imp_bl, "S", "N")
        Grid_flete.Rows = Grid_flete.Rows + 1
    End If
End Sub

Private Sub Cmd_primer_det_Click()
    If Lbl_no_det_tot.Caption <> "0" Then
        Lbl_no_det_act.Caption = "1"
        Call muestra_detalle
    End If
End Sub

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub Cmd_siguiente_det_Click()
If CDbl(Lbl_no_det_act.Caption) < CDbl(Lbl_no_det_tot.Caption) Then
    Lbl_no_det_act.Caption = Str(CDbl(Lbl_no_det_act.Caption) + 1)
    Call muestra_detalle
End If
End Sub

Private Sub Cmd_ultimo_det_Click()
   If Lbl_no_det_tot.Caption <> "0" Then
        Lbl_no_det_act.Caption = Lbl_no_det_tot.Caption
        Call muestra_detalle
   End If
End Sub

Private Function valida_Datos() As Boolean
On Error GoTo ERROR_VALIDA

   valida_Datos = True
   If Trim$(cmb_linea.BoundText) = "" Then
      MsgBox "Ingrese el Valor de Linea", vbInformation, App.Title
      valida_Datos = False
      If cmb_linea.Enabled Then cmb_linea.SetFocus
      Exit Function
   End If
   
   If Trim$(Cmb_buque.BoundText) = "" Then
      MsgBox "Ingrese el Valor de buque", vbInformation, App.Title
      valida_Datos = False
      If Cmb_buque.Enabled Then Cmb_buque.SetFocus
      Exit Function
   End If
   If Trim$(Cmb_viaje.BoundText) = "" Then
      MsgBox "Ingrese el Código de Viaje", vbInformation, App.Title
      valida_Datos = False
      If Cmb_viaje.Enabled Then Cmb_viaje.SetFocus
      Exit Function
   End If
   If Trim$(Txt_pto_emb.Text) = "" Then
      MsgBox "Ingrese el Código del puerto embarque", vbInformation, App.Title
      valida_Datos = False
      If Txt_pto_emb.Enabled Then Txt_pto_emb.SetFocus
      Exit Function
   End If
   If Trim$(Cmb_oper_carga.BoundText) = "" Then
      MsgBox "Ingrese el Código del operador carga ", vbInformation, App.Title
      valida_Datos = False
      If Cmb_oper_carga.Enabled Then Cmb_oper_carga.SetFocus
      Exit Function
   End If
   If Trim$(Cmb_tipo_id_consignatario.BoundText) = "" Then
      MsgBox "Ingrese el Código de identificador del consignatario", vbInformation, App.Title
      valida_Datos = False
      If Cmb_tipo_id_cons.Enabled Then Cmb_tipo_id_cons.SetFocus
      Exit Function
   End If
   If Trim$(Cmb_tipo_carga.BoundText) = "" Then
      MsgBox "Ingrese el Código del tipo de carga", vbInformation, App.Title
      valida_Datos = False
      If Cmb_tipo_carga.Enabled Then Cmb_tipo_carga.SetFocus
      Exit Function
   End If
   If Trim$(Cmb_embalaje.BoundText) = "" Then
      MsgBox "Ingrese el Código del embalaje", vbInformation, App.Title
      valida_Datos = False
      If Cmb_embalaje.Enabled Then Cmb_embalaje.SetFocus
      Exit Function
   End If
   If Trim$(Cmb_IMO.BoundText) = "" Then
      MsgBox "Ingrese el Código de la peligrosidad", vbInformation, App.Title
      valida_Datos = False
      If Cmb_IMO.Enabled Then Cmb_IMO.SetFocus
      Exit Function
   End If
   If Trim$(Txt_bl.Text) = "" Then
      MsgBox "Ingrese el número del BL", vbInformation, App.Title
      valida_Datos = False
      If Txt_bl.Enabled Then Txt_bl.SetFocus
      Exit Function
   End If
   If Trim$(Txt_pto_origen.Text) = "" Then
      MsgBox "Ingrese el Código del puerto de origen", vbInformation, App.Title
      valida_Datos = False
      If Txt_pto_origen.Enabled Then Txt_pto_origen.SetFocus
      Exit Function
   End If
   If Trim$(Txt_pto_descarga.Text) = "" Then
      MsgBox "Ingrese el Código del puerto de descarga", vbInformation, App.Title
      valida_Datos = False
      If Txt_pto_descarga.Enabled Then Txt_pto_descarga.SetFocus
      Exit Function
   End If
   If Trim$(Txt_pto_destino.Text) = "" Then
      MsgBox "Ingrese el Código del puerto de destino", vbInformation, App.Title
      valida_Datos = False
      If Txt_pto_destino.Enabled Then Txt_pto_destino.SetFocus
      Exit Function
   End If
'   If Trim$(Txt_id_consignatario.Text) = "" Then
'      MsgBox "Ingrese el tipo de identificacion del consignatario", vbInformation, App.Title
'      valida_Datos = False
'      If Txt_id_consignatario.Enabled Then Txt_id_consignatario.SetFocus
'      Exit Function
'   End If
   If Trim$(Txt_nombre_consignatario.Text) = "" Then
      MsgBox "Ingrese el nombre del consignatario", vbInformation, App.Title
      valida_Datos = False
      If Txt_nombre_cons.Enabled Then Txt_nombre_cons.SetFocus
      Exit Function
   End If
   If Trim$(Cmb_tipo_id_consignatario.BoundText) <> "5" Then
      If Trim$(Txt_nac_consignatario.Text) = "" Then
         MsgBox "Ingrese la nacionalidad del consignatario", vbInformation, App.Title
         valida_Datos = False
         If Txt_nac_consigna.Enabled Then Txt_nac_consigna.SetFocus
         Exit Function
      End If
   End If
   If Trim$(Txt_direc_consignatario.Text) = "" Then
      MsgBox "Ingrese la direccion del consignatario", vbInformation, App.Title
      valida_Datos = False
      If Txt_direc_consig.Enabled Then Txt_direc_consig.SetFocus
      Exit Function
   End If
   If Trim$(Txt_nombre_embarcador.Text) = "" Then
      MsgBox "Ingrese el nombre del embarcador", vbInformation, App.Title
      valida_Datos = False
      If Txt_nombre_emb.Enabled Then Txt_nombre_emb.SetFocus
      Exit Function
   End If
   If Trim$(Txt_direc_embarcador.Text) = "" Then
      MsgBox "Ingrese la direccion del embarcador", vbInformation, App.Title
      valida_Datos = False
      If Txt_direc_embarc.Enabled Then Txt_direc_embarc.SetFocus
      Exit Function
   End If
   ''''''''''''''''''''''''''''''''''''''''''''''''''
   'TL Tengo que hacer un lazo para revisar la carga'
   ''''''''''''''''''''''''''''''''''''''''''''''''''
   For i = 1 To UBound(mani_bl_det)
        If Trim$(mani_bl_det(i).desc_carga) = "" Then
           MsgBox "Ingrese la descripcion de la carga del detalle: " & i, vbInformation, App.Title
           valida_Datos = False
           Exit Function
        End If
   Next i
      
Exit Function
ERROR_VALIDA:
MsgBox "Error :" & Err.Description, vbInformation, App.Title
Exit Function
Resume 0
End Function

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    carga_datacombo data_linea, cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
    If Not data_linea.Recordset.EOF Then data_linea.Recordset.MoveFirst
    carga_datacombo data_tipo_id_consignatario, Cmb_tipo_id_consignatario, "Select * from docu_tipo_id", "tipo_id_codigo", "tipo_id_descripcion"
    carga_datacombo data_tipo_id_notificador, Cmb_tipo_id_notificador, "Select * from docu_tipo_id", "tipo_id_codigo", "tipo_id_descripcion"
    Call inicia_controles
    If Gstr_opc_bl = "A" Then
        habilita_detalle False
    Else
        habilita_detalle True
    End If
    Carga_cabecera_fletes
End Sub

Private Sub Grid_flete_DblClick()
    If Grid_flete.Row <= 0 Then Exit Sub
    Load doc_aux_flete_bl
    'doc_aux_flete_bl.Form_Load
    doc_aux_flete_bl.cmb_linea.BoundText = cmb_linea.BoundText
    Grid_flete.Col = 0: doc_aux_flete_bl.Cmb_cargo.BoundText = Grid_flete.Text
    Grid_flete.Col = 2: doc_aux_flete_bl.Cmb_debtor.BoundText = Grid_flete.Text
    Grid_flete.Col = 4: doc_aux_flete_bl.GTM_flete.Value = CDbl(Grid_flete.Text)
    Grid_flete.Col = 5: doc_aux_flete_bl.Cmb_signo.ListIndex = IIf(Grid_flete.Text = "+", 0, IIf(Grid_flete.Text = "-", 1, -1))
    Grid_flete.Col = 6: doc_aux_flete_bl.Cmb_moneda.BoundText = Grid_flete.Text
    Grid_flete.Col = 7: doc_aux_flete_bl.Txt_marcas.Text = Grid_flete.Text
    Grid_flete.Col = 8: doc_aux_flete_bl.Txt_pto.Text = Grid_flete.Text
    Grid_flete.Col = 9: doc_aux_flete_bl.Chk_impr.Value = IIf(Grid_flete.Text = "S", 1, 0)
    
    
    doc_aux_flete_bl.Show 1
    If Gtyp_flete.reg_cancel = False Then
        Grid_flete.Col = 0: Grid_flete.Text = Gtyp_flete.cargo_cod
        Grid_flete.Col = 1: Grid_flete.Text = Gtyp_flete.cargo
        Grid_flete.Col = 2: Grid_flete.Text = Gtyp_flete.deptor_cod
        Grid_flete.Col = 3: Grid_flete.Text = Gtyp_flete.debtor
        Grid_flete.Col = 4: Grid_flete.Text = Format(Gtyp_flete.valor, "###,###,##0.00")
        Grid_flete.Col = 5: Grid_flete.Text = Gtyp_flete.SIGNO
        Grid_flete.Col = 6: Grid_flete.Text = Gtyp_flete.MONEDA
        Grid_flete.Col = 7: Grid_flete.Text = Gtyp_flete.marcas
        Grid_flete.Col = 8: Grid_flete.Text = Gtyp_flete.lugar
        Grid_flete.Col = 9: Grid_flete.Text = IIf(Gtyp_flete.imp_bl, "S", "N")
        'Grid_flete.Rows = Grid_flete.Rows + 1
    End If
    
End Sub

Private Sub GTM_bultos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_bultos_LostFocus()
   mani_bl_det(CDbl(Lbl_no_det_act.Caption)).no_bultos = GTM_bultos.Value
   actualiza_totales_bl
End Sub

Private Sub GTM_peso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_peso_LostFocus()
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).peso = GTM_peso.Value
    actualiza_totales_bl
End Sub

Private Sub GTM_volumen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_volumen_LostFocus()
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).volumen = GTM_volumen.Value
    actualiza_totales_bl
End Sub

Private Sub Msk_fecha_emb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_bl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_contenedor_KeyPress(KeyAscii As Integer)
    If Len(Trim$(Txt_contenedor.Text)) >= 4 And Not InStr(1, "0123456789" & Chr(32) & Chr(8), Chr(KeyAscii)) <> 0 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_contenedor_LostFocus()
    If Trim(Txt_contenedor.Text) <> "" Then
        mani_bl_det(CDbl(Lbl_no_det_act.Caption)).CONTEINER = Txt_contenedor.Text
        'Call ubica_contenedor
    End If
    actualiza_totales_bl
End Sub

Private Sub Txt_desc_carga_LostFocus()
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).desc_carga = Txt_desc_carga.Text
End Sub

Private Sub Txt_direc_consignatario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_direc_embarcador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_direc_notificador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_id_consignatario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_id_consignatario_LostFocus()
On Error GoTo Errores
    Dim sql As String
    Dim reg As New ADODB.Recordset
    If Trim$(Txt_id_consignatario.Text) <> "" Then
        sql = "Select * from docu_consignatario " & _
              " Where consig_tipo_id = '" & Cmb_tipo_id_consignatario.BoundText & "'" & _
              "   And consig_codigo  = '" & Trim$(Txt_id_consignatario.Text) & "'"
        
        reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        If Not reg.EOF Then
            Txt_nac_consignatario.Text = reg.Fields("consig_nacional")
            Txt_nombre_consignatario = reg.Fields("consig_nombre")
            Txt_direc_consignatario = reg.Fields("consig_direccion")
        End If
        reg.Close
    End If
    Exit Sub
Errores:
    MsgBox "Ocurrio el Error " & Err.Description, vbInformation
    Exit Sub
    Resume 0
End Sub

Private Sub Txt_id_notificador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_id_notificador_LostFocus()
On Error GoTo Errores
    Dim sql As String
    Dim reg As New ADODB.Recordset
    If Trim$(Txt_id_notificador.Text) <> "" Then
        sql = "Select * from docu_consignatario " & _
              " Where consig_tipo_id = '" & Cmb_tipo_id_notificador.BoundText & "'" & _
              "   And consig_codigo  = '" & Trim$(Txt_id_notificador.Text) & "'"
        
        reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        If Not reg.EOF Then
            Txt_nombre_notificador.Text = reg.Fields("consig_nombre")
            Txt_direc_notificador.Text = reg.Fields("consig_direccion")
        End If
        reg.Close
    End If
    Exit Sub
Errores:
    MsgBox "Ocurrio el Error " & Err.Description, vbInformation
    Exit Sub
    Resume 0
End Sub

Private Sub limpia_detalle()
    Cmb_tipo_carga.BoundText = ""
    Txt_contenedor.Text = ""
    GTM_bultos.Value = 0
    Cmb_embalaje.BoundText = ""
    GTM_peso.Value = 0
    GTM_volumen.Value = 0
    Txt_marcas.Text = ""
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'TL pongo en comentarios para mantener la descripcion  ingresada
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Txt_desc_carga.Text = ""
    
    Cmb_IMO.BoundText = ""
    Cmb_tam_Equip.BoundText = ""
    cmb_condicion.BoundText = ""
    Txt_sello1.Text = ""
End Sub

Private Sub Txt_marcas_LostFocus()
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).marca_num = Txt_marcas.Text
End Sub

Private Sub muestra_detalle()
    Cmb_tipo_carga.BoundText = Trim(mani_bl_det(CDbl(Lbl_no_det_act.Caption)).tipo_carga)
    Txt_contenedor.Text = mani_bl_det(CDbl(Lbl_no_det_act.Caption)).CONTEINER
    GTM_bultos.Value = mani_bl_det(CDbl(Lbl_no_det_act.Caption)).no_bultos
    Cmb_embalaje.BoundText = (mani_bl_det(CDbl(Lbl_no_det_act.Caption)).embalaje)
    GTM_peso.Value = mani_bl_det(CDbl(Lbl_no_det_act.Caption)).peso
    GTM_volumen.Value = mani_bl_det(CDbl(Lbl_no_det_act.Caption)).volumen
    Txt_marcas.Text = mani_bl_det(CDbl(Lbl_no_det_act.Caption)).marca_num
    Txt_desc_carga.Text = mani_bl_det(CDbl(Lbl_no_det_act.Caption)).desc_carga
    Cmb_IMO.BoundText = Trim(mani_bl_det(CDbl(Lbl_no_det_act.Caption)).IMO)
    Cmb_tam_Equip.BoundText = Trim(mani_bl_det(CDbl(Lbl_no_det_act.Caption)).tamano_equip)
    cmb_condicion.BoundText = Trim(mani_bl_det(CDbl(Lbl_no_det_act.Caption)).cond_cont)
    Txt_sello1.Text = Trim(mani_bl_det(CDbl(Lbl_no_det_act.Caption)).Sello)
End Sub

Public Sub habilita_detalle(opc As Boolean)
    Cmb_tipo_carga.Enabled = opc
    Txt_contenedor.Enabled = opc
    GTM_bultos.Enabled = opc
    Cmb_embalaje.Enabled = opc
    GTM_peso.Enabled = opc
    GTM_volumen.Enabled = opc
    Txt_marcas.Enabled = opc
    Txt_desc_carga.Enabled = opc
    Cmb_IMO.Enabled = opc
    Cmb_tam_Equip.Enabled = opc
    cmb_condicion.Enabled = opc
    Txt_sello1.Enabled = opc

End Sub
' PROCEDIMIENTO QUE SIRVE PARA CARGAR EL BL Y PRESENTARLO, ESTO DEBE SER INVOCADO DESDE LA PANTALLA DE BL STATUS
Public Sub carga_bl(linea As String, buque As String, viaje As String, _
                    pto_emb As String, pto_desc As String, bl As String)
    Dim sql         As String
    Dim reg         As New ADODB.Recordset
    Dim REG_det     As New ADODB.Recordset
    Dim REG_flete   As New ADODB.Recordset
    Dim i           As Integer
    
    sql = "Select * " & _
               " from docu_blcab_exp " & _
               " Where blcab_linea       ='" & linea & "'" & _
               "   and blcab_buque       ='" & buque & "'" & _
               "   and blcab_viaje       ='" & viaje & "'" & _
               "   and blcab_puer_emb    ='" & pto_emb & "'" & _
               "   and blcab_puer_des    ='" & pto_desc & "'" & _
               "   and blcab_cono_emb    ='" & bl & "'"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    If Not reg.EOF Then
        cmb_linea.BoundText = reg("blcab_linea")
        Cmb_buque.BoundText = reg("blcab_buque")
        Cmb_viaje.BoundText = reg("blcab_viaje")
        
        Txt_pto_emb.Text = reg("blcab_puer_emb")
        Txt_bl.Text = reg("blcab_cono_emb")
        Txt_pto_origen.Text = reg("blcab_puer_origen")
        Txt_pto_descarga.Text = reg("blcab_puer_des")
        Txt_pto_destino.Text = reg("blcab_pto_dest_final")
        Msk_fecha_emb.Text = reg("blcab_fecha_emb")
        Cmb_tipo_id_consignatario.BoundText = IIf(Not IsNull(reg("blcab_tipo_ide_cons")), reg("blcab_tipo_ide_cons"), "")
        Txt_id_consignatario.Text = IIf(Not IsNull(reg("blcab_ide_cons")), reg("blcab_ide_cons"), "")
        Txt_nombre_consignatario.Text = IIf(Not IsNull(reg("blcab_consig")), reg("blcab_consig"), "")
        Txt_nac_consignatario.Text = IIf(Not IsNull(reg("blcab_nacio_consig")), reg("blcab_nacio_consig"), "")
        Txt_direc_consignatario.Text = IIf(Not IsNull(reg("blcab_direc_consig")), reg("blcab_direc_consig"), "")
        
        Cmb_tipo_id_notificador.BoundText = IIf(Not IsNull(reg("blcab_tipo_ide_notif")), reg("blcab_tipo_ide_notif"), "")
        Txt_id_notificador.Text = IIf(Not IsNull(reg("blcab_ide_notif")), reg("blcab_ide_notif"), "")
        Txt_nombre_notificador.Text = IIf(Not IsNull(reg("blcab_notif")), reg("blcab_notif"), "")
        Txt_direc_notificador.Text = IIf(Not IsNull(reg("blcab_direc_notif")), reg("blcab_direc_notif"), "")
        
        Txt_nombre_embarcador.Text = IIf(Not IsNull(reg("blcab_embarcador")), reg("blcab_embarcador"), "")
        Txt_direc_embarcador.Text = IIf(Not IsNull(reg("blcab_direc_embarc")), reg("blcab_direc_embarc"), "")
        
        Txt_peso_tot.Text = Format(IIf(Not IsNull(reg("blcab_peso_tot")), reg("blcab_peso_tot"), "0"), "###,###,###.00")
        Txt_volumen_tot.Text = Format(IIf(Not IsNull(reg("blcab_volumen_tot")), reg("blcab_volumen_tot"), "0"), "###,###,###.00")
        Txt_bultos_tot.Text = Format(IIf(Not IsNull(reg("blcab_bultos_tot")), reg("blcab_bultos_tot"), "0"), "###,###,###.00")
        Txt_cont_tot.Text = Format(IIf(Not IsNull(reg("blcab_no_cont")), reg("blcab_no_cont"), "0"), "###,###,###.00")
        Cmb_oper_carga.BoundText = IIf(Not IsNull(reg("blcab_emp_resp")), reg("blcab_emp_resp"), "")

        sql = "Select * " & _
               " from docu_bldet_exp " & _
               " Where bldet_linea       ='" & linea & "'" & _
               "   and bldet_buque       ='" & buque & "'" & _
               "   and bldet_viaje       ='" & viaje & "'" & _
               "   and bldet_puer_emb    ='" & pto_emb & "'" & _
               "   and bldet_puer_des    ='" & pto_desc & "'" & _
               "   and bldet_cono_emb    ='" & bl & "'"
        REG_det.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        Erase mani_bl_det
        i = 1
        Do While Not REG_det.EOF
            ReDim Preserve mani_bl_det(1 To i)
            inicia_arreglo_detalle i
            mani_bl_det(i).tipo_carga = IIf(Not IsNull(REG_det("bldet_tipo_Carga")), REG_det("bldet_tipo_Carga"), "")
            mani_bl_det(i).CONTEINER = IIf(Not IsNull(REG_det("bldet_contenedor")), REG_det("bldet_contenedor"), "")
            mani_bl_det(i).no_bultos = IIf(Not IsNull(REG_det("bldet_bultos")), REG_det("bldet_bultos"), "0")
            mani_bl_det(i).embalaje = IIf(Not IsNull(REG_det("bldet_embalaje")), REG_det("bldet_embalaje"), "")
            mani_bl_det(i).peso = IIf(Not IsNull(REG_det("bldet_peso")), REG_det("bldet_peso"), "0")
            mani_bl_det(i).volumen = IIf(Not IsNull(REG_det("bldet_volumen")), REG_det("bldet_volumen"), "")
            mani_bl_det(i).marca_num = IIf(Not IsNull(REG_det("bldet_marcas_num")), REG_det("bldet_marcas_num"), "")
            mani_bl_det(i).desc_carga = IIf(Not IsNull(REG_det("bldet_desc_carg")), REG_det("bldet_desc_carg"), "")
            mani_bl_det(i).IMO = IIf(Not IsNull(REG_det("bldet_imo")), REG_det("bldet_imo"), "")
            mani_bl_det(i).cond_cont = IIf(Not IsNull(REG_det("bldet_cond_cont")), REG_det("bldet_cond_cont"), "")
            mani_bl_det(i).Sello = IIf(Not IsNull(REG_det("bldet_sello")), REG_det("bldet_sello"), "")
            mani_bl_det(i).tamano_equip = IIf(Not IsNull(REG_det("bldet_tam_equi")), REG_det("bldet_tam_equi"), "")
            REG_det.MoveNext
            i = i + 1
        Loop
        If i > 1 Then
            Lbl_no_det_tot.Caption = i - 1
            Lbl_no_det_act.Caption = 1
            muestra_detalle
        Else
            Lbl_no_det_tot.Caption = 0
            Lbl_no_det_act.Caption = 0
        End If
        
        sql = "Select * " & _
               " from docu_blflete_exp " & _
               " Where flete_linea       ='" & linea & "'" & _
               "   and flete_buque       ='" & buque & "'" & _
               "   and flete_viaje       ='" & viaje & "'" & _
               "   and flete_puer_emb    ='" & pto_emb & "'" & _
               "   and flete_puer_des    ='" & pto_desc & "'" & _
               "   and flete_cono_emb    ='" & bl & "'"
        REG_flete.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        Do While Not REG_flete.EOF
            Grid_flete.Row = Grid_flete.Rows - 1
            If Not IsNull(REG_flete("flete_cargo")) Then
                Grid_flete.Col = 0: Grid_flete.Text = REG_flete("flete_cargo")
                Grid_flete.Col = 1: Grid_flete.Text = devuelve_desc("docu_cargo", "cargo_codigo", "cargo_descripcion", REG_flete("flete_cargo"), "cargo_linea = '" & cmb_linea.BoundText & "'")
            Else
                Grid_flete.Col = 0: Grid_flete.Text = ""
                Grid_flete.Col = 1: Grid_flete.Text = ""
            End If
            If Not IsNull(REG_flete("flete_debtor")) Then
                Grid_flete.Col = 2: Grid_flete.Text = REG_flete("flete_debtor")
                Grid_flete.Col = 3: Grid_flete.Text = devuelve_desc("docu_debtor", "debtor_codigo", "debtor_descripcion", REG_flete("flete_debtor"), "debtor_linea = '" & cmb_linea.BoundText & "'")
            Else
                Grid_flete.Col = 2: Grid_flete.Text = ""
                Grid_flete.Col = 3: Grid_flete.Text = ""
            End If
            Grid_flete.Col = 4: Grid_flete.Text = Format(REG_flete("flete_valor"), "###,###,##0.00")
            Grid_flete.Col = 5: Grid_flete.Text = REG_flete("flete_signo")
            Grid_flete.Col = 6: Grid_flete.Text = IIf(Not IsNull(REG_flete("flete_moneda")), REG_flete("flete_moneda"), "")
            Grid_flete.Col = 7: Grid_flete.Text = IIf(Not IsNull(REG_flete("flete_marcas")), REG_flete("flete_marcas"), "")
            Grid_flete.Col = 8: Grid_flete.Text = IIf(Not IsNull(REG_flete("flete_lugar")), REG_flete("flete_lugar"), "")
            Grid_flete.Col = 9: Grid_flete.Text = REG_flete("flete_imp_bl")
            Grid_flete.Rows = Grid_flete.Rows + 1
            REG_flete.MoveNext
        Loop
    End If
End Sub

Public Sub inicia_arreglo_detalle(i As Integer)
    mani_bl_det(i).cond_cont = ""
    mani_bl_det(i).CONTEINER = ""
    mani_bl_det(i).desc_carga = ""
    mani_bl_det(i).embalaje = ""
    mani_bl_det(i).IMO = ""
    mani_bl_det(i).just_cambio = ""
    mani_bl_det(i).marca_num = ""
    mani_bl_det(i).no_bultos = 0
    mani_bl_det(i).pelig_ONU = ""
    mani_bl_det(i).peso = 0
    mani_bl_det(i).SECUENCIA = 0
    mani_bl_det(i).Sello = ""
    mani_bl_det(i).tamano_equip = ""
    mani_bl_det(i).tipo_equip = ""
    mani_bl_det(i).tipo_carga = ""
    mani_bl_det(i).Tipo_cont = ""
    mani_bl_det(i).volumen = 0
End Sub

Public Sub inicia_controles()
    For i% = 0 To Me.Count - 1
        If TypeOf Me.Controls(i%) Is TextBox Then
            Me.Controls(i%).Text = ""
        ElseIf TypeOf Me.Controls(i%) Is DataCombo Then
            Me.Controls(i%).BoundText = ""
            Me.Controls(i%).Text = ""
        End If
    Next i%
    
    GTM_bultos.Value = 0
    GTM_peso.Value = 0
    GTM_volumen.Value = 0
        
    Txt_cont_tot = "0"
    Me.Txt_peso_tot = "0"
    Me.Txt_volumen_tot = "0"
    Me.Txt_bultos_tot = "0"
    
    Erase mani_bl_det
    Lbl_no_det_act.Caption = "0"
    Lbl_no_det_tot.Caption = "0"
End Sub

Private Sub actualiza_totales_bl()
Dim peso    As Double
Dim volumen As Double
Dim cont    As Double
Dim bultos  As Double

    If Lbl_no_det_tot.Caption <> "0" Then
        peso = 0
        volumen = 0
        cont = 0
        bultos = 0
       For i = 1 To UBound(mani_bl_det)
          If mani_bl_det(i).elim <> "E" Then
              peso = peso + mani_bl_det(i).peso
              volumen = volumen + mani_bl_det(i).volumen
              bultos = bultos + mani_bl_det(i).no_bultos
              If Trim$(mani_bl_det(i).CONTEINER) <> "" Then cont = cont + 1
          End If
       Next i
       Txt_bultos_tot.Text = Format(bultos, "###,###,###.00")
       Txt_peso_tot.Text = Format(peso, "###,###,###.00")
       Txt_volumen_tot.Text = Format(volumen, "###,###,###.00")
       Txt_cont_tot.Text = Format(cont, "###,###,###.00")
       
    End If
End Sub


Private Sub Txt_nac_consignatario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_nombre_consignatario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_nombre_embarcador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_nombre_notificador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_pto_descarga_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_pto_destino_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_pto_origen_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Sub Carga_cabecera_fletes()
    Grid_flete.Cols = 10
    Grid_flete.Row = 0
    Grid_flete.Col = 0: Grid_flete.Text = "Cargo Cod":   Grid_flete.ColWidth(0) = 600
    Grid_flete.Col = 1: Grid_flete.Text = "Cargo Desc":  Grid_flete.ColWidth(1) = 1000
    Grid_flete.Col = 2: Grid_flete.Text = "Debtor Cod":  Grid_flete.ColWidth(2) = 600
    Grid_flete.Col = 3: Grid_flete.Text = "Debtor Desc": Grid_flete.ColWidth(3) = 1000
    Grid_flete.Col = 4: Grid_flete.Text = "Valor":       Grid_flete.ColWidth(4) = 900
    Grid_flete.Col = 5: Grid_flete.Text = "Signo":       Grid_flete.ColWidth(5) = 600
    Grid_flete.Col = 6: Grid_flete.Text = "Moneda":      Grid_flete.ColWidth(6) = 700
    Grid_flete.Col = 7: Grid_flete.Text = "Marcas":      Grid_flete.ColWidth(7) = 2000
    Grid_flete.Col = 8: Grid_flete.Text = "Lugar Pago":  Grid_flete.ColWidth(8) = 900
    Grid_flete.Col = 9: Grid_flete.Text = "Impr Bl":     Grid_flete.ColWidth(9) = 700
End Sub

Sub graba_fletes()
On Error GoTo Errores
    
    If Grid_flete.Rows <= 2 Then Exit Sub
    
    For i = 1 To Grid_flete.Rows - 2
        sql = "Insert into docu_blflete_exp values(" & _
              "'" & cmb_linea.BoundText & "',  '" & Cmb_buque.BoundText & "'," & _
              "'" & Cmb_viaje.BoundText & "',  '" & Txt_pto_emb.Text & "'," & _
              "'" & Txt_pto_descarga.Text & "','" & Cmb_oper_carga.BoundText & "', " & _
              "'" & Txt_bl.Text & "'," & i & ","
        Grid_flete.Col = 0: sql = sql & "'" & Grid_flete.Text & "', "
        Grid_flete.Col = 2: sql = sql & "'" & Grid_flete.Text & "', "
        Grid_flete.Col = 4: sql = sql & " " & CDbl(Grid_flete.Text) & ",  "
        Grid_flete.Col = 5: sql = sql & "'" & Grid_flete.Text & "', "
        Grid_flete.Col = 6: sql = sql & "'" & Grid_flete.Text & "', "
        Grid_flete.Col = 7: sql = sql & "'" & Grid_flete.Text & "', "
        Grid_flete.Col = 8: sql = sql & "'" & Grid_flete.Text & "', "
        Grid_flete.Col = 9: sql = sql & "'" & Grid_flete.Text & "') "
        Gcon_docu.Execute sql
    Next i
    
    Exit Sub
Errores:
    MsgBox "Ocurrio el Error " & Err.Description
    Exit Sub
    Resume 0
End Sub

Private Sub Txt_sello1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_sello1_LostFocus()
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).Sello = Txt_sello1.Text
End Sub
