VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "GTDATE32.OCX"
Object = "{667E8C40-F9B5-11CF-90AB-444553540000}#1.0#0"; "GTNUM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form doc_bl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill of lading de Importación"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   Icon            =   "Doc_bl.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   10425
   Begin VB.CommandButton Cmd_salir 
      Height          =   465
      Left            =   1875
      Picture         =   "Doc_bl.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   0
      Width           =   465
   End
   Begin VB.CommandButton Cmd_cancelar 
      Height          =   465
      Left            =   1410
      Picture         =   "Doc_bl.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   0
      Width           =   465
   End
   Begin VB.CommandButton Cmd_grabar 
      Height          =   465
      Left            =   945
      Picture         =   "Doc_bl.frx":123E
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   0
      Width           =   465
   End
   Begin VB.CommandButton Cmd_eliminar 
      Height          =   465
      Left            =   465
      Picture         =   "Doc_bl.frx":18A8
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   0
      Width           =   465
   End
   Begin VB.CommandButton Cmd_nuevo 
      Height          =   465
      Left            =   0
      Picture         =   "Doc_bl.frx":6D32
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   0
      Width           =   465
   End
   Begin TabDlg.SSTab STab_bl 
      Height          =   7035
      Left            =   60
      TabIndex        =   82
      Top             =   525
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   12409
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   503
      TabCaption(0)   =   "Datos Generales Bl"
      TabPicture(0)   =   "Doc_bl.frx":739C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame12"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Datos Aduaneros"
      TabPicture(1)   =   "Doc_bl.frx":73B8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame10"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Fletes"
      TabPicture(2)   =   "Doc_bl.frx":73D4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).Control(1)=   "Frame21"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Detalle de Bl"
      TabPicture(3)   =   "Doc_bl.frx":73F0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Cmd_ultimo_det"
      Tab(3).Control(1)=   "Cmd_siguiente_det"
      Tab(3).Control(2)=   "Cmd_anterior_det"
      Tab(3).Control(3)=   "Cmd_primer_det"
      Tab(3).Control(4)=   "Cmd_eliminar_det"
      Tab(3).Control(5)=   "Cmd_nuevo_det"
      Tab(3).Control(6)=   "Frame29"
      Tab(3).Control(7)=   "Label47"
      Tab(3).Control(8)=   "Lbl_no_det_tot"
      Tab(3).Control(9)=   "Lbl_no_det_act"
      Tab(3).ControlCount=   10
      Begin VB.Frame Frame12 
         Height          =   1335
         Left            =   7080
         TabIndex        =   200
         Top             =   960
         Width           =   2295
         Begin VB.Frame Frame18 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   120
            TabIndex        =   203
            Top             =   120
            Width           =   1590
            Begin VB.CommandButton Cmd_hlp_pto_trasb 
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
               Picture         =   "Doc_bl.frx":740C
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Buscar"
               Top             =   255
               Width           =   300
            End
            Begin VB.TextBox Txt_pto_trasb 
               Height          =   285
               Left            =   120
               MaxLength       =   5
               TabIndex        =   13
               Top             =   255
               Width           =   975
            End
            Begin VB.Label Label62 
               AutoSize        =   -1  'True
               Caption         =   "Puerto Trasbordo"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   204
               Top             =   15
               Width           =   1230
            End
         End
         Begin VB.Frame Frame17 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   585
            Left            =   270
            TabIndex        =   201
            Top             =   720
            Width           =   1380
            Begin GTMaskDate.GTMaskDate Msk_fecha_trasb 
               Height          =   285
               Left            =   30
               TabIndex        =   15
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
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Trasb"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   30
               TabIndex        =   202
               Top             =   30
               Width           =   900
            End
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   " Flete Detallado "
         Height          =   3150
         Left            =   -74790
         TabIndex        =   196
         Top             =   2985
         Width           =   9750
         Begin VB.CommandButton Cmd_nuevo_flete 
            Height          =   465
            Left            =   8565
            Picture         =   "Doc_bl.frx":793E
            Style           =   1  'Graphical
            TabIndex        =   199
            Top             =   2625
            Width           =   465
         End
         Begin VB.CommandButton Cmd_eliminar_flete 
            Height          =   465
            Left            =   9030
            Picture         =   "Doc_bl.frx":7FA8
            Style           =   1  'Graphical
            TabIndex        =   198
            Top             =   2625
            Width           =   465
         End
         Begin MSGrid.Grid Grid_flete 
            Height          =   2250
            Left            =   150
            TabIndex        =   197
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
      Begin VB.CommandButton Cmd_ultimo_det 
         Height          =   465
         Left            =   -72345
         Picture         =   "Doc_bl.frx":83EA
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   615
         Width           =   465
      End
      Begin VB.CommandButton Cmd_siguiente_det 
         Height          =   465
         Left            =   -72810
         Picture         =   "Doc_bl.frx":84EC
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   615
         Width           =   465
      End
      Begin VB.CommandButton Cmd_anterior_det 
         Height          =   465
         Left            =   -73275
         Picture         =   "Doc_bl.frx":85EE
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   615
         Width           =   465
      End
      Begin VB.CommandButton Cmd_primer_det 
         Height          =   465
         Left            =   -73740
         Picture         =   "Doc_bl.frx":86F0
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   615
         Width           =   465
      End
      Begin VB.CommandButton Cmd_eliminar_det 
         Height          =   465
         Left            =   -74205
         Picture         =   "Doc_bl.frx":87F2
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   615
         Width           =   465
      End
      Begin VB.CommandButton Cmd_nuevo_det 
         Height          =   465
         Left            =   -74670
         Picture         =   "Doc_bl.frx":8C34
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   615
         Width           =   465
      End
      Begin VB.Frame Frame29 
         Height          =   5535
         Left            =   -74640
         TabIndex        =   155
         Top             =   1095
         Width           =   9585
         Begin VB.Frame Frame19 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   120
            TabIndex        =   207
            Top             =   3240
            Width           =   3705
            Begin MSDataListLib.DataCombo Cmb_tipo_prod 
               Height          =   315
               Left            =   105
               TabIndex        =   63
               Top             =   225
               Width           =   3360
               _ExtentX        =   5927
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin MSAdodcLib.Adodc data_tipo_prod 
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
            Begin VB.Label Label65 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Producto"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   208
               Top             =   0
               Width           =   1230
            End
         End
         Begin VB.Frame Frame42 
            BorderStyle     =   0  'None
            Height          =   1575
            Left            =   3960
            TabIndex        =   176
            Top             =   1680
            Width           =   3810
            Begin VB.TextBox Txt_desc_carga 
               Height          =   1260
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   62
               Top             =   195
               Width           =   3585
            End
            Begin VB.Label Label58 
               AutoSize        =   -1  'True
               Caption         =   "Descripción carga"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   177
               Top             =   -30
               Width           =   1290
            End
         End
         Begin VB.Frame Frame43 
            BorderStyle     =   0  'None
            Height          =   840
            Left            =   75
            TabIndex        =   174
            Top             =   3885
            Width           =   5610
            Begin VB.TextBox Txt_just_cambio 
               Height          =   600
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   64
               Top             =   195
               Width           =   5385
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               Caption         =   "Justificación cambio/Observación"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   175
               Top             =   -30
               Width           =   2400
            End
         End
         Begin VB.Frame Frame41 
            BorderStyle     =   0  'None
            Height          =   1575
            Left            =   75
            TabIndex        =   172
            Top             =   1665
            Width           =   3810
            Begin VB.TextBox Txt_marcas 
               Height          =   1260
               Left            =   135
               MultiLine       =   -1  'True
               TabIndex        =   61
               Top             =   195
               Width           =   3585
            End
            Begin VB.Label Label57 
               AutoSize        =   -1  'True
               Caption         =   "Marcas/Números"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   173
               Top             =   -30
               Width           =   1230
            End
         End
         Begin VB.Frame Frame40 
            BorderStyle     =   0  'None
            Height          =   510
            Left            =   3855
            TabIndex        =   170
            Top             =   270
            Width           =   1830
            Begin VB.CommandButton Cmd_hlp_cont 
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
               Left            =   1470
               Picture         =   "Doc_bl.frx":929E
               Style           =   1  'Graphical
               TabIndex        =   195
               ToolTipText     =   "Buscar"
               Top             =   195
               Width           =   300
            End
            Begin VB.TextBox Txt_contenedor 
               Height          =   285
               Left            =   105
               MaxLength       =   11
               TabIndex        =   56
               Top             =   195
               Width           =   1305
            End
            Begin VB.Label Label56 
               AutoSize        =   -1  'True
               Caption         =   "Contenedor"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   171
               Top             =   -30
               Width           =   825
            End
         End
         Begin VB.Frame Frame39 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   150
            TabIndex        =   168
            Top             =   180
            Width           =   3750
            Begin MSDataListLib.DataCombo Cmb_tipo_carga 
               Height          =   315
               Left            =   105
               TabIndex        =   55
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
               TabIndex        =   169
               Top             =   0
               Width           =   780
            End
         End
         Begin VB.Frame Frame38 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   120
            TabIndex        =   166
            Top             =   975
            Width           =   3705
            Begin MSDataListLib.DataCombo Cmb_embalaje 
               Height          =   315
               Left            =   105
               TabIndex        =   58
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
               TabIndex        =   167
               Top             =   0
               Width           =   645
            End
         End
         Begin VB.Frame Frame37 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   3885
            TabIndex        =   164
            Top             =   990
            Width           =   1695
            Begin GTMaskNum.GTMaskNum GTM_peso 
               Height          =   315
               Left            =   75
               TabIndex        =   59
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
               TabIndex        =   165
               Top             =   0
               Width           =   360
            End
         End
         Begin VB.Frame Frame35 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   120
            TabIndex        =   162
            Top             =   4815
            Width           =   3705
            Begin MSDataListLib.DataCombo Cmb_IMO 
               Height          =   315
               Left            =   105
               TabIndex        =   65
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
               TabIndex        =   163
               Top             =   -30
               Width           =   300
            End
         End
         Begin VB.Frame Frame34 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   5715
            TabIndex        =   160
            Top             =   240
            Width           =   1695
            Begin GTMaskNum.GTMaskNum GTM_bultos 
               Height          =   315
               Left            =   75
               TabIndex        =   57
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
               TabIndex        =   161
               Top             =   -15
               Width           =   720
            End
         End
         Begin VB.Frame Frame33 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   4080
            TabIndex        =   158
            Top             =   4875
            Width           =   3780
            Begin MSDataListLib.DataCombo Cmb_ONU 
               Height          =   315
               Left            =   105
               TabIndex        =   66
               Top             =   240
               Width           =   3345
               _ExtentX        =   5900
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin MSAdodcLib.Adodc data_ONU 
               Height          =   330
               Left            =   2595
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
            Begin VB.Label Label49 
               AutoSize        =   -1  'True
               Caption         =   "Peligrosidad ONU"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   159
               Top             =   0
               Width           =   1260
            End
         End
         Begin VB.Frame Frame32 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   5715
            TabIndex        =   156
            Top             =   1020
            Width           =   1695
            Begin GTMaskNum.GTMaskNum GTM_volumen 
               Height          =   315
               Left            =   75
               TabIndex        =   60
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
               TabIndex        =   157
               Top             =   0
               Width           =   615
            End
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   " Flete General "
         Height          =   2340
         Left            =   -74775
         TabIndex        =   148
         Top             =   495
         Width           =   8910
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Enabled         =   0   'False
            Height          =   540
            Left            =   3225
            TabIndex        =   191
            Top             =   1530
            Width           =   2835
            Begin MSDataListLib.DataCombo Cmb_moneda_seguro 
               Height          =   315
               Left            =   105
               TabIndex        =   53
               Top             =   240
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin MSAdodcLib.Adodc data_moneda_seguro 
               Height          =   330
               Left            =   1590
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
            Begin VB.Label Label60 
               AutoSize        =   -1  'True
               Caption         =   "Moneda Seguro"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   192
               Top             =   0
               Width           =   1140
            End
         End
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Enabled         =   0   'False
            Height          =   540
            Left            =   6420
            TabIndex        =   189
            Top             =   1530
            Width           =   1695
            Begin GTMaskNum.GTMaskNum GTM_seguro 
               Height          =   315
               Left            =   75
               TabIndex        =   54
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
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               Caption         =   "Valor Seguro"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   90
               TabIndex        =   190
               Top             =   0
               Width           =   915
            End
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   180
            TabIndex        =   187
            Top             =   870
            Width           =   3075
            Begin MSDataListLib.DataCombo Cmb_metodo_pago_otros 
               Height          =   315
               Left            =   105
               TabIndex        =   50
               Top             =   240
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin MSAdodcLib.Adodc data_metodo_pago_otros 
               Height          =   330
               Left            =   1770
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
            Begin VB.Label Label48 
               AutoSize        =   -1  'True
               Caption         =   "Método Pago Otros"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   188
               Top             =   0
               Width           =   1380
            End
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   3225
            TabIndex        =   185
            Top             =   885
            Width           =   2835
            Begin MSDataListLib.DataCombo Cmb_moneda_otros 
               Height          =   315
               Left            =   105
               TabIndex        =   51
               Top             =   240
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin MSAdodcLib.Adodc data_moneda_otros 
               Height          =   330
               Left            =   1620
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
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               Caption         =   "Moneda Otros"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   186
               Top             =   0
               Width           =   1005
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   6420
            TabIndex        =   183
            Top             =   885
            Width           =   1695
            Begin GTMaskNum.GTMaskNum GTM_flete_otros 
               Height          =   315
               Left            =   75
               TabIndex        =   52
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
               Caption         =   "Valor Flete Otros"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   90
               TabIndex        =   184
               Top             =   0
               Width           =   1170
            End
         End
         Begin VB.Frame Frame25 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   6390
            TabIndex        =   153
            Top             =   210
            Width           =   1695
            Begin GTMaskNum.GTMaskNum GTM_flete 
               Height          =   315
               Left            =   75
               TabIndex        =   49
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
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               Caption         =   "Valor Flete"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   90
               TabIndex        =   154
               Top             =   0
               Width           =   750
            End
         End
         Begin VB.Frame Frame24 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   3195
            TabIndex        =   151
            Top             =   210
            Width           =   2835
            Begin MSDataListLib.DataCombo Cmb_moneda 
               Height          =   315
               Left            =   105
               TabIndex        =   48
               Top             =   240
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin MSAdodcLib.Adodc data_moneda 
               Height          =   330
               Left            =   1605
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
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   152
               Top             =   0
               Width           =   585
            End
         End
         Begin VB.Frame Frame22 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   150
            TabIndex        =   149
            Top             =   195
            Width           =   3075
            Begin MSDataListLib.DataCombo Cmb_metodo_pago 
               Height          =   315
               Left            =   105
               TabIndex        =   47
               Top             =   240
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin MSAdodcLib.Adodc data_metodo_pago 
               Height          =   330
               Left            =   1860
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
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               Caption         =   "Método Pago"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   150
               Top             =   0
               Width           =   960
            End
         End
      End
      Begin VB.Frame Frame11 
         Enabled         =   0   'False
         Height          =   1455
         Left            =   7005
         TabIndex        =   124
         Top             =   855
         Visible         =   0   'False
         Width           =   3105
         Begin VB.Frame Fra_bl_master 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   30
            TabIndex        =   127
            Top             =   765
            Width           =   2970
            Begin VB.TextBox Txt_bl_master 
               Height          =   285
               Left            =   120
               MaxLength       =   25
               TabIndex        =   69
               Top             =   270
               Width           =   2520
            End
            Begin VB.CommandButton Cmd_hlp_bl_master 
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
               Left            =   2685
               Picture         =   "Doc_bl.frx":97D0
               Style           =   1  'Graphical
               TabIndex        =   70
               ToolTipText     =   "Buscar"
               Top             =   255
               Width           =   300
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Bl Master"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   135
               TabIndex        =   128
               Top             =   30
               Width           =   660
            End
         End
         Begin VB.Frame Fra_consolidaora 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   30
            TabIndex        =   125
            Top             =   165
            Width           =   1410
            Begin VB.CommandButton Cmd_hlp_consolidadora 
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
               Left            =   1020
               Picture         =   "Doc_bl.frx":9D02
               Style           =   1  'Graphical
               TabIndex        =   68
               ToolTipText     =   "Buscar"
               Top             =   255
               Width           =   300
            End
            Begin VB.TextBox Txt_consolidadora 
               Height          =   285
               Left            =   120
               MaxLength       =   5
               TabIndex        =   67
               Top             =   270
               Width           =   795
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               Caption         =   "Consolidadora"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   126
               Top             =   30
               Width           =   1005
            End
         End
      End
      Begin VB.Frame Frame10 
         Height          =   4590
         Left            =   -74775
         TabIndex        =   112
         Top             =   450
         Width           =   9540
         Begin VB.TextBox txt_docum_auto 
            Height          =   285
            Left            =   3840
            MaxLength       =   80
            TabIndex        =   43
            Top             =   3540
            Width           =   1740
         End
         Begin VB.Frame Fra_Date 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   585
            Left            =   150
            TabIndex        =   113
            Top             =   2595
            Visible         =   0   'False
            Width           =   2025
            Begin GTMaskDate.GTMaskDate GTD_date 
               Height          =   285
               Left            =   30
               TabIndex        =   40
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
            Begin GTMaskDate.GTMaskDate GTD_hora 
               Height          =   285
               Left            =   1380
               TabIndex        =   41
               Top             =   255
               Width           =   570
               _Version        =   65537
               _ExtentX        =   1005
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
               Alignment       =   1
               MaskType        =   1
               MaskTimeSeparator=   ":"
               Mask24HourTime  =   2
               MaskShowTimeMarker=   0
               CalDropDown     =   0   'False
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
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               Caption         =   "Fecha/Hora Descarga"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   30
               TabIndex        =   114
               Top             =   30
               Width           =   1605
            End
         End
         Begin GTMaskNum.GTMaskNum GTM_peso_mal 
            Height          =   315
            Left            =   5700
            TabIndex        =   44
            Top             =   3540
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
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
         Begin GTMaskNum.GTMaskNum GTM_bultos_mal 
            Height          =   315
            Left            =   7185
            TabIndex        =   45
            Top             =   3540
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
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
         Begin MSDataListLib.DataCombo Cmb_tipo_docum 
            Height          =   315
            Left            =   165
            TabIndex        =   35
            Top             =   600
            Width           =   4080
            _ExtentX        =   7197
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_tipo_docum 
            Height          =   330
            Left            =   1665
            Top             =   600
            Visible         =   0   'False
            Width           =   2775
            _ExtentX        =   4895
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
         Begin MSDataListLib.DataCombo Cmb_almacen_temp 
            Height          =   315
            Left            =   4500
            TabIndex        =   36
            Top             =   600
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_almacen_temp 
            Height          =   330
            Left            =   5265
            Top             =   735
            Visible         =   0   'False
            Width           =   1635
            _ExtentX        =   2884
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
         Begin MSDataListLib.DataCombo Cmb_tipo_deposito 
            Height          =   315
            Left            =   180
            TabIndex        =   37
            Top             =   1350
            Width           =   4080
            _ExtentX        =   7197
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_tipo_deposito 
            Height          =   330
            Left            =   1725
            Top             =   1350
            Visible         =   0   'False
            Width           =   2775
            _ExtentX        =   4895
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
         Begin MSDataListLib.DataCombo Cmb_deposito 
            Height          =   315
            Left            =   4515
            TabIndex        =   38
            Top             =   1350
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_deposito 
            Height          =   330
            Left            =   6075
            Top             =   1440
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
         Begin MSDataListLib.DataCombo Cmb_regimen_adua 
            Height          =   315
            Left            =   180
            TabIndex        =   39
            Top             =   2160
            Width           =   4080
            _ExtentX        =   7197
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_regimen_adua 
            Height          =   330
            Left            =   1800
            Top             =   2175
            Visible         =   0   'False
            Width           =   2775
            _ExtentX        =   4895
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
         Begin MSDataListLib.DataCombo Cmb_tipo_salida 
            Height          =   315
            Left            =   210
            TabIndex        =   42
            Top             =   3540
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_tipo_salida 
            Height          =   330
            Left            =   1815
            Top             =   3555
            Visible         =   0   'False
            Width           =   1965
            _ExtentX        =   3466
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
         Begin MSDataListLib.DataCombo Cmb_nav_carga 
            Height          =   315
            Left            =   240
            TabIndex        =   46
            Top             =   4155
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_nav_carga 
            Height          =   330
            Left            =   1710
            Top             =   4155
            Visible         =   0   'False
            Width           =   1950
            _ExtentX        =   3440
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
            Caption         =   "Ag. Naviera de Carga(Linea)"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   210
            TabIndex        =   179
            Top             =   3915
            Width           =   2010
         End
         Begin VB.Label Label35 
            Caption         =   "Docum. Autorización"
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   3840
            TabIndex        =   123
            Top             =   3330
            Width           =   1530
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Salida"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   180
            TabIndex        =   122
            Top             =   3300
            Width           =   1020
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Régimen aduanero"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   180
            TabIndex        =   121
            Top             =   1890
            Width           =   1350
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Depósito aduanero / comerc"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   4515
            TabIndex        =   120
            Top             =   1110
            Width           =   2040
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Depósito"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   165
            TabIndex        =   119
            Top             =   1110
            Width           =   990
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Almacen Temp."
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   4470
            TabIndex        =   118
            Top             =   360
            Width           =   1545
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   135
            TabIndex        =   117
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "No. bultos Mal estado"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   7170
            TabIndex        =   116
            Top             =   3330
            Width           =   1545
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Peso Mal estado"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   5670
            TabIndex        =   115
            Top             =   3330
            Width           =   1185
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1965
         Left            =   150
         TabIndex        =   83
         Top             =   345
         Width           =   6765
         Begin VB.Frame Frame7 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   75
            TabIndex        =   193
            Top             =   765
            Width           =   1500
            Begin MSDataListLib.DataCombo Cmb_pto_emb 
               Height          =   315
               Left            =   135
               TabIndex        =   3
               Top             =   240
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               ListField       =   ""
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
               TabIndex        =   194
               Top             =   0
               Width           =   1230
            End
         End
         Begin VB.Frame Fra_fecha_emb 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   585
            Left            =   4875
            TabIndex        =   133
            Top             =   1320
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
               Caption         =   "Fecha Emb"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   30
               TabIndex        =   134
               Top             =   30
               Width           =   810
            End
         End
         Begin VB.Frame Fra_pto_origen 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   105
            TabIndex        =   131
            Top             =   1320
            Width           =   1590
            Begin VB.TextBox Txt_pto_origen 
               Height          =   285
               Left            =   105
               MaxLength       =   5
               TabIndex        =   6
               Top             =   270
               Width           =   975
            End
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
               Picture         =   "Doc_bl.frx":A234
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "Buscar"
               Top             =   255
               Width           =   300
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "Puerto Origen"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   132
               Top             =   15
               Width           =   975
            End
         End
         Begin VB.Frame Fra_pto_destino 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   3165
            TabIndex        =   129
            Top             =   1320
            Width           =   1590
            Begin VB.TextBox Txt_pto_destino 
               Height          =   285
               Left            =   120
               MaxLength       =   5
               TabIndex        =   10
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
               Picture         =   "Doc_bl.frx":A766
               Style           =   1  'Graphical
               TabIndex        =   11
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
               TabIndex        =   130
               Top             =   15
               Width           =   1380
            End
         End
         Begin VB.Frame Fra_pto_descarga 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   1635
            TabIndex        =   92
            Top             =   1320
            Width           =   1590
            Begin VB.TextBox Txt_pto_descarga 
               Height          =   285
               Left            =   120
               MaxLength       =   5
               TabIndex        =   8
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
               Picture         =   "Doc_bl.frx":AC98
               Style           =   1  'Graphical
               TabIndex        =   9
               ToolTipText     =   "Buscar"
               Top             =   255
               Width           =   300
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Puerto Descarga"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   93
               Top             =   15
               Width           =   1200
            End
         End
         Begin VB.Frame Fra_linea 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   75
            TabIndex        =   90
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
               TabIndex        =   91
               Top             =   0
               Width           =   825
            End
         End
         Begin VB.Frame Fra_bl 
            BorderStyle     =   0  'None
            Height          =   510
            Left            =   1650
            TabIndex        =   88
            Top             =   825
            Width           =   2535
            Begin VB.TextBox Txt_bl 
               Height          =   285
               Left            =   120
               MaxLength       =   25
               TabIndex        =   4
               Top             =   180
               Width           =   2385
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Bill of lading"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   89
               Top             =   -30
               Width           =   840
            End
         End
         Begin VB.Frame Fra_buque 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   1605
            TabIndex        =   86
            Top             =   210
            Width           =   2640
            Begin MSDataListLib.DataCombo Cmb_buque 
               Height          =   315
               Left            =   120
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
               TabIndex        =   87
               Top             =   0
               Width           =   465
            End
         End
         Begin VB.Frame Fra_viaje 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   4170
            TabIndex        =   84
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
               TabIndex        =   85
               Top             =   0
               Width           =   345
            End
         End
         Begin MSDataListLib.DataCombo Cmb_oper_carga 
            Height          =   315
            Left            =   4305
            TabIndex        =   5
            Top             =   990
            Width           =   2295
            _ExtentX        =   4048
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
            TabIndex        =   182
            Top             =   780
            Width           =   2085
         End
      End
      Begin VB.Frame Frame9 
         Height          =   4620
         Left            =   135
         TabIndex        =   94
         Top             =   2265
         Width           =   10005
         Begin VB.Frame Fra_consignatario 
            Caption         =   "  Consignatario "
            Height          =   1965
            Left            =   120
            TabIndex        =   106
            Top             =   135
            Width           =   4845
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
               Picture         =   "Doc_bl.frx":B1CA
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Buscar"
               Top             =   450
               Width           =   300
            End
            Begin VB.TextBox Txt_id_consignatario 
               Height          =   285
               Left            =   2025
               MaxLength       =   15
               TabIndex        =   17
               Top             =   465
               Width           =   1500
            End
            Begin VB.TextBox Txt_nombre_consignatario 
               Height          =   285
               Left            =   135
               MaxLength       =   80
               TabIndex        =   19
               Top             =   1035
               Width           =   3360
            End
            Begin VB.TextBox Txt_direc_consignatario 
               Height          =   285
               Left            =   135
               MaxLength       =   80
               TabIndex        =   22
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
               Picture         =   "Doc_bl.frx":B6FC
               Style           =   1  'Graphical
               TabIndex        =   21
               ToolTipText     =   "Buscar"
               Top             =   1020
               Width           =   300
            End
            Begin VB.TextBox Txt_nac_consignatario 
               Height          =   285
               Left            =   3570
               MaxLength       =   5
               TabIndex        =   20
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
               TabIndex        =   16
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
               TabIndex        =   111
               Top             =   210
               Width           =   945
            End
            Begin VB.Label Label4 
               Caption         =   "Nombre"
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   135
               TabIndex        =   110
               Top             =   780
               Width           =   1260
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Dirección"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   150
               TabIndex        =   109
               Top             =   1350
               Width           =   675
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Nacionalidad"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   3585
               TabIndex        =   108
               Top             =   780
               Width           =   930
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Identif."
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   107
               Top             =   195
               Width           =   1065
            End
         End
         Begin VB.Frame Fra_notificador 
            Caption         =   " Notificador "
            Height          =   1965
            Left            =   5055
            TabIndex        =   101
            Top             =   120
            Width           =   4845
            Begin VB.TextBox Txt_direc_notificador 
               Height          =   285
               Left            =   150
               MaxLength       =   80
               TabIndex        =   27
               Top             =   1575
               Width           =   4500
            End
            Begin VB.TextBox Txt_nombre_notificador 
               Height          =   285
               Left            =   135
               MaxLength       =   80
               TabIndex        =   26
               Top             =   1035
               Width           =   3360
            End
            Begin VB.TextBox Txt_id_notificador 
               Height          =   285
               Left            =   2025
               MaxLength       =   15
               TabIndex        =   24
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
               Picture         =   "Doc_bl.frx":BC2E
               Style           =   1  'Graphical
               TabIndex        =   25
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
               TabIndex        =   23
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
               TabIndex        =   105
               Top             =   195
               Width           =   1065
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Dirección"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   150
               TabIndex        =   104
               Top             =   1350
               Width           =   675
            End
            Begin VB.Label Label21 
               Caption         =   "Nombre"
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   135
               TabIndex        =   103
               Top             =   780
               Width           =   1260
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Identificación"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   2025
               TabIndex        =   102
               Top             =   210
               Width           =   945
            End
         End
         Begin VB.Frame Fra_embarcador 
            Caption         =   " Embarcador "
            Height          =   1905
            Left            =   135
            TabIndex        =   98
            Top             =   2085
            Width           =   4845
            Begin VB.TextBox Txt_id_emb 
               Height          =   285
               Left            =   2025
               MaxLength       =   15
               TabIndex        =   29
               Top             =   510
               Width           =   1500
            End
            Begin VB.CommandButton Cmd_hlp_emb 
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
               Picture         =   "Doc_bl.frx":C160
               Style           =   1  'Graphical
               TabIndex        =   30
               ToolTipText     =   "Buscar"
               Top             =   495
               Width           =   300
            End
            Begin VB.TextBox Txt_nombre_embarcador 
               Height          =   285
               Left            =   135
               MaxLength       =   80
               TabIndex        =   31
               Top             =   1005
               Width           =   3360
            End
            Begin VB.TextBox Txt_direc_embarcador 
               Height          =   285
               Left            =   135
               MaxLength       =   80
               TabIndex        =   32
               Top             =   1530
               Width           =   4470
            End
            Begin MSAdodcLib.Adodc data_tipo_id_emb 
               Height          =   330
               Left            =   3870
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
            Begin MSDataListLib.DataCombo Cmb_tipo_id_emb 
               Height          =   315
               Left            =   135
               TabIndex        =   28
               Top             =   480
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
            End
            Begin VB.Label Label64 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Identif."
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   206
               Top             =   240
               Width           =   1065
            End
            Begin VB.Label Label63 
               AutoSize        =   -1  'True
               Caption         =   "Identificación"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   2025
               TabIndex        =   205
               Top             =   255
               Width           =   945
            End
            Begin VB.Label Label23 
               Caption         =   "Nombre"
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   135
               TabIndex        =   100
               Top             =   810
               Width           =   1260
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Dirección"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   150
               TabIndex        =   99
               Top             =   1320
               Width           =   675
            End
         End
         Begin VB.Frame Fra_notificador2 
            Caption         =   " Notificador 2 "
            Height          =   1305
            Left            =   5055
            TabIndex        =   95
            Top             =   2085
            Width           =   4845
            Begin VB.TextBox Txt_direc_notificador2 
               Height          =   285
               Left            =   135
               MaxLength       =   80
               TabIndex        =   34
               Top             =   945
               Width           =   4470
            End
            Begin VB.TextBox Txt_nombre_notificador2 
               Height          =   285
               Left            =   135
               MaxLength       =   80
               TabIndex        =   33
               Top             =   405
               Width           =   3360
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Dirección"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   150
               TabIndex        =   97
               Top             =   720
               Width           =   675
            End
            Begin VB.Label Label25 
               Caption         =   "Nombre"
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   135
               TabIndex        =   96
               Top             =   210
               Width           =   1260
            End
         End
         Begin VB.Frame Frame20 
            Height          =   570
            Left            =   135
            TabIndex        =   135
            Top             =   3945
            Width           =   9765
            Begin VB.Frame Frame13 
               BorderStyle     =   0  'None
               Height          =   360
               Left            =   60
               TabIndex        =   136
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
                  TabIndex        =   137
                  Top             =   15
                  Width           =   1155
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  Caption         =   "Peso Total"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   105
                  TabIndex        =   138
                  Top             =   75
                  Width           =   765
               End
            End
            Begin VB.Frame Frame15 
               BorderStyle     =   0  'None
               Height          =   360
               Left            =   7290
               TabIndex        =   145
               Top             =   165
               Width           =   2370
               Begin VB.TextBox Txt_cont_tot 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000016&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1110
                  Locked          =   -1  'True
                  MaxLength       =   80
                  TabIndex        =   146
                  Top             =   15
                  Width           =   1155
               End
               Begin VB.Label Label26 
                  AutoSize        =   -1  'True
                  Caption         =   "No.Contens."
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   105
                  TabIndex        =   147
                  Top             =   75
                  Width           =   885
               End
            End
            Begin VB.Frame Frame16 
               BorderStyle     =   0  'None
               Height          =   360
               Left            =   2475
               TabIndex        =   142
               Top             =   165
               Width           =   2370
               Begin VB.TextBox Txt_volumen_tot 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000016&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1110
                  Locked          =   -1  'True
                  MaxLength       =   80
                  TabIndex        =   143
                  Top             =   15
                  Width           =   1155
               End
               Begin VB.Label Label17 
                  Caption         =   "Volum Total"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   135
                  TabIndex        =   144
                  Top             =   75
                  Width           =   1020
               End
            End
            Begin VB.Frame Frame14 
               BorderStyle     =   0  'None
               Height          =   360
               Left            =   4950
               TabIndex        =   139
               Top             =   165
               Width           =   2370
               Begin VB.TextBox Txt_bultos_tot 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000016&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1110
                  Locked          =   -1  'True
                  MaxLength       =   80
                  TabIndex        =   140
                  Top             =   15
                  Width           =   1155
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  Caption         =   "BultosTotal"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   105
                  TabIndex        =   141
                  Top             =   75
                  Width           =   795
               End
            End
         End
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "de"
         Height          =   195
         Left            =   -71115
         TabIndex        =   181
         Top             =   750
         Width           =   180
      End
      Begin VB.Label Lbl_no_det_tot 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   -70860
         TabIndex        =   180
         Top             =   750
         Width           =   90
      End
      Begin VB.Label Lbl_no_det_act 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   -71265
         TabIndex        =   178
         Top             =   750
         Width           =   90
      End
   End
End
Attribute VB_Name = "doc_bl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Gstr_opc_bl As String
Private Sub Cmb_almacen_temp_KeyPress(KeyAscii As Integer)
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

Private Sub Cmb_deposito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_embalaje_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_embalaje_LostFocus()
If CDbl(Lbl_no_det_act.Caption) <> 0 Then
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).embalaje = Cmb_embalaje.BoundText
End If
End Sub

Private Sub Cmb_IMO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_IMO_LostFocus()
If CDbl(Lbl_no_det_act.Caption) <> 0 Then
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).Imo = Cmb_IMO.BoundText
End If
End Sub

Private Sub cmb_linea_Change()
If Not Cmb_linea.MatchedWithList Then Exit Sub
carga_datacombo data_buque, Cmb_buque, "Select * from docu_buque                                Where buqu_linea        = '" & Cmb_linea.BoundText & "' and buqu_status      ='A'", "buqu_codigo", "buqu_descripcion"
carga_datacombo data_tipo_docum, Cmb_tipo_docum, "Select * from docu_tipo_documento             Where tipo_doc_linea    = '" & Cmb_linea.BoundText & "' and tipo_doc_status  ='A'", "tipo_doc_codigo", "tipo_doc_descripcion"
carga_datacombo data_almacen_temp, Cmb_almacen_temp, "Select * from docu_almacen                Where almac_linea       = '" & Cmb_linea.BoundText & "' and almac_status     ='A'", "almac_codigo", "almac_descripcion"
'carga_datacombo data_deposito, Cmb_deposito, "Select * from docu_deposito                       Where depo_linea        = '" & Cmb_linea.BoundText & "' and depo_status      ='A'", "depo_codigo", "depo_descripcion"

carga_datacombo data_tipo_deposito, Cmb_tipo_deposito, "Select * from docu_tipo_deposito        Where tipo_dep_linea    = '" & Cmb_linea.BoundText & "' and tipo_dep_status  ='A'", "tipo_dep_codigo", "tipo_dep_descripcion"
carga_datacombo data_regimen_adua, Cmb_regimen_adua, "Select * from docu_regimen                Where regi_linea        = '" & Cmb_linea.BoundText & "' and regi_status      ='A'", "regi_codigo", "regi_descripcion"
carga_datacombo data_tipo_salida, Cmb_tipo_salida, "Select * from docu_tipo_salida              Where tipo_sal_linea    = '" & Cmb_linea.BoundText & "' and tipo_sal_status  ='A'", "tipo_sal_codigo", "tipo_sal_descripcion"
carga_datacombo data_metodo_pago, Cmb_metodo_pago, "Select * from docu_metodo_pago              Where meto_pago_linea   = '" & Cmb_linea.BoundText & "' and meto_pago_status ='A'", "meto_pago_codigo", "meto_pago_descripcion"
carga_datacombo data_moneda, Cmb_moneda, "Select * from docu_moneda                             Where mone_linea        = '" & Cmb_linea.BoundText & "' and mone_status      ='A'", "mone_codigo", "mone_descripcion"
carga_datacombo data_metodo_pago_otros, Cmb_metodo_pago_otros, "Select * from docu_metodo_pago  Where meto_pago_linea   = '" & Cmb_linea.BoundText & "' and meto_pago_status ='A'", "meto_pago_codigo", "meto_pago_descripcion"
carga_datacombo data_moneda_otros, Cmb_moneda_otros, "Select * from docu_moneda                 Where mone_linea        = '" & Cmb_linea.BoundText & "' and mone_status      ='A'", "mone_codigo", "mone_descripcion"
carga_datacombo data_moneda_seguro, Cmb_moneda_seguro, "Select * from docu_moneda               Where mone_linea        = '" & Cmb_linea.BoundText & "' and mone_status      ='A'", "mone_codigo", "mone_descripcion"
carga_datacombo data_tipo_carga, Cmb_tipo_carga, "Select * from docu_tipo_carga                 Where tipo_carga_linea  = '" & Cmb_linea.BoundText & "' and tipo_carga_status='A'", "tipo_carga_codigo", "tipo_carga_descripcion"
carga_datacombo data_tipo_prod, Cmb_tipo_prod, "Select * from docu_tipo_producto                Where tipo_prod_linea   = '" & Cmb_linea.BoundText & "' and tipo_prod_status ='A'", "tipo_prod_codigo", "tipo_prod_descripcion"
carga_datacombo data_embalaje, Cmb_embalaje, "Select * from docu_embalaje                       Where emba_linea        = '" & Cmb_linea.BoundText & "' and emba_status      ='A'", "emba_codigo", "emba_descripcion"
carga_datacombo data_IMO, Cmb_IMO, "Select * from docu_imo                                      Where imo_linea         = '" & Cmb_linea.BoundText & "' and imo_status       ='A'", "imo_codigo", "imo_descripcion"
carga_datacombo data_ONU, Cmb_ONU, "Select * from docu_imo_onu                                  Where imo_onu_linea     = '" & Cmb_linea.BoundText & "' and imo_onu_status   ='A'", "imo_onu_codigo", "imo_onu_descripcion"
carga_datacombo data_oper_carga, Cmb_oper_carga, "Select * from docu_emp_resp                   Where emp_resp_linea    = '" & Cmb_linea.BoundText & "' and emp_resp_status  ='A'", "emp_resp_codigo", "emp_resp_descripcion"
carga_datacombo data_nav_carga, Cmb_nav_carga, "Select * from docu_emp_anav                     Where emp_anav_linea    = '" & Cmb_linea.BoundText & "' and emp_anav_status  ='A'", "emp_anav_codigo", "emp_anav_descripcion"
carga_parametros Cmb_linea.BoundText

Txt_pto_descarga.Text = Gstr_param.pto_descarga
Txt_pto_destino.Text = Gstr_param.pto_descarga
Txt_pto_trasb.Text = ""
Cmb_oper_carga.BoundText = Gstr_param.naviera
Cmb_moneda.BoundText = Gstr_param.MONEDA
'Cmb_moneda_otros.BoundText = Gstr_param.MONEDA
Cmb_tipo_docum.BoundText = Gstr_param.tipo_documento
Me.Cmb_tipo_id_consignatario.BoundText = Gstr_param.tipo_id
Me.Cmb_tipo_id_notificador.BoundText = Gstr_param.tipo_id
Me.Cmb_tipo_id_emb.BoundText = Gstr_param.tipo_id
Me.Txt_nac_consignatario.Text = Gstr_param.nacionalidad
Me.Cmb_metodo_pago.BoundText = Gstr_param.metodo_pago
'Me.Cmb_metodo_pago_otros.BoundText = Gstr_param.metodo_pago
Me.Cmb_regimen_adua.BoundText = Gstr_param.regimen_adua
Me.Cmb_nav_carga.BoundText = Gstr_param.Linea_Nav
Me.Cmb_almacen_temp.BoundText = Gstr_param.almacen_temporal
End Sub

Private Sub Cmb_linea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_metodo_pago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_metodo_pago_otros_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_moneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub Cmb_moneda_otros_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_moneda_seguro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_nav_carga_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    STab_bl.Tab = 2
End If
End Sub

Private Sub Cmb_ONU_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_ONU_LostFocus()
If CDbl(Lbl_no_det_act.Caption) <> 0 Then
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).pelig_ONU = Cmb_ONU.BoundText
End If
End Sub

Private Sub Cmb_oper_carga_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_pto_emb_Change()
Dim sql As String
Dim reg As New ADODB.Recordset

On Error GoTo Errores
If Cmb_pto_emb.MatchedWithList Then
   sql = "Select * from docu_manifiesto " & _
              " Where mani_linea = '" & Cmb_linea.BoundText & "'" & _
              "   And mani_buque  = '" & Cmb_buque.BoundText & "'" & _
              "   And mani_viaje  = '" & Cmb_viaje.BoundText & "'" & _
              "   And mani_puer_emb ='" & Cmb_pto_emb.BoundText & "'"
        reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        If Not reg.EOF Then
           Txt_pto_descarga.Text = reg("mani_puer_des")
           Msk_fecha_emb.Text = reg("mani_fecha_zarp_ini")
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'TL quiero repetir lo que escoge en puerto de embarque en origen
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Txt_pto_origen.Text = Cmb_pto_emb.BoundText
End If
Exit Sub
Errores:
   MsgBox "Ocurrio el error " & Err.Description, vbInformation
End Sub

Private Sub Cmb_pto_emb_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_regimen_adua_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_tipo_carga_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_tipo_carga_LostFocus()
If CDbl(Lbl_no_det_act.Caption) <> 0 Then
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).tipo_carga = Cmb_tipo_carga.BoundText
End If
End Sub

Private Sub Cmb_tipo_deposito_Change()
If Me.Cmb_linea.MatchedWithList And Me.Cmb_tipo_deposito.MatchedWithList Then
    carga_datacombo data_deposito, Cmb_deposito, "Select * from docu_deposito Where depo_linea = '" & Cmb_linea.BoundText & "' and depo_tipo_depo ='" & Cmb_tipo_deposito.BoundText & "'", "depo_codigo", "depo_descripcion"
End If
End Sub


Private Sub Cmb_tipo_deposito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_tipo_docum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_tipo_id_consignatario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_tipo_id_notificador_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_tipo_prod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_tipo_prod_LostFocus()
If CDbl(Lbl_no_det_act.Caption) <> 0 Then
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).tipo_prod = Cmb_tipo_prod.BoundText
End If

End Sub

Private Sub Cmb_tipo_salida_KeyPress(KeyAscii As Integer)
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

Private Sub Cmd_anterior_det_Click()
If CDbl(Lbl_no_det_act.Caption) > 1 Then
    Lbl_no_det_act.Caption = Str(CDbl(Lbl_no_det_act.Caption) - 1)
    Call muestra_detalle
End If
End Sub
Public Sub cmd_eliminar_Click()
If MsgBox("Esta seguro que desea eliminar el bl " & txt_bl.Text, vbYesNo + vbQuestion) = vbYes Then
    If Elimina_bl(Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, Cmb_pto_emb.BoundText, Txt_pto_descarga.Text, Cmb_oper_carga.BoundText, txt_bl.Text) Then
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
        If Grid_flete.Cols > 10 Then
            Grid_flete.Col = 10: temp1 = Grid_flete.Text
            Grid_flete.Col = 11: temp2 = Grid_flete.Text
        End If
        
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
        If Grid_flete.Cols > 10 Then
            Grid_flete.Col = 10: Grid_flete.Text = temp1
            Grid_flete.Col = 11: Grid_flete.Text = temp2
        End If
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
        sql = "Insert into docu_blcab (blcab_linea,     blcab_buque,    blcab_viaje, " & _
                                     " blcab_puer_emb,  blcab_puer_des, blcab_emp_resp, " & _
                                     " blcab_cono_emb,  blcab_emp_anav, blcab_tipo_doc, " & _
                                     " blcab_cono_master,   blcab_peso_tot, blcab_bultos_tot, " & _
                                     " blcab_tipo_ide_cons, blcab_ide_cons,     blcab_consig, " & _
                                     " blcab_direc_consig,  blcab_nacio_consig, blcab_tipo_ide_notif," & _
                                     " blcab_ide_notif,     blcab_notif,        blcab_direc_notif, " & _
                                     " blcab_tipo_ide_emb,  blcab_ide_emb,      " & _
                                     " blcab_embarcador,    blcab_direc_embarc, blcab_puer_origen," & _
                                     " blcab_puer_trasb,    blcab_fecha_trasb, " & _
                                     " blcab_almac_temp,    blcab_tipo_depos,   blcab_deposito_adua, " & _
                                     " blcab_regimen_adu,   blcab_consolidad,    " & _
                                     " blcab_tipo_salida,  blcab_docum_aut, " & _
                                     " blcab_pesob_mal_est, blcab_bultos_mal_est, blcab_notif2," & _
                                     " blcab_direc_notif2,  blcab_no_cont,      blcab_volumen_tot," & _
                                     " blcab_pto_dest_final,blcab_metod_pago,   blcab_moneda, " & _
                                     " blcab_flete,         blcab_otros_gastos, blcab_moneda_otros, " & _
                                     " blcab_seguro,        blcab_moneda_seguro,blcab_metod_pago_otros," & _
                                     " blcab_fecha_emb,     blcab_user_ing,     blcab_fecha_ing, " & _
                                     " blcab_hora_ing)"
         sql = sql & "VALUES ('" & Cmb_linea.BoundText & "',    '" & Cmb_buque.BoundText & "',       '" & Cmb_viaje.BoundText & "', " & _
                        "'" & Cmb_pto_emb.BoundText & "',       '" & Txt_pto_descarga.Text & "',     '" & Cmb_oper_carga.BoundText & "', " & _
                        "'" & txt_bl.Text & "',            '" & Cmb_nav_carga.BoundText & "',    '" & Cmb_tipo_docum.BoundText & "', " & _
                        "'" & Txt_bl_master.Text & "',      " & CDbl(Txt_peso_tot.Text) & ",     " & CDbl(Txt_bultos_tot.Text) & ", " & _
                        "'" & Cmb_tipo_id_consignatario.BoundText & "', '" & Txt_id_consignatario.Text & "',  '" & Txt_nombre_consignatario.Text & "', " & _
                        "'" & Txt_direc_consignatario.Text & "',        '" & Txt_nac_consignatario.Text & "', '" & Cmb_tipo_id_notificador.BoundText & "', " & _
                        "'" & Txt_id_notificador.Text & "',             '" & Txt_nombre_notificador.Text & "','" & Txt_direc_notificador.Text & "', " & _
                        "'" & Cmb_tipo_id_emb.BoundText & "',           '" & Txt_id_emb.Text & "'," & _
                        "'" & Txt_nombre_embarcador.Text & "',          '" & Txt_direc_embarcador.Text & "',  '" & Txt_pto_origen.Text & "', " & _
                        "'" & Txt_pto_trasb.Text & "',                  '" & Format(Msk_fecha_trasb.Text, GSTR_FORMATO_FECHA) & "', "
                        
         sql = sql & "   '" & Cmb_almacen_temp.BoundText & "',          '" & Cmb_tipo_deposito.BoundText & "','" & Cmb_deposito.BoundText & "', " & _
                        "'" & Cmb_regimen_adua.BoundText & "',          '" & Txt_consolidadora.Text & "',     " & _
                        "'" & Cmb_tipo_salida.BoundText & "','" & txt_docum_auto.Text & "', " & _
                        " " & GTM_peso_mal.Value & ",                    " & GTM_bultos_mal.Value & ",        '" & Txt_nombre_notificador2.Text & "', " & _
                        "'" & Txt_direc_notificador2.Text & "',          " & CDbl(Txt_cont_tot.Text) & ",      " & CDbl(Txt_volumen_tot.Text) & ", " & _
                        "'" & Txt_pto_destino.Text & "',                '" & Cmb_metodo_pago.BoundText & "',  '" & Cmb_moneda.BoundText & "', " & _
                        " " & CDbl(GTM_flete.Value) & ",                 " & CDbl(GTM_flete_otros.Value) & ", '" & Cmb_moneda_otros.BoundText & "', " & _
                        " " & CDbl(GTM_seguro.Value) & ",               '" & Cmb_moneda_seguro.BoundText & "','" & Cmb_metodo_pago_otros.BoundText & "', " & _
                        "'" & Format(Msk_fecha_emb.Text, GSTR_FORMATO_FECHA) & "',                  '" & Gstr_user & "',                  '" & Format(Now, GSTR_FORMATO_FECHA) & "', " & _
                        "'" & Format(Time, "hh:mm:ss") & "')"

Else
        sql = "UPDATE docu_blcab SET blcab_emp_anav ='" & Cmb_nav_carga.BoundText & "', " & _
                                    "blcab_tipo_doc ='" & Cmb_tipo_docum.BoundText & "', " & _
                                    "blcab_cono_master='" & Txt_bl_master.Text & "', " & _
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
                                    "blcab_tipo_ide_emb  ='" & Cmb_tipo_id_emb.BoundText & "', " & _
                                    "blcab_ide_emb       ='" & Trim$(Txt_id_emb.Text) & "', " & _
                                    "blcab_embarcador    ='" & Trim$(Txt_nombre_embarcador.Text) & "', " & _
                                    "blcab_direc_embarc  ='" & Trim$(Txt_direc_embarcador.Text) & "', " & _
                                    "blcab_puer_origen   ='" & Txt_pto_origen.Text & "', " & _
                                    "blcab_puer_trasb    ='" & Txt_pto_trasb.Text & "', " & _
                                    "blcab_fecha_trasb   = " & IIf(IsNull(Msk_fecha_trasb.DateValue), "null", "'" & Format(Msk_fecha_trasb.Text, GSTR_FORMATO_FECHA) & "'") & ", " & _
                                    "blcab_almac_temp    ='" & Cmb_almacen_temp.BoundText & "', " & _
                                    "blcab_tipo_depos    ='" & Cmb_tipo_deposito.BoundText & "', "
                        sql = sql & "blcab_deposito_adua ='" & Cmb_deposito.BoundText & "', " & _
                                    "blcab_regimen_adu   ='" & Cmb_regimen_adua.BoundText & "', " & _
                                    "blcab_consolidad    ='" & Txt_consolidadora.Text & "'," & _
                                    "" & _
                                    "" & _
                                    "blcab_tipo_salida   ='" & Cmb_tipo_salida.BoundText & "', " & _
                                    "blcab_docum_aut     ='" & txt_docum_auto.Text & "', " & _
                                    "blcab_pesob_mal_est = " & GTM_peso_mal.Value & ", " & _
                                    "blcab_bultos_mal_est= " & GTM_bultos_mal.Value & ", " & _
                                    "blcab_notif2        ='" & Trim$(Txt_nombre_notificador2.Text) & "'," & _
                                    "blcab_direc_notif2  ='" & Trim$(Txt_direc_notificador2.Text) & "', " & _
                                    "blcab_no_cont       = " & CDbl(Txt_cont_tot.Text) & ", " & _
                                    "blcab_volumen_tot   = " & CDbl(Txt_volumen_tot.Text) & "," & _
                                    "blcab_pto_dest_final='" & Txt_pto_destino.Text & "', " & _
                                    "blcab_metod_pago    ='" & Cmb_metodo_pago.BoundText & "', "
                        sql = sql & "blcab_moneda        ='" & Cmb_moneda.BoundText & "', " & _
                                    "blcab_flete         = " & CDbl(GTM_flete.Value) & ", " & _
                                    "blcab_otros_gastos  = " & CDbl(GTM_flete_otros.Value) & ", " & _
                                    "blcab_moneda_otros  ='" & Cmb_moneda_otros.BoundText & "', " & _
                                    "blcab_seguro        = " & CDbl(GTM_seguro.Value) & ", " & _
                                    "blcab_moneda_seguro ='" & Cmb_moneda_seguro.BoundText & "', " & _
                                    "blcab_metod_pago_otros='" & Cmb_metodo_pago_otros.BoundText & " ', " & _
                                    "blcab_fecha_emb     ='" & Format(Msk_fecha_emb.Text, GSTR_FORMATO_FECHA) & "', " & _
                                    "blcab_user_mod      ='" & Gstr_user & "', " & _
                                    "blcab_fecha_mod     ='" & Format(Now, GSTR_FORMATO_FECHA) & "', " & _
                                    "blcab_hora_mod      ='" & Format(Time, "hh:mm:ss") & "'"
                                    
        sql = sql & "  Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
                        "   and blcab_buque    ='" & Cmb_buque.BoundText & "'" & _
                        "   and blcab_viaje    ='" & Cmb_viaje.BoundText & "'" & _
                        "   and blcab_puer_emb ='" & Cmb_pto_emb.BoundText & "'" & _
                        "   and blcab_puer_des ='" & Txt_pto_descarga.Text & "'" & _
                        "   and blcab_emp_resp ='" & Cmb_oper_carga.BoundText & "'" & _
                        "   and blcab_cono_emb ='" & txt_bl.Text & "'"
      Gcon_docu.Execute sql
      
      sql = "Delete from docu_bldet " & _
              "Where bldet_linea      ='" & Cmb_linea.BoundText & "'" & _
               "   and bldet_buque    ='" & Cmb_buque.BoundText & "'" & _
               "   and bldet_viaje    ='" & Cmb_viaje.BoundText & "'" & _
               "   and bldet_puer_emb ='" & Cmb_pto_emb.BoundText & "'" & _
               "   and bldet_puer_des ='" & Txt_pto_descarga.Text & "'" & _
               "   and bldet_emp_resp ='" & Cmb_oper_carga.BoundText & "'" & _
               "   and bldet_cono_emb ='" & txt_bl.Text & "'"
      
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
            'sql = "insert into docu_bldet (bldet_linea, bldet_buque, bldet_viaje," & _
                  " bldet_puer_emb,     bldet_puer_des,     bldet_emp_resp, " & _
                  " bldet_cono_emb,     bldet_secuencia,    bldet_tipo_carga, " & _
                  " bldet_peso,         bldet_bultos,       bldet_embalaje, " & _
                  " bldet_marcas_num,   bldet_desc_carg,    bldet_volumen," & _
                  " bldet_contenedor,   bldet_just_cambio,  bldet_imo, " & _
                  " bldet_pelig_onu,    bldet_sello,        bldet_tam_equi, " & _
                  " bldet_tip_equi,     bldet_cond_cont)"
                  
          sql = "insert into docu_bldet " & _
            " VALUES ('" & Cmb_linea.BoundText & "',    '" & Cmb_buque.BoundText & "',      '" & Cmb_viaje.BoundText & "', " & _
                  "'" & Cmb_pto_emb.BoundText & "',       '" & Txt_pto_descarga.Text & "',     '" & Cmb_oper_carga.BoundText & "', " & _
                  "'" & txt_bl.Text & "',             " & i & ",                          '" & mani_bl_det(i).tipo_carga & "', " & _
                  " " & mani_bl_det(i).peso & ",              " & mani_bl_det(i).no_bultos & ",          '" & Trim(mani_bl_det(i).embalaje) & "', " & _
                  "'" & Trim(mani_bl_det(i).marca_num) & "', '" & Trim(mani_bl_det(i).desc_carga) & "',   '" & mani_bl_det(i).tipo_prod & "', " & mani_bl_det(i).volumen & ", " & _
                  "'" & Trim$(mani_bl_det(i).CONTEINER) & "','" & Trim(mani_bl_det(i).just_cambio) & "', '" & Trim(mani_bl_det(i).Imo) & "'," & _
                  "'" & Trim(mani_bl_det(i).pelig_ONU) & "', '" & Trim(mani_bl_det(i).Sello) & "',       '" & Trim(mani_bl_det(i).tamano_equip) & "'," & _
                  "'" & Trim(mani_bl_det(i).tipo_equip) & "', '" & Trim(mani_bl_det(i).cond_cont) & "')"


                  
            Gcon_docu.Execute sql
        End If
    Next i
End If
Call actualiza_totales_manifiesto(Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, Cmb_pto_emb.BoundText, Txt_pto_descarga.Text, Cmb_oper_carga.BoundText)
' ------ seccion de fletes ------
 sql = "Delete from docu_blflete " & _
              "Where flete_linea      ='" & Cmb_linea.BoundText & "'" & _
               "   and flete_buque    ='" & Cmb_buque.BoundText & "'" & _
               "   and flete_viaje    ='" & Cmb_viaje.BoundText & "'" & _
               "   and flete_puer_emb ='" & Cmb_pto_emb.BoundText & "'" & _
               "   and flete_puer_des ='" & Txt_pto_descarga.Text & "'" & _
               "   and flete_emp_resp ='" & Cmb_oper_carga.BoundText & "'" & _
               "   and flete_cono_emb ='" & txt_bl.Text & "'"
Gcon_docu.Execute sql, i

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

Private Sub Cmd_hlp_consolidadora_Click()
   Txt_consolidadora.Text = Ayuda("docu_consolidadora", "cons_codigo", "cons_descripcion", " cons_linea ='" & Cmb_linea.BoundText & "' and cons_status ='A' ", "Consolidadora")
End Sub

Private Sub Cmd_hlp_cont_Click()
     Txt_contenedor.Text = Ayuda("docu_cont", "cont_contenedor", "cont_tam_equip", " cont_linea ='" & Cmb_linea.BoundText & "' and cont_buque ='" & Cmb_buque.BoundText & "' and cont_viaje ='" & Cmb_viaje.BoundText & "' and cont_puer_emb ='" & Cmb_pto_emb.BoundText & "' and cont_puer_des ='" & Txt_pto_descarga.Text & "'", "Contenedor")
End Sub

Private Sub Cmd_hlp_emb_Click()
   Txt_id_emb.Text = Ayuda("docu_consignatario", "consig_codigo", "consig_nombre", " consig_tipo_id = '" & Cmb_tipo_id_emb.BoundText & "' and consig_status ='A' ", "Embarcador")
   Txt_id_emb.SetFocus
End Sub

Private Sub Cmd_hlp_nac_consignatario_Click()
    Txt_nac_consignatario.Text = Ayuda("docu_nacionalidad", "naci_codigo", "naci_descripcion", " naci_status ='A' ", "Nacionalidad")
End Sub

Private Sub Cmd_hlp_notificador_Click()
    Txt_id_notificador.Text = Ayuda("docu_consignatario", "consig_codigo", "consig_nombre", " consig_tipo_id = '" & Cmb_tipo_id_notificador.BoundText & "' and consig_status ='A' ", "Notificador")
    Txt_id_notificador.SetFocus
End Sub

Private Sub Cmd_hlp_pto_descarga_Click()
   Txt_pto_descarga.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Descarga")
End Sub

Private Sub Cmd_hlp_pto_destino_Click()
    Txt_pto_destino.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Destino")
End Sub

Private Sub Cmd_hlp_pto_origen_Click()
   Txt_pto_origen.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Origen")
End Sub

Private Sub Cmd_hlp_pto_trasb_Click()
    Txt_pto_trasb.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Trasbordo")
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
    If Not Cmb_linea.MatchedWithList Then Exit Sub
    Load doc_aux_flete_bl
    'doc_aux_flete_bl.Form_Load
    doc_aux_flete_bl.Cmb_linea.BoundText = Cmb_linea.BoundText
    doc_aux_flete_bl.Show 1
    If Gtyp_flete.reg_cancel = False Then
        Grid_flete.Col = 0:  If Grid_flete.Text = "" Then Exit Sub
        Grid_flete.Row = Grid_flete.Rows - 1
        If Cmb_linea.BoundText = "KHS" Then
            Grid_flete.Col = 0:  Grid_flete.Text = Gtyp_flete.cargo_cod
            Grid_flete.Col = 1:  Grid_flete.Text = Gtyp_flete.cargo
            Grid_flete.Col = 2:  Grid_flete.Text = Gtyp_flete.deptor_cod
            Grid_flete.Col = 3:  Grid_flete.Text = Gtyp_flete.debtor
            Grid_flete.Col = 4:  Grid_flete.Text = Format(Gtyp_flete.valor_unit, "###,###,##0.00")
            Grid_flete.Col = 5:  Grid_flete.Text = Gtyp_flete.nivel_cargo
            Grid_flete.Col = 6:  Grid_flete.Text = Format(Gtyp_flete.cantidad, "###,###,##0.00")
            Grid_flete.Col = 7:  Grid_flete.Text = Format(Gtyp_flete.valor, "###,###,##0.00")
            Grid_flete.Col = 8:  Grid_flete.Text = Gtyp_flete.MONEDA
            Grid_flete.Col = 9:  Grid_flete.Text = Gtyp_flete.marcas
            Grid_flete.Col = 10: Grid_flete.Text = Gtyp_flete.lugar
            Grid_flete.Col = 11: Grid_flete.Text = IIf(Gtyp_flete.imp_bl, "S", "N")
        Else
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
        End If
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
   If Trim$(Cmb_linea.BoundText) = "" Then
      MsgBox "Ingrese el Valor de Linea", vbInformation, App.Title
      valida_Datos = False
      If Cmb_linea.Enabled Then Cmb_linea.SetFocus
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
   If Trim$(Cmb_pto_emb.BoundText) = "" Then
      MsgBox "Ingrese el Código del puerto embarque", vbInformation, App.Title
      valida_Datos = False
      If Cmb_pto_emb.Enabled Then Cmb_pto_emb.SetFocus
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
   If Trim$(Cmb_metodo_pago.BoundText) = "" Then
      MsgBox "Ingrese el Código del metodo de pago", vbInformation, App.Title
      valida_Datos = False
      If Cmb_metodo_pago.Enabled Then Cmb_metodo_pago.SetFocus
      Exit Function
   End If
   If Cmb_metodo_pago.BoundText <> "NC" Then
      If Trim$(Cmb_moneda.BoundText) = "" Then
         MsgBox "Ingrese el Código de la moneda", vbInformation, App.Title
         valida_Datos = False
         If Cmb_moneda.Enabled Then Cmb_moneda.SetFocus
         Exit Function
      End If
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
   If Trim$(txt_bl.Text) = "" Then
      MsgBox "Ingrese el número del BL", vbInformation, App.Title
      valida_Datos = False
      If txt_bl.Enabled Then txt_bl.SetFocus
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
'   If Trim$(Txt_pto_destino.Text) = "" Then
'      MsgBox "Ingrese el Código del puerto de destino", vbInformation, App.Title
'      valida_Datos = False
'      If Txt_pto_destino.Enabled Then Txt_pto_destino.SetFocus
'      Exit Function
'   End If
   If Trim$(Txt_id_consignatario.Text) = "" Then
      MsgBox "Ingrese el tipo de identificacion del consignatario", vbInformation, App.Title
      valida_Datos = False
      If Txt_id_consignatario.Enabled Then Txt_id_consignatario.SetFocus
      Exit Function
   End If
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
   
   ' validacion incluida por publicacion del boletin 113 y 114 de la ADUANA
   If Cmb_regimen_adua.BoundText = "" And Cmb_almacen_temp.BoundText = "" Then
      MsgBox "El campo Régimen aduanero o el Almacen temporal debe ser ingresado, las Autoridades aduaneras rechazarán este envio.", vbInformation, App.Title
      valida_Datos = False
      If Cmb_regimen_adua.Enabled Then Cmb_regimen_adua.SetFocus
      Exit Function
   End If
   
   If Cmb_almacen_temp.BoundText = "" Then
      MsgBox "El campo Almacen temporal no fue ingresado (no es un problema pero las Autoridades aduaneras asignaran automáticamente un almacen cuando se realice el envio).", vbInformation, App.Title
      'valida_Datos = False
      'If Cmb_almacen_temp.Enabled Then Cmb_almacen_temp.SetFocus
      'Exit Function
   End If
   '---------------------------------------------------------------------------
   
   
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
    carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
    If Not data_linea.Recordset.EOF Then data_linea.Recordset.MoveFirst
    carga_datacombo data_tipo_id_consignatario, Cmb_tipo_id_consignatario, "Select * from docu_tipo_id", "tipo_id_codigo", "tipo_id_descripcion"
    carga_datacombo data_tipo_id_notificador, Cmb_tipo_id_notificador, "Select * from docu_tipo_id", "tipo_id_codigo", "tipo_id_descripcion"
    carga_datacombo data_tipo_id_emb, Cmb_tipo_id_emb, "Select * from docu_tipo_id", "tipo_id_codigo", "tipo_id_descripcion"
    
    
    
    Call inicia_controles
    If Gstr_opc_bl = "A" Then
        habilita_detalle False
    Else
        habilita_detalle True
    End If
    Carga_cabecera_fletes " "
End Sub

Private Sub Grid_flete_DblClick()
    If Grid_flete.Row <= 0 Then Exit Sub
    Load doc_aux_flete_bl
    'doc_aux_flete_bl.Form_Load
    doc_aux_flete_bl.Cmb_linea.BoundText = Cmb_linea.BoundText
    Grid_flete.Col = 0: doc_aux_flete_bl.Cmb_cargo.BoundText = Grid_flete.Text
    Grid_flete.Col = 2: doc_aux_flete_bl.Cmb_debtor.BoundText = Grid_flete.Text
    If Cmb_linea.BoundText = "KHS" Then
        Grid_flete.Col = 4:  doc_aux_flete_bl.GTM_valor_unit.Value = CDbl(Grid_flete.Text)
        Grid_flete.Col = 5:  doc_aux_flete_bl.Cmb_nivel_cargo.BoundText = Grid_flete.Text
        Grid_flete.Col = 6:  doc_aux_flete_bl.GTM_cant.Value = CDbl(Grid_flete.Text)
        Grid_flete.Col = 7:  doc_aux_flete_bl.GTM_flete.Value = CDbl(Grid_flete.Text)
        Grid_flete.Col = 8:  doc_aux_flete_bl.Cmb_moneda.BoundText = Grid_flete.Text
        Grid_flete.Col = 9:  doc_aux_flete_bl.Txt_marcas.Text = Grid_flete.Text
        Grid_flete.Col = 10: doc_aux_flete_bl.Txt_pto.Text = Grid_flete.Text
        Grid_flete.Col = 11: doc_aux_flete_bl.Chk_impr.Value = IIf(Grid_flete.Text = "S", 1, 0)
    Else
        Grid_flete.Col = 4: doc_aux_flete_bl.GTM_flete.Value = CDbl(Grid_flete.Text)
        Grid_flete.Col = 5: doc_aux_flete_bl.Cmb_signo.ListIndex = IIf(Grid_flete.Text = "+", 0, IIf(Grid_flete.Text = "-", 1, -1))
        Grid_flete.Col = 6: doc_aux_flete_bl.Cmb_moneda.BoundText = Grid_flete.Text
        Grid_flete.Col = 7: doc_aux_flete_bl.Txt_marcas.Text = Grid_flete.Text
        Grid_flete.Col = 8: doc_aux_flete_bl.Txt_pto.Text = Grid_flete.Text
        Grid_flete.Col = 9: doc_aux_flete_bl.Chk_impr.Value = IIf(Grid_flete.Text = "S", 1, 0)
    End If
    
    doc_aux_flete_bl.Show 1
    If Gtyp_flete.reg_cancel = False Then
        Grid_flete.Col = 0: Grid_flete.Text = Gtyp_flete.cargo_cod
        Grid_flete.Col = 1: Grid_flete.Text = Gtyp_flete.cargo
        Grid_flete.Col = 2: Grid_flete.Text = Gtyp_flete.deptor_cod
        Grid_flete.Col = 3: Grid_flete.Text = Gtyp_flete.debtor
        If Cmb_linea.BoundText = "KHS" Then
            Grid_flete.Col = 4:  Grid_flete.Text = Format(Gtyp_flete.valor_unit, "###,###,##0.00")
            Grid_flete.Col = 5:  Grid_flete.Text = Gtyp_flete.nivel_cargo
            Grid_flete.Col = 6:  Grid_flete.Text = Format(Gtyp_flete.cantidad, "###,###,##0.00")
            Grid_flete.Col = 7:  Grid_flete.Text = Format(Gtyp_flete.valor, "###,###,##0.00")
            Grid_flete.Col = 8:  Grid_flete.Text = Gtyp_flete.MONEDA
            Grid_flete.Col = 9:  Grid_flete.Text = Gtyp_flete.marcas
            Grid_flete.Col = 10: Grid_flete.Text = Gtyp_flete.lugar
            Grid_flete.Col = 11: Grid_flete.Text = IIf(Gtyp_flete.imp_bl, "S", "N")
        Else
            Grid_flete.Col = 4: Grid_flete.Text = Format(Gtyp_flete.valor, "###,###,##0.00")
            Grid_flete.Col = 5: Grid_flete.Text = Gtyp_flete.SIGNO
            Grid_flete.Col = 6: Grid_flete.Text = Gtyp_flete.MONEDA
            Grid_flete.Col = 7: Grid_flete.Text = Gtyp_flete.marcas
            Grid_flete.Col = 8: Grid_flete.Text = Gtyp_flete.lugar
            Grid_flete.Col = 9: Grid_flete.Text = IIf(Gtyp_flete.imp_bl, "S", "N")
        End If
        'Grid_flete.Rows = Grid_flete.Rows + 1
    End If
    
End Sub

Private Sub GTD_date_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTD_hora_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_bultos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_bultos_LostFocus()
If CDbl(Lbl_no_det_act.Caption) <> 0 Then
   mani_bl_det(CDbl(Lbl_no_det_act.Caption)).no_bultos = GTM_bultos.Value
   actualiza_totales_bl
End If
End Sub

Private Sub GTM_bultos_mal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_flete_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_flete_otros_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{TAB}"
If KeyAscii = 13 Then Cmd_nuevo_det.SetFocus 'SendKeys "{TAB}"
STab_bl.Tab = 3
End Sub

Private Sub GTM_peso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_peso_LostFocus()
If CDbl(Lbl_no_det_act.Caption) <> 0 Then
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).peso = GTM_peso.Value
    actualiza_totales_bl
End If
End Sub

Private Sub GTM_peso_mal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_seguro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Cmd_nuevo_det.SetFocus 'SendKeys "{TAB}"
STab_bl.Tab = 3
End Sub

Private Sub GTM_volumen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub GTM_volumen_LostFocus()
If CDbl(Lbl_no_det_act.Caption) <> 0 Then
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).volumen = GTM_volumen.Value
    actualiza_totales_bl
End If
End Sub

Private Sub Msk_fecha_emb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Msk_fecha_trasb_KeyPress(KeyAscii As Integer)
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
 If Lbl_no_det_act.Caption <> "0" Then
    If Trim(Txt_contenedor.Text) <> "" Then
        mani_bl_det(CDbl(Lbl_no_det_act.Caption)).CONTEINER = Txt_contenedor.Text
        Call ubica_contenedor
    Else
        mani_bl_det(CDbl(Lbl_no_det_act.Caption)).CONTEINER = ""
    End If
    actualiza_totales_bl
End If
End Sub

Private Sub Txt_desc_carga_LostFocus()
If CDbl(Lbl_no_det_act.Caption) <> 0 Then
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).desc_carga = Txt_desc_carga.Text
End If
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

Private Sub Txt_direc_notificador2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{TAB}"
   STab_bl.Tab = 1
End If
End Sub

Private Sub txt_docum_auto_KeyPress(KeyAscii As Integer)
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

Private Sub Txt_id_emb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Txt_id_emb_LostFocus()
On Error GoTo Errores
    Dim sql As String
    Dim reg As New ADODB.Recordset
    If Trim$(Txt_id_emb.Text) <> "" Then
        sql = "Select * from docu_consignatario " & _
              " Where consig_tipo_id = '" & Cmb_tipo_id_emb.BoundText & "'" & _
              "   And consig_codigo  = '" & Trim$(Txt_id_emb.Text) & "'"
        
        reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        If Not reg.EOF Then
            
            Txt_nombre_embarcador = reg.Fields("consig_nombre")
            Txt_direc_embarcador = reg.Fields("consig_direccion")
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
    Cmb_tipo_prod.BoundText = ""
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'TL pongo en comentarios para mantener la descripcion  ingresada
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Txt_desc_carga.Text = ""
    Txt_just_cambio.Text = ""
    Cmb_IMO.BoundText = ""
    Cmb_ONU.BoundText = ""
End Sub

Private Sub Txt_just_cambio_LostFocus()
If CDbl(Lbl_no_det_act.Caption) <> 0 Then
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).just_cambio = Txt_just_cambio.Text
End If
End Sub

Private Sub Txt_marcas_LostFocus()
If CDbl(Lbl_no_det_act.Caption) <> 0 Then
    mani_bl_det(CDbl(Lbl_no_det_act.Caption)).marca_num = Txt_marcas.Text
End If
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
    Cmb_tipo_prod.BoundText = mani_bl_det(CDbl(Lbl_no_det_act.Caption)).tipo_prod
    Txt_just_cambio.Text = mani_bl_det(CDbl(Lbl_no_det_act.Caption)).just_cambio
    Cmb_IMO.BoundText = Trim(mani_bl_det(CDbl(Lbl_no_det_act.Caption)).Imo)
    Cmb_ONU.BoundText = Trim(mani_bl_det(CDbl(Lbl_no_det_act.Caption)).pelig_ONU)
End Sub

Public Sub habilita_detalle(opc As Boolean)
    Cmb_tipo_carga.Enabled = opc
    Txt_contenedor.Enabled = opc
    GTM_bultos.Enabled = opc
    Cmb_embalaje.Enabled = opc
    GTM_peso.Enabled = opc
    GTM_volumen.Enabled = opc
    Txt_marcas.Enabled = opc
    Cmb_tipo_prod.Enabled = opc
    Txt_desc_carga.Enabled = opc
    Txt_just_cambio.Enabled = opc
    Cmb_IMO.Enabled = opc
    Cmb_ONU.Enabled = opc
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
               " from docu_blcab " & _
               " Where blcab_linea       ='" & linea & "'" & _
               "   and blcab_buque       ='" & buque & "'" & _
               "   and blcab_viaje       ='" & viaje & "'" & _
               "   and blcab_puer_emb    ='" & pto_emb & "'" & _
               "   and blcab_puer_des    ='" & pto_desc & "'" & _
               "   and blcab_cono_emb    ='" & bl & "'"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    If Not reg.EOF Then
        Cmb_linea.BoundText = reg("blcab_linea")
        Cmb_buque.BoundText = reg("blcab_buque")
        Cmb_viaje.BoundText = reg("blcab_viaje")
        Call Cmb_viaje_Change
        Cmb_pto_emb.BoundText = reg("blcab_puer_emb")
        txt_bl.Text = reg("blcab_cono_emb")
        Txt_pto_origen.Text = reg("blcab_puer_origen")
        Txt_pto_descarga.Text = reg("blcab_puer_des")
        Txt_pto_destino.Text = reg("blcab_pto_dest_final")
        Msk_fecha_emb.Text = reg("blcab_fecha_emb")
        Txt_pto_trasb.Text = IIf(Not IsNull(reg("blcab_puer_trasb")), reg("blcab_puer_trasb"), "")
        Msk_fecha_trasb.Text = IIf(Not IsNull(reg("blcab_fecha_trasb")), reg("blcab_fecha_trasb"), "")
        If Msk_fecha_trasb.Text = "01/01/1900" Then Msk_fecha_trasb.Text = ""
        
        Cmb_tipo_id_consignatario.BoundText = IIf(Not IsNull(reg("blcab_tipo_ide_cons")), reg("blcab_tipo_ide_cons"), "")
        Txt_id_consignatario.Text = IIf(Not IsNull(reg("blcab_ide_cons")), reg("blcab_ide_cons"), "")
        Txt_nombre_consignatario.Text = IIf(Not IsNull(reg("blcab_consig")), reg("blcab_consig"), "")
        Txt_nac_consignatario.Text = IIf(Not IsNull(reg("blcab_nacio_consig")), reg("blcab_nacio_consig"), "")
        Txt_direc_consignatario.Text = IIf(Not IsNull(reg("blcab_direc_consig")), reg("blcab_direc_consig"), "")
        
        Cmb_tipo_id_notificador.BoundText = IIf(Not IsNull(reg("blcab_tipo_ide_notif")), reg("blcab_tipo_ide_notif"), "")
        Txt_id_notificador.Text = IIf(Not IsNull(reg("blcab_ide_notif")), reg("blcab_ide_notif"), "")
        Txt_nombre_notificador.Text = IIf(Not IsNull(reg("blcab_notif")), reg("blcab_notif"), "")
        Txt_direc_notificador.Text = IIf(Not IsNull(reg("blcab_direc_notif")), reg("blcab_direc_notif"), "")
        
        Cmb_tipo_id_emb.BoundText = IIf(Not IsNull(reg("blcab_tipo_ide_emb")), reg("blcab_tipo_ide_emb"), "")
        Txt_id_emb.Text = IIf(Not IsNull(reg("blcab_ide_emb")), reg("blcab_ide_emb"), "")
        Txt_nombre_embarcador.Text = IIf(Not IsNull(reg("blcab_embarcador")), reg("blcab_embarcador"), "")
        Txt_direc_embarcador.Text = IIf(Not IsNull(reg("blcab_direc_embarc")), reg("blcab_direc_embarc"), "")
        
        Txt_nombre_notificador2.Text = IIf(Not IsNull(reg("blcab_notif2")), reg("blcab_notif2"), "")
        Txt_direc_notificador2.Text = IIf(Not IsNull(reg("blcab_direc_notif2")), reg("blcab_direc_notif2"), "")
        
        Txt_peso_tot.Text = Format(IIf(Not IsNull(reg("blcab_peso_tot")), reg("blcab_peso_tot"), "0"), "###,###,###.00")
        Txt_volumen_tot.Text = Format(IIf(Not IsNull(reg("blcab_volumen_tot")), reg("blcab_volumen_tot"), "0"), "###,###,###.00")
        Txt_bultos_tot.Text = Format(IIf(Not IsNull(reg("blcab_bultos_tot")), reg("blcab_bultos_tot"), "0"), "###,###,###.00")
        Txt_cont_tot.Text = Format(IIf(Not IsNull(reg("blcab_no_cont")), reg("blcab_no_cont"), "0"), "###,###,###.00")
        Cmb_tipo_docum.BoundText = IIf(Not IsNull(reg("blcab_tipo_doc")), reg("blcab_tipo_doc"), "")
        Cmb_almacen_temp.BoundText = IIf(Not IsNull(reg("blcab_almac_temp")), reg("blcab_almac_temp"), "")
        Cmb_tipo_deposito.BoundText = IIf(Not IsNull(reg("blcab_tipo_depos")), reg("blcab_tipo_depos"), "")
        Cmb_deposito.BoundText = IIf(Not IsNull(reg("blcab_deposito_adua")), reg("blcab_deposito_adua"), "")
        Cmb_regimen_adua.BoundText = IIf(Not IsNull(reg("blcab_regimen_adu")), reg("blcab_regimen_adu"), "")
        GTD_date.DateValue = IIf(Not IsNull(reg("blcab_fecha_trans")), reg("blcab_fecha_trans"), Null)
        If IsDate(reg("blcab_hora_trans")) Then GTD_hora.DateValue = CDate(reg("blcab_hora_trans"))
        Cmb_tipo_salida.BoundText = IIf(Not IsNull(reg("blcab_tipo_salida")), reg("blcab_tipo_salida"), "")
        txt_docum_auto.Text = IIf(Not IsNull(reg("blcab_docum_aut")), reg("blcab_docum_aut"), "")
        GTM_peso_mal.Value = CDbl(IIf(Not IsNull(reg("blcab_pesob_mal_est")), reg("blcab_pesob_mal_est"), 0))
        GTM_bultos_mal.Value = CDbl(IIf(Not IsNull(reg("blcab_bultos_mal_est")), reg("blcab_bultos_mal_est"), 0))
        Cmb_oper_carga.BoundText = IIf(Not IsNull(reg("blcab_emp_resp")), reg("blcab_emp_resp"), "")
        Cmb_nav_carga.BoundText = IIf(Not IsNull(reg("blcab_emp_anav")), reg("blcab_emp_anav"), "")
        Cmb_metodo_pago.BoundText = IIf(Not IsNull(reg("blcab_metod_pago")), reg("blcab_metod_pago"), "")
        Cmb_moneda.BoundText = IIf(Not IsNull(reg("blcab_moneda")), reg("blcab_moneda"), "")
        GTM_flete.Value = CDbl(IIf(Not IsNull(reg("blcab_flete")), reg("blcab_flete"), 0))
        Cmb_metodo_pago_otros.BoundText = IIf(Not IsNull(reg("blcab_metod_pago_otros")), reg("blcab_metod_pago_otros"), "")
        Cmb_moneda_otros.BoundText = IIf(Not IsNull(reg("blcab_moneda_otros")), reg("blcab_moneda_otros"), "")
        GTM_flete_otros.Value = CDbl(IIf(Not IsNull(reg("blcab_otros_gastos")), reg("blcab_otros_gastos"), 0))
        Cmb_moneda_seguro.BoundText = IIf(Not IsNull(reg("blcab_moneda_seguro")), reg("blcab_moneda_seguro"), "")
        GTM_seguro.Value = CDbl(IIf(Not IsNull(reg("blcab_seguro")), reg("blcab_seguro"), 0))

        sql = "Select * " & _
               " from docu_bldet " & _
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
            mani_bl_det(i).just_cambio = IIf(Not IsNull(REG_det("bldet_just_cambio")), REG_det("bldet_just_cambio"), "")
            mani_bl_det(i).Imo = IIf(Not IsNull(REG_det("bldet_imo")), REG_det("bldet_imo"), "")
            mani_bl_det(i).pelig_ONU = IIf(Not IsNull(REG_det("bldet_pelig_onu")), REG_det("bldet_pelig_onu"), "")
            mani_bl_det(i).tipo_prod = IIf(Not IsNull(REG_det("bldet_tipo_prod")), REG_det("bldet_tipo_prod"), "")
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
               " from docu_blflete " & _
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
                Grid_flete.Col = 1: Grid_flete.Text = devuelve_desc("docu_cargo", "cargo_codigo", "cargo_descripcion", REG_flete("flete_cargo"), "cargo_linea = '" & Cmb_linea.BoundText & "'")
            Else
                Grid_flete.Col = 0: Grid_flete.Text = ""
                Grid_flete.Col = 1: Grid_flete.Text = ""
            End If
            If Not IsNull(REG_flete("flete_debtor")) Then
                Grid_flete.Col = 2: Grid_flete.Text = REG_flete("flete_debtor")
                Grid_flete.Col = 3: Grid_flete.Text = devuelve_desc("docu_debtor", "debtor_codigo", "debtor_descripcion", REG_flete("flete_debtor"), "debtor_linea = '" & Cmb_linea.BoundText & "'")
            Else
                Grid_flete.Col = 2: Grid_flete.Text = ""
                Grid_flete.Col = 3: Grid_flete.Text = ""
            End If
            If Cmb_linea.BoundText = "KHS" Then
                Grid_flete.Col = 4:  Grid_flete.Text = Format(REG_flete("flete_valor_unit"), "###,###,##0.00")
                Grid_flete.Col = 5:  Grid_flete.Text = IIf(Not IsNull(REG_flete("flete_per")), REG_flete("flete_per"), "")
                Grid_flete.Col = 6:  Grid_flete.Text = REG_flete("flete_Cantidad")
                Grid_flete.Col = 7:  Grid_flete.Text = Format(REG_flete("flete_valor"), "###,###,##0.00")
                Grid_flete.Col = 8:  Grid_flete.Text = IIf(Not IsNull(REG_flete("flete_moneda")), REG_flete("flete_moneda"), "")
                Grid_flete.Col = 9:  Grid_flete.Text = IIf(Not IsNull(REG_flete("flete_marcas")), REG_flete("flete_marcas"), "")
                Grid_flete.Col = 10: Grid_flete.Text = IIf(Not IsNull(REG_flete("flete_lugar")), REG_flete("flete_lugar"), "")
                Grid_flete.Col = 11: Grid_flete.Text = REG_flete("flete_imp_bl")
            Else
                Grid_flete.Col = 4: Grid_flete.Text = Format(REG_flete("flete_valor"), "###,###,##0.00")
                Grid_flete.Col = 5: Grid_flete.Text = REG_flete("flete_signo")
                Grid_flete.Col = 6: Grid_flete.Text = IIf(Not IsNull(REG_flete("flete_moneda")), REG_flete("flete_moneda"), "")
                Grid_flete.Col = 7: Grid_flete.Text = IIf(Not IsNull(REG_flete("flete_marcas")), REG_flete("flete_marcas"), "")
                Grid_flete.Col = 8: Grid_flete.Text = IIf(Not IsNull(REG_flete("flete_lugar")), REG_flete("flete_lugar"), "")
                Grid_flete.Col = 9: Grid_flete.Text = REG_flete("flete_imp_bl")
            End If
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
mani_bl_det(i).Imo = ""
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
mani_bl_det(i).tipo_prod = ""
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
    GTD_date.Text = ""
    GTD_hora.Text = ""
    GTM_peso_mal.Value = 0
    GTM_bultos_mal.Value = 0
    GTM_flete.Value = 0
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

Public Sub ubica_contenedor()
Dim reg As New ADODB.Recordset

sql = "Select * from docu_cont " & _
               " Where cont_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and cont_buque    ='" & Cmb_buque.BoundText & "'" & _
               "   and cont_viaje    ='" & Cmb_viaje.BoundText & "'" & _
               "   and cont_puer_emb ='" & Cmb_pto_emb.BoundText & "'" & _
               "   and cont_puer_des ='" & Txt_pto_descarga.Text & "'" & _
               "   and cont_contenedor='" & Txt_contenedor.Text & "'"

reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
If Not reg.EOF Then
      mani_bl_det(CDbl(Lbl_no_det_act.Caption)).Sello = reg("cont_sello1")
      mani_bl_det(CDbl(Lbl_no_det_act.Caption)).Tipo_cont = "" 'reg("cont_tipo_cont")
      mani_bl_det(CDbl(Lbl_no_det_act.Caption)).tipo_equip = reg("cont_tipo_equip")
      mani_bl_det(CDbl(Lbl_no_det_act.Caption)).tamano_equip = reg("cont_tam_equip")
      mani_bl_det(CDbl(Lbl_no_det_act.Caption)).cond_cont = reg("cont_condicion")
Else
    Dim frm_cont As New doc_contenedor
    frm_cont.Show
    frm_cont.Gbol_CONT_desde_menu = False
    frm_cont.cmd_agregar_Click
    frm_cont.Cmb_linea.BoundText = Cmb_linea.BoundText
    frm_cont.Cmb_buque.BoundText = Cmb_buque.BoundText
    frm_cont.Cmb_viaje.BoundText = Cmb_viaje.BoundText
    frm_cont.Cmb_pto_emb.BoundText = Cmb_pto_emb.BoundText
    frm_cont.Txt_pto_descarga.Text = Txt_pto_descarga.Text
    frm_cont.Txt_contenedor.Text = Txt_contenedor.Text
    Set frm_cont.forma_bl = Me
    frm_cont.Caption = Txt_contenedor.Text
End If
reg.Close
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

Private Sub Txt_nombre_notificador2_KeyPress(KeyAscii As Integer)
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

Sub Carga_cabecera_fletes(linea As String)
If linea = "KHS" Then
    Grid_flete.Cols = 12
    Grid_flete.Row = 0
    Grid_flete.Col = 0:  Grid_flete.Text = "Cargo Cod":   Grid_flete.ColWidth(0) = 600
    Grid_flete.Col = 1:  Grid_flete.Text = "Cargo Desc":  Grid_flete.ColWidth(1) = 1000
    Grid_flete.Col = 2:  Grid_flete.Text = "Debtor Cod":  Grid_flete.ColWidth(2) = 600
    Grid_flete.Col = 3:  Grid_flete.Text = "Debtor Desc": Grid_flete.ColWidth(3) = 1000
    Grid_flete.Col = 4:  Grid_flete.Text = "Valor Unit":  Grid_flete.ColWidth(4) = 900
    Grid_flete.Col = 5:  Grid_flete.Text = "Nivel Cargo": Grid_flete.ColWidth(5) = 900
    Grid_flete.Col = 6:  Grid_flete.Text = "Cantidad":    Grid_flete.ColWidth(6) = 900
    Grid_flete.Col = 7:  Grid_flete.Text = "Valor Total": Grid_flete.ColWidth(7) = 900
    Grid_flete.Col = 8:  Grid_flete.Text = "Moneda":      Grid_flete.ColWidth(8) = 700
    Grid_flete.Col = 9:  Grid_flete.Text = "Marcas":      Grid_flete.ColWidth(9) = 2000
    Grid_flete.Col = 10: Grid_flete.Text = "Lugar Pago":  Grid_flete.ColWidth(10) = 900
    Grid_flete.Col = 11: Grid_flete.Text = "Impr Bl":     Grid_flete.ColWidth(11) = 700
Else
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
End If

End Sub

' procedimiento que realiza la grabacion de los fletes de la carga
' para la linea KHS se usan mas campos que para MSC por lo cual se
' hace un if de la linea ya que se tienen otras columnas asignadas
' para cada linea, por lo que en algunos  casos se graba en blanco o cero
Sub graba_fletes()
On Error GoTo Errores
    
    If Grid_flete.Rows <= 2 Then Exit Sub
    
    For i = 1 To Grid_flete.Rows - 2
        Grid_flete.Row = i
        sql = "Insert into docu_blflete values(" & _
              "'" & Cmb_linea.BoundText & "',  '" & Cmb_buque.BoundText & "'," & _
              "'" & Cmb_viaje.BoundText & "',  '" & Cmb_pto_emb.BoundText & "'," & _
              "'" & Txt_pto_descarga.Text & "','" & Cmb_oper_carga.BoundText & "', " & _
              "'" & txt_bl.Text & "'," & i & ","
        Grid_flete.Col = 0: sql = sql & "'" & Grid_flete.Text & "', "
        Grid_flete.Col = 2: sql = sql & "'" & Grid_flete.Text & "', "
        If Cmb_linea.BoundText = "KHS" Then
            Grid_flete.Col = 7:  sql = sql & " " & CDbl(Grid_flete.Text) & ",  "    ' valor total
                                 sql = sql & "'+',"                                 ' signo
            Grid_flete.Col = 8:  sql = sql & "'" & Grid_flete.Text & "', "          ' moneda
            Grid_flete.Col = 9:  sql = sql & "'" & Grid_flete.Text & "', "          ' marcas
            Grid_flete.Col = 10: sql = sql & "'" & Grid_flete.Text & "', "          ' lugar
            Grid_flete.Col = 11: sql = sql & "'" & Grid_flete.Text & "', "          ' impr bl
            Grid_flete.Col = 4:  sql = sql & " " & CDbl(Grid_flete.Text) & ",  "    ' valor unit
            Grid_flete.Col = 5:  sql = sql & "'" & chk_str(Grid_flete.Text, "S") & "', "          ' nivel de cargo (per)
            Grid_flete.Col = 6:  sql = sql & " " & CDbl(Grid_flete.Text) & ")"      ' cantidad
        Else
            Grid_flete.Col = 4: sql = sql & " " & CDbl(Grid_flete.Text) & ",  "     ' valor flete
            Grid_flete.Col = 5: sql = sql & "'" & Grid_flete.Text & "', "           ' signo
            Grid_flete.Col = 6: sql = sql & "'" & Grid_flete.Text & "', "           ' moneda
            Grid_flete.Col = 7: sql = sql & "'" & Grid_flete.Text & "', "           ' marcas
            Grid_flete.Col = 8: sql = sql & "'" & Grid_flete.Text & "', "           ' lugar
            Grid_flete.Col = 9: sql = sql & "'" & Grid_flete.Text & "')" ', 0,'',0)"    ' impr_bl, valor unit, nivel cargo, cantidad
            ' OJO ESTOS CAMPOS SI DEBERIAN EXISTIR EN LA TABLA
        End If
        Gcon_docu.Execute sql
    Next i
    
    Exit Sub
Errores:
    MsgBox "Ocurrio el Error " & Err.Description
    Exit Sub
    Resume 0
End Sub

Private Sub Txt_pto_trasb_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
