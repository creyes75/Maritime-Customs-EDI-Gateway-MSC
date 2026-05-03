VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "GTDATE32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form doc_archivo_CAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación Archivo"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "doc_archivo_CAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6210
   Begin VB.Frame Frame1 
      Height          =   6060
      Left            =   45
      TabIndex        =   10
      Top             =   15
      Width           =   6030
      Begin VB.CheckBox Chk_envio_pto 
         Caption         =   " Envío por Puerto "
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Frame fra_puerto 
         Height          =   2415
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   5775
         Begin VB.Frame Fra_pto_descarga 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   1695
            TabIndex        =   31
            Top             =   255
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
               Picture         =   "doc_archivo_CAE.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   33
               ToolTipText     =   "Buscar"
               Top             =   255
               Width           =   300
            End
            Begin VB.TextBox Txt_pto_descarga 
               Height          =   285
               Left            =   120
               MaxLength       =   5
               TabIndex        =   32
               Top             =   270
               Width           =   975
            End
            Begin VB.Label Lbl_pto_desc 
               AutoSize        =   -1  'True
               Caption         =   "Puerto Descarga"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   34
               Top             =   15
               Width           =   1200
            End
         End
         Begin VB.Frame Frame7 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   120
            TabIndex        =   28
            Top             =   285
            Width           =   1500
            Begin MSDataListLib.DataCombo Cmb_pto_emb 
               Height          =   315
               Left            =   135
               TabIndex        =   29
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
            Begin VB.Label Lbl_pto_emb 
               AutoSize        =   -1  'True
               Caption         =   "Puerto Embarque"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   105
               TabIndex        =   30
               Top             =   -15
               Width           =   1230
            End
         End
         Begin VB.ListBox Lst_bl 
            Height          =   1410
            Left            =   1695
            Style           =   1  'Checkbox
            TabIndex        =   27
            Top             =   945
            Width           =   3930
         End
         Begin VB.CheckBox Chk_todos 
            Caption         =   "Todos BLs"
            Height          =   255
            Left            =   135
            TabIndex        =   26
            Top             =   1290
            Width           =   1410
         End
         Begin MSDataListLib.DataCombo Cmb_oper_carga 
            Height          =   315
            Left            =   3285
            TabIndex        =   35
            Top             =   480
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
            Left            =   4710
            Top             =   480
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
            Caption         =   "Agen. Oper. de Carg"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   3315
            TabIndex        =   37
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Bill of Lading"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   135
            TabIndex        =   36
            Top             =   960
            Width           =   900
         End
      End
      Begin VB.OptionButton Opt_tipo_impo 
         Caption         =   "Arch. Importación"
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
         Left            =   1335
         TabIndex        =   24
         Top             =   240
         Width           =   2115
      End
      Begin VB.OptionButton Opt_tipo_expo 
         Caption         =   "Arch. Exportación"
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
         Left            =   3495
         TabIndex        =   23
         Top             =   255
         Width           =   2205
      End
      Begin VB.CheckBox Chk_mail 
         Caption         =   "Enviar Correo Automáticamente"
         Height          =   195
         Left            =   2790
         TabIndex        =   22
         Top             =   5625
         Width           =   2925
      End
      Begin VB.Frame Fra_Date 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   585
         Left            =   195
         TabIndex        =   20
         Top             =   5385
         Width           =   2025
         Begin GTMaskDate.GTMaskDate GTD_date 
            Height          =   285
            Left            =   30
            TabIndex        =   6
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
            TabIndex        =   7
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
            Caption         =   "Fecha/Hora Transacción"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   30
            TabIndex        =   21
            Top             =   30
            Width           =   1800
         End
      End
      Begin VB.TextBox txt_just 
         Height          =   285
         Left            =   225
         MaxLength       =   60
         TabIndex        =   5
         Top             =   5055
         Width           =   5580
      End
      Begin VB.Frame Fra_viaje 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   3255
         TabIndex        =   15
         Top             =   1050
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
            TabIndex        =   16
            Top             =   0
            Width           =   345
         End
      End
      Begin VB.Frame Fra_buque 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   135
         TabIndex        =   13
         Top             =   1050
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
            TabIndex        =   14
            Top             =   0
            Width           =   465
         End
      End
      Begin VB.Frame Fra_linea 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   1200
         TabIndex        =   11
         Top             =   480
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
            TabIndex        =   12
            Top             =   0
            Width           =   825
         End
      End
      Begin MSDataListLib.DataCombo Cmb_tipo_transaccion 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   4410
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc data_tipo_transaccion 
         Height          =   330
         Left            =   2205
         Top             =   4395
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
      Begin MSDataListLib.DataCombo Cmb_tipo_envio 
         Height          =   315
         Left            =   3390
         TabIndex        =   4
         Top             =   4410
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc data_tipo_envio 
         Height          =   330
         Left            =   4710
         Top             =   4395
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
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
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "doc_archivo_CAE.frx":083C
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Justificación Cambios:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   19
         Top             =   4830
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Transacción :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   225
         TabIndex        =   18
         Top             =   4185
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Envío :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   3390
         TabIndex        =   17
         Top             =   4185
         Width           =   1245
      End
   End
   Begin Threed.SSCommand cmd_Proceso 
      Height          =   615
      Left            =   4440
      TabIndex        =   8
      Top             =   6120
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   1085
      _StockProps     =   78
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
      Picture         =   "doc_archivo_CAE.frx":6AC6
   End
   Begin Threed.SSCommand cmd_salir 
      Height          =   615
      Left            =   5280
      TabIndex        =   9
      Top             =   6120
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   1085
      _StockProps     =   78
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
      Picture         =   "doc_archivo_CAE.frx":6DE0
   End
   Begin MSMAPI.MAPIMessages mapMess 
      Left            =   660
      Top             =   6090
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
      Top             =   6045
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "doc_archivo_CAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bNewSession     As Boolean ' utilizada para saber si se dio logon o no a la conexion de correo de outlook
Dim mail_to         As String
Dim mail_subject    As String
Dim mail_text       As String

Private Sub Chk_envio_pto_Click()
If Chk_envio_pto = 1 Then
    fra_puerto.Enabled = True
Else
    fra_puerto.Enabled = False
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

Private Sub cmb_linea_Change()
carga_datacombo data_buque, Cmb_buque, "Select * from docu_buque Where buqu_linea = '" & Cmb_linea.BoundText & "' and buqu_status ='A'", "buqu_codigo", "buqu_descripcion"
carga_datacombo data_oper_carga, Cmb_oper_carga, "Select * from docu_emp_resp Where emp_resp_linea = '" & Cmb_linea.BoundText & "' and emp_resp_status ='A'", "emp_resp_codigo", "emp_resp_descripcion"
carga_datacombo data_tipo_transaccion, Cmb_tipo_transaccion, "Select * from docu_tipo_transaccion WHERE tipo_tran_aplica ='S'", "tipo_tran_codigo", "tipo_tran_descripcion"
'carga_datacombo data_tipo_envio, Cmb_tipo_envio, "Select * from docu_tipo_envio Where tipo_env_linea ='" & Cmb_linea.BoundText & "'", "tipo_env_codigo", "tipo_env_descripcion"

carga_parametros Cmb_linea.BoundText
Txt_pto_descarga.Text = Gstr_param.pto_descarga
Cmb_oper_carga.BoundText = Gstr_param.naviera

End Sub

Private Sub Cmb_pto_emb_Change()
Dim reg As New ADODB.Recordset

'If Opt_bl.Value Then
    If Opt_tipo_impo.Value Then
        sql = "Select blcab_cono_emb " & _
                   " from docu_blcab " & _
                   " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
                   "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
                   "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'" & _
                   "   and blcab_puer_emb    ='" & Cmb_pto_emb.BoundText & "'"
    Else
        sql = "Select blcab_cono_emb " & _
                   " from docu_blcab_exp " & _
                   " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
                   "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
                   "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'" & _
                   "   and blcab_puer_des    ='" & Cmb_pto_emb.BoundText & "'"
    End If
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Lst_bl.Clear
    Do While Not reg.EOF
        Lst_bl.AddItem reg("blcab_cono_emb")
        reg.MoveNext
    Loop
'End If

End Sub

Private Sub Cmb_tipo_transaccion_Change()
    Select Case Cmb_tipo_transaccion.BoundText
       Case "00": carga_datacombo data_tipo_envio, Cmb_tipo_envio, "Select * from docu_tipo_envio Where tipo_env_linea ='" & Cmb_linea.BoundText & "' and tipo_env_status='A'", "tipo_env_codigo", "tipo_env_descripcion"
       Case "30": carga_datacombo data_tipo_envio, Cmb_tipo_envio, "Select * from docu_tipo_envio Where tipo_env_linea ='" & Cmb_linea.BoundText & "'AND tipo_env_codigo in ('21') and tipo_env_status='A'", "tipo_env_codigo", "tipo_env_descripcion"
       Case "03": carga_datacombo data_tipo_envio, Cmb_tipo_envio, "Select * from docu_tipo_envio Where tipo_env_linea ='" & Cmb_linea.BoundText & "'AND tipo_env_codigo in ('09','21') and tipo_env_status='A'", "tipo_env_codigo", "tipo_env_descripcion"
       Case "06": carga_datacombo data_tipo_envio, Cmb_tipo_envio, "Select * from docu_tipo_envio Where tipo_env_linea ='" & Cmb_linea.BoundText & "'AND tipo_env_codigo in ('09','21') and tipo_env_status='A'", "tipo_env_codigo", "tipo_env_descripcion"
    End Select
End Sub

Private Sub Cmb_viaje_Change()
   Dim sql As String
   If Opt_tipo_impo.Value Then
        sql = "Select * from docu_manifiesto " & _
              " Where mani_linea ='" & Cmb_linea.BoundText & "'" & _
              "   And mani_buque ='" & Cmb_buque.BoundText & "'" & _
              "   And mani_viaje ='" & Cmb_viaje.BoundText & "'"
              '"   And buqv_status = 'A'"
        carga_datacombo data_pto_emb, Cmb_pto_emb, sql, "mani_puer_emb", "mani_puer_emb"
   Else
        sql = "Select * from docu_manifiesto_exp " & _
              " Where mani_linea ='" & Cmb_linea.BoundText & "'" & _
              "   And mani_buque ='" & Cmb_buque.BoundText & "'" & _
              "   And mani_viaje ='" & Cmb_viaje.BoundText & "'"
              '"   And buqv_status = 'A'"
        carga_datacombo data_pto_emb, Cmb_pto_emb, sql, "mani_puer_des", "mani_puer_des"
   End If

End Sub

Private Sub Cmd_hlp_pto_descarga_Click()
If Opt_tipo_impo.Value Then
   Txt_pto_descarga.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Descarga")
Else
   Txt_pto_descarga.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Embarque")
End If
End Sub

Private Sub cmd_Proceso_Click()

If Cmb_linea.MatchedWithList And Cmb_buque.MatchedWithList And Cmb_viaje.MatchedWithList And _
   Txt_pto_descarga.Text <> "" And Cmb_oper_carga.MatchedWithList And Cmb_tipo_transaccion.MatchedWithList And _
   Cmb_tipo_envio.MatchedWithList And IsDate(GTD_date.DateValue) And IsDate(GTD_hora.DateValue) Then
              
        If (Not Cmb_pto_emb.MatchedWithList And Chk_envio_pto.Value = 1) Then GoTo Error_parametro:
        
        
        If Opt_tipo_impo.Value Then
            If Not Valida_datos_CAE(Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, Cmb_pto_emb.BoundText) Then
                If MsgBox("Existen errores en los datos, por lo cual se pueden presentar problemas con la CAE, Desea continuar de todas maneras?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
            End If
            Call actualiza_totales_manifiesto(Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, Cmb_pto_emb.BoundText, Txt_pto_descarga.Text, Cmb_oper_carga.BoundText) ' se pone para que siempre esten actualizados los totales antes del envio
            Call exporta_datos(Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, Cmb_pto_emb.BoundText, Txt_pto_descarga.Text, Cmb_oper_carga.BoundText)
        Else
            If Not Valida_datos_CAE_expo(Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, Cmb_pto_emb.BoundText) Then
                If MsgBox("Existen errores en los datos, por lo cual se pueden presentar problemas con la CAE, Desea continuar de todas maneras?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
            End If
            Call actualiza_totales_manifiesto_expo(Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, Txt_pto_descarga.Text, Cmb_pto_emb.BoundText, Cmb_oper_carga.BoundText)   ' se pone para que siempre esten actualizados los totales antes del envio
            Call exporta_datos(Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, Txt_pto_descarga.Text, Cmb_pto_emb.BoundText, Cmb_oper_carga.BoundText)
        End If
        
   Else
Error_parametro:
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
    Opt_tipo_impo.Value = True
    carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
    If Not data_linea.Recordset.EOF Then data_linea.Recordset.MoveFirst
    GTD_date.DateValue = Date
    GTD_hora.DateValue = Time
    Carga_parametros_MAIL
    Chk_envio_pto.Value = 1
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TIPO TRANSACCION              ARCHIVOS INCLUIDOS                                                  FORMATO
' 00 MANIFIESTO COMPLETO        MANHDR01,       MANDET01,      MANDET02,   MANCON01  Y ENVCTROL     CUSCAR
' 30 CORRECTOR MANIFIESTO       MANHDR01,       MANDET01,      MANDET02,  POR CONFIR Y ENVCTROL     CARDAT
' 03 COMPLEMENTO DE MANIF       MANHDR01(*),    MANDET01(*)    --------    --------  Y ENVCTROL     CARDAT
' 06 FALTANTES Y SOBRANTE       MANHDR01(*),    MANDET01(*),   MANDET02    --------  Y ENVCTROL
'(*) Menos campos que el archivo original.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub exporta_datos(linea As String, buque As String, viaje As String, pto_emb As String, pto_descarga As String, oper_carga As String)
Dim sql         As String
Dim reg_mani    As New ADODB.Recordset
Dim reg_bls     As New ADODB.Recordset
Dim reg_blcab   As New ADODB.Recordset
Dim reg_bldet   As New ADODB.Recordset
Dim reg_cont    As New ADODB.Recordset
Dim reg_cont_dt As New ADODB.Recordset
Dim REG_viaje   As New ADODB.Recordset
Dim reg_clave   As New ADODB.Recordset
Dim emp_anav    As String
Dim sep         As String
Dim del         As String
Dim x
Dim cadena      As String
Dim No_bls      As Integer
Dim descrip_marcas As String
Dim descrip_carga  As String
Dim viaje_trunc    As String
Dim ruta_archivo   As String
Dim total_docs     As Double

On Error GoTo errores

carga_parametros linea

sep = Gstr_param.separador '","
del = Gstr_param.delimitador ' "|"

If Me.Lst_bl.SelCount = 0 Then
   If Cmb_tipo_transaccion.BoundText = "00" Then
    If Chk_envio_pto.Value = 1 Then
        If MsgBox("No se ha seleccionado ningun bl, desea generar todo el Puerto.", vbYesNo + vbQuestion) = vbYes Then
           Chk_todos.Value = 1
        Else
            Exit Sub
        End If
     End If
   Else
        MsgBox "No se ha seleccionado ningun bl", vbQuestion
        Exit Sub
   End If
End If

    bls = ""
    For i = 0 To Lst_bl.ListCount - 1
        If Lst_bl.Selected(i) Then
           Lst_bl.ListIndex = i
           bls = bls & "'" & Lst_bl.Text & "', "
        End If
    Next i
    If Len(bls) > 0 Then bls = Mid(bls, 1, Len(Trim(bls)) - 1)    ' se quita la ultima coma

     x = FreeFile
     If Dir(Gstr_Ruta_files & "\" & buque, vbDirectory) = "" Then MkDir (Gstr_Ruta_files & "\" & buque)
     If Dir(Gstr_Ruta_files & "\" & buque & "\" & viaje, vbDirectory) = "" Then MkDir (Gstr_Ruta_files & "\" & buque & "\" & viaje)
     
     If Opt_tipo_impo.Value Then
        If Dir(Gstr_Ruta_files & "\" & buque & "\" & viaje & "\Import", vbDirectory) = "" Then MkDir (Gstr_Ruta_files & "\" & buque & "\" & viaje & "\Import")
        If Dir(Gstr_Ruta_files & "\" & buque & "\" & viaje & "\Import\" & pto_emb, vbDirectory) = "" Then MkDir (Gstr_Ruta_files & "\" & buque & "\" & viaje & "\Import\" & pto_emb)
        ruta_archivo = Gstr_Ruta_files & "\" & buque & "\" & viaje & "\Import\" & pto_emb
     Else
        If Dir(Gstr_Ruta_files & "\" & buque & "\" & viaje & "\Export", vbDirectory) = "" Then MkDir (Gstr_Ruta_files & "\" & buque & "\" & viaje & "\Export")
        If Dir(Gstr_Ruta_files & "\" & buque & "\" & viaje & "\Export\" & pto_descarga, vbDirectory) = "" Then MkDir (Gstr_Ruta_files & "\" & buque & "\" & viaje & "\Export\" & pto_descarga)
        ruta_archivo = Gstr_Ruta_files & "\" & buque & "\" & viaje & "\Export\" & pto_descarga
     End If
     
     If Dir(Gstr_Ruta_files & "\TMP", vbDirectory) = "" Then MkDir (Gstr_Ruta_files & "\TMP")
     
     If Dir(Gstr_Ruta_files & "TMP\MANHDR01.txt", vbArchive) <> "" Then Kill Gstr_Ruta_files & "TMP\MANHDR01.txt"
     If Dir(Gstr_Ruta_files & "TMP\MANDET01.txt", vbArchive) <> "" Then Kill Gstr_Ruta_files & "TMP\MANDET01.txt"
     If Dir(Gstr_Ruta_files & "TMP\MANDET02.txt", vbArchive) <> "" Then Kill Gstr_Ruta_files & "TMP\MANDET02.txt"
     If Dir(Gstr_Ruta_files & "TMP\MANCON01.txt", vbArchive) <> "" Then Kill Gstr_Ruta_files & "TMP\MANCON01.txt"
     
     If Dir(Gstr_Ruta_files & "TMP\", vbArchive) <> "" Then Kill Gstr_Ruta_files & "TMP\*.*"
    
     sql = "Select * from " & _
                    " docu_buque_viaje, docu_buque " & _
                    " Where buqv_linea    ='" & linea & "'" & _
                    "   and buqv_buque    ='" & buque & "'" & _
                    "   and buqv_viaje    ='" & viaje & "'" & _
                    "   and buqv_linea    = buqu_linea" & _
                    "   and buqv_buque    = buqu_codigo"
     REG_viaje.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
     
     
     '************************* M A N H D R 0 1 ********************************************
     Open Gstr_Ruta_files & "TMP\MANHDR01.txt" For Output As #x
      If Opt_tipo_impo.Value Then
            sql = "Select * from docu_manifiesto "
      Else
            sql = "Select * from docu_manifiesto_exp "
      End If
      sql = sql & "   Where mani_linea    ='" & linea & "'" & _
                    "   and mani_buque    ='" & buque & "'" & _
                    "   and mani_viaje    ='" & viaje & "'"
      If Chk_envio_pto.Value = 1 Then
          sql = sql & " and mani_puer_emb ='" & pto_emb & "'" & _
                    "   and mani_puer_des ='" & pto_descarga & "'"
      End If
     
     reg_mani.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
     If Not reg_mani.EOF Then
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       'TL revisar esto porque cuando se genera para determinados BL's y no todo el
       '   puerto esta totalizando mal, y esto traera el resultado correcto
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
           sql = "Select sum(blcab_peso_tot) as peso," & _
                       " sum(blcab_bultos_tot) bultos,sum(blcab_volumen_tot) volumen," & _
                       " sum(blcab_no_cont) cont, count(blcab_cono_emb) no_bls from "
           If Opt_tipo_impo.Value Then sql = sql & " docu_blcab " Else sql = sql & " docu_blcab_exp "
           sql = sql & " Where blcab_linea    ='" & linea & "'" & _
                       "   and blcab_buque    ='" & buque & "'" & _
                       "   and blcab_viaje    ='" & viaje & "'"
           If Chk_envio_pto.Value = 1 Then
               sql = sql & "   and blcab_puer_emb ='" & pto_emb & "'" & _
                           "   and blcab_puer_des ='" & pto_descarga & "'"
           End If
           If Chk_todos.Value = 0 And Chk_envio_pto.Value = 1 Then sql = sql & "   and blcab_cono_emb in (" & bls & ")"
           reg_bls.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
     
     
       cadena = del & reg_mani("mani_adua_ing") & del & sep & _
                 del & reg_mani("mani_ano_car") & del & sep & _
                 del & reg_mani("mani_tip_mani") & del & sep & _
                 del & reg_mani("mani_mani_car") & del & sep
       If Opt_tipo_impo.Value Then
            cadena = cadena & _
                 del & reg_mani("mani_puer_des") & del & sep
       Else
            cadena = cadena & _
                 del & reg_mani("mani_puer_emb") & del & sep
       End If
       cadena = cadena & _
                 del & reg_mani("mani_digi_ver") & del & sep
                 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                 'TL  Pongo este if para que las correcciones suban ya que estan dando'
                 '    error por el campo mani_ano_ord, que es el 7 campo del manhdr01 '
                 '    esta nulo, para depsues revisar como se debe llenar             '
                 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                 If Cmb_tipo_transaccion.BoundText = "00" Then
                    cadena = Trim$(cadena) & del & reg_mani("mani_ano_ord") & del & sep
                 Else
                    cadena = Trim$(cadena) & del & Year(Now) & del & sep
                 End If
                 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                 'TL aqui falta condicionar cuando toma los totales de docu_manifiesto y
                 '   cuando los toma del sum que debe hacer a la cabecera
                 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                 cadena = Trim$(cadena) & del & reg_mani("mani_num_ord") & del & sep & _
                       del & Cmb_tipo_transaccion.BoundText & del & sep & _
                       reg_bls("no_bls") & sep
'                 If Chk_todos.Value = 0 Then 'cra feb 2004
                       cadena = Trim$(cadena) & reg_bls("peso") & sep & _
                       reg_bls("bultos") & sep & _
                       reg_bls("volumen") & sep & _
                       reg_bls("cont") & sep
                       'reg_bls.Close
'                 Else      'cra feb 2004
'                       cadena = Trim$(cadena) & reg_mani("mani_tot_peso") & sep & _
'                       reg_mani("mani_tot_bult") & sep & _
'                       reg_mani("mani_tot_vol") & sep & _
'                       reg_mani("mani_tot_con") & sep
'                 End If
                 cadena = Trim$(cadena) & del & Cmb_tipo_envio.BoundText & del & sep & _
                 del & txt_just.Text & del

         'solo para correctores de exportaacion se envian estos datos o para manifiesto original en ambos casos (19 enero 2003)
         If Cmb_tipo_transaccion.BoundText = "00" Or (Cmb_tipo_transaccion.BoundText = "30" And Opt_tipo_expo.Value) Then    ' 19 sep 2002 cra es extraño ya que nunca se requeria esta inform. para correctores, solo para manif original
            'esto es para poder eliminar los ceros a la izquierda que la CAE no requiere, gana de ....molestar
            viaje_trunc = REG_viaje("buqv_viaje")
            Do While Mid(viaje_trunc, 1, 1) = "0"
                viaje_trunc = Mid(viaje_trunc, 2)
            Loop
            'OJO CAMBIAR esto antes de generar exe
            'para la agencia se utiliza "mani_emp_prop" y "mani_emp_nave" taylor
            'para la linea se utiliza   "mani_emp_prop" y "mani_emp_prop" msc
            
            cadena = cadena & sep & _
                 del & reg_mani("mani_emp_prop") & del & sep & _
                 del & reg_mani("mani_emp_prop") & del & sep & _
                 del & viaje_trunc & del & sep & _
                 del & Format(REG_viaje("buqv_fecha_arribo"), "yyyymmdd") & del & sep & _
                 del & REG_viaje("buqu_descripcion") & del & sep & _
                 del & REG_viaje("buqu_nacio") & del & sep & _
                 del & reg_mani("mani_capitan") & del & sep & _
                 del & reg_mani("mani_puer_ini") & del & sep & _
                 del & Format(reg_mani("mani_fecha_zarp_ini"), "yyyymmdd") & del & sep & _
                 del & Format(Now, "yyyymmdd") & del & sep & _
                 del & Format(Now, "hh:mm:ss") & del
         End If

         Print #x, cadena
     End If
     Close #x


     '************************* M A N D E T 0 1 ********************************************
     x = FreeFile
     Open Gstr_Ruta_files & "TMP\MANDET01.txt" For Output As #x
     If Opt_tipo_impo.Value Then
        sql = "Select * from docu_blcab "
     Else
        sql = "Select * from docu_blcab_exp "
     End If
     sql = sql & "    Where blcab_linea      ='" & linea & "'" & _
                    "   and blcab_buque    ='" & buque & "'" & _
                    "   and blcab_viaje    ='" & viaje & "'"
     If Chk_envio_pto.Value = 1 Then
               sql = sql & "   and blcab_puer_emb ='" & pto_emb & "'" & _
                           "   and blcab_puer_des ='" & pto_descarga & "'" & _
                           "   and blcab_emp_resp ='" & oper_carga & "'"
     End If
     
     If Chk_todos.Value = 0 And Chk_envio_pto.Value = 1 Then sql = sql & "   and blcab_cono_emb in (" & bls & ")"
     sql = sql & " Order by  blcab_cono_emb"
     
     reg_blcab.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
     No_bls = 0
     Do While Not reg_blcab.EOF
        No_bls = No_bls + 1
        
        cadena = del & reg_mani("mani_adua_ing") & del & sep & _
                 del & reg_mani("mani_ano_car") & del & sep & _
                 del & reg_mani("mani_tip_mani") & del & sep & _
                 del & reg_mani("mani_mani_car") & del & sep & _
                 del & reg_blcab("blcab_puer_des") & del & sep & _
                 del & reg_blcab("blcab_puer_emb") & del & sep & _
                 del & reg_blcab("blcab_emp_anav") & del & sep & _
                 del & reg_blcab("blcab_cono_emb") & del & sep & _
                 del & reg_blcab("blcab_emp_resp") & del & sep & _
                 del & reg_blcab("blcab_tipo_doc") & del & sep & _
                 del & reg_blcab("blcab_cono_master") & del & sep & _
                       reg_blcab("blcab_peso_tot") & sep & _
                       reg_blcab("blcab_bultos_tot") & sep & _
                 del & reg_blcab("blcab_tipo_ide_cons") & del & sep & _
                 del & reg_blcab("blcab_ide_cons") & del & sep & _
                 del & reg_blcab("blcab_consig") & del & sep & _
                 del & reg_blcab("blcab_direc_consig") & del & sep & _
                 del & reg_blcab("blcab_nacio_consig") & del & sep & _
                 del & reg_blcab("blcab_tipo_ide_notif") & del & sep & _
                 del & reg_blcab("blcab_ide_notif") & del & sep & _
                 del & reg_blcab("blcab_notif") & del & sep & _
                 del & reg_blcab("blcab_direc_notif") & del & sep & _
                 del & reg_blcab("blcab_embarcador") & del & sep & _
                 del & reg_blcab("blcab_direc_embarc") & del & sep & _
                 del & reg_blcab("blcab_puer_origen") & del & sep
        cadena = cadena & _
                 del & reg_blcab("blcab_almac_temp") & del & sep & _
                 del & reg_blcab("blcab_tipo_depos") & del & sep & _
                 del & reg_blcab("blcab_deposito_adua") & del & sep & _
                 del & reg_blcab("blcab_regimen_adu") & del & sep & _
                 del & reg_blcab("blcab_consolidad") & del & sep & _
                 del & Format(GTD_date.DateValue, "YYYYMMDD") & del & sep & _
                 del & Format(GTD_hora.DateValue, "hh:mm:ss") & del & sep & _
                 del & reg_blcab("blcab_tipo_salida") & del & sep & _
                 del & reg_blcab("blcab_docum_aut") & del & sep & _
                       reg_blcab("blcab_pesob_mal_est") & sep & _
                       reg_blcab("blcab_bultos_mal_est") & sep & _
                 del & reg_blcab("blcab_notif2") & del & sep & _
                 del & reg_blcab("blcab_direc_notif2") & del & sep & _
                       reg_blcab("blcab_no_cont") & sep & _
                       reg_blcab("blcab_volumen_tot") & sep & _
                 del & reg_blcab("blcab_pto_dest_final") & del & sep & _
                 del & reg_blcab("blcab_metod_pago") & del & sep & _
                 del & reg_blcab("blcab_moneda") & del & sep & _
                       reg_blcab("blcab_flete")
         
         'solo en manifiesto completo o correccion
         ' 19 sept 2002 hasta ahora no se requeria esto mas que para manif. original, se agrego a correctores también
         If Cmb_tipo_transaccion.BoundText = "00" Or Cmb_tipo_transaccion.BoundText = "30" Then
            cadena = cadena & sep & _
                       reg_blcab("blcab_otros_gastos") & sep & _
                 del & reg_blcab("blcab_moneda_otros") & del & sep & _
                       reg_blcab("blcab_seguro") & sep & _
                 del & reg_blcab("blcab_moneda_seguro") & del & sep & _
                 del & reg_blcab("blcab_metod_pago_otros") & del & sep & _
                 del & Format(reg_blcab("blcab_fecha_emb"), "YYYYMMDD") & del
         Else ' segun boletin # 84  del 20 sept/2005 se debe incluir todos los campos aunque esten vacios
            cadena = cadena & sep & _
                       "" & sep & _
                 del & "" & del & sep & _
                       "" & sep & _
                 del & "" & del & sep & _
                 del & "" & del & sep & _
                 del & "" & del
         End If
         ' cambio solicitador por la CAE 15 Agosto 2005 Boletin No 67 se adiciona puerto y fecha de trasbordo y tipo y id de embarcador
         ' solo se debe enviar si es transaccion de envio de manifiesto (tipo envio Original y reemplazo detalle) o en transac corrector o trasacc complemento
         ''esto se cambio por el boletin #84 que indica que se incluyan todos los campo esten o no llenos
         If (Cmb_tipo_transaccion.BoundText = "00" And (Cmb_tipo_envio.BoundText = "09" Or Cmb_tipo_envio.BoundText = "21")) Or _
             Cmb_tipo_transaccion.BoundText = "30" Or Cmb_tipo_transaccion.BoundText = "03" Then
            cadena = cadena & sep & _
                   del & IIf(IsNull(reg_blcab("blcab_puer_trasb")), "", reg_blcab("blcab_puer_trasb")) & del & sep & _
                   del & IIf(IsNull(reg_blcab("blcab_fecha_trasb")) Or reg_blcab("blcab_fecha_trasb") = "1/1/1900" Or reg_blcab("blcab_fecha_trasb") = "01/01/1900", "", Format(reg_blcab("blcab_fecha_trasb"), "YYYYMMDD")) & del & sep & _
                   del & IIf(IsNull(reg_blcab("blcab_tipo_ide_emb")), "", reg_blcab("blcab_tipo_ide_emb")) & del & sep & _
                   del & IIf(IsNull(reg_blcab("blcab_ide_emb")), "", reg_blcab("blcab_ide_emb")) & del
         Else ''esto se cambio por el boletin #84 que indica que se incluyan todos los campo esten o no llenos
            cadena = cadena & sep & _
                   del & "" & del & sep & _
                   del & "" & del & sep & _
                   del & "" & del & sep & _
                   del & "" & del
         End If
        Print #x, cadena
        reg_blcab.MoveNext
     Loop
     Close #x
     
    '************************* M A N D E T 0 2 ********************************************
    ' en manifiesto completo, correcion o sobrantes y faltantes
     If Cmb_tipo_transaccion.BoundText = "00" Or Cmb_tipo_transaccion.BoundText = "30" Or Cmb_tipo_transaccion.BoundText = "06" Then
         x = FreeFile
         Open Gstr_Ruta_files & "TMP\MANDET02.txt" For Output As #x
         If Opt_tipo_impo.Value Then
            sql = "Select docu_bldet.*, docu_blcab.blcab_emp_anav from " & _
                        " docu_bldet,   docu_blcab  "
         Else
            sql = "Select docu_bldet_exp.*, docu_blcab_exp.blcab_emp_anav from " & _
                        " docu_bldet_exp,   docu_blcab_exp "
         End If
            sql = sql & " Where bldet_linea      ='" & linea & "'" & _
                        "   and bldet_buque    ='" & buque & "'" & _
                        "   and bldet_viaje    ='" & viaje & "'"

            If Chk_envio_pto.Value = 1 Then
               sql = sql & "   and bldet_puer_emb ='" & pto_emb & "'" & _
                           "   and bldet_puer_des ='" & pto_descarga & "'" & _
                           "   and bldet_emp_resp ='" & oper_carga & "'"
            End If
            If Chk_todos.Value = 0 And Chk_envio_pto.Value = 1 Then sql = sql & "   and bldet_cono_emb in (" & bls & ")"
            sql = sql & "   and blcab_linea      = bldet_linea" & _
                        "   and blcab_buque    = bldet_buque " & _
                        "   and blcab_viaje    = bldet_viaje " & _
                        "   and blcab_puer_emb = bldet_puer_emb " & _
                        "   and blcab_puer_des = bldet_puer_des " & _
                        "   and blcab_emp_resp = bldet_emp_resp " & _
                        "   and blcab_cono_emb = bldet_cono_emb " & _
                        " Order by  bldet_cono_emb, bldet_secuencia "
                
         reg_bldet.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
         
         Do While Not reg_bldet.EOF
            cadena = del & reg_mani("mani_adua_ing") & del & sep & _
                     del & reg_mani("mani_ano_car") & del & sep & _
                     del & reg_mani("mani_tip_mani") & del & sep & _
                     del & reg_mani("mani_mani_car") & del & sep & _
                     del & reg_bldet("bldet_puer_des") & del & sep & _
                     del & reg_bldet("bldet_puer_emb") & del & sep & _
                     del & reg_bldet("blcab_emp_anav") & del & sep & _
                     del & reg_bldet("bldet_cono_emb") & del & sep & _
                           reg_bldet("bldet_secuencia") & sep & _
                     del & reg_bldet("bldet_tipo_carga") & del & sep & _
                           reg_bldet("bldet_peso") & sep & _
                           reg_bldet("bldet_bultos") & sep & _
                     del & reg_bldet("bldet_embalaje") & del & sep
                     '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                     'si las marcas son mayores a 210 solo toma los primeros
                     '210 sino la toma completa para eso es el if
                     '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                     descrip_marcas = ""
                     varias_marcas = 0
                     If Len(Trim$(reg_bldet("bldet_marcas_num"))) > 210 Then
                        descrip_marcas = Mid(reg_bldet("bldet_marcas_num"), 1, 210)
                        descrip_marcas = chk_str(descrip_marcas)
                        varias_marcas = 1
                        cadena = Trim$(cadena) & del & Trim$(descrip_marcas) & del & sep
                     Else
                        cadena = Trim$(cadena) & del & chk_str(Trim$(reg_bldet("bldet_marcas_num"))) & del & sep
                        varias_marcas = 0
                     End If
                     '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                     'si la descripcion es mayor a 210 solo toma los primeros
                     '210 sino la toma completa para eso es el if
                     '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                     descrip_carga = ""
                     varias_descrip = 0
                     If Len(Trim$(reg_bldet("bldet_desc_carg"))) > 210 Then
                        descrip_carga = Mid(reg_bldet("bldet_desc_carg"), 1, 210)
                        descrip_carga = chk_str(descrip_carga)
                        varias_descrip = 1
                        cadena = Trim$(cadena) & del & Trim$(descrip_carga) & del & sep
                     Else
                        cadena = Trim$(cadena) & del & chk_str(Trim$(reg_bldet("bldet_desc_carg"))) & del & sep
                        varias_descrip = 0
                     End If
                     cadena = Trim$(cadena) & reg_bldet("bldet_volumen") & sep & _
                     del & reg_bldet("bldet_contenedor") & del & sep & _
                     del & reg_bldet("bldet_just_cambio") & del & sep & _
                     del & reg_bldet("bldet_imo") & del & sep & _
                     del & reg_bldet("bldet_pelig_ONU") & del ' campo 20
                     ' 19 septiembre 2002 se aumento por solicitud de la CAE para correctores 30 y para 06 (de importacion unicamente en expo no se requiere aparentemente 19 ene 2003)
                     If (Cmb_tipo_transaccion.BoundText = "30" Or Cmb_tipo_transaccion.BoundText = "06") And Opt_tipo_impo.Value Then
                        If reg_bldet("bldet_contenedor") <> "" Then
                            If Opt_tipo_impo.Value Then
                                sql = "Select docu_cont.* from docu_cont  "
                            Else
                                sql = "Select docu_cont_exp.* from docu_cont_exp  "
                            End If
                            sql = sql & " Where cont_linea     ='" & linea & "'" & _
                                  "   and cont_buque     ='" & buque & "'" & _
                                  "   and cont_viaje     ='" & viaje & "'" & _
                                  "   and cont_puer_emb  ='" & reg_bldet("bldet_puer_emb") & "'" & _
                                  "   and cont_puer_des  ='" & reg_bldet("bldet_puer_des") & "'" & _
                                  "   and cont_contenedor='" & reg_bldet("bldet_contenedor") & "'"
                             reg_cont_dt.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                             If Not reg_cont_dt.EOF Then
                                cadena = cadena & sep & _
                                         del & reg_cont_dt("cont_sello1") & del & sep & _
                                         del & reg_cont_dt("cont_tam_equip") & del & sep & _
                                         del & reg_cont_dt("cont_tipo_equip") & del & sep & _
                                         del & reg_cont_dt("cont_condicion") & del
                             Else
                                cadena = cadena & sep & _
                                         del & "" & del & sep & _
                                         del & "" & del & sep & _
                                         del & "" & del & sep & _
                                         del & "" & del
                             End If
                             reg_cont_dt.Close
                        Else 'contenedor <> ""
                             cadena = cadena & sep & _
                                         del & "" & del & sep & _
                                         del & "" & del & sep & _
                                         del & "" & del & sep & _
                                         del & "" & del
                        End If
                     Else 'tipo transaccion
                        cadena = cadena & sep & _
                                    del & "" & del & sep & _
                                    del & "" & del & sep & _
                                    del & "" & del & sep & _
                                    del & "" & del
                     End If ' 19 sep 2002
                     
                     ' esto es boletin #84
                        cadena = cadena & sep & _
                                    del & "" & del & sep & _
                                    del & "" & del & sep & _
                                    del & "" & del & sep & _
                                    del & "" & del & sep & _
                                    del & "" & del & sep & _
                                    del & "" & del & sep & _
                                    del & "" & del & sep & _
                                    del & "" & del & sep & _
                                    del & "" & del & sep & _
                                    del & "" & del            ' 34
                     
                     
                     ' cambio solicitador por la CAE 15 Agosto 2005 Boletin No 67 se adiciona tipo de producto
                     ' solo se debe enviar si es transaccion de envio de manifiesto (tipo envio Original y reemplazo detalle) o en transac corrector o trasacc complemento
                     ' para Envio de Manifiesto Original y Complemento el campo es el No. 21 para los demas tipos de envio es el No 35(Se incluye inform del contenedor antes).
                     If (Cmb_tipo_transaccion.BoundText = "00" And (Cmb_tipo_envio.BoundText = "09" Or Cmb_tipo_envio.BoundText = "21")) Or _
                        Cmb_tipo_transaccion.BoundText = "30" Or Cmb_tipo_transaccion.BoundText = "03" Then
                            cadena = cadena & sep & _
                                    del & reg_bldet("bldet_tipo_prod") & del
                     Else
                            cadena = cadena & sep & _
                                    del & "" & del
                     End If
                     
                     
            Print #x, cadena
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''TL Esto lo aumente para poner las descripciones en varias lineas
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'no_concatena1 = 0
            'no_concatena2 = 0
            'hasta1 = 0
            'hasta2 = 0
            'hasta = 0
            'If varias_descrip = 1 Or varias_marcas = 1 Then
            '    'se debe poner varios registros solo con la clave primaria hasta
            '    'completar toda la descripcion de la carga que quiero enviar
            '     If varias_marcas = 1 Then
            '        hasta1 = Round((Len(reg_bldet("bldet_marcas_num")) - 210) / 70)
            '        If ((Len(reg_bldet("bldet_marcas_num")) - 210) Mod 70) > 0 Then
            '            hasta1 = hasta1 + 1
            '        End If
            '     End If
            '     If varias_descrip = 1 Then
            '        hasta2 = Round((Len(reg_bldet("bldet_desc_carg")) - 210) / 70)
            '        If ((Len(reg_bldet("bldet_desc_carg")) - 210) Mod 70) > 0 Then
            '            hasta2 = hasta2 + 1
            '        End If
            '     End If
            '     If hasta1 <= hasta2 Then
            '        hasta = hasta2
            '     Else
            '        hasta = hasta1
            '     End If
            '     For desde = 2 To hasta
            '        'reg_bldet("bldet_secuencia") & sep
            '        detalle_adicional = detalle_adicional + 1
            '        cadena = del & reg_mani("mani_adua_ing") & del & sep & _
            '                 del & reg_mani("mani_ano_car") & del & sep & _
            '                 del & reg_mani("mani_tip_mani") & del & sep & _
            '                 del & reg_mani("mani_mani_car") & del & sep & _
            '                 del & reg_bldet("bldet_puer_des") & del & sep & _
            '                 del & reg_bldet("bldet_puer_emb") & del & sep & _
            '                 del & reg_bldet("bldet_emp_resp") & del & sep & _
            '                 del & reg_bldet("bldet_cono_emb") & del & sep & _
            '                       reg_bldet("bldet_secuencia") & sep & _
            '                 del & "" & del & sep & _
            '                        0 & sep & _
            '                        0 & sep & _
            '                 del & "" & del & sep
            '                 'codigo para descripcion de marcas
            '                 If varias_marcas = 1 Then
            '                    descrip_marcas = Mid$(reg_bldet("bldet_marcas_num"), ((210 + (70 * (desde - 2))) + (1)), 70)
            '                    descrip_marcas = chk_str(descrip_marcas)
            '                    cadena = Trim$(cadena) & del & Trim$(descrip_marcas) & del & sep
            '                 Else
            '                    cadena = Trim$(cadena) & del & del & sep
            '                 End If

            '                 'codigo para descripcion de carga
            '                 If varias_descrip = 1 Then
            '                    descrip_carga = Mid$(reg_bldet("bldet_desc_carg"), ((210 + (70 * (desde - 2))) + (1)), 70)
            '                    descrip_carga = chk_str(descrip_carga)
            '                    cadena = Trim$(cadena) & del & Trim$(descrip_carga) & del & sep
            '                 Else
            '                    cadena = Trim$(cadena) & del & del & sep
            '                 End If
            '                 cadena = Trim$(cadena) & 0 & sep & _
            '                 del & "" & del & sep & _
            '                 del & "" & del & sep & _
            '                 del & "" & del & sep & _
            '                 del & "" & del
                    ' 19 septiembre 2002 se aumento por solicitud de la CAE para correctores 30 y para 06
            '         If Cmb_tipo_transaccion.BoundText = "30" Or Cmb_tipo_transaccion.BoundText = "06" Then
            '                   cadena = cadena & sep & _
            '                             del & "" & del & sep & _
            '                             del & "" & del & sep & _
            '                             del & "" & del & sep & _
            '                             del & "" & del
            '         end if
                    ' 19 septiembre 2002
            '        Print #x, cadena
            '     Next desde
            'End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''TL  hasta aqui
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            reg_bldet.MoveNext
         Loop
    
         Close #x
     End If

     '************************* M A N C O N 0 1 ********************************************
     If Cmb_tipo_transaccion.BoundText = "00" Or (Cmb_tipo_transaccion.BoundText = "30" And Opt_tipo_expo.Value) Then
         x = FreeFile
         Open Gstr_Ruta_files & "TMP\MANCON01.txt" For Output As #x
         If Opt_tipo_impo.Value Then
            sql = "Select docu_cont.* from docu_bldet,docu_cont  "
         Else
            sql = "Select docu_cont_exp.* from docu_bldet_exp, docu_cont_exp "
         End If
         sql = sql & " Where bldet_linea   ='" & linea & "'" & _
                        "   and bldet_buque    ='" & buque & "'" & _
                        "   and bldet_viaje    ='" & viaje & "'"
         If Chk_envio_pto.Value = 1 Then
            sql = sql & "   and bldet_puer_emb ='" & pto_emb & "'" & _
                        "   and bldet_puer_des ='" & pto_descarga & "'"
         End If
         If Chk_todos.Value = 0 And Chk_envio_pto.Value = 1 Then sql = sql & "   and bldet_cono_emb in (" & bls & ") "
         sql = sql & "   and cont_linea     = bldet_linea " & _
                     "   and cont_buque     = bldet_buque " & _
                     "   and cont_viaje     = bldet_viaje " & _
                     "   and cont_puer_emb  = bldet_puer_emb " & _
                     "   and cont_puer_des  = bldet_puer_des " & _
                     "   and cont_contenedor= bldet_contenedor " & _
                     " Order by  cont_contenedor "
         reg_cont.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    
         Do While Not reg_cont.EOF
            cadena = del & reg_mani("mani_adua_ing") & del & sep & _
                     del & reg_mani("mani_ano_car") & del & sep & _
                     del & reg_mani("mani_tip_mani") & del & sep & _
                     del & reg_mani("mani_mani_car") & del & sep
            If Opt_tipo_impo.Value Then
                cadena = cadena & _
                     del & reg_cont("cont_puer_des") & del & sep
            Else
                cadena = cadena & _
                     del & reg_cont("cont_puer_emb") & del & sep
            End If
            cadena = cadena & _
                     del & reg_cont("cont_contenedor") & del & sep & _
                     del & reg_cont("cont_tam_equip") & del & sep & _
                     del & reg_cont("cont_condicion") & del & sep & _
                           reg_cont("cont_tara") & sep & _
                     del & reg_cont("cont_tipo_equip") & del & sep & _
                     del & reg_cont("cont_proved_equip") & del & sep & _
                     del & reg_cont("cont_estad_equip") & del & sep & _
                     del & reg_cont("cont_indic_llen_vac") & del & sep & _
                     del & reg_cont("cont_cond_contrato") & del & sep & _
                     del & reg_cont("cont_reg_serv_transp") & del & sep & _
                     del & reg_cont("cont_natu_carg") & del & sep & _
                           reg_cont("cont_peso") & sep & _
                     del & reg_cont("cont_tipo_bien") & del & sep & _
                     del & reg_cont("cont_sello1") & del & sep & _
                     del & reg_cont("cont_sello2") & del & sep & _
                     del & reg_cont("cont_sello3") & del & sep & _
                     del & reg_cont("cont_sello4") & del & sep & _
                           reg_cont("cont_temp_dur_almac") & sep
             cadena = cadena & reg_cont("cont_temp_dur_trans") & sep & _
                           reg_cont("cont_temp_dur_manip") & sep & _
                           reg_cont("cont_temp_min_almac") & sep & _
                           reg_cont("cont_temp_max_almac") & sep & _
                           reg_cont("cont_temp_min_trans") & sep & _
                           reg_cont("cont_temp_max_trans") & sep & _
                           reg_cont("cont_temp_min_manip") & sep & _
                           reg_cont("cont_temp_max_manip")
            Print #x, cadena
            reg_cont.MoveNext
         Loop
    
         Close #x
     End If
     
     reg_blcab.MoveFirst
     x = FreeFile
     Open Gstr_Ruta_files & "TMP\ENVCTROL.txt" For Output As #x
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     'TL aumento para que de acuerdo al tipo de transaccion se envie '
     '   el codigo de la linea o de la agencia                       '
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Cmb_tipo_transaccion.BoundText = "00" Then ' solo si es envio original
        'siempre envio original de linea
        cadena = del & Gstr_param.tipo_oper & del & sep & _
                 del & reg_blcab("blcab_emp_anav") & del & sep
     Else
        'siempre envio adicional naviera
        cadena = del & "07" & del & sep & _
                 del & reg_blcab("blcab_emp_resp") & del & sep
     End If
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     cadena = Trim$(cadena) & 1 & sep & _
                    "0" & sep & _
                    reg_bls("no_bls") & sep & _
                    "0" & sep
     cadena = Trim$(cadena) & reg_bls("peso") & sep
     reg_bls.Close
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     'TL aumento para que de acuerdo al tipo de transaccion se envie '
     '   la clave de la linea o de la agencia                        '
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     ' se aumento un select a la tabla de parametros para buscar la clave
     ' de la Linea en caso de existir mas de 1 linea anteriormente se podia tener máximo 2 lineas
     ' con este select se puede tener n lineas, creando nuevas entrada de parámetros
     '****************************************************************
     If Cmb_tipo_transaccion.BoundText = "00" Then
        reg_clave.Open "Select * from docu_parametro where para_codigo ='CLAVE_" & reg_blcab("blcab_emp_anav") & "' and para_linea ='" & reg_blcab("blcab_linea") & "'", Gcon_docu
        If Not reg_clave.EOF Then
            cadena = Trim$(cadena) & del & reg_clave("para_valor") & del & sep
            'If reg_blcab("blcab_emp_anav") = "1186" Then 'CLAVE DE KHS
            '    cadena = Trim$(cadena) & del & Gstr_param.clave_KHS & del & sep
        Else                                         'CLAVE DE MSC
            cadena = Trim$(cadena) & del & Gstr_param.clave & del & sep
        End If
        reg_clave.Close
     Else
        cadena = Trim$(cadena) & del & Gstr_param.clave_TAYLOR & del & sep
     End If
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     cadena = Trim$(cadena) & "0" & sep & _
              del & "" & del & sep & _
              del & "" & del
     Print #x, cadena
     Close #x
     
     'CUSCAR se usa para envio original y para correctores de Exportacion lo demas se va por CARDAT (19 ene 2003)
     If Cmb_tipo_transaccion.BoundText = "00" Or (Cmb_tipo_transaccion.BoundText = "30" And Opt_tipo_expo.Value) Then
        If Dir(ruta_archivo & "\CUSCAR.ZIP", vbArchive) <> "" Then Kill (ruta_archivo & "\CUSCAR.ZIP")
        Shell (Gstr_Ruta_files & "pkzip " & ruta_archivo & "\CUSCAR.ZIP " & Gstr_Ruta_files & "TMP\*.*")
        MsgBox "Generación de archivo de " & Cmb_tipo_transaccion.Text & " fue realizada con éxito. En el archivo " & ruta_archivo & "\CUSCAR.ZIP", vbInformation
        If Chk_mail.Value = 1 Then Call envia_mail(ruta_archivo & "\CUSCAR.ZIP")
     Else
        If Dir(ruta_archivo & "\CARDAT.ZIP", vbArchive) <> "" Then Kill (ruta_archivo & "\CARDAT.ZIP")
        Shell (Gstr_Ruta_files & "pkzip " & ruta_archivo & "\CARDAT.ZIP " & Gstr_Ruta_files & "TMP\*.*")
        MsgBox "Generación de archivo de " & Cmb_tipo_transaccion.Text & " fue realizada con éxito. En el archivo " & ruta_archivo & "\CARDAT.ZIP", vbInformation
        If Chk_mail.Value = 1 Then Call envia_mail(ruta_archivo & "\CARDAT.ZIP")
     End If

     Exit Sub
errores:
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
    On Error GoTo errores
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
errores:
    MsgBox "Ocurrio el error " & Err.Description
End Sub

Private Sub Opt_tipo_expo_Click()
If Opt_tipo_expo.Value Then
    Lbl_pto_emb.Caption = "Puerto Descarga"
    Lbl_pto_desc.Caption = "Puerto Embarque"
End If
End Sub

Private Sub Opt_tipo_impo_Click()
If Opt_tipo_impo.Value Then
    Lbl_pto_emb.Caption = "Puerto Embarque"
    Lbl_pto_desc.Caption = "Puerto Descarga"
End If
End Sub
