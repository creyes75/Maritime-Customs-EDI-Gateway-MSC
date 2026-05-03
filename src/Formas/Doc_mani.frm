VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8850
   LinkTopic       =   "Form2"
   ScaleHeight     =   3075
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8355
      Begin VB.Frame Fra_depot_port 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   150
         TabIndex        =   32
         Top             =   870
         Width           =   1590
         Begin VB.TextBox Txt_depot_port 
            Height          =   285
            Left            =   120
            MaxLength       =   5
            TabIndex        =   34
            Top             =   270
            Width           =   975
         End
         Begin VB.CommandButton Cmd_depot_port 
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
            Picture         =   "Doc_mani.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Buscar"
            Top             =   255
            Width           =   300
         End
         Begin VB.Label Label1 
            Caption         =   "Puerto Descarga"
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   120
            TabIndex        =   35
            Top             =   15
            Width           =   1260
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   1710
         TabIndex        =   28
         Top             =   900
         Width           =   1590
         Begin VB.CommandButton Command1 
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
            Picture         =   "Doc_mani.frx":0532
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Buscar"
            Top             =   255
            Width           =   300
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            MaxLength       =   5
            TabIndex        =   29
            Top             =   270
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Puerto Embarque"
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   120
            TabIndex        =   31
            Top             =   15
            Width           =   1260
         End
      End
      Begin VB.Frame Fra_respons 
         BorderStyle     =   0  'None
         Height          =   510
         Left            =   3390
         TabIndex        =   25
         Top             =   255
         Width           =   1095
         Begin VB.TextBox Txt_respons 
            Height          =   285
            Left            =   105
            MaxLength       =   80
            TabIndex        =   26
            Top             =   195
            Width           =   765
         End
         Begin VB.Label Label31 
            Caption         =   "Año"
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   90
            TabIndex        =   27
            Top             =   -15
            Width           =   1530
         End
      End
      Begin VB.Frame Fra_voyage 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   75
         TabIndex        =   22
         Top             =   195
         Width           =   1395
         Begin MSDataListLib.DataCombo Cmb_voyage1 
            Height          =   315
            Left            =   135
            TabIndex        =   23
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_voyage1 
            Height          =   330
            Left            =   480
            Top             =   120
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
            Left            =   90
            TabIndex        =   24
            Top             =   0
            Width           =   825
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   1635
         TabIndex        =   19
         Top             =   195
         Width           =   1500
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   315
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   495
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
         Begin VB.Label Label3 
            Caption         =   "Aduana Ing"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   21
            Top             =   0
            Width           =   825
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   4500
         TabIndex        =   16
         Top             =   240
         Width           =   1500
         Begin MSDataListLib.DataCombo DataCombo2 
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   495
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Manifiesto"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   90
            TabIndex        =   18
            Top             =   0
            Width           =   1080
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   510
         Left            =   6150
         TabIndex        =   13
         Top             =   255
         Width           =   1320
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   105
            MaxLength       =   80
            TabIndex        =   14
            Top             =   195
            Width           =   945
         End
         Begin VB.Label Label6 
            Caption         =   "Manifiesto"
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   90
            TabIndex        =   15
            Top             =   -15
            Width           =   1530
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   3315
         TabIndex        =   10
         Top             =   870
         Width           =   1500
         Begin MSDataListLib.DataCombo DataCombo3 
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc Adodc3 
            Height          =   330
            Left            =   495
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
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Empresa Resp."
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   90
            TabIndex        =   12
            Top             =   0
            Width           =   1080
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   510
         Left            =   5220
         TabIndex        =   7
         Top             =   930
         Width           =   2160
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   105
            MaxLength       =   80
            TabIndex        =   8
            Top             =   195
            Width           =   1740
         End
         Begin VB.Label Label8 
            Caption         =   "Bill of lading"
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   90
            TabIndex        =   9
            Top             =   -15
            Width           =   1530
         End
      End
      Begin VB.Frame Fra_vissel_name 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   135
         TabIndex        =   4
         Top             =   1440
         Width           =   2925
         Begin MSDataListLib.DataCombo Cmb_vessel_name 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_vessel 
            Height          =   330
            Left            =   1440
            Top             =   240
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
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
            Caption         =   "Vessel name"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   6
            Top             =   0
            Width           =   1035
         End
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   3210
         TabIndex        =   1
         Top             =   1440
         Width           =   1395
         Begin MSDataListLib.DataCombo Cmb_voyage 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_voyage 
            Height          =   330
            Left            =   480
            Top             =   120
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
            Caption         =   "Voyage"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   3
            Top             =   0
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
