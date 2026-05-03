VERSION 5.00
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "GTDATE32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Cont_vessel_voyage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Veseel_Voyage"
   ClientHeight    =   3345
   ClientLeft      =   3330
   ClientTop       =   4110
   ClientWidth     =   6510
   Icon            =   "Cont_vessel_voyage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3345
   ScaleWidth      =   6510
   Begin TabDlg.SSTab Tab_dato 
      Height          =   2775
      Left            =   75
      TabIndex        =   4
      Top             =   60
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Listado"
      TabPicture(0)   =   "Cont_vessel_voyage.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "RDC_grid_dato"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Grid_dato"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalles"
      TabPicture(1)   =   "Cont_vessel_voyage.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_tipo_rol"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "RDC_dato"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin MSDataGridLib.DataGrid Grid_dato 
         Height          =   2175
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3836
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
         Top             =   2520
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
         Top             =   2280
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
         Height          =   1920
         Left            =   -74700
         TabIndex        =   5
         Top             =   345
         Width           =   5805
         _Version        =   65536
         _ExtentX        =   10239
         _ExtentY        =   3387
         _StockProps     =   14
         Caption         =   " Datos de Vessel / Voyage"
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
            Caption         =   " Direction "
            ForeColor       =   &H00800000&
            Height          =   540
            Left            =   3390
            TabIndex        =   14
            Top             =   825
            Width           =   1725
            Begin VB.OptionButton Opt_R 
               Caption         =   "R"
               Height          =   210
               Left            =   915
               TabIndex        =   16
               Top             =   255
               Width           =   555
            End
            Begin VB.OptionButton Opt_A 
               Caption         =   "A"
               Height          =   210
               Left            =   135
               TabIndex        =   15
               Top             =   255
               Width           =   555
            End
         End
         Begin VB.TextBox Txt_voyage 
            Height          =   285
            Left            =   3390
            MaxLength       =   5
            TabIndex        =   10
            Top             =   495
            Width           =   1230
         End
         Begin VB.TextBox txt_desc 
            Height          =   285
            Left            =   1650
            MaxLength       =   60
            TabIndex        =   7
            Top             =   1500
            Width           =   4005
         End
         Begin VB.TextBox txt_vessel 
            Height          =   285
            Left            =   1380
            MaxLength       =   5
            TabIndex        =   6
            Top             =   195
            Visible         =   0   'False
            Width           =   840
         End
         Begin GTMaskDate.GTMaskDate GTD_date 
            Height          =   300
            Left            =   405
            TabIndex        =   12
            Top             =   1065
            Width           =   1545
            _Version        =   65537
            _ExtentX        =   2725
            _ExtentY        =   529
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
            Left            =   480
            TabIndex        =   18
            Top             =   480
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_vessel 
            Height          =   330
            Left            =   1920
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sailing/Arrival Date :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   435
            TabIndex        =   13
            Top             =   855
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Voyage :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   3435
            TabIndex        =   11
            Top             =   270
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comments :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   405
            TabIndex        =   9
            Top             =   1530
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vessel :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   465
            TabIndex        =   8
            Top             =   270
            Width           =   555
         End
      End
   End
   Begin Threed.SSCommand cmd_agregar 
      Height          =   375
      Left            =   1635
      TabIndex        =   0
      Top             =   2910
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
      Top             =   2910
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
      Top             =   2910
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
      Top             =   2910
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
      TabIndex        =   17
      Top             =   2910
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
   Begin Crystal.CrystalReport CReport 
      Left            =   90
      Top             =   2940
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "Cont_vessel_voyage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lstr_opc As String * 1

Public Sub habilita(ByVal opc As Boolean)
   If opc Then
      If Lstr_opc = "M" Then
         txt_vessel.Enabled = False
         cmb_vessel.Enabled = False
         Txt_voyage.Enabled = False
         'txt_desc.Enabled = False
      Else
         txt_vessel.Enabled = opc
         cmb_vessel.Enabled = opc
         Txt_voyage.Enabled = opc
         GTD_date.Enabled = opc
         txt_desc.Enabled = opc
      End If
   Else
      txt_vessel.Enabled = True
      cmb_vessel.Enabled = True
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
   'Txt_part.Text = ""
   'Txt_dias_vaca.Text = ""
   'Chk_apli.Value = 0
End Sub
Private Function Valida_Datos() As Boolean
On Error GoTo ERROR_VALIDA
Dim REG As New ADODB.Recordset

Valida_Datos = True
   'If Trim$(txt_vessel.Text) = "" Then
   If cmb_vessel.BoundText = "" Then
      MsgBox "Ingrese el Valor de Vessel", vbInformation, App.Title
      Valida_Datos = False
      'If txt_vessel.Enabled Then txt_vessel.SetFocus
      If cmb_vessel.Enabled Then cmb_vessel.SetFocus
      Exit Function
   End If
   If Trim$(Txt_voyage.Text) = "" Then
      MsgBox "Ingrese el Código de Voyage", vbInformation, App.Title
      Valida_Datos = False
      If Txt_voyage.Enabled Then Txt_voyage.SetFocus
      Exit Function
   End If
   If Trim$(GTD_date.Text) = "" Then
      MsgBox "Ingrese el Fecha ", vbInformation, App.Title
      Valida_Datos = False
      If GTD_date.Enabled Then GTD_date.SetFocus
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
   Dim REG          As New ADODB.Recordset
   Dim reg_max      As New ADODB.Recordset
   Dim maximo       As Integer
   
   On Error GoTo ERROR_GRABACION
   
   Gcon_Cont.BeginTrans
   If Opt_A.Value Then Direction = "A" Else Direction = "R"
   
   If Lstr_opc = "A" Then
     sql = "Select * from cont_Vessel_Voyage Where vess_vessel = '" & cmb_vessel.BoundText & "' and vess_voyage = '" & Txt_voyage.Text & "'"
     REG.Open sql, Gcon_Cont, adOpenDynamic, adLockOptimistic
     If REG.EOF Then
        sql = "INSERT INTO cont_Vessel_Voyage (vess_Vessel, vess_Voyage, vess_date_arraive, vess_direction, vess_comment, vess_status, vess_fech_ing, vess_user_ing)" & _
             " VALUES ('" & Trim$(cmb_vessel.BoundText) & "','" & Trim$(Txt_voyage.Text) & "','" & Format(GTD_date.Text, "dd/mm/yyyy") & "','" & Direction & "','" & Trim$(txt_desc.Text) & "', 'A', '" & Format$(Date, "dd/mm/yyyy") & "','" & Gstr_user & "')"
     Else
        MsgBox "Este Voyage Header ya existe, entre por la opción modificación", vbExclamation, App.Title
        Exit Sub
     End If
   ElseIf Lstr_opc = "M" Then
      sql = "UPDATE cont_Vessel_Voyage SET vess_date_arraive = '" & GTD_date.Text & "', " & _
                                          "vess_direction    = '" & Direction & "'," & _
                                          "vess_comment      = '" & Trim$(txt_desc.Text) & "'," & _
                                          "vess_status       = 'A' " & _
                                    "WHERE vess_vessel  = '" & Trim$(cmb_vessel.BoundText) & "'" & _
                                    "  and vess_voyage  = '" & Trim$(Txt_voyage.Text) & "'"
   ElseIf Lstr_opc = "E" Then
'      sql = "Select trcg_cargo from rhutrol_cargo where trcg_tiporol='" & txt_codigo.Text & "' and trcg_st_tupla = 'A'"
'      Set reg = Gcon_Cont.OpenRecordset(sql, rdOpenKeyset)
'      If Not reg.EOF Then
'         MsgBox "Este Tipo de Rol tiene Cargos asociados, NO es posible Eliminarlo", vbInformation, App.Title
'         Gcon_Cont.RollbackTrans
'         Exit Sub
'      End If
        sql = "UPDATE cont_Vessel_Voyage SET " & _
                                            "vess_status       = 'E' " & _
                                      "WHERE vess_vessel  = '" & Trim$(cmb_vessel.BoundText) & "'" & _
                                      "  and vess_voyage  = '" & Trim$(Txt_voyage.Text) & "'"
                                                 
   End If
   If Lstr_opc = "E" Then
      resp = MsgBox("Esta seguro que desea eliminar este Vessel/Voyyage", vbQuestion + vbYesNo, App.Title)
   Else
      resp = vbYes
   End If
   If resp = vbYes Then
      Gcon_Cont.Execute sql
      MsgBox "Transacción Realizada con éxito", vbInformation, App.Title
   Else
      Exit Sub
      'MsgBox "Transacción cancelada "
   End If
   limpia
   If Lstr_opc <> "A" Then
     habilita False
   End If
   Gcon_Cont.CommitTrans
   Exit Sub
   
ERROR_GRABACION:
   MsgBox "Ocurrio el Error " & Err.Description, vbInformation, App.Title
   Gcon_Cont.RollbackTrans
   Exit Sub
   Resume 0
End Sub


Private Sub cmd_agregar_Click()
Select Case cmd_agregar.Caption
  Case "&Agregar"
      Lstr_opc = "A"
      'cmd_agregar.ToolTipText = "Grabar"
      cmd_agregar.Caption = "&Grabar"
      cmd_salir.Caption = "&Cancelar"
      cmd_modificar.Enabled = False
      cmd_eliminar.Enabled = False
      RDC_dato.Caption = ""
      RDC_dato.Enabled = False
      habilita True
      limpia
      cmb_vessel.SetFocus
      'cmb_vessel.lis = "" 'Nuevo_Tipo_Rol
      Txt_voyage.Text = ""
      GTD_date.Text = ""
      txt_desc.Text = ""
      Tab_dato.Tab = 1
      Tab_dato.TabEnabled(0) = False
  Case "&Grabar"
      If Not Valida_Datos Then Exit Sub
      Graba_trol
      If Lstr_opc = "M" Or Lstr_opc = "E" Then
         Lstr_opc = "C"
         cmd_salir_Click
      Else
         txt_vessel.Text = "" 'Nuevo_Tipo_Rol
      End If
      'If txt_vessel.Enabled Then txt_vessel.SetFocus
      If cmb_vessel.Enabled Then cmb_vessel.SetFocus
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

Private Sub cmd_eliminar_Click()
    Lstr_opc = "E"
    cmd_agregar.Caption = "&Grabar"
    cmd_salir.Caption = "&Cancelar"
    cmd_eliminar.Enabled = False
    cmd_modificar.Enabled = False
    RDC_dato.Enabled = False
    habilita True
    cmb_vessel.Enabled = False
    Txt_voyage.Enabled = False
    Tab_dato.Tab = 1
    Tab_dato.TabEnabled(0) = False
End Sub

Private Sub cmd_imprimir_Click()
On Error GoTo errores
    CReport.ReportFileName = Gstr_ruta_reportes & "cont030.rpt"
    CReport.Connect = Gstr_conex_cont
    CReport.SelectionFormula = "{cont_vessel_voyage.vess_status}='A'"
    CReport.WindowTitle = App.Title
    CReport.Action = 1
    Exit Sub
errores:
    MsgBox "Ocurrio el error " & Err.Description, vbInformation, App.Title

End Sub

Private Sub cmd_modificar_Click()
    Lstr_opc = "M"
    cmd_agregar.Caption = "&Grabar"
    cmd_salir.Caption = "&Cancelar"
    cmd_modificar.Enabled = False
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
       cmd_modificar.Enabled = False
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
   Me.Height = 3750
   Me.Width = 6630
   Centra_Pantalla Me
   
   'RDC_dato.ConnectionString = Gstr_conex_cont
   
   RDC_grid_dato.ConnectionString = Gstr_conex_cont
   RDC_grid_dato.RecordSource = "Select vess_Vessel as Vessel, vess_Voyage as Voyage, vess_direction as Direction, vess_date_arraive as SailArrivDate, vess_comment as Comments from cont_Vessel_Voyage Where vess_status = 'A'"
   RDC_grid_dato.Refresh
   Set Grid_dato.DataSource = RDC_grid_dato
   Grid_dato.Refresh
   
   
   data_vessel.ConnectionString = Gstr_conex_cont
   data_vessel.RecordSource = "Select * from cont_vessel where vessel_status ='A'"
   data_vessel.Refresh
   cmb_vessel.BoundColumn = "vessel_codigo"
   cmb_vessel.ListField = "vessel_descripcion"
   Set cmb_vessel.RowSource = data_vessel
End Sub

Private Sub Grid_dato_DblClick()
   If Not RDC_grid_dato.Recordset.EOF Then
      cmb_vessel.BoundText = Trim$(RDC_grid_dato.Recordset.Fields("vessel"))
      txt_vessel.Text = Trim$(RDC_grid_dato.Recordset.Fields("Vessel"))
      Txt_voyage.Text = Trim$(RDC_grid_dato.Recordset.Fields("Voyage"))
      txt_desc.Text = Trim$(RDC_grid_dato.Recordset.Fields("Comments"))
      If Trim$(RDC_grid_dato.Recordset.Fields("Direction")) = "A" Then
         Opt_A.Value = True
      ElseIf Trim$(RDC_grid_dato.Recordset.Fields("Direction")) = "R" Then
         Opt_R.Value = True
      End If
      GTD_date.Text = Trim$(RDC_grid_dato.Recordset.Fields("SailArrivDate"))
      cmd_modificar.Enabled = True
      cmd_eliminar.Enabled = True
      'Call txt_codigo_LostFocus
      Tab_dato.Tab = 1
   End If
End Sub

Private Sub Grid_dato_HeadClick(ByVal ColIndex As Integer)
Me.MousePointer = 11
If ColIndex = 0 Then
   RDC_grid_dato.RecordSource = "Select vess_Vessel as Vessel, vess_Voyage as Voyage, vess_direction as Direction, vess_date_arraive as SailArrivDate, vess_comment as Comments from cont_Vessel_Voyage Where vess_status = 'A' ORDER BY vess_Vessel"
ElseIf ColIndex = 1 Then
   RDC_grid_dato.RecordSource = "Select vess_Vessel as Vessel, vess_Voyage as Voyage, vess_direction as Direction, vess_date_arraive as SailArrivDate, vess_comment as Comments from cont_Vessel_Voyage Where vess_status = 'A' ORDER BY vess_Voyage"
ElseIf ColIndex = 2 Then
   RDC_grid_dato.RecordSource = "Select vess_Vessel as Vessel, vess_Voyage as Voyage, vess_direction as Direction, vess_date_arraive as SailArrivDate, vess_comment as Comments from cont_Vessel_Voyage Where vess_status = 'A' ORDER BY vess_direction"
ElseIf ColIndex = 3 Then
   RDC_grid_dato.RecordSource = "Select vess_Vessel as Vessel, vess_Voyage as Voyage, vess_direction as Direction, vess_date_arraive as SailArrivDate, vess_comment as Comments from cont_Vessel_Voyage Where vess_status = 'A' ORDER BY vess_date_arraive"
End If
RDC_grid_dato.Refresh
Grid_dato.Refresh
Me.MousePointer = 0
End Sub

Private Sub rdc_dato_Reposition()
If Not RDC_dato.Recordset.EOF Then
   cmb_vessel.BoundText = Trim$(RDC_dato.Recordset.Fields("vess_vessel"))
   txt_vessel.Text = Trim$(RDC_dato.Recordset.Fields("vess_vessel"))
   Txt_voyage.Text = Trim$(RDC_dato.Recordset.Fields("vess_voyage"))
   txt_desc.Text = Trim$(RDC_dato.Recordset.Fields("vess_comment"))
   If Trim$(RDC_dato.Recordset.Fields("vess_direction")) = "A" Then
      Opt_A.Value = True
   ElseIf Trim$(RDC_dato.Recordset.Fields("vess_direction")) = "R" Then
      Opt_R.Value = True
   End If
   GTD_date.Text = Trim$(RDC_dato.Recordset.Fields("vess_date_arraive"))
'   Txt_part.Text = Trim$(RDC_dato.Recordset.Fields("trol_partida"))
'   Txt_dias_vaca.Text = Trim$(RDC_dato.Recordset.Fields("trol_dias_vac"))

'   If RDC_dato.Recordset.Fields("trol_st_aplica") = "S" Then
'      Chk_apli.Value = 1
'   ElseIf RDC_dato.Recordset.Fields("trol_st_aplica") = "N" Then
'      Chk_apli.Value = 0
'   End If
   RDC_dato.Caption = Trim(Str(RDC_dato.Recordset.AbsolutePosition)) & " De " & Str(RDC_dato.Recordset.RecordCount) & " Registros Afectados"
End If
End Sub

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txt_codigo_LostFocus()
Dim REG     As New ADODB.Recordset
Dim sql     As String
Dim i       As Integer
If Trim$(txt_codigo.Text) = "" Then
   limpia
   cmd_modificar.Enabled = False
   cmd_eliminar.Enabled = False
   RDC_dato.Enabled = False
   Exit Sub
ElseIf Lstr_opc = "A" Or Lstr_opc = "M" Or Lstr_opc = "E" Then
   cmd_modificar.Enabled = False
   cmd_eliminar.Enabled = False
   RDC_dato.Enabled = False
   Exit Sub
End If

sql = "Select * FROM cont_Ownership WHERE Owner_codigo like '" & txt_codigo.Text & "'" & _
      " and Owner_status = 'A'"
REG.Open sql, Gcon_Cont, adOpenDynamic, adLockOptimistic
'Set REG = Gcon_Cont.OpenRecordset(sql)
If Not REG.EOF Then
   RDC_dato.RecordSource = sql
   RDC_dato.Refresh
   rdc_dato_Reposition
   
   cmd_modificar.Enabled = True
   cmd_eliminar.Enabled = True
   RDC_dato.Enabled = True
Else
   limpia
   RDC_dato.Caption = "0 Registros Afectados"
   If txt_codigo.Enabled Then txt_codigo.SetFocus
   RDC_dato.Enabled = False
   cmd_modificar.Enabled = False
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
'Set REG = Gcon_Cont.OpenRecordset(sql)
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

