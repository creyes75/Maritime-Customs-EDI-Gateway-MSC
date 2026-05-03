VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form doc_repo_correc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Corección de Bls de importación"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   5310
   Begin VB.CommandButton Cmd_plantilla 
      Height          =   615
      Left            =   120
      Picture         =   "doc_repo_corecc.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Abre Documento Plantilla"
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton Cmd_procesar 
      Height          =   615
      Left            =   3840
      Picture         =   "doc_repo_corecc.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton Cmd_cancel 
      Height          =   615
      Left            =   4500
      Picture         =   "doc_repo_corecc.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   5145
      Left            =   75
      TabIndex        =   13
      Top             =   75
      Width           =   5160
      Begin VB.CommandButton Cmd_cons_bl 
         Caption         =   "&Consulta Bls"
         Height          =   675
         Left            =   4080
         Picture         =   "doc_repo_corecc.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2760
         Width           =   990
      End
      Begin VB.Frame Fra_pto_descarga 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   75
         TabIndex        =   28
         Top             =   2040
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
            Left            =   1125
            Picture         =   "doc_repo_corecc.frx":2618
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Buscar"
            Top             =   255
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.TextBox Txt_pto_descarga 
            Height          =   285
            Left            =   105
            MaxLength       =   5
            TabIndex        =   9
            Top             =   255
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Puerto Descarga"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   15
            Visible         =   0   'False
            Width           =   1200
         End
      End
      Begin VB.TextBox txt_bl 
         Height          =   315
         Left            =   1980
         MaxLength       =   25
         TabIndex        =   4
         Top             =   1665
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   45
         TabIndex        =   25
         Top             =   1455
         Width           =   1500
         Begin MSDataListLib.DataCombo Cmb_pto_emb 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
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
            TabIndex        =   26
            Top             =   0
            Width           =   1230
         End
      End
      Begin VB.Frame Fra_viaje 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   3090
         TabIndex        =   23
         Top             =   795
         Width           =   1965
         Begin MSDataListLib.DataCombo Cmb_viaje 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_viaje 
            Height          =   330
            Left            =   705
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
            TabIndex        =   24
            Top             =   0
            Width           =   345
         End
      End
      Begin VB.Frame Fra_buque 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   75
         TabIndex        =   21
         Top             =   795
         Width           =   2970
         Begin MSDataListLib.DataCombo Cmb_buque 
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_buque 
            Height          =   330
            Left            =   1845
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
            TabIndex        =   22
            Top             =   0
            Width           =   465
         End
      End
      Begin VB.Frame Fra_linea 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   75
         TabIndex        =   19
         Top             =   195
         Width           =   2820
         Begin MSDataListLib.DataCombo Cmb_linea 
            Height          =   315
            Left            =   165
            TabIndex        =   0
            Top             =   255
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_linea 
            Height          =   330
            Left            =   1770
            Top             =   255
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
            Caption         =   "Linea"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   105
            TabIndex        =   20
            Top             =   0
            Width           =   825
         End
      End
      Begin VB.TextBox Txt_fecha 
         Height          =   285
         Left            =   330
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2955
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox Chk_todos 
         Caption         =   "Seleccionar Todos"
         Height          =   255
         Left            =   1140
         TabIndex        =   11
         Top             =   4815
         Width           =   2895
      End
      Begin VB.ListBox Lst_bl 
         Height          =   2085
         Left            =   1125
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   2670
         Width           =   2895
      End
      Begin VB.ComboBox Cmb_formato 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   5175
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo Cmb_oper_carga 
         Height          =   315
         Left            =   1965
         TabIndex        =   10
         Top             =   2205
         Visible         =   0   'False
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
         Left            =   3930
         Top             =   2250
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
         Caption         =   "Ag. Oper. de Carga (Línea)"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1995
         TabIndex        =   31
         Top             =   2010
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.Label Label3 
         Caption         =   "Bill of Lading"
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   1995
         TabIndex        =   27
         Top             =   1455
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Bill of Lading"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   165
         TabIndex        =   17
         Top             =   2685
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Formato"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   5235
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.OptionButton Opt_tarja 
      Caption         =   "Tarja de Importación"
      Height          =   255
      Left            =   2145
      TabIndex        =   15
      Top             =   105
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.OptionButton Opt_bl 
      Caption         =   "Impresión de Bls"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   75
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   360
      Top             =   6015
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "doc_repo_correc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oWord    As Object
Public oCell    As Object
Public oMerge   As Object
Public doc_origen       As String
Public doc_plantilla    As String
' VARIABLES PARA EXCEL
Public wb_est  As Object 'Excel.Workbook
Public sht_est As Excel.Worksheet
Public rng_est As Excel.Range

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

Private Sub cmb_linea_Change()
carga_datacombo data_buque, Cmb_buque, "Select * from docu_buque Where buqu_linea = '" & Cmb_linea.BoundText & "' and buqu_status ='A'", "buqu_codigo", "buqu_descripcion"
carga_datacombo data_oper_carga, Cmb_oper_carga, "Select * from docu_emp_resp Where emp_resp_linea = '" & Cmb_linea.BoundText & "' and emp_resp_status ='A'", "emp_resp_codigo", "emp_resp_descripcion"
carga_parametros Cmb_linea.BoundText

Txt_pto_descarga.Text = Gstr_param.pto_descarga
Cmb_oper_carga.BoundText = Gstr_param.naviera

End Sub

Private Sub Cmb_linea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_pto_emb_Change()
Dim reg As New ADODB.Recordset

If Opt_bl.Value Then
    sql = "Select blcab_cono_emb " & _
               " from docu_blcab " & _
               " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'" & _
               "   and blcab_puer_emb    ='" & Cmb_pto_emb.BoundText & "'"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Lst_bl.Clear
    Do While Not reg.EOF
        Lst_bl.AddItem reg("blcab_cono_emb")
        reg.MoveNext
    Loop
    reg.Close
    sql = "Select * from docu_buque_viaje " & _
          " Where buqv_linea ='" & Cmb_linea.BoundText & "'" & _
          "   and buqv_buque ='" & Cmb_buque.BoundText & "'" & _
          "   and buqv_viaje ='" & Cmb_viaje.BoundText & "'"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Txt_fecha.Text = reg("buqv_fecha_arribo")
End If

End Sub

Private Sub Cmb_pto_emb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
End If
End Sub

Private Sub Cmb_viaje_Change()
   Dim sql As String
   Dim reg As New ADODB.Recordset
   
   sql = "Select * from docu_manifiesto " & _
         " Where mani_linea ='" & Cmb_linea.BoundText & "'" & _
         "   And mani_buque ='" & Cmb_buque.BoundText & "'" & _
         "   And mani_viaje ='" & Cmb_viaje.BoundText & "'"
         '"   And buqv_status = 'A'"
   
   carga_datacombo data_pto_emb, Cmb_pto_emb, sql, "mani_puer_emb", "mani_puer_emb"
   
'If Opt_tarja.Value Then
    sql = "Select distinct blcab_cono_emb " & _
               " from docu_blcab " & _
               " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Lst_bl.Clear
    Do While Not reg.EOF
        Lst_bl.AddItem reg("blcab_cono_emb")
        reg.MoveNext
    Loop
    reg.Close
    sql = "Select * from docu_buque_viaje " & _
          " Where buqv_linea ='" & Cmb_linea.BoundText & "'" & _
          "   and buqv_buque ='" & Cmb_buque.BoundText & "'" & _
          "   and buqv_viaje ='" & Cmb_viaje.BoundText & "'"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Txt_fecha.Text = reg("buqv_fecha_arribo")
'End If

End Sub

Private Sub Cmb_viaje_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
End If
End Sub

Public Sub Cmd_cons_bl_Click()
Dim reg As New ADODB.Recordset

    sql = "Select blcab_cono_emb " & _
               " from docu_blcab " & _
               " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'" & _
               "   and blcab_puer_emb    ='" & Cmb_pto_emb.BoundText & "'" & _
               "   and blcab_puer_des    ='" & Txt_pto_descarga.Text & "'" & _
               "   and blcab_emp_resp    ='" & Me.Cmb_oper_carga.BoundText & "'"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Lst_bl.Clear
    Do While Not reg.EOF
        Lst_bl.AddItem reg("blcab_cono_emb")
        reg.MoveNext
    Loop
    reg.Close
    sql = "Select * from docu_buque_viaje " & _
          " Where buqv_linea ='" & Cmb_linea.BoundText & "'" & _
          "   and buqv_buque ='" & Cmb_buque.BoundText & "'" & _
          "   and buqv_viaje ='" & Cmb_viaje.BoundText & "'"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Txt_fecha.Text = reg("buqv_fecha_arribo")
End Sub

Private Sub Cmd_hlp_pto_descarga_Click()
   Txt_pto_descarga.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Descarga")
End Sub


Private Sub Cmd_plantilla_Click()
On Error GoTo fin
    Dim objWord As Object
    doc_plantilla = App.Path + "\Plantillas\" + "PlantillaCorrec.doc"
    Set objWord = CreateObject("Word.Application")
    objWord.Documents.Open doc_plantilla
    objWord.Visible = True
    Set objWord = Nothing
fin:
End Sub

Private Sub Cmd_procesar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Chk_todos_Click()
Dim valor As Boolean
valor = Chk_todos.Value
For i = 0 To Lst_bl.ListCount - 1
    Lst_bl.Selected(i) = valor
Next i
End Sub

Private Sub Cmd_procesar_Click()
Dim Ruta As String

Dim i As Integer
Set oWord = Nothing

    If Lst_bl.SelCount = 0 Then
        MsgBox "No existe ningun bl seleccionado", vbInformation
        Exit Sub
    End If
    Set oWord = CreateObject("Word.Application")
    Ruta = App.Path + "\Plantillas\"
    doc_plantilla = Ruta + "PlantillaCorrec.doc"
    doc_origen = Ruta + "Origen.doc"
    
    If Dir(doc_origen, vbArchive) = "" Then
       MsgBox "No se encontró archivo Origen de Datos " + doc_origen, vbExclamation
       Exit Sub
    End If
    If Dir(doc_plantilla, vbArchive) = "" Then
       MsgBox "No se encontro archivo Plantilla " + doc_plantilla, vbExclamation
       Exit Sub
    End If
    Me.MousePointer = 11
    Call Imprime_bls
    Me.MousePointer = 0
'End If
End Sub

Private Sub Cmd_cancel_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
    If Not data_linea.Recordset.EOF Then data_linea.Recordset.MoveFirst
    Opt_bl.Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Set oWord = Nothing
        Set oCell = Nothing
        Set oMerge = Nothing
        Set sht_est = Nothing
        Set wb_est = Nothing
End Sub

Private Sub Opt_bl_Click()
    'Cmb_pto.Enabled = True
    Cmd_cons_bl.Enabled = False
    Cmb_pto_emb.Enabled = True
    txt_bl.Enabled = True
    Lst_bl.Enabled = True
    Chk_todos.Enabled = True
'    Cmb_formato.Enabled = True
End Sub


Public Sub Imprime_bls()
Dim msg         As String
Dim nombre      As String
Dim ls_sql      As String
Dim condicion   As String
Dim desc        As String
Dim fecha       As Date
'*******************
Dim NomArchivo  As String
Dim reg         As New ADODB.Recordset 'variables que trae el grupo de datos seleccionados
Dim REG_det     As New ADODB.Recordset
Dim reg_cont    As New ADODB.Recordset
Dim reg_descrip As New ADODB.Recordset
Dim LWS_Carpeta As String
Dim bls         As String
Dim i           As Integer
Dim marcas      As String
Dim Tipo_cont   As String
Dim Sello       As String
Dim no_detalles As Integer
Dim desc_carga  As String

On Error GoTo CheckError

    '************************************************************
    oWord.Documents.Open doc_plantilla
    oWord.Documents.Open doc_origen
    '************************************************************
    bls = ""
    oWord.Visible = True
    For i = 0 To Lst_bl.ListCount - 1
        If Lst_bl.Selected(i) Then
           Lst_bl.ListIndex = i
           bls = bls & "'" & Lst_bl.Text & "', "
        End If
    Next i
    bls = Mid(bls, 1, Len(Trim(bls)) - 1) ' se quita la ultima coma
    
    sql = "Select * " & _
               " from docu_blcab, docu_buque_viaje " & _
               " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'" & _
               "   and blcab_linea = buqv_linea " & _
               "   and blcab_buque = buqv_buque " & _
               "   and blcab_viaje = buqv_viaje "
               
               If Cmb_pto_emb.VisibleCount <> 0 Then
                  sql = Trim$(sql) & "   and blcab_puer_emb    ='" & Cmb_pto_emb.BoundText & "'"
               End If
               sql = Trim$(sql) & "   and blcab_cono_emb in (" & bls & ") " & _
                     " and blcab_puer_des = '" & Gstr_param.pto_descarga & "'"

    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        
    If reg.EOF Then
        MsgBox "No existen estos Bls", vbExclamation
        Exit Sub
    End If
    reg.MoveFirst
    i = 2 '2
    Do While Not reg.EOF
    '****************************************************************************
       'LWS_Carpeta = reg("puerto")
       If Dir(Gstr_Ruta_BlWord, vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord)
       If Dir(Gstr_Ruta_BlWord & Trim(Cmb_buque.Text), vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord & Trim(Cmb_buque.Text))
       If Dir(Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text), vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text))
              
       NomArchivo = Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text) & "\" & Trim$(reg("BLCAB_CONO_EMB")) & "(corrector).doc"
       oWord.Documents(doc_origen).Activate
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=1) 'shipper 1
       oCell.Range.Text = Trim(reg("blcab_embarcador"))
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=2) 'shipper 2
       oCell.Range.Text = Trim(reg("blcab_direc_embarc"))
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=3) 'Consignee1
       oCell.Range.Text = reg("blcab_consig")
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=4) 'Consignee2
       oCell.Range.Text = Trim(reg("blcab_ide_cons"))
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=5) 'Consignee3
       oCell.Range.Text = Mid(reg("blcab_direc_consig"), 61)
              
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=6)  'Notify 1
       oCell.Range.Text = Trim(reg("blcab_notif"))
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=7)  'Notify 2
       oCell.Range.Text = Trim(reg("blcab_direc_notif"))
              
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=8) 'viaje
       oCell.Range.Text = Cmb_viaje.Text
       nombre = ""
                          
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=9) 'Bl1
       oCell.Range.Text = reg("blcab_cono_emb")
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=10) 'Bl2
       oCell.Range.Text = ""
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=11) 'buque
       oCell.Range.Text = Trim(Cmb_buque.Text)
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=58) ' fecha embarque
       oCell.Range.Text = Format(reg("buqv_fecha_arribo"), "dd/mmm/yyyy")
       
       '========================== EMPIEZA LA IMPRESION ============================
       Call Print_Merge(NomArchivo, i - 1)
       '-----------------------------------------------------
       reg.MoveNext: i = i + 1
    Loop
    
    oWord.Documents(doc_origen).Close savechanges:=wdDoNotSaveChanges
    oWord.Documents(doc_plantilla).Close savechanges:=wdDoNotSaveChanges
    oWord.Visible = True
    Exit Sub
    
CANCELAR:
    MsgBox msg + Chr(13) + Chr(13) + "Se cancela la Carga de Datos", vbExclamation, App.Title
    oWord.Documents(param_doc_origen).Close savechanges:=wdDoNotSaveChanges
    Cmb_trol.SetFocus
    Exit Sub
    Resume
 
CheckError:
   
   MsgBox "Ocurrio el Error " & Err.Description
   Exit Sub
   Resume 0
End Sub

Sub Print_Merge(NomArchivo As String, no_registo As Integer)
Dim ls_desc As String
    ls_desc = ""
    '************************************************************
    'oWord.Documents.Open doc_plantilla
    '************************************************************
    oWord.Documents(doc_plantilla).Activate
    Set oMerge = oWord.ActiveDocument.MailMerge
       With oMerge.DataSource
                   .FirstRecord = no_registo
                   .LastRecord = no_registo
       End With
    With oMerge
         .Destination = wdSendToNewDocument
         .Execute
         oWord.ActiveDocument.SaveAs FileName:=NomArchivo
         
    End With
End Sub

Private Sub Txt_bl_KeyPress(KeyAscii As Integer)
Dim reg As New ADODB.Recordset

If KeyAscii = 13 Then
If Opt_bl.Value Then
    sql = "Select blcab_cono_emb " & _
               " from docu_blcab " & _
               " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'" & _
               "   and blcab_puer_emb    ='" & Cmb_pto_emb.BoundText & "'" & _
               "   and blcab_cono_emb    ='" & txt_bl.Text & "'"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Lst_bl.Clear
    Do While Not reg.EOF
        Lst_bl.AddItem reg("blcab_cono_emb")
        reg.MoveNext
    Loop
    reg.Close
    sql = "Select * from docu_buque_viaje " & _
          " Where buqv_linea ='" & Cmb_linea.BoundText & "'" & _
          "   and buqv_buque ='" & Cmb_buque.BoundText & "'" & _
          "   and buqv_viaje ='" & Cmb_viaje.BoundText & "'"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Txt_fecha.Text = reg("buqv_fecha_arribo")
End If
SendKeys "{TAB}"
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Public Sub txt_bl_LostFocus()
Dim reg As New ADODB.Recordset

If Opt_bl.Value Then
    sql = "Select blcab_cono_emb " & _
               " from docu_blcab " & _
               " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'" & _
               "   and blcab_puer_emb    ='" & Cmb_pto_emb.BoundText & "'" & _
               "   and blcab_cono_emb    ='" & txt_bl.Text & "'"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Lst_bl.Clear
    Do While Not reg.EOF
        Lst_bl.AddItem reg("blcab_cono_emb")
        reg.MoveNext
    Loop
    reg.Close
End If
End Sub
