VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form doc_repo_bl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Bls de importación"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   5940
   Begin VB.Frame Frame1 
      Height          =   5625
      Left            =   75
      TabIndex        =   13
      Top             =   315
      Width           =   5760
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
            Picture         =   "doc_copia_bl.frx":0000
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
      Begin VB.CommandButton Cmd_cons_bl 
         Caption         =   "&Consulta Bls"
         Height          =   315
         Left            =   4365
         TabIndex        =   6
         Top             =   2700
         Width           =   1110
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
         Width           =   3135
      End
      Begin VB.ComboBox Cmb_formato 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   5175
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
         Width           =   855
      End
   End
   Begin VB.OptionButton Opt_tarja 
      Caption         =   "Tarja de Importación"
      Height          =   255
      Left            =   2145
      TabIndex        =   15
      Top             =   105
      Width           =   1815
   End
   Begin VB.OptionButton Opt_bl 
      Caption         =   "Impresión de Bls"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   75
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   360
      Top             =   6015
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   4725
      TabIndex        =   8
      Top             =   6045
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   495
      Left            =   3705
      TabIndex        =   7
      Top             =   6045
      Width           =   975
   End
End
Attribute VB_Name = "doc_repo_bl"
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

Private Sub Cmd_cons_bl_Click()
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


Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Chk_todos_Click()
Dim valor As Boolean
valor = Chk_todos.Value
For i = 0 To Lst_bl.ListCount - 1
    Lst_bl.Selected(i) = valor
Next i
End Sub

Private Sub Command1_Click()
Dim Ruta As String

Dim i As Integer
Set oWord = Nothing

If Opt_tarja.Value Then
    Set wb_est = CreateObject("excel.Application")
    DoEvents
    Me.MousePointer = 11
    Call LISTADO_EXCEL
    Me.MousePointer = 0
Else
    If Lst_bl.SelCount = 0 Then
        MsgBox "No existe ningun bl seleccionado", vbInformation
        Exit Sub
    End If
    Set oWord = CreateObject("Word.Application")
    Ruta = App.Path + "\Plantillas\"
    If Cmb_formato.Text = "MSC" Then
       doc_plantilla = Ruta + "PlantillaMSC.doc"
    ElseIf Cmb_formato.Text = "MSC 53" Then
       doc_plantilla = Ruta + "PlantillaMSC53.doc"
    ElseIf Cmb_formato.Text = "KHS" Then
       doc_plantilla = Ruta + "PlantillaKHS.doc"
    End If
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
End If
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
    If Not data_linea.Recordset.EOF Then data_linea.Recordset.MoveFirst
    Cmb_formato.AddItem "KHS"
    Cmb_formato.AddItem "MSC"
    Cmb_formato.AddItem "MSC 53"
    Cmb_formato.ListIndex = 0
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
    Cmb_formato.Enabled = True
End Sub

Private Sub Opt_tarja_Click()
    Cmd_cons_bl.Enabled = False
    Cmb_pto_emb.Enabled = False
    txt_bl.Enabled = False
    Cmb_formato.Enabled = False
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

'    If MsgBox("Este proceso tomará algunos minutos. Desea continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    '************************************************************
    oWord.Documents.Open doc_plantilla
    oWord.Documents.Open doc_origen
    '************************************************************
    bls = ""
    For i = 0 To Lst_bl.ListCount - 1
        If Lst_bl.Selected(i) Then
           Lst_bl.ListIndex = i
           bls = bls & "'" & Lst_bl.Text & "', "
        End If
    Next i
    bls = Mid(bls, 1, Len(Trim(bls)) - 1) ' se quita la ultima coma
    
    sql = "Select * " & _
               " from docu_blcab " & _
               " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'"
               If Cmb_pto_emb.VisibleCount <> 0 Then
                  sql = Trim$(sql) & "   and blcab_puer_emb    ='" & Cmb_pto_emb.BoundText & "'"
               End If
               sql = Trim$(sql) & "   and blcab_cono_emb in (" & bls & ") " & _
                     " and blcab_puer_des = 'ECGYE' "

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
              
       NomArchivo = Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text) & "\" & Trim$(reg("BLCAB_CONO_EMB")) & ".doc"
       oWord.Documents(doc_origen).Activate
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=1) 'shipper 1
       oCell.Range.Text = Trim(reg("blcab_embarcador"))
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=2) 'shipper 2
       oCell.Range.Text = Trim(reg("blcab_direc_embarc"))
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=3) 'Consignee1
       oCell.Range.Text = reg("blcab_ide_cons")
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=4) 'Consignee2
       oCell.Range.Text = Trim(reg("blcab_consig"))
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=5) 'Consignee3
       oCell.Range.Text = Mid(reg("blcab_direc_consig"), 61)
              
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=6)  'Notify 1
       oCell.Range.Text = Trim(reg("blcab_notif"))
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=7)  'Notify 2
       oCell.Range.Text = Trim(reg("blcab_direc_notif"))
              
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=8) 'viaje
       oCell.Range.Text = Cmb_viaje.Text
       nombre = ""
              
       If Cmb_formato.Text = "MSC" Or Cmb_formato.Text = "MSC 53" Then
          For k = 1 To Len(Trim(reg("blcab_cono_emb")))
             If IsNumeric(Mid(reg("blcab_cono_emb"), k, 1)) Then Exit For
          Next k
          cod_bl1 = Mid(reg("blcab_cono_emb"), 1, k - 1):
          cod_bl2 = Mid(reg("blcab_cono_emb"), k)
       ElseIf Cmb_formato.Text = "KHS" Then
          cod_bl1 = reg("blcab_cono_emb"):  cod_bl2 = ""
       End If
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=9) 'Bl1
       oCell.Range.Text = cod_bl1
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=10) 'Bl2
       oCell.Range.Text = cod_bl2
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=11) 'buque
       oCell.Range.Text = Trim(Cmb_buque.Text)
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=12) 'pto embarque
       oCell.Range.Text = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", reg("blcab_puer_emb"), " puer_linea ='" & Cmb_linea.BoundText & "'")
         
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=13) 'pto descarga
       oCell.Range.Text = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", reg("blcab_puer_des"), " puer_linea ='" & Cmb_linea.BoundText & "'")
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=14) 'pto destino final
       oCell.Range.Text = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", reg("blcab_pto_dest_final"), " puer_linea ='" & Cmb_linea.BoundText & "'")
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=15) 'bultos
       'If reg("blcab_no_cont") = 0 Then
            oCell.Range.Text = reg("blcab_bultos_tot")
       'Else
       '     oCell.Range.Text = reg("blcab_no_cont")
       'End If
       
       ' se buscan los detalles de la carga
       sql = "Select count(*) " & _
               " from docu_bldet " & _
               " Where bldet_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and bldet_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and bldet_viaje       ='" & Cmb_viaje.BoundText & "'" & _
               "   and bldet_puer_emb    ='" & reg("blcab_puer_emb") & "'" & _
               "   and bldet_puer_des    ='" & reg("blcab_puer_des") & "'" & _
               "   and bldet_emp_resp    ='" & reg("blcab_emp_resp") & "'" & _
               "   and bldet_cono_emb    ='" & reg("blcab_cono_emb") & "'"
       REG_det.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
       no_detalles = REG_det(0)
       REG_det.Close
       
       sql = "Select * " & _
               " from docu_bldet " & _
               " Where bldet_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and bldet_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and bldet_viaje       ='" & Cmb_viaje.BoundText & "'" & _
               "   and bldet_puer_emb    ='" & reg("blcab_puer_emb") & "'" & _
               "   and bldet_puer_des    ='" & reg("blcab_puer_des") & "'" & _
               "   and bldet_emp_resp    ='" & reg("blcab_emp_resp") & "'" & _
               "   and bldet_cono_emb    ='" & reg("blcab_cono_emb") & "'"

       REG_det.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
       If Not REG_det.EOF Then
          oCell.Range.Text = reg("blcab_bultos_tot") & " " & REG_det("bldet_embalaje") ' nuevo se agrega el embalaje    16/sep/2002
          For x = 1 To 10
              If REG_det.EOF Then Exit For
              desc_carga = REG_det("bldet_desc_carg")
              If Trim(REG_det("bldet_contenedor")) <> "" Then
                 marcas = REG_det("bldet_contenedor")
                 
                 ls_sql = "Select * from docu_cont " & _
                            " Where cont_linea       ='" & Cmb_linea.BoundText & "'" & _
                            "   and cont_buque       ='" & Cmb_buque.BoundText & "'" & _
                            "   and cont_viaje       ='" & Cmb_viaje.BoundText & "'" & _
                            "   and cont_puer_emb    ='" & REG_det("bldet_puer_emb") & "'" & _
                            "   and cont_puer_des    ='" & REG_det("bldet_puer_des") & "'" & _
                            "   and cont_contenedor  ='" & REG_det("bldet_contenedor") & "'"
                 reg_cont.Open ls_sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                 If Not reg_cont.EOF Then
                    marcas = marcas & Chr(13) & "     Seal :" & reg_cont("cont_sello1") & Chr(13) & "     Type : " & devuelve_desc("docu_tamano_equip", "tama_equip_codigo", "tama_equip_descripcion", reg_cont("cont_tam_equip"), "tama_equip_linea = '" & Cmb_linea.BoundText & "'")
                 End If
                 reg_cont.Close
              Else
                 marcas = chk_str(REG_det("bldet_marcas_num"))
              End If
              Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=15 + x) 'marcas
              oCell.Range.Text = marcas
              REG_det.MoveNext
          Next x
       End If
       'If reg_det.EOF Then reg_det.MovePrevious
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=26) ' desc. carga
       oCell.Range.Text = chk_str(desc_carga) 'reg_det("bldet_desc_carg"))
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=36) ' peso
       oCell.Range.Text = reg("blcab_peso_tot")
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=46) ' volumen
       oCell.Range.Text = reg("blcab_volumen_tot")
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=56) ' total bultos
       If reg("blcab_no_cont") = 0 Then
            oCell.Range.Text = reg("blcab_bultos_tot")
       Else
            oCell.Range.Text = reg("blcab_no_cont")
       End If

       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=57) ' total peso
       oCell.Range.Text = reg("blcab_peso_tot")
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=58) ' pto, fecha embarque
       oCell.Range.Text = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", reg("blcab_puer_emb"), " puer_linea ='" & Cmb_linea.BoundText & "'") & _
                          ", " & Format(reg("blcab_fecha_emb"), "dd/mmm/yyyy")
       '****************************************************************************
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=59) 'pto origen
       oCell.Range.Text = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", reg("blcab_puer_origen"), " puer_linea ='" & Cmb_linea.BoundText & "'")
       '========================== EMPIEZA LA IMPRESION ============================
       Call Print_Merge(NomArchivo, i - 1)
            '-----------------------------------------------------
            'impresion de anexo de contenedores por si acaso no entran en las linas del bl
            If Not REG_det.EOF Then
               Call anexo_contenedores(no_detalles - 9, reg("blcab_cono_emb")) ' es - 9 por la fila del titulo que ocupa un espacio
               linea = 2
               Do While Not REG_det.EOF
                   If Trim(REG_det("bldet_contenedor")) <> "" Then
                      marcas = REG_det("bldet_contenedor")
                      ls_sql = "Select * from docu_cont " & _
                            " Where cont_linea       ='" & Cmb_linea.BoundText & "'" & _
                            "   and cont_buque       ='" & Cmb_buque.BoundText & "'" & _
                            "   and cont_viaje       ='" & Cmb_viaje.BoundText & "'" & _
                            "   and cont_puer_emb    ='" & REG_det("bldet_puer_emb") & "'" & _
                            "   and cont_puer_des    ='" & REG_det("bldet_puer_des") & "'" & _
                            "   and cont_contenedor  ='" & REG_det("bldet_contenedor") & "'"
                     reg_cont.Open ls_sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                     If Not reg_cont.EOF Then
                         Sello = reg_cont("cont_sello1")
                         ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                         'TL para traer la descripcion del tamaño del contenedor
                         ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                         If reg("blcab_emp_anav") = "1186" Then 'variable reemplazo
                            sql = "Select descripcion from dtx_tipo_contein " & _
                                  "Where apg ='" & reg_cont("cont_tam_equip") & "' " & _
                                  "and linea = 'KHS'"    ' TL revisar esto
                         Else
                            sql = "Select descripcion from dtx_tipo_contein " & _
                                  "Where apg ='" & reg_cont("cont_tam_equip") & "'" & _
                                  "and linea = 'MSC'"    ' TL revisar esto
                         End If
                         reg_descrip.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                         If Not reg_descrip.EOF Then
                            Tipo_cont = Mid$(Trim$(reg_descrip("descripcion")), 1, 20)
                         Else
                            Tipo_cont = reg_cont("cont_tam_equip")
                         End If
                         reg_descrip.Close
                         '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                      End If
                      reg_cont.Close
                   Else
                      marcas = chk_str(REG_det("bldet_marcas_num"))
                   End If
                   Set oCell = oWord.ActiveDocument.Tables(3).Cell(Row:=linea, Column:=1) 'contenedor
                   oCell.Range.Text = marcas
                   Set oCell = oWord.ActiveDocument.Tables(3).Cell(Row:=linea, Column:=2)  'sello
                   oCell.Range.Text = Sello
                   Set oCell = oWord.ActiveDocument.Tables(3).Cell(Row:=linea, Column:=3)  'tipo
                   oCell.Range.Text = Tipo_cont
                   Set oCell = oWord.ActiveDocument.Tables(3).Cell(Row:=linea, Column:=4)  'PESO
                   oCell.Range.Text = Format$(REG_det("bldet_PESO"), "#,###,###,##0.00")
                   Set oCell = oWord.ActiveDocument.Tables(3).Cell(Row:=linea, Column:=5)  'NO BULTOS
                   oCell.Range.Text = REG_det("bldet_BULTOS")
                   linea = linea + 1: REG_det.MoveNext
               Loop
               oWord.ActiveDocument.SaveAs FileName:=NomArchivo
            End If
            REG_det.Close
            'fin de anexo de contenedores
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

Function anexo_contenedores(filas As Integer, COD_BL As String)
On Error GoTo errores_anexo
Dim campo As String

    campo = "460": oWord.Selection.MoveDown Unit:=wdLine, Count:=80
    campo = "461": oWord.Selection.MoveRight Unit:=wdCharacter, Count:=1
    campo = "462": oWord.Selection.InsertBreak Type:=wdPageBreak
    campo = "463": oWord.Selection.TypeParagraph
    campo = "464": oWord.Selection.TypeParagraph
    campo = "465": oWord.Selection.TypeParagraph
    campo = "466": oWord.Selection.TypeText Text:="CONTEINER LIST"
    campo = "467": oWord.Selection.HomeKey Unit:=wdLine, Extend:=wdMove
    campo = "468": oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    campo = "469": oWord.Selection.Font.Bold = wdToggle
    campo = "470": oWord.Selection.Font.Size = 12
    campo = "471": oWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    campo = "473": oWord.Selection.EndKey Unit:=wdLine, Extend:=wdMove
    campo = "474": oWord.Selection.TypeParagraph
    campo = "475": oWord.Selection.TypeParagraph
    campo = "476": oWord.Selection.TypeParagraph
    campo = "477": oWord.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    campo = "478": oWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    campo = "479": oWord.Selection.Font.Bold = wdToggle
    campo = "480": oWord.Selection.Font.Size = 10
    campo = "481": oWord.Selection.TypeParagraph
    campo = "494": oWord.Selection.MoveDown Unit:=wdLine, Count:=1
    campo = "481": oWord.Selection.TypeParagraph
    
    campo = "483": oWord.ActiveDocument.Tables.Add Range:=oWord.Selection.Range, NumRows:=3, NumColumns:=4
    campo = "484": oWord.Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    campo = "485": oWord.Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    campo = "486": oWord.Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    campo = "487": oWord.Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    campo = "488": oWord.Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    campo = "489": oWord.Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    campo = "490": oWord.Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    campo = "491": oWord.Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
    campo = "492": oWord.Selection.Borders(wdBorderVertical).LineStyle = wdLineStyleNone
    
    campo = "494": oWord.Selection.MoveDown Unit:=wdLine, Count:=1
    campo = "495": Set oCell = oWord.ActiveDocument.Tables(2).Cell(Row:=1, Column:=1)
    campo = "496": oCell.Range.Font.Bold = wdToggle:    oCell.Range.Font.Size = 10: oCell.Range.Text = "VESSEL :":
    campo = "497": Set oCell = oWord.ActiveDocument.Tables(2).Cell(Row:=1, Column:=3)
    campo = "498": oCell.Range.Font.Bold = wdToggle:    oCell.Range.Font.Size = 10: oCell.Range.Text = "VOYAGE :":
    campo = "499": Set oCell = oWord.ActiveDocument.Tables(2).Cell(Row:=2, Column:=1)
    campo = "500": oCell.Range.Font.Bold = wdToggle:    oCell.Range.Font.Size = 10: oCell.Range.Text = "PORT OF LOADING :":
    campo = "501": Set oCell = oWord.ActiveDocument.Tables(2).Cell(Row:=2, Column:=3)
    campo = "502": oCell.Range.Font.Bold = wdToggle:   oCell.Range.Font.Size = 10:  oCell.Range.Text = "PORT OF DISCHARGE :":
    campo = "503": Set oCell = oWord.ActiveDocument.Tables(2).Cell(Row:=3, Column:=1)
    campo = "504": oCell.Range.Font.Bold = wdToggle:    oCell.Range.Font.Size = 10: oCell.Range.Text = "BILL OF LADING :":
    
    campo = "506": Set oCell = oWord.ActiveDocument.Tables(2).Cell(Row:=1, Column:=2)
    campo = "507": oCell.Range.Text = Trim$(Cmb_buque.Text)
    campo = "508": Set oCell = oWord.ActiveDocument.Tables(2).Cell(Row:=1, Column:=4)
    campo = "509": oCell.Range.Text = Trim$(Cmb_viaje.Text)
    campo = "510": Set oCell = oWord.ActiveDocument.Tables(2).Cell(Row:=2, Column:=2)
    campo = "511": oCell.Range.Text = Trim$(Cmb_pto_emb.Text) 'TL confirmar que puerto es
    campo = "512": Set oCell = oWord.ActiveDocument.Tables(2).Cell(Row:=2, Column:=4)
    campo = "513": oCell.Range.Text = "GUAYAQUIL"
    campo = "514": Set oCell = oWord.ActiveDocument.Tables(2).Cell(Row:=3, Column:=2)
    campo = "515": oCell.Range.Text = Trim$(COD_BL)
    
    campo = "517": oWord.Selection.EndKey Unit:=wdLine, Extend:=wdMove
    campo = "518": oWord.Selection.MoveDown Unit:=wdLine, Count:=15, Extend:=wdMove
    campo = "519": oWord.Selection.TypeParagraph
    campo = "520": oWord.Selection.TypeParagraph
    campo = "521": oWord.ActiveDocument.Tables.Add Range:=oWord.Selection.Range, NumRows:=filas, NumColumns:=5
    campo = "522": oWord.Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    campo = "523": oWord.Selection.MoveRight Unit:=wdCharacter, Count:=4, Extend:=wdExtend
    campo = "524": oWord.Selection.MoveDown Unit:=wdLine, Count:=filas - 2, Extend:=wdExtend
    campo = "525": oWord.Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    campo = "526": oWord.Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    campo = "527": oWord.Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    campo = "528": oWord.Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    campo = "529": oWord.Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
    campo = "530": oWord.Selection.Borders(wdBorderVertical).LineStyle = wdLineStyleNone
    
    campo = "532": Set oCell = oWord.ActiveDocument.Tables(3).Cell(Row:=1, Column:=1)
    campo = "533": oCell.Range.Text = "CONTEINER No.": oCell.Range.Font.Bold = wdToggle:    oCell.Range.Font.Size = 10
    campo = "534": Set oCell = oWord.ActiveDocument.Tables(3).Cell(Row:=1, Column:=2)
    campo = "535": oCell.Range.Text = "SEAL":          oCell.Range.Font.Bold = wdToggle:    oCell.Range.Font.Size = 10
    campo = "536": Set oCell = oWord.ActiveDocument.Tables(3).Cell(Row:=1, Column:=3)
    campo = "537": oCell.Range.Text = "TYPE":          oCell.Range.Font.Bold = wdToggle:    oCell.Range.Font.Size = 10
    campo = "538": Set oCell = oWord.ActiveDocument.Tables(3).Cell(Row:=1, Column:=4)
    campo = "539": oCell.Range.Text = "WEIGHT":        oCell.Range.Font.Bold = wdToggle:    oCell.Range.Font.Size = 10
    campo = "540": Set oCell = oWord.ActiveDocument.Tables(3).Cell(Row:=1, Column:=5)
    campo = "541": oCell.Range.Text = "PACKAGES":      oCell.Range.Font.Bold = wdToggle:    oCell.Range.Font.Size = 10
    
    Exit Function
    
errores_anexo:
   MsgBox "Ocurrio el error " & Err.Description & " en la linea: " & campo
   Exit Function
   Resume 0
End Function

Public Function LISTADO_EXCEL()
On Error GoTo Errores
Dim REG_det As New ADODB.Recordset
Dim reg_descrip As New ADODB.Recordset
Dim i As Integer
Dim cont_20, cont_40 As Integer
Dim linea As Integer
      
      
      wb_est.Workbooks.Add ' aume nta un libro dentro de la aplicacion excel
      '*** crea nuevas hojas dentro del mismo libro *******
      'For i = 1 To 2  'hasta el numero de reportes
      '   wb_est.Sheets.Add
      '   DoEvents
      'Next i
      pagina% = 1
         'Set sht_est = wb_est.ActiveWorkbook.Sheets("Hoja1")   'setea la hoja actual
         'Set sht_est = wb_est.ActiveWorkbook.Sheets("Sheet1")   'setea la hoja actual
         Set sht_est = wb_est.ActiveWorkbook.Sheets(1)   'setea la hoja actual
         '================>>>>>>>   VERSION_EXCEL  <<<<<<<===================
         'wb_est.Application.Visible = True ' visible la aplicación
         
         sht_est.Name = "Discharge List"         'asigna un nombre a la hoja
         sht_est.PageSetup.Orientation = 2
         sht_est.PageSetup.PaperSize = xlPaperA4
         sht_est.PageSetup.BottomMargin = 40
         sht_est.PageSetup.TopMargin = 40
         sht_est.PageSetup.FooterMargin = 20
         sht_est.PageSetup.HeaderMargin = 20
         sht_est.Activate
      
   sht_est.Cells(2, 5).Value = "DISCHARGE LIST"
   sht_est.Cells(4, 3).Value = "VESSEL :": sht_est.Cells(4, 4).Value = Cmb_buque.Text
   sht_est.Cells(5, 3).Value = "VOYAGE :": sht_est.Cells(5, 4).Value = Cmb_viaje.Text
   sht_est.Cells(6, 3).Value = "DATE   :": sht_est.Cells(6, 4).Value = Txt_fecha.Text
   
   sht_est.Cells(9, 1).Value = "POL"
   sht_est.Cells(9, 2).Value = "BILL NO"
   sht_est.Cells(9, 3).Value = "CNR NO"
   sht_est.Cells(9, 4).Value = "TYPE"
   sht_est.Cells(9, 5).Value = "CARGO"
   sht_est.Cells(9, 6).Value = "SEAL"
   sht_est.Cells(9, 7).Value = "GROSS"
   sht_est.Cells(9, 8).Value = "PACKING"
   sht_est.Cells(9, 9).Value = "IMO"
   sht_est.Cells(9, 10).Value = "CONSIGNEE"
   
   sht_est.Rows(2).Font.Name = "moderm"
   sht_est.Rows(2).Font.Bold = True
   sht_est.Rows(2).Font.Italic = True
   sht_est.Rows(2).Font.Color = 8000000
   sht_est.Rows(2).Font.Size = 14
   
   sht_est.Rows(9).Font.Name = "moderm"
   sht_est.Rows(9).Font.Bold = True
   sht_est.Rows(9).Font.Italic = True
   
   sht_est.Cells(4, 3).Font.Bold = True:   sht_est.Cells(4, 3).Font.Color = 8000000
   sht_est.Cells(5, 3).Font.Bold = True:   sht_est.Cells(5, 3).Font.Color = 8000000
   sht_est.Cells(6, 3).Font.Bold = True:   sht_est.Cells(6, 3).Font.Color = 8000000
   sht_est.Cells(7, 3).Font.Bold = True:   sht_est.Cells(7, 3).Font.Color = 8000000
   
   DoEvents
   
   Set rng_est = sht_est.Range(sht_est.Cells(4, 1), sht_est.Cells(4, 1)):   rng_est.ColumnWidth = 8 'POL
   Set rng_est = sht_est.Range(sht_est.Cells(4, 2), sht_est.Cells(4, 2)):   rng_est.ColumnWidth = 13 ' BILL NO
   Set rng_est = sht_est.Range(sht_est.Cells(1, 3), sht_est.Cells(1, 3)):   rng_est.ColumnWidth = 10 ' CNT
   Set rng_est = sht_est.Range(sht_est.Cells(1, 4), sht_est.Cells(1, 4)):   rng_est.ColumnWidth = 8 ' TYPE
   Set rng_est = sht_est.Range(sht_est.Cells(1, 5), sht_est.Cells(1, 5)):   rng_est.ColumnWidth = 30 'CARGO
   Set rng_est = sht_est.Range(sht_est.Cells(1, 6), sht_est.Cells(1, 6)):   rng_est.ColumnWidth = 8 'SEAL
   Set rng_est = sht_est.Range(sht_est.Cells(1, 7), sht_est.Cells(1, 7)):   rng_est.ColumnWidth = 8 'GROSS
   Set rng_est = sht_est.Range(sht_est.Cells(1, 8), sht_est.Cells(1, 8)):   rng_est.ColumnWidth = 10 'PACKING
   Set rng_est = sht_est.Range(sht_est.Cells(1, 9), sht_est.Cells(1, 9)):   rng_est.ColumnWidth = 7 'IMO
   Set rng_est = sht_est.Range(sht_est.Cells(1, 10), sht_est.Cells(1, 10)): rng_est.ColumnWidth = 30 'CONSIGNEE
      
   sht_est.Cells(9, 1).Borders.LineStyle = 1
   sht_est.Cells(9, 2).Borders.LineStyle = 1
   sht_est.Cells(9, 3).Borders.LineStyle = 1
   sht_est.Cells(9, 4).Borders.LineStyle = 1
   sht_est.Cells(9, 5).Borders.LineStyle = 1
   sht_est.Cells(9, 6).Borders.LineStyle = 1
   sht_est.Cells(9, 7).Borders.LineStyle = 1
   sht_est.Cells(9, 8).Borders.LineStyle = 1
   sht_est.Cells(9, 9).Borders.LineStyle = 1
   sht_est.Cells(9, 10).Borders.LineStyle = 1
 
   
       sql = "Select * " & _
               " from docu_blcab, docu_bldet, docu_cont " & _
               " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'" & _
               "   and blcab_linea       = bldet_linea" & _
               "   and blcab_buque       = bldet_buque" & _
               "   and blcab_viaje       = bldet_viaje" & _
               "   and blcab_puer_emb    = bldet_puer_emb" & _
               "   and blcab_puer_des    = bldet_puer_des" & _
               "   and blcab_emp_resp    = bldet_emp_resp" & _
               "   and blcab_cono_emb    = bldet_cono_emb" & _
               "   and blcab_linea       = cont_linea" & _
               "   and blcab_buque       = cont_buque" & _
               "   and blcab_viaje       = cont_viaje" & _
               "   and blcab_puer_emb    = cont_puer_emb" & _
               "   and blcab_puer_des    = cont_puer_des" & _
               "   and bldet_contenedor  = cont_contenedor" & _
               "   and bldet_contenedor <> ''"

       REG_det.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
           
       cont_20 = 0
       cont_40 = 0
       linea = 10
       Do While Not REG_det.EOF
            sht_est.Rows(linea).Font.Size = 7
            sht_est.Cells(linea, 1).Value = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", REG_det("blcab_puer_emb"), " puer_linea ='" & Cmb_linea.BoundText & "'")
            sht_est.Cells(linea, 2).Value = REG_det("blcab_cono_emb")
            sht_est.Cells(linea, 3).Value = REG_det("bldet_contenedor")
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'TL aumento esto para sacar la descripcion del tipo de contenedor
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If REG_det("blcab_emp_anav") = "1186" Then
               sql = "Select size as tama,descripcion from dtx_tipo_contein " & _
                     "Where apg ='" & REG_det("cont_tam_equip") & "' " & _
                     "and linea = 'KHS'"    ' TL revisar esto
            Else
               sql = "Select size as tama,descripcion from dtx_tipo_contein " & _
                     "Where apg ='" & REG_det("cont_tam_equip") & "'" & _
                     "and linea = 'MSC'"    ' TL revisar esto
            End If
            reg_descrip.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
            If Not reg_descrip.EOF Then
               sht_est.Cells(linea, 4).Value = Mid$(Trim$(reg_descrip("descripcion")), 1, 20)
               If reg_descrip("tama") = "20" Then
                  cont_20 = cont_20 + 1
               Else
                  cont_40 = cont_40 + 1
               End If
            Else
               sht_est.Cells(linea, 4).Value = REG_det("cont_tam_equip")
            End If
            reg_descrip.Close
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            sht_est.Cells(linea, 5).Value = chk_str(REG_det("bldet_desc_carg"))
            sht_est.Cells(linea, 6).Value = " " & REG_det("cont_sello1")
            
            If IsNull(REG_det("bldet_peso")) Then
               sht_est.Cells(linea, 7).Value = 0
            Else
              sht_est.Cells(linea, 7).Value = CDbl(REG_det("bldet_peso"))
            End If
            If IsNull(REG_det("bldet_bultos")) Then
               sht_est.Cells(linea, 8).Value = 0
            Else
               sht_est.Cells(linea, 8).Value = CDbl(REG_det("bldet_bultos"))
            End If
            sht_est.Cells(linea, 9).Value = Trim(REG_det("bldet_IMO"))
            sht_est.Cells(linea, 10).Value = Trim(REG_det("blcab_consig"))
            REG_det.MoveNext
            linea = linea + 1
       Loop
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       'TL aumento de total de contenedores al reporte
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       linea = linea + 2
       sht_est.Cells(linea, 2).Value = "Total de Contenedores:"
       sht_est.Cells(linea, 4).Value = CDbl(linea - 12)
       linea = linea + 1
       sht_est.Cells(linea, 2).Value = "Total de Contenedores 20:"
       sht_est.Cells(linea, 4).Value = CDbl(cont_20)
       linea = linea + 1
       sht_est.Cells(linea, 2).Value = "Total de Contenedores 40:"
       sht_est.Cells(linea, 4).Value = CDbl(cont_40)
      '*************************************************************

      '*************************************************************

       If Dir(Gstr_Ruta_BlWord, vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord)
       If Dir(Gstr_Ruta_BlWord & Trim(Cmb_buque.Text), vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord & Trim(Cmb_buque.Text))
       If Dir(Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text), vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text))
       
       NomArchivo = Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text) & "\" & "Discharge List.XLS"
       If Dir(NomArchivo, vbArchive) <> "" Then Kill (NomArchivo)
       
       wb_est.ActiveWorkbook.SaveAs FileName:= _
        NomArchivo, FileFormat:= _
        xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
        , CreateBackup:=False
       
       'Set sht_est = Nothing
       wb_est.Application.Visible = True ' visible la aplicación
       'Set wb_est.Application
       'wb_est.Application.Quit
       ' Release the object variable.
       'Set wb_est = Nothing
       Exit Function

Errores:
MsgBox "Ocurrio el error " & Err.Description
       

End Function

Private Sub txt_bl_KeyPress(KeyAscii As Integer)
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

Private Sub txt_bl_LostFocus()
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
