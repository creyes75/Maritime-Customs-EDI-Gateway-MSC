VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form doc_repo_aviso_llegada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Avisos de Llegada"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   Icon            =   "doc_repo_aviso_llegada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5940
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   120
      Picture         =   "doc_repo_aviso_llegada.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Abre Documento Plantilla"
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   4320
      Picture         =   "doc_repo_aviso_llegada.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   5100
      Picture         =   "doc_repo_aviso_llegada.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   5265
      Left            =   75
      TabIndex        =   13
      Top             =   30
      Width           =   5760
      Begin VB.CommandButton Cmd_cons_bl 
         Caption         =   "&Consulta Bls"
         Height          =   675
         Left            =   4440
         Picture         =   "doc_repo_aviso_llegada.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2640
         Width           =   990
      End
      Begin VB.Frame Fra_pto_descarga 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   75
         TabIndex        =   26
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
            Picture         =   "doc_repo_aviso_llegada.frx":2922
            Style           =   1  'Graphical
            TabIndex        =   27
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
            TabIndex        =   28
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
         TabIndex        =   23
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
            TabIndex        =   24
            Top             =   0
            Width           =   1230
         End
      End
      Begin VB.Frame Fra_viaje 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   3090
         TabIndex        =   21
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
            TabIndex        =   22
            Top             =   0
            Width           =   345
         End
      End
      Begin VB.Frame Fra_buque 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   75
         TabIndex        =   19
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
            TabIndex        =   20
            Top             =   0
            Width           =   465
         End
      End
      Begin VB.Frame Fra_linea 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   75
         TabIndex        =   17
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
            TabIndex        =   18
            Top             =   0
            Width           =   825
         End
      End
      Begin VB.TextBox Txt_fecha 
         Height          =   285
         Left            =   330
         Locked          =   -1  'True
         TabIndex        =   16
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
         Width           =   2010
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
         Left            =   3990
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   4785
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
         TabIndex        =   29
         Top             =   2010
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.Label Label3 
         Caption         =   "Bill of Lading"
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   1995
         TabIndex        =   25
         Top             =   1455
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Bill of Lading"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   165
         TabIndex        =   15
         Top             =   2685
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Formato"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   3030
         TabIndex        =   14
         Top             =   4845
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   2760
      Top             =   5415
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "doc_repo_aviso_llegada"
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

    If Lst_bl.SelCount = 0 Then
        MsgBox "No existe ningun bl seleccionado", vbInformation
        Exit Sub
    End If
    Set oWord = CreateObject("Word.Application")
    Ruta = App.Path + "\Plantillas\"
'    If Cmb_formato.Text = "MSC" Then
'       doc_plantilla = Ruta + "PlantillaMSC.doc"
'    ElseIf Cmb_formato.Text = "MSC 53" Then
'       doc_plantilla = Ruta + "PlantillaMSC53.doc"
'    ElseIf Cmb_formato.Text = "KHS" Then
       doc_plantilla = Ruta + "PlantillaAviso.doc"
'    End If
    doc_origen = Ruta + "Origen Aviso.doc"
    
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

End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Command3_Click()
On Error GoTo fin
    Dim objWord As Object
    doc_plantilla = """" + App.Path + "\Plantillas\" + "PlantillaAviso.doc" + """"
    Set objWord = CreateObject("Word.Application")
    objWord.Documents.Open doc_plantilla
    objWord.Visible = True
    Set objWord = Nothing
    'x = Shell("WinWord " & doc_plantilla, vbMaximizedFocus)
fin:
End Sub

Private Sub Form_Load()
    carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
    If Not data_linea.Recordset.EOF Then data_linea.Recordset.MoveFirst
    'Cmb_formato.AddItem "KHS"
    Cmb_formato.AddItem "MSC"
    Cmb_formato.AddItem "MSC 53"
    Cmb_formato.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
        Set oWord = Nothing
        Set oCell = Nothing
        Set oMerge = Nothing
        Set sht_est = Nothing
        Set wb_est = Nothing
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
Dim reg_cont    As New ADODB.Recordset
Dim reg_buque   As New ADODB.Recordset
Dim reg_consig  As New ADODB.Recordset
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
    sql = "Select * " & _
               " from docu_buque_viaje " & _
               " Where buqv_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and buqv_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and buqv_viaje       ='" & Cmb_viaje.BoundText & "'"
    reg_buque.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    
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
              
       NomArchivo = Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text) & "\" & "Aviso " & Trim$(reg("BLCAB_CONO_EMB")) & ".doc"
       oWord.Documents(doc_origen).Activate
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=1) 'buque
       oCell.Range.Text = Trim(Cmb_buque.Text) & " V " & Trim(Cmb_viaje.Text)
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=2) 'fecha arribo
       oCell.Range.Text = Format(reg_buque("buqv_fecha_arribo"), "mmm dd/yyyy") & " APROX."
    
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=3) 'Bl
       oCell.Range.Text = reg("blcab_cono_emb")
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=4) 'cont
       sql = "Select count(*), cont_tam_equip from docu_cont, docu_bldet " & _
             " Where bldet_linea       ='" & Cmb_linea.BoundText & "'" & _
             "   and bldet_buque       ='" & Cmb_buque.BoundText & "'" & _
             "   and bldet_viaje       ='" & Cmb_viaje.BoundText & "'" & _
             "   and bldet_puer_emb    ='" & reg("blcab_puer_emb") & "'" & _
             "   and bldet_puer_des    ='" & reg("blcab_puer_des") & "'" & _
             "   and bldet_emp_resp    ='" & reg("blcab_emp_resp") & "'" & _
             "   and bldet_cono_emb    ='" & reg("blcab_cono_emb") & "'" & _
             "   and bldet_linea       = cont_linea  " & _
             "   and bldet_buque       = cont_buque  " & _
             "   and bldet_viaje       = cont_viaje  " & _
             "   and bldet_puer_emb    = cont_puer_emb " & _
             "   and bldet_puer_des    = cont_puer_des " & _
             "   and bldet_contenedor  = cont_contenedor "
       sql = sql & " group by cont_tam_equip"
       reg_cont.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                oCell.Range.Text = ""
                Do While Not reg_cont.EOF
                    tipo_cont_abrev = devuelve_desc("docu_tamano_Equip", "tama_equip_codigo", "tama_equip_abrev", reg_cont(1), "tama_equip_linea ='" & Cmb_linea.BoundText & "'")
                    If tipo_cont_abrev = "" Then tipo_cont_abrev = reg_cont(1)
                    If Asc(oCell.Range.Text) = 32 Or Asc(oCell.Range.Text) = 13 Or Trim(oCell.Range.Text) = "" Then
                        oCell.Range.Text = reg_cont(0) & " x " & tipo_cont_abrev
                    Else
                        oCell.Range.Text = oCell.Range.Text & "; " & reg_cont(0) & " x " & tipo_cont_abrev
                    End If
                    reg_cont.MoveNext
                Loop
                reg_cont.Close
       
       
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=5) 'shipper 1
       oCell.Range.Text = Trim(reg("blcab_embarcador"))
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=6) 'Consignee1
       oCell.Range.Text = Trim(reg("blcab_consig"))
       
       'telefono y fax
       sql = "select * from docu_consignatario where consig_codigo = '" & reg("blcab_ide_cons") & "'"
       reg_consig.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
       If Not reg_consig.EOF Then
            Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=7) 'telefono
            oCell.Range.Text = Trim(reg_consig("consig_fono"))
            
            Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=8) 'fax
            oCell.Range.Text = Trim(reg_consig("consig_fax"))
       End If
       reg_consig.Close
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=9) 'pto embarque
       oCell.Range.Text = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", reg("blcab_puer_emb"), " puer_linea ='" & Cmb_linea.BoundText & "'")
         
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=10) 'pto descarga
       oCell.Range.Text = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", reg("blcab_puer_des"), " puer_linea ='" & Cmb_linea.BoundText & "'")
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=11) 'pto destino final
       oCell.Range.Text = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", reg("blcab_pto_dest_final"), " puer_linea ='" & Cmb_linea.BoundText & "'")
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=12) 'flete prep
       oCell.Range.Text = "USD" & reg("blcab_flete")
            
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=14) 'peso
       oCell.Range.Text = reg("blcab_peso_tot")
       
       Set oCell = oWord.ActiveDocument.Tables(1).Cell(Row:=i, Column:=15) 'pto origen
       oCell.Range.Text = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", reg("blcab_puer_origen"), " puer_linea ='" & Cmb_linea.BoundText & "'")
     
       '========================== EMPIEZA LA IMPRESION ============================
       Call Print_Merge(NomArchivo, i - 1)
       reg.MoveNext: i = i + 1
    Loop
    
    oWord.Documents(doc_origen).Close savechanges:=wdDoNotSaveChanges
    oWord.Documents(doc_plantilla).Close savechanges:=wdDoNotSaveChanges
    oWord.Visible = True
    Exit Sub
    
CANCELAR:
    MsgBox msg + Chr(13) + Chr(13) + "Se cancela la Carga de Datos", vbExclamation, App.Title
    oWord.Documents(param_doc_origen).Close savechanges:=wdDoNotSaveChanges
    
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


Private Sub Txt_bl_KeyPress(KeyAscii As Integer)
Dim reg As New ADODB.Recordset

If KeyAscii = 13 Then
'If Opt_bl.Value Then
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
'End If
SendKeys "{TAB}"
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Public Sub txt_bl_LostFocus()
Dim reg As New ADODB.Recordset

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
End Sub
