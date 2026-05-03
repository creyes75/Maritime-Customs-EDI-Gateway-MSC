VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form doc_repo_expo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Bls de Exportación"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   6195
   Begin VB.PictureBox CReport 
      Height          =   480
      Left            =   165
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   25
      Top             =   5400
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   5160
      Left            =   60
      TabIndex        =   2
      Top             =   -30
      Width           =   6030
      Begin VB.ListBox Lst_bl 
         Enabled         =   0   'False
         Height          =   2085
         Left            =   1200
         Style           =   1  'Checkbox
         TabIndex        =   21
         Top             =   2595
         Width           =   3135
      End
      Begin VB.CheckBox Chk_todos 
         Caption         =   "Por Bls"
         Height          =   255
         Left            =   225
         TabIndex        =   20
         Top             =   2235
         Width           =   2895
      End
      Begin VB.Frame Fra_pto_descarga 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   1665
         TabIndex        =   14
         Top             =   1530
         Width           =   1590
         Begin VB.TextBox Txt_pto_descarga 
            Height          =   285
            Left            =   120
            MaxLength       =   5
            TabIndex        =   16
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
            Picture         =   "doc_repo_expo.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   15
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
            TabIndex        =   17
            Top             =   15
            Width           =   1200
         End
      End
      Begin VB.Frame Fra_viaje 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   3255
         TabIndex        =   11
         Top             =   840
         Width           =   2145
         Begin MSDataListLib.DataCombo Cmb_viaje 
            Height          =   315
            Left            =   135
            TabIndex        =   12
            Top             =   225
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
            TabIndex        =   13
            Top             =   0
            Width           =   345
         End
      End
      Begin VB.Frame Fra_buque 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   135
         TabIndex        =   8
         Top             =   840
         Width           =   3090
         Begin MSDataListLib.DataCombo Cmb_buque 
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Top             =   225
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
            TabIndex        =   10
            Top             =   0
            Width           =   465
         End
      End
      Begin VB.Frame Fra_linea 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   120
         TabIndex        =   5
         Top             =   165
         Width           =   2865
         Begin MSDataListLib.DataCombo Cmb_linea 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   225
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
            TabIndex        =   7
            Top             =   0
            Width           =   825
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   105
         TabIndex        =   3
         Top             =   1560
         Width           =   1500
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
            Left            =   1170
            Picture         =   "doc_repo_expo.frx":0532
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Buscar"
            Top             =   210
            Width           =   300
         End
         Begin VB.TextBox Txt_pto_emb 
            Height          =   285
            Left            =   150
            MaxLength       =   5
            TabIndex        =   23
            Top             =   225
            Width           =   975
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "Puerto Embarque"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   105
            TabIndex        =   4
            Top             =   -15
            Width           =   1230
         End
      End
      Begin MSDataListLib.DataCombo Cmb_oper_carga 
         Height          =   315
         Left            =   3375
         TabIndex        =   18
         Top             =   1755
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
         Left            =   4800
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bill of Lading"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   2610
         Width           =   900
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Agen. Oper. de Carg"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3405
         TabIndex        =   19
         Top             =   1515
         Width           =   1455
      End
   End
   Begin Threed.SSCommand cmd_Proceso 
      Height          =   645
      Left            =   4650
      TabIndex        =   0
      Top             =   5145
      Width           =   645
      _Version        =   65536
      _ExtentX        =   1138
      _ExtentY        =   1138
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
      Picture         =   "doc_repo_expo.frx":0A64
   End
   Begin Threed.SSCommand cmd_salir 
      Height          =   645
      Left            =   5385
      TabIndex        =   1
      Top             =   5145
      Width           =   645
      _Version        =   65536
      _ExtentX        =   1138
      _ExtentY        =   1138
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
      Picture         =   "doc_repo_expo.frx":0D7E
   End
End
Attribute VB_Name = "doc_repo_expo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VARIABLES PARA EXCEL
Public wb_est  As Object 'Excel.Workbook
Public sht_est As Excel.Worksheet
Public rng_est As Excel.Range
Dim bls As String

Private Sub Chk_todos_Click()
If Chk_todos.Value = 1 Then
    Me.Lst_bl.Enabled = True
Else
    Me.Lst_bl.Enabled = False
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

carga_parametros Cmb_linea.BoundText

Txt_pto_emb.Text = Gstr_param.pto_descarga
Cmb_oper_carga.BoundText = Gstr_param.naviera
End Sub

Private Sub Cmb_pto_emb_Change()
Dim reg As New ADODB.Recordset

    sql = "Select blcab_cono_emb " & _
               " from docu_blcab " & _
               " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'" & _
               "   and blcab_puer_emb    ='" & Cmb_pto_emb.BoundText & "'" & _
               "   and blcab_puer_des    ='" & Txt_pto_descarga.Text & "'" & _
               "   and blcab_emp_resp    ='" & Cmb_oper_carga.BoundText & "'"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Lst_bl.Clear
    Do While Not reg.EOF
        Lst_bl.AddItem reg("blcab_cono_emb")
        reg.MoveNext
    Loop
    reg.Close

End Sub

Private Sub Cmb_viaje_Change()
'   Dim sql As String
'   sql = "Select * from docu_manifiesto " & _
'         " Where mani_linea ='" & Cmb_linea.BoundText & "'" & _
'         "   And mani_buque ='" & Cmb_buque.BoundText & "'" & _
'         "   And mani_viaje ='" & Cmb_viaje.BoundText & "'"
'         '"   And buqv_status = 'A'"
'
'   carga_datacombo data_pto_emb, Cmb_pto_emb, sql, "mani_puer_emb", "mani_puer_emb"

End Sub

Private Sub Cmd_hlp_pto_descarga_Click()
   Txt_pto_descarga.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Descarga")
End Sub

Private Sub Cmd_hlp_pto_emb_Click()
   Txt_pto_emb.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Embarque")
End Sub

Private Sub cmd_Proceso_Click()
On Error GoTo Errores
    Dim sql2 As String
    
    
    Set wb_est = CreateObject("excel.Application")
    'wb_est.Application.Caption = "DISCHARGE LIST"
    DoEvents
    Me.MousePointer = 11: DoEvents
    If Chk_todos.Value Then
        bls = ""
        For i = 0 To Lst_bl.ListCount - 1
            If Lst_bl.Selected(i) Then
               Lst_bl.ListIndex = i
               bls = bls & "'" & Lst_bl.Text & "', "
            End If
        Next i
        bls = Mid(bls, 1, Len(Trim(bls)) - 1) ' se quita la ultima coma
    End If

    Call LISTADO_EXCEL
    Me.MousePointer = 0: DoEvents
    Exit Sub
    'este codigo si vale solo que debido a incompatibilidad con Win2000 se tuvo que llevar a EXCEL
'       sql2 = " {docu_blcab.blcab_linea} ='" & Cmb_linea.BoundText & "'" & _
'              " and {docu_blcab.blcab_buque} ='" & Cmb_buque.BoundText & "'" & _
'              " and {docu_blcab.blcab_viaje} ='" & Cmb_viaje.BoundText & "'" & _
'              " and {docu_blcab.blcab_puer_emb} ='" & Cmb_pto_emb.BoundText & "'" & _
'              " and {docu_blcab.blcab_puer_des} ='" & Txt_pto_descarga.Text & "'" & _
'              " and {docu_blcab.blcab_emp_resp} ='" & Cmb_oper_carga.BoundText & "'"
'    If Chk_todos.Value Then
'        bls = ""
'        sql2 = sql2 & " and ("
'        For i = 0 To Lst_bl.ListCount - 1
'            If Lst_bl.Selected(i) Then
'               Lst_bl.ListIndex = i
'               sql2 = sql2 & " {docu_blcab.blcab_cono_emb} ='" & Lst_bl.Text & "' or"
'               'bls = bls & "'" & Lst_bl.Text & "', "
'            End If
'        Next i
'        'bls = Mid(bls, 1, Len(Trim(bls)) - 1) ' se quita la ultima coma
'        sql2 = Mid(sql2, 1, Len(Trim(sql2)) - 2) ' se quita la ultima coma
'        'sql2 = sql2 & " and {docu_blcab.blcab_cono_emb} in (" & bls & ")"
'        sql2 = sql2 & ")"
'    End If
'
'    CReport.ReportFileName = Gstr_ruta_reportes & "cargo manifiest.rpt"
'    CReport.Connect = Gstr_conex_docu
'    CReport.SelectionFormula = sql2
'    CReport.Action = 1
    Exit Sub
    
Errores:
    MsgBox "Ocurrio el error " & Err.Description
    Exit Sub
    Resume 0
End Sub

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
    If Not data_linea.Recordset.EOF Then data_linea.Recordset.MoveFirst
End Sub

Public Function LISTADO_EXCEL()
Dim REG_det As New ADODB.Recordset
Dim reg As New ADODB.Recordset
Dim i As Integer
Dim linea As Integer
On Error GoTo errores_Excel
      
      
      wb_est.Workbooks.Add ' aumenta un libro dentro de la aplicacion excel
      pagina% = 1
         Set sht_est = wb_est.ActiveWorkbook.Sheets(1)   'setea la hoja actual
         'parametros generales de la pagina (Sheet)
         sht_est.Name = "EXPORT"         'asigna un nombre a la hoja
         sht_est.PageSetup.Orientation = xlLandscape
         sht_est.PageSetup.PaperSize = xlPaperA4
         sht_est.PageSetup.BottomMargin = 20
         sht_est.PageSetup.TopMargin = 20
         sht_est.PageSetup.FooterMargin = 10
         sht_est.PageSetup.HeaderMargin = 10
         sht_est.PageSetup.Zoom = 75
         sht_est.Activate
         
   sql = "SELECT  * FROM DOCU_BUQUE_VIAJE, docu_buque " & _
         " Where buqv_linea       ='" & Cmb_linea.BoundText & "'" & _
         "   and buqv_buque       ='" & Cmb_buque.BoundText & "'" & _
         "   and buqv_viaje       ='" & Cmb_viaje.BoundText & "'" & _
         "   and buqv_linea       = buqu_linea " & _
         "   and buqv_buque       = buqu_codigo"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
      
   sht_est.Cells(2, 3).Value = "EXPORT"
   sht_est.Cells(4, 3).Value = "CLASS AND NAME VESSEL"
   'sht_est.Cells(4, 5).Value = "PORT OF ISSUE"
   sht_est.Cells(4, 7).Value = "VOYAGE_NO."
   sht_est.Cells(4, 9).Value = "ARRIVED"
   sht_est.Cells(4, 10).Value = "SAILED"
   
   sht_est.Cells(5, 3).Value = Cmb_buque.Text
   'sht_est.Cells(5, 5).Value = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", Cmb_pto_emb.BoundText, " puer_linea ='" & Cmb_linea.BoundText & "'")
   sht_est.Cells(5, 7).Value = Cmb_viaje.Text
   
   sht_est.Cells(6, 3).Value = "NACIONALITY"
   sht_est.Cells(6, 4).Value = "NAME OF MASTER"
   'sht_est.Cells(6, 5).Value = "LOADING PORT"
   'sht_est.Cells(6, 6).Value = "DISCHARGE PORT"
   'sht_est.Cells(6, 7).Value = "FINAL DESTINATION"
   'sht_est.Cells(6, 9).Value = "SAILING DATE LOAD PORT"
   
   sht_est.Cells(7, 3).Value = reg("buqu_nacio")
   'sht_est.Cells(7, 5).Value = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", reg("mani_puer_ini"), " puer_linea ='" & Cmb_linea.BoundText & "'")
   'sht_est.Cells(7, 6).Value = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", reg("mani_puer_des"), " puer_linea ='" & Cmb_linea.BoundText & "'")
   'sht_est.Cells(7, 7).Value = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", reg("mani_puer_des"), " puer_linea ='" & Cmb_linea.BoundText & "'")
   'sht_est.Cells(7, 9).Value = CDate(reg("mani_fecha_zarp_ini"))
   
   sht_est.Cells(8, 3).Value = "DISCHARGE PORT":
   sht_est.Cells(8, 4).Value = "SHIPPER/ADRESS":
   sht_est.Cells(8, 5).Value = "CONSIGNEES/ADRESS":
   sht_est.Cells(8, 6).Value = "BILL OF LADING":
   sht_est.Cells(8, 7).Value = "NUMBER AND MARKS":
   sht_est.Cells(8, 8).Value = "TYPE/":
   sht_est.Cells(8, 9).Value = "PACKAGES":
   sht_est.Cells(8, 10).Value = "QUANTITY AND DESCRIPTION OF GOODS":
   'sht_est.Cells(8, 11).Value = "CHARGE"
   sht_est.Cells(8, 12).Value = "GROSS WEIGHT"
   sht_est.Cells(8, 13).Value = "VOLUMEN"
   
   sht_est.Cells(9, 5).Value = "NOTIFY PARTY / ADRESS"
   sht_est.Cells(9, 8).Value = "COND CONT"
   'fuente de encabezados
   Call Excel_setea_fuente_fila(sht_est, 2, 14, "moderm", True, True, 8000000)
   Call Excel_setea_fuente_fila(sht_est, 4, 8, "moderm", True, False)
   Call Excel_setea_fuente_fila(sht_est, 6, 8, "moderm", True, False)
   Call Excel_setea_fuente_fila(sht_est, 8, 8, "moderm", True, False)
   Call Excel_setea_fuente_fila(sht_est, 9, 8, "moderm", True, False)
   'fuente de texto
   Call Excel_setea_fuente_fila(sht_est, 5, 8) '"moderm",  False, False)
   Call Excel_setea_fuente_fila(sht_est, 7, 8) ', "moderm", True, True)
   Call Excel_setea_fuente_fila(sht_est, 9, 8) ', "moderm", True, True)
   Call Excel_setea_fuente_fila(sht_est, 10, 8) ', "moderm", True, True)
               
   DoEvents
   'ancho de las columnas
   Call Excel_setea_ancho_columna(sht_est, 4, 1, 4, 2, 1)
   Call Excel_setea_ancho_columna(sht_est, 4, 3, 4, 3, 12)
   Call Excel_setea_ancho_columna(sht_est, 4, 4, 4, 5, 25)
   Call Excel_setea_ancho_columna(sht_est, 4, 6, 4, 6, 15)
   Call Excel_setea_ancho_columna(sht_est, 4, 7, 4, 8, 15)
   Call Excel_setea_ancho_columna(sht_est, 4, 9, 4, 9, 8)
   Call Excel_setea_ancho_columna(sht_est, 4, 10, 4, 10, 30)
   Call Excel_setea_ancho_columna(sht_est, 4, 11, 4, 13, 10)
   'Set rng_est = sht_est.Range(sht_est.Cells(4, 10), sht_est.Cells(4, 12)): rng_est.ColumnWidth = 10
   
   'seteo de bordes
   Call Excel_setea_borde(sht_est, 4, 3, 5, 4, 1, False)
   Call Excel_setea_borde(sht_est, 4, 5, 5, 6, 1, False)
   Call Excel_setea_borde(sht_est, 4, 7, 5, 8, 1, False)
   Call Excel_setea_borde(sht_est, 4, 9, 5, 9, 1, False)
   Call Excel_setea_borde(sht_est, 4, 10, 5, 12, 1, False)
   Call Excel_setea_borde(sht_est, 6, 3, 7, 3, 1, False)
   Call Excel_setea_borde(sht_est, 6, 4, 7, 4, 1, False)
   Call Excel_setea_borde(sht_est, 6, 5, 7, 5, 1, False)
   Call Excel_setea_borde(sht_est, 6, 6, 7, 6, 1, False)
   Call Excel_setea_borde(sht_est, 6, 7, 7, 8, 1, False)
   Call Excel_setea_borde(sht_est, 6, 9, 7, 12, 1, False)
   Call Excel_setea_borde(sht_est, 8, 3, 9, 3, 1, False)
   Call Excel_setea_borde(sht_est, 8, 4, 9, 4, 1, False)
   Call Excel_setea_borde(sht_est, 8, 5, 9, 5, 1, False)
   Call Excel_setea_borde(sht_est, 8, 6, 9, 6, 1, False)
   Call Excel_setea_borde(sht_est, 8, 7, 9, 7, 1, False)
   Call Excel_setea_borde(sht_est, 8, 8, 9, 8, 1, False)
   Call Excel_setea_borde(sht_est, 8, 9, 9, 9, 1, False)
   Call Excel_setea_borde(sht_est, 8, 10, 9, 10, 1, False)
   Call Excel_setea_borde(sht_est, 8, 11, 9, 11, 1, False)
   Call Excel_setea_borde(sht_est, 8, 12, 9, 12, 1, False)
   Call Excel_setea_borde(sht_est, 8, 13, 9, 13, 1, False)
   'Set rng_est = sht_est.Range(sht_est.Cells(4, 3), sht_est.Cells(5, 4)):   rng_est.Borders.LineStyle = 1: rng_est.Borders(xlInsideVertical).LineStyle = xlNone:  rng_est.Borders(xlInsideHorizontal).LineStyle = xlNone
'-------------------------------------------------------------------------------------------------------------
' inicio del proceso de lectura de datos y llenado en excel
       sql = "Select docu_blcab_exp.*, docu_bldet_exp.*, tama_equip_descripcion, cond_cont_descripcion " & _
               " from docu_blcab_exp, docu_bldet_exp, docu_cont_exp, docu_tamano_equip, docu_cond_contenedor " & _
               " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'" & _
               "   and blcab_puer_emb    ='" & Txt_pto_emb.Text & "'"
       If Chk_todos.Value Then
            sql = sql & " and blcab_cono_emb  in (" & bls & ")"
       End If
       sql = sql & " and blcab_linea       = bldet_linea" & _
               "   and blcab_buque       = bldet_buque" & _
               "   and blcab_viaje       = bldet_viaje" & _
               "   and blcab_puer_emb    = bldet_puer_emb" & _
               "   and blcab_puer_des    = bldet_puer_des" & _
               "   and blcab_emp_resp    = bldet_emp_resp" & _
               "   and blcab_cono_emb    = bldet_cono_emb" & _
               "   and cont_linea       = tama_equip_linea " & _
               "   and cont_tam_equip   = tama_equip_codigo" & _
               "   and cont_linea       = cond_cont_linea    " & _
               "   and cont_condicion   = cond_cont_codigo   " & _
               "   and bldet_linea      = cont_linea " & _
               "   and bldet_buque      = cont_buque " & _
               "   and bldet_viaje      = cont_viaje " & _
               "   and bldet_puer_emb   = cont_puer_emb " & _
               "   and bldet_puer_des   = cont_puer_des " & _
               "   and bldet_contenedor = cont_contenedor "
               '& _
               " Order by blcab_linea,    blcab_buque,    blcab_viaje," & _
               "          blcab_puer_des, blcab_cono_emb, bldet_contenedor"
        sql = sql & " UNION Select docu_blcab_exp.*, docu_bldet_exp.*, '' as tama_equip_descripcion, '' as cond_cont_descripcion " & _
               " from docu_blcab_exp, docu_bldet_exp " & _
               " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'" & _
               "   and blcab_puer_emb    ='" & Txt_pto_emb.Text & "'"
       If Chk_todos.Value Then
            sql = sql & " and blcab_cono_emb  in (" & bls & ")"
       End If
       sql = sql & " and blcab_linea       = bldet_linea" & _
               "   and blcab_buque       = bldet_buque" & _
               "   and blcab_viaje       = bldet_viaje" & _
               "   and blcab_puer_emb    = bldet_puer_emb" & _
               "   and blcab_puer_des    = bldet_puer_des" & _
               "   and blcab_emp_resp    = bldet_emp_resp" & _
               "   and blcab_cono_emb    = bldet_cono_emb" & _
               "   and bldet_contenedor  = ''" & _
               " Order by blcab_linea,    blcab_buque,    blcab_viaje," & _
               "          blcab_puer_des, blcab_cono_emb, bldet_contenedor"

       REG_det.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
           
       linea = 11
       bl = ""
       pto = ""
       Do While Not REG_det.EOF
            If bl <> REG_det("blcab_cono_emb") Or pto <> REG_det("blcab_puer_des") Then
                Set rng_est = sht_est.Range(sht_est.Cells(linea, 3), sht_est.Cells(linea, 12)): rng_est.Borders(xlDiagonalDown).LineStyle = xlNone: rng_est.Borders(xlEdgeBottom).LineStyle = xlContinuous
                linea = linea + 1:      Call Excel_setea_fuente_fila(sht_est, linea, 7)
                sht_est.Cells(linea, 3).Value = REG_det("blcab_puer_des")
                sht_est.Cells(linea, 4).Value = REG_det("blcab_embarcador") & Chr(10) & " / " & REG_det("blcab_direc_embarc")
                sht_est.Cells(linea, 5).Value = REG_det("blcab_consig") & Chr(10) & " / " & REG_det("blcab_direc_consig")
                sht_est.Cells(linea, 6).Value = REG_det("blcab_cono_emb")
                'sht_est.Cells(linea, 10).Value = REG_det("bldet_desc_carg") ' 16/enero/2007
                'sht_est.Cells(linea, 11).Value = CDbl(REG_det("blcab_flete"))
                'sht_est.Cells(linea, 12).Value = CDbl(REG_det("bldet_peso")) ' 16/ene/2007
                'sht_est.Cells(linea, 13).Value = CDbl(REG_det("bldet_volumen")) ' 16/ene/2007
                linea = linea + 1:      Call Excel_setea_fuente_fila(sht_est, linea, 7)
                sht_est.Cells(linea, 5).Value = REG_det("blcab_notif") & Chr(10) & " / " & REG_det("blcab_direc_notif")
                linea = linea + 1
                pto = REG_det("blcab_puer_des") 'ojo esta asignacion no esta en el report de import posible error de quiebre
                bl = REG_det("blcab_cono_emb")
            End If
            Call Excel_setea_fuente_fila(sht_est, linea, 6)
            sht_est.Cells(linea, 7).Value = Trim(REG_det("bldet_contenedor")) & "    " & Trim(REG_det("bldet_sello"))
            sht_est.Cells(linea, 8).Value = REG_det("tama_equip_descripcion") & Chr(10) & REG_det("cond_cont_descripcion")
            sht_est.Cells(linea, 9).Value = CDbl(REG_det("bldet_bultos")) & " " & REG_det("bldet_embalaje")
            sht_est.Cells(linea, 10).Value = CStr(REG_det("bldet_desc_carg"))     ' 16/ene/2007
            sht_est.Cells(linea, 12).Value = CDbl(REG_det("bldet_peso"))    ' 16/ene/2007
            sht_est.Cells(linea, 13).Value = CDbl(REG_det("bldet_volumen")) ' 16/ene/2007
            REG_det.MoveNext
            linea = linea + 1
       Loop
        'esto hace que el texto que no entra en una celda se acomode aumentando el espacio de la misma
        Set rng_est = sht_est.Range(sht_est.Cells(11, 3), sht_est.Cells(linea + 1, 10))
        rng_est.HorizontalAlignment = xlGeneral
        rng_est.VerticalAlignment = xlBottom
        rng_est.WrapText = True
        rng_est.Orientation = 0
       
'*************************************************************

      'If Dir(Grabar.filename) <> "" Then
       If Dir(Gstr_Ruta_BlWord, vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord)
       If Dir(Gstr_Ruta_BlWord & Trim(Cmb_buque.Text), vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord & Trim(Cmb_buque.Text))
       If Dir(Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text), vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text))
       
       NomArchivo = Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text) & "\" & "Export Cargo.XLS"
       If Dir(NomArchivo, vbArchive) <> "" Then Kill (NomArchivo)
       
       wb_est.ActiveWorkbook.SaveAs FileName:= _
        NomArchivo, FileFormat:= _
        xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
        , CreateBackup:=False
       
       'wb_est.SaveAs Trim$(NomArchivo)
       wb_est.Application.Visible = True ' visible la aplicación
       Set wb_est = Nothing
Exit Function
errores_Excel:
    MsgBox "Error " & Err.Description
    Exit Function
    Resume 0
End Function

Private Sub Txt_pto_descarga_LostFocus()
Dim reg As New ADODB.Recordset

    sql = "Select blcab_cono_emb " & _
               " from docu_blcab_exp " & _
               " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'" & _
               "   and blcab_puer_emb    ='" & Txt_pto_emb.Text & "'" & _
               "   and blcab_puer_des    ='" & Txt_pto_descarga.Text & "'" & _
               "   and blcab_emp_resp    ='" & Cmb_oper_carga.BoundText & "'"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Lst_bl.Clear
    Do While Not reg.EOF
        Lst_bl.AddItem reg("blcab_cono_emb")
        reg.MoveNext
    Loop
    reg.Close
End Sub
