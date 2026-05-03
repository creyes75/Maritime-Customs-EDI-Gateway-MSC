VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form doc_repo_mani 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Manifiestos"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   6195
   Begin VB.PictureBox CReport 
      Height          =   480
      Left            =   165
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   26
      Top             =   5400
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   4920
      Left            =   60
      TabIndex        =   2
      Top             =   -30
      Width           =   6030
      Begin VB.CheckBox Chk_envio_pto 
         Caption         =   " Por Puerto "
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Frame Fra_PUERTO 
         Enabled         =   0   'False
         Height          =   3375
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   5775
         Begin VB.Frame Frame7 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   540
            Left            =   120
            TabIndex        =   19
            Top             =   285
            Width           =   1500
            Begin MSDataListLib.DataCombo Cmb_pto_emb 
               Height          =   315
               Left            =   135
               TabIndex        =   20
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
               TabIndex        =   21
               Top             =   -15
               Width           =   1230
            End
         End
         Begin VB.Frame Fra_pto_descarga 
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   1680
            TabIndex        =   15
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
               Picture         =   "doc_repo_mani.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   17
               ToolTipText     =   "Buscar"
               Top             =   255
               Width           =   300
            End
            Begin VB.TextBox Txt_pto_descarga 
               Height          =   285
               Left            =   120
               MaxLength       =   5
               TabIndex        =   16
               Top             =   270
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Puerto Descarga"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   18
               Top             =   15
               Width           =   1200
            End
         End
         Begin VB.CheckBox Chk_todos 
            Caption         =   "Manifiesto por Bls"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   960
            Width           =   2895
         End
         Begin VB.ListBox Lst_bl 
            Enabled         =   0   'False
            Height          =   1860
            Left            =   1215
            Style           =   1  'Checkbox
            TabIndex        =   13
            Top             =   1320
            Width           =   3135
         End
         Begin MSDataListLib.DataCombo Cmb_oper_carga 
            Height          =   315
            Left            =   3390
            TabIndex        =   22
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
            Left            =   4815
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
            Left            =   3420
            TabIndex        =   24
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Bill of Lading"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   255
            TabIndex        =   23
            Top             =   1335
            Width           =   900
         End
      End
      Begin VB.Frame Fra_viaje 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   3255
         TabIndex        =   9
         Top             =   840
         Width           =   2145
         Begin MSDataListLib.DataCombo Cmb_viaje 
            Height          =   315
            Left            =   120
            TabIndex        =   10
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
            TabIndex        =   11
            Top             =   0
            Width           =   345
         End
      End
      Begin VB.Frame Fra_buque 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   135
         TabIndex        =   6
         Top             =   840
         Width           =   3090
         Begin MSDataListLib.DataCombo Cmb_buque 
            Height          =   315
            Left            =   120
            TabIndex        =   7
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
            TabIndex        =   8
            Top             =   0
            Width           =   465
         End
      End
      Begin VB.Frame Fra_linea 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   120
         TabIndex        =   3
         Top             =   165
         Width           =   2865
         Begin MSDataListLib.DataCombo Cmb_linea 
            Height          =   315
            Left            =   120
            TabIndex        =   4
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
            TabIndex        =   5
            Top             =   0
            Width           =   825
         End
      End
   End
   Begin Threed.SSCommand cmd_Proceso 
      Height          =   645
      Left            =   4635
      TabIndex        =   0
      Top             =   4905
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
      Picture         =   "doc_repo_mani.frx":0532
   End
   Begin Threed.SSCommand cmd_salir 
      Height          =   645
      Left            =   5385
      TabIndex        =   1
      Top             =   4905
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
      Picture         =   "doc_repo_mani.frx":084C
   End
End
Attribute VB_Name = "doc_repo_mani"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VARIABLES PARA EXCEL
Public wb_est  As Object 'Excel.Workbook
Public sht_est As Excel.Worksheet
Public rng_est As Excel.Range
Dim bls As String

Private Sub Chk_envio_pto_Click()
If Chk_envio_pto = 1 Then
    Fra_puerto.Enabled = True
Else
    Fra_puerto.Enabled = False
End If
End Sub

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

Txt_pto_descarga.Text = Gstr_param.pto_descarga
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
   Dim sql As String
   sql = "Select * from docu_manifiesto " & _
         " Where mani_linea ='" & Cmb_linea.BoundText & "'" & _
         "   And mani_buque ='" & Cmb_buque.BoundText & "'" & _
         "   And mani_viaje ='" & Cmb_viaje.BoundText & "'"
         '"   And buqv_status = 'A'"
   
   carga_datacombo data_pto_emb, Cmb_pto_emb, sql, "mani_puer_emb", "mani_puer_emb"

End Sub

Private Sub Cmd_hlp_pto_descarga_Click()
   Txt_pto_descarga.Text = Ayuda("docu_puerto", "puer_codigo", "puer_descripcion", " puer_linea ='" & Cmb_linea.BoundText & "' and puer_status ='A' ", "Puerto de Descarga")
End Sub

Private Sub cmd_Proceso_Click()
On Error GoTo Errores
    Dim sql2 As String
    
    
    Set wb_est = CreateObject("excel.Application")
    'wb_est.Application.Caption = "DISCHARGE LIST"
    DoEvents
    Me.MousePointer = 11: DoEvents
    If Chk_envio_pto Then
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
      
      Me.MousePointer = 11
      DoEvents
      wb_est.Workbooks.Add ' aumenta un libro dentro de la aplicacion excel
      pagina% = 1
         Set sht_est = wb_est.ActiveWorkbook.Sheets(1)   'setea la hoja actual
         'parametros generales de la pagina (Sheet)
         sht_est.Name = "CARGO MANIFIEST"         'asigna un nombre a la hoja
         sht_est.PageSetup.Orientation = xlLandscape
         sht_est.PageSetup.PaperSize = xlPaperA4
         sht_est.PageSetup.BottomMargin = 20
         sht_est.PageSetup.TopMargin = 20
         sht_est.PageSetup.FooterMargin = 10
         sht_est.PageSetup.HeaderMargin = 10
         sht_est.PageSetup.Zoom = 75
         sht_est.Activate
         
   sql = "SELECT  * FROM DOCU_BUQUE_VIAJE, DOCU_MANIFIESTO, docu_buque " & _
         " Where buqv_linea       ='" & Cmb_linea.BoundText & "'" & _
         "   and buqv_buque       ='" & Cmb_buque.BoundText & "'" & _
         "   and buqv_viaje       ='" & Cmb_viaje.BoundText & "'" & _
         "   and buqv_linea       = mani_linea " & _
         "   and buqv_buque       = mani_buque " & _
         "   and buqv_viaje       = mani_viaje "
    If Chk_envio_pto.Value = 1 Then
         sql = sql + "   and mani_puer_emb    ='" & Cmb_pto_emb.BoundText & "'"
    End If
    sql = sql & _
         "   and buqv_linea       = buqu_linea " & _
         "   and buqv_buque       = buqu_codigo"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
      
   If reg.EOF Then
        MsgBox "No existen datos para este buque viaje", vbInformation
        GoTo fin
   End If
   sht_est.Cells(2, 6).Value = "CARGO MANIFEST"
   
   sht_est.Cells(3, 1).Value = "1.MSC GUAYAQUIL"
   sht_est.Cells(4, 1).Value = "AVE FCO DE ORELLANA Y ALBERTO BORGES"
   sht_est.Cells(5, 1).Value = "GUAYAQUIL"
   sht_est.Cells(6, 1).Value = "mscecuador@mscecuador.com"
   
   sht_est.Cells(4, 4).Value = "2.VESSEL:":             sht_est.Cells(4, 4).Font.Bold = True:
   sht_est.Cells(4, 8).Value = "3.VOYAGE:":             sht_est.Cells(4, 8).Font.Bold = True:
   sht_est.Cells(4, 10).Value = "4.SAILING DATE:":      sht_est.Cells(4, 10).Font.Bold = True:
   
   sht_est.Cells(5, 4).Value = "5.NATIONALITY":         sht_est.Cells(5, 4).Font.Bold = True:
   sht_est.Cells(5, 6).Value = "6.ORIGIN":              sht_est.Cells(5, 6).Font.Bold = True:
   sht_est.Cells(5, 8).Value = "7.PORT OF LOADING":     sht_est.Cells(5, 8).Font.Bold = True:
   sht_est.Cells(5, 10).Value = "8.PORT OF DISCHARGE":  sht_est.Cells(5, 10).Font.Bold = True:
   sht_est.Cells(5, 12).Value = "9.FINAL DESTINATION":  sht_est.Cells(5, 12).Font.Bold = True:
   
   sht_est.Cells(7, 1).Value = "+593 4269 3292":        sht_est.Cells(7, 1).Font.Bold = True:
   sht_est.Cells(7, 3).Value = "13.B/L NO":             sht_est.Cells(7, 3).Font.Bold = True:
   sht_est.Cells(7, 5).Value = "14.MARKS AND NOS":      sht_est.Cells(7, 5).Font.Bold = True:
   sht_est.Cells(7, 8).Value = "17.NO. OF CNTRS":       sht_est.Cells(7, 8).Font.Bold = True:
   sht_est.Cells(7, 11).Value = "20.NET":               sht_est.Cells(7, 11).Font.Bold = True:
   sht_est.Cells(7, 12).Value = "23.B.REF":             sht_est.Cells(7, 12).Font.Bold = True:
   sht_est.Cells(7, 13).Value = "24.REMARKS":           sht_est.Cells(7, 13).Font.Bold = True:
   
   sht_est.Cells(8, 1).Value = "+593 4269 3292":        sht_est.Cells(8, 1).Font.Bold = True:
   sht_est.Cells(8, 5).Value = "15.CNTR NOS/TYPE":      sht_est.Cells(8, 5).Font.Bold = True:
   sht_est.Cells(8, 8).Value = "18.DESCRIPTION OF GOODS": sht_est.Cells(8, 8).Font.Bold = True:
   sht_est.Cells(8, 11).Value = "21.GROSS":             sht_est.Cells(8, 11).Font.Bold = True:
   
   sht_est.Cells(9, 1).Value = "12.NOTIFY(N)":          sht_est.Cells(9, 1).Font.Bold = True:
   sht_est.Cells(9, 5).Value = "16.SEAL":               sht_est.Cells(9, 5).Font.Bold = True:
   sht_est.Cells(9, 8).Value = "19.SHIPPING TERMS":     sht_est.Cells(9, 8).Font.Bold = True:
   sht_est.Cells(9, 11).Value = "22.TARE":              sht_est.Cells(9, 11).Font.Bold = True:
      
   sht_est.Cells(4, 5).Value = Cmb_buque.Text
   sht_est.Cells(4, 9).Value = Cmb_viaje.Text
   
   sht_est.Cells(4, 12).Value = CDate(reg("mani_fecha_zarp_ini"))
   If Chk_envio_pto.Value = 1 Then
    sht_est.Cells(6, 8).Value = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", reg("mani_puer_des"), " puer_linea ='" & Cmb_linea.BoundText & "'")
    sht_est.Cells(6, 10).Value = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", Cmb_pto_emb.BoundText, " puer_linea ='" & Cmb_linea.BoundText & "'")
    sht_est.Cells(6, 12).Value = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", reg("mani_puer_des"), " puer_linea ='" & Cmb_linea.BoundText & "'")
   End If
   
   
   'fuente de encabezados
'   sht_est.Cells(4, 3).Font.Bold = True:   sht_est.Cells(4, 3).Font.color = 8000000
   
   DoEvents
   'ancho de las columnas
   sht_est.Range("A1", "N9").Font.Name = "Courier New"
   sht_est.Range("A1", "N9").Font.Size = 8
   sht_est.Range("A10", "N600").Font.Name = "Arial"
   sht_est.Range("A10", "N600").Font.Size = 8
   Call Excel_setea_ancho_columna(sht_est, 4, 1, 4, 1, 10.17)
   Call Excel_setea_ancho_columna(sht_est, 4, 2, 4, 13, 8.33)
   Call Excel_setea_ancho_columna(sht_est, 4, 14, 4, 14, 9.17)
   
   'seteo de bordes
   Call Excel_setea_borde(sht_est, 4, 4, 4, 14, 1, False)
   Call Excel_setea_borde(sht_est, 5, 4, 6, 5, 1, False)
   Call Excel_setea_borde(sht_est, 5, 6, 6, 7, 1, False)
   Call Excel_setea_borde(sht_est, 5, 8, 6, 9, 1, False)
   Call Excel_setea_borde(sht_est, 5, 10, 6, 11, 1, False)
   Call Excel_setea_borde(sht_est, 5, 12, 6, 14, 1, False)
   
   Call Excel_setea_borde(sht_est, 7, 1, 9, 2, 1, False)
   Call Excel_setea_borde(sht_est, 7, 3, 9, 4, 1, False)
   Call Excel_setea_borde(sht_est, 7, 5, 9, 7, 1, False)
   Call Excel_setea_borde(sht_est, 7, 8, 9, 10, 1, False)
   Call Excel_setea_borde(sht_est, 7, 11, 9, 11, 1, False)
   Call Excel_setea_borde(sht_est, 7, 12, 9, 12, 1, False)
   Call Excel_setea_borde(sht_est, 7, 13, 9, 14, 1, False)
   
   'Set rng_est = sht_est.Range(sht_est.Cells(4, 3), sht_est.Cells(5, 4)):   rng_est.Borders.LineStyle = 1: rng_est.Borders(xlInsideVertical).LineStyle = xlNone:  rng_est.Borders(xlInsideHorizontal).LineStyle = xlNone
'-------------------------------------------------------------------------------------------------------------
' inicio del proceso de lectura de datos y llenado en excel
       sql = "Select docu_blcab.*, docu_bldet.*, tama_equip_abrev as tama_equip_descripcion, cond_cont_descripcion, cont_sello1, cont_tara " & _
               " from docu_blcab, docu_bldet, docu_cont, docu_tamano_equip, docu_cond_contenedor " & _
               " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'"
       If Chk_envio_pto.Value = 1 Then
          sql = sql + "   and blcab_puer_emb    ='" & Cmb_pto_emb.BoundText & "'"
       If Chk_todos.Value Then
            sql = sql & " and blcab_cono_emb  in (" & bls & ")"
       End If
       End If
       sql = sql & " and blcab_linea       = bldet_linea" & _
               "   and blcab_buque       = bldet_buque" & _
               "   and blcab_viaje       = bldet_viaje" & _
               "   and blcab_puer_emb    = bldet_puer_emb" & _
               "   and blcab_puer_des    = bldet_puer_des" & _
               "   and blcab_emp_resp    = bldet_emp_resp" & _
               "   and blcab_cono_emb    = bldet_cono_emb" & _
               "   and bldet_linea       = cont_linea" & _
               "   and bldet_buque       = cont_buque" & _
               "   and bldet_viaje       = cont_viaje" & _
               "   and bldet_puer_emb    = cont_puer_emb" & _
               "   and bldet_puer_des    = cont_puer_des" & _
               "   and bldet_contenedor  = cont_contenedor" & _
               "   and cont_linea        = tama_equip_linea " & _
               "   and cont_tam_equip    = tama_equip_codigo" & _
               "   and cond_cont_linea   = cont_linea       " & _
               "   and cond_cont_codigo  = cont_condicion   " '& _
               " Order by blcab_linea,    blcab_buque,    blcab_viaje," & _
               "          blcab_puer_emb, blcab_cono_emb, bldet_contenedor "

' // carga general o granel
sql = sql & "UNION Select docu_blcab.*, docu_bldet.*, '' as tama_equip_descripcion, '' as cond_cont_descripcion, '' as cont_sello1, 0 as cont_tara  " & _
               " from docu_blcab, docu_bldet " & _
               " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'"
       If Chk_envio_pto.Value = 1 Then
          sql = sql + "   and blcab_puer_emb    ='" & Cmb_pto_emb.BoundText & "'"
       If Chk_todos.Value Then
            sql = sql & " and blcab_cono_emb  in (" & bls & ")"
       End If
       End If
       sql = sql & " and blcab_linea       = bldet_linea" & _
               "   and blcab_buque       = bldet_buque" & _
               "   and blcab_viaje       = bldet_viaje" & _
               "   and blcab_puer_emb    = bldet_puer_emb" & _
               "   and blcab_puer_des    = bldet_puer_des" & _
               "   and blcab_emp_resp    = bldet_emp_resp" & _
               "   and blcab_cono_emb    = bldet_cono_emb" & _
               "   and bldet_contenedor  = '' " & _
               " Order by blcab_linea,    blcab_buque,    blcab_viaje," & _
               "          blcab_puer_emb, blcab_cono_emb, bldet_contenedor"


       REG_det.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
           
       linea = 11
       linea_ini = linea
       bl = ""
       Do While Not REG_det.EOF
            Me.MousePointer = 11
            DoEvents
            If bl <> REG_det("blcab_cono_emb") Then
                'Set rng_est = sht_est.Range(sht_est.Cells(linea, 3), sht_est.Cells(linea, 12)): rng_est.Borders(xlDiagonalDown).LineStyle = xlNone: rng_est.Borders(xlEdgeBottom).LineStyle = xlContinuous
                If linea_ini + 8 > linea And linea <> 11 Then linea = linea_ini + 10 Else linea = linea + 3
                linea_ini = linea
                sht_est.Cells(linea, 1).Value = "(S)" & REG_det("blcab_embarcador")
                sht_est.Cells(linea + 1, 1).Value = REG_det("blcab_direc_embarc")
                sht_est.Cells(linea + 3, 1).Value = "(C)" & REG_det("blcab_consig")
                sht_est.Cells(linea + 4, 1).Value = REG_det("blcab_direc_consig")
                sht_est.Cells(linea + 6, 1).Value = "(N)" & REG_det("blcab_notif")
                sht_est.Cells(linea + 7, 1).Value = REG_det("blcab_direc_notif")
                sht_est.Cells(linea - 1, 3).Value = REG_det("blcab_cono_emb"): sht_est.Cells(linea - 1, 3).Font.Bold = True:
                If Chk_envio_pto.Value = 0 Then
                    sht_est.Cells(linea - 1, 1).Value = devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", REG_det("blcab_puer_emb"), " puer_linea ='" & Cmb_linea.BoundText & "'"): sht_est.Cells(linea - 1, 1).Font.Bold = True:
                End If
                'If "MSCUBU107146" = REG_det("blcab_cono_emb") Then MsgBox "aqui"
                sht_est.Cells(linea + 1, 8).Value = CStr(chk_str(REG_det("bldet_desc_carg")))
                bl = REG_det("blcab_cono_emb")
            
            End If
            
            sht_est.Cells(linea, 11).Value = "Gross: " & CDbl(REG_det("blcab_peso_tot"))
            sht_est.Cells(linea + 1, 11).Value = "Tare: " & CDbl(REG_det("cont_tara"))
            sht_est.Cells(linea + 2, 11).Value = "Volume :" & CDbl(REG_det("blcab_volumen_tot"))
                        
            sht_est.Cells(linea, 5).Value = Trim(REG_det("bldet_contenedor"))
            sht_est.Cells(linea + 1, 5).Value = REG_det("tama_equip_descripcion") & Chr(10) & REG_det("cond_cont_descripcion")
            sht_est.Cells(linea + 2, 5).Value = "SEAL: " & Trim(REG_det("cont_sello1"))
            sht_est.Cells(linea, 8).Value = CDbl(REG_det("bldet_bultos")) & " " & REG_det("bldet_embalaje")
            REG_det.MoveNext
            Me.MousePointer = 0
            DoEvents
            linea = linea + 3
       Loop
               
'*************************************************************

      'If Dir(Grabar.filename) <> "" Then
       If Dir(Gstr_Ruta_BlWord, vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord)
       If Dir(Gstr_Ruta_BlWord & Trim(Cmb_buque.Text), vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord & Trim(Cmb_buque.Text))
       If Dir(Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text), vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text))
       
       NomArchivo = Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text) & "\" & "Manifest Cargo.XLS"
       If Dir(NomArchivo, vbArchive) <> "" Then Kill (NomArchivo)
       Me.MousePointer = 0
       DoEvents
       
       wb_est.ActiveWorkbook.SaveAs FileName:= _
        NomArchivo, FileFormat:= _
        xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
        , CreateBackup:=False
fin:
       'wb_est.SaveAs Trim$(NomArchivo)
       wb_est.Application.Visible = True ' visible la aplicación
       Set wb_est = Nothing
Exit Function
errores_Excel:
    Me.MousePointer = 0
    DoEvents
    If Err = 1004 Then
        Resume Next
    Else
        MsgBox "Error " & Err.Description
        Exit Function
        Resume 0
    End If
End Function


