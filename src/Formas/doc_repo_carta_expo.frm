VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form doc_repo_carta_expo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carta de Exportacion"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   1830
   ClientWidth     =   5070
   Icon            =   "doc_repo_carta_expo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5070
   Begin VB.CommandButton Cmd_plantilla 
      Height          =   615
      Left            =   120
      Picture         =   "doc_repo_carta_expo.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Abre Documento Plantilla"
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Cmd_Procesar 
      Height          =   615
      Left            =   3720
      Picture         =   "doc_repo_carta_expo.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Cmd_cancel 
      Height          =   615
      Left            =   4380
      Picture         =   "doc_repo_carta_expo.frx":14DE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   825
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   90
      TabIndex        =   0
      Top             =   15
      Width           =   4905
      Begin MSAdodcLib.Adodc data_viaje 
         Height          =   330
         Left            =   4200
         Top             =   720
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
         Caption         =   "Adodc3"
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
      Begin MSAdodcLib.Adodc data_buque 
         Height          =   330
         Left            =   4080
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
         Caption         =   "Adodc2"
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
      Begin MSAdodcLib.Adodc data_linea 
         Height          =   330
         Left            =   3120
         Top             =   390
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
      Begin MSDataListLib.DataCombo Cmb_linea 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   405
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo Cmb_buque 
         Height          =   315
         Left            =   135
         TabIndex        =   3
         Top             =   1065
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo Cmb_viaje 
         Height          =   315
         Left            =   3120
         TabIndex        =   5
         Top             =   1080
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Viaje"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3120
         TabIndex        =   6
         Top             =   840
         Width           =   345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Buque"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   165
         TabIndex        =   4
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Linea"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   165
         TabIndex        =   2
         Top             =   150
         Width           =   825
      End
   End
End
Attribute VB_Name = "doc_repo_carta_expo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VARIABLES PARA EXCEL
Public wb_est  As Object 'Excel.Workbook
Public sht_est As Excel.Worksheet
Public rng_est As Excel.Range
Public oWord    As Object


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
End Sub



Private Sub Cmd_plantilla_Click()
On Error GoTo fin
    Dim objWord As Object
    For i = 1 To 3
        doc_plantilla = App.Path + "\Plantillas\" + "CartaExport" & i & ".doc"
        Set objWord = CreateObject("Word.Application")
        objWord.Documents.Open doc_plantilla
        objWord.Visible = True
        Set objWord = Nothing
    Next i
    'x = Shell("WinWord " & doc_plantilla, vbMaximizedFocus)
fin:
End Sub

Private Sub Cmd_procesar_Click()
    
    Set oWord = CreateObject("Word.Application")
    Ruta = App.Path + "\Plantillas\"
    For i = 1 To 3
        
        doc_plantilla = Ruta + "CartaExport" & i & ".doc"
        
        If Dir(doc_plantilla, vbArchive) = "" Then
           MsgBox "No se encontro archivo Plantilla " + doc_plantilla, vbExclamation
           Exit Sub
        End If
        oWord.Documents.Open doc_plantilla
        oWord.Visible = True
        Me.MousePointer = 11
        Call Imprime_bls(Str(i))
        Me.MousePointer = 0
        'oWord.Documents.Close doc_plantilla
    Next i
    Set oWord = Nothing
End Sub

Private Sub Cmd_cancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Opt_cont.Value = True
carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
If Not data_linea.Recordset.EOF Then data_linea.Recordset.MoveFirst
    
End Sub

Public Function LISTADO_EXCEL_CONT()
Dim REG_det As New ADODB.Recordset
Dim i As Integer
Dim linea As Integer
      
Me.MousePointer = 11
      wb_est.Workbooks.Add
      
      pagina% = 1
         Set sht_est = wb_est.ActiveWorkbook.Sheets(1)   'setea la hoja actual
         sht_est.Name = "Listado de contenedores"         'asigna un nombre a la hoja
         sht_est.PageSetup.Orientation = 2
         sht_est.PageSetup.PaperSize = xlPaperA4
         sht_est.PageSetup.BottomMargin = 40
         sht_est.PageSetup.TopMargin = 40
         sht_est.PageSetup.FooterMargin = 20
         sht_est.PageSetup.HeaderMargin = 20
         sht_est.Activate
      
   sht_est.Cells(2, 5).Value = "Listado de Contenedores"
   sht_est.Cells(4, 3).Value = "BUQUE :": sht_est.Cells(4, 4).Value = Cmb_buque.Text
   sht_est.Cells(5, 3).Value = "VIAJE :": sht_est.Cells(5, 4).Value = Cmb_viaje.Text
      
   sht_est.Cells(9, 1).Value = "Contenedor"
   sht_est.Cells(9, 2).Value = "Tamaño"
   sht_est.Cells(9, 3).Value = "Condicion"
   sht_est.Cells(9, 4).Value = "Sello1"
   sht_est.Cells(9, 5).Value = "Sello2"
   sht_est.Cells(9, 6).Value = "Sello3"
   sht_est.Cells(9, 7).Value = "Sello4"
   
   sht_est.Rows(2).Font.Name = "moderm"
   sht_est.Rows(2).Font.Bold = True
   sht_est.Rows(2).Font.Italic = True
   sht_est.Rows(2).Font.color = 8000000
   sht_est.Rows(2).Font.Size = 14
   
   sht_est.Rows(9).Font.Name = "moderm"
   sht_est.Rows(9).Font.Bold = True
   sht_est.Rows(9).Font.Italic = True
   
   sht_est.Cells(4, 3).Font.Bold = True:   sht_est.Cells(4, 3).Font.color = 8000000
   sht_est.Cells(5, 3).Font.Bold = True:   sht_est.Cells(5, 3).Font.color = 8000000
   sht_est.Cells(6, 3).Font.Bold = True:   sht_est.Cells(6, 3).Font.color = 8000000
   sht_est.Cells(7, 3).Font.Bold = True:   sht_est.Cells(7, 3).Font.color = 8000000
   
   DoEvents
   
   Set rng_est = sht_est.Range(sht_est.Cells(4, 1), sht_est.Cells(4, 1)):   rng_est.ColumnWidth = 12 'CNT
   Set rng_est = sht_est.Range(sht_est.Cells(4, 2), sht_est.Cells(4, 2)):   rng_est.ColumnWidth = 20 'TAMAÑO
   Set rng_est = sht_est.Range(sht_est.Cells(1, 3), sht_est.Cells(1, 3)):   rng_est.ColumnWidth = 8  'CONDICION
   Set rng_est = sht_est.Range(sht_est.Cells(1, 4), sht_est.Cells(1, 4)):   rng_est.ColumnWidth = 15 'SELLO1
   Set rng_est = sht_est.Range(sht_est.Cells(1, 5), sht_est.Cells(1, 5)):   rng_est.ColumnWidth = 15 'SELLO2
   Set rng_est = sht_est.Range(sht_est.Cells(1, 6), sht_est.Cells(1, 6)):   rng_est.ColumnWidth = 15 'SELLO3
   Set rng_est = sht_est.Range(sht_est.Cells(1, 7), sht_est.Cells(1, 7)):   rng_est.ColumnWidth = 15 'SELLO4
      
 
   sht_est.Cells(9, 1).Borders.LineStyle = 1
   sht_est.Cells(9, 2).Borders.LineStyle = 1
   sht_est.Cells(9, 3).Borders.LineStyle = 1
   sht_est.Cells(9, 4).Borders.LineStyle = 1
   sht_est.Cells(9, 5).Borders.LineStyle = 1
   sht_est.Cells(9, 6).Borders.LineStyle = 1
   sht_est.Cells(9, 7).Borders.LineStyle = 1
   
      sql = "Select * " & _
               " from docu_cont " & _
               " Where cont_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and cont_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and cont_viaje       ='" & Cmb_viaje.BoundText & "'"


       REG_det.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
           
       linea = 10
       Do While Not REG_det.EOF
            sht_est.Rows(linea).Font.Size = 7
            sht_est.Cells(linea, 1).Value = REG_det("cont_contenedor") 'devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", REG_det("blcab_puer_emb"), " puer_linea ='" & Cmb_linea.BoundText & "'")
            sht_est.Cells(linea, 2).Value = devuelve_desc("docu_tamano_equip", "tama_equip_codigo", "tama_equip_descripcion", REG_det("cont_tam_equip"), " tama_equip_linea ='" & Cmb_linea.BoundText & "'")
            sht_est.Cells(linea, 3).Value = devuelve_desc("docu_cond_contenedor", "cond_cont_codigo", "cond_cont_descripcion", REG_det("cont_condicion"), " cond_cont_linea ='" & Cmb_linea.BoundText & "' and cond_cont_status ='A'")
            sht_est.Cells(linea, 4).Value = REG_det("cont_sello1")
            sht_est.Cells(linea, 5).Value = REG_det("cont_sello2")
            sht_est.Cells(linea, 6).Value = REG_det("cont_sello3")
            sht_est.Cells(linea, 7).Value = REG_det("cont_sello4")
            'sht_est.Cells(linea, 8).Value = REG_det("bldet_bultos")
            'sht_est.Cells(linea, 9).Value = REG_det("bldet_IMO")
            'sht_est.Cells(linea, 10).Value = REG_det("blcab_consig")
            REG_det.MoveNext
            linea = linea + 1
       Loop
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       'TL aumento de total de contenedor al reporte
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       linea = linea + 2
       sht_est.Cells(linea, 2).Value = "Total de Contenedores:"
       sht_est.Cells(linea, 4).Value = (linea - 12)
      '*************************************************************

      'If Dir(Grabar.filename) <> "" Then
       If Dir(Gstr_Ruta_BlWord, vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord)
       If Dir(Gstr_Ruta_BlWord & Trim(Cmb_buque.Text), vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord & Trim(Cmb_buque.Text))
       If Dir(Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text), vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text))
       
       NomArchivo = Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text) & "\" & "ContImpo.XLS"
       If Dir(NomArchivo, vbArchive) <> "" Then Kill (NomArchivo)
       
       wb_est.ActiveWorkbook.SaveAs FileName:= _
        NomArchivo, FileFormat:= _
        xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
        , CreateBackup:=False
       
       'wb_est.SaveAs Trim$(NomArchivo)
       wb_est.Application.Visible = True ' visible la aplicación
       'Set wb_est = Nothing
Me.MousePointer = 0
End Function

Public Function LISTADO_EXCEL_CONT_BL()
Dim REG_det As New ADODB.Recordset
Dim i As Integer
Dim linea As Integer
      
Me.MousePointer = 11
      wb_est.Workbooks.Add
      
      pagina% = 1
         Set sht_est = wb_est.ActiveWorkbook.Sheets(1)   'setea la hoja actual
         sht_est.Name = "Listado de contenedores x BL"         'asigna un nombre a la hoja
         sht_est.PageSetup.Orientation = 2
         sht_est.PageSetup.PaperSize = xlPaperA4
         sht_est.PageSetup.BottomMargin = 40
         sht_est.PageSetup.TopMargin = 40
         sht_est.PageSetup.FooterMargin = 20
         sht_est.PageSetup.HeaderMargin = 20
         sht_est.Activate
      
   sht_est.Cells(2, 5).Value = "Listado de Contenedores por BL"
   sht_est.Cells(4, 3).Value = "BUQUE :": sht_est.Cells(4, 4).Value = Cmb_buque.Text
   sht_est.Cells(5, 3).Value = "VIAJE :": sht_est.Cells(5, 4).Value = Cmb_viaje.Text
      
   sht_est.Cells(9, 1).Value = "BL"
   sht_est.Cells(9, 2).Value = "Contenedor"
   sht_est.Cells(9, 3).Value = "Tamaño"
   sht_est.Cells(9, 4).Value = "Condicion"
   sht_est.Cells(9, 5).Value = "Sello1"
   sht_est.Cells(9, 6).Value = "Sello2"
   sht_est.Cells(9, 7).Value = "Sello3"
   sht_est.Cells(9, 8).Value = "Sello4"
   
   sht_est.Rows(2).Font.Name = "moderm"
   sht_est.Rows(2).Font.Bold = True
   sht_est.Rows(2).Font.Italic = True
   sht_est.Rows(2).Font.color = 8000000
   sht_est.Rows(2).Font.Size = 14
   
   sht_est.Rows(9).Font.Name = "moderm"
   sht_est.Rows(9).Font.Bold = True
   sht_est.Rows(9).Font.Italic = True
   
   sht_est.Cells(4, 3).Font.Bold = True:   sht_est.Cells(4, 3).Font.color = 8000000
   sht_est.Cells(5, 3).Font.Bold = True:   sht_est.Cells(5, 3).Font.color = 8000000
   sht_est.Cells(6, 3).Font.Bold = True:   sht_est.Cells(6, 3).Font.color = 8000000
   sht_est.Cells(7, 3).Font.Bold = True:   sht_est.Cells(7, 3).Font.color = 8000000
   sht_est.Cells(8, 3).Font.Bold = True:   sht_est.Cells(8, 3).Font.color = 8000000
   
   DoEvents
   
   Set rng_est = sht_est.Range(sht_est.Cells(4, 1), sht_est.Cells(4, 1)):   rng_est.ColumnWidth = 12 'BL
   Set rng_est = sht_est.Range(sht_est.Cells(4, 2), sht_est.Cells(4, 2)):   rng_est.ColumnWidth = 12 'TAMAÑO
   Set rng_est = sht_est.Range(sht_est.Cells(1, 3), sht_est.Cells(1, 3)):   rng_est.ColumnWidth = 20  'TAMAÑO
   Set rng_est = sht_est.Range(sht_est.Cells(1, 4), sht_est.Cells(1, 4)):   rng_est.ColumnWidth = 8 'CONDICION
   Set rng_est = sht_est.Range(sht_est.Cells(1, 5), sht_est.Cells(1, 5)):   rng_est.ColumnWidth = 15 'SELLO1
   Set rng_est = sht_est.Range(sht_est.Cells(1, 6), sht_est.Cells(1, 6)):   rng_est.ColumnWidth = 15 'SELLO2
   Set rng_est = sht_est.Range(sht_est.Cells(1, 7), sht_est.Cells(1, 7)):   rng_est.ColumnWidth = 15 'SELLO3
   Set rng_est = sht_est.Range(sht_est.Cells(1, 8), sht_est.Cells(1, 8)):   rng_est.ColumnWidth = 15 'SELLO4
      
 
   sht_est.Cells(9, 1).Borders.LineStyle = 1
   sht_est.Cells(9, 2).Borders.LineStyle = 1
   sht_est.Cells(9, 3).Borders.LineStyle = 1
   sht_est.Cells(9, 4).Borders.LineStyle = 1
   sht_est.Cells(9, 5).Borders.LineStyle = 1
   sht_est.Cells(9, 6).Borders.LineStyle = 1
   sht_est.Cells(9, 7).Borders.LineStyle = 1
   sht_est.Cells(9, 8).Borders.LineStyle = 1
   
      sql = "Select * " & _
               " from docu_bldet, docu_cont " & _
               " Where cont_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and cont_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and cont_viaje       ='" & Cmb_viaje.BoundText & "'" & _
               "   and cont_linea = bldet_linea " & _
               "   and cont_buque = bldet_buque " & _
               "   and cont_viaje = bldet_viaje " & _
               "   and cont_puer_emb = bldet_puer_emb " & _
               "   and cont_puer_des = bldet_puer_des " & _
               "   and cont_contenedor = bldet_contenedor"


       REG_det.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
           
       linea = 10
       Do While Not REG_det.EOF
            sht_est.Rows(linea).Font.Size = 7
            sht_est.Cells(linea, 1).Value = REG_det("bldet_cono_emb")
            sht_est.Cells(linea, 2).Value = REG_det("cont_contenedor") 'devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", REG_det("blcab_puer_emb"), " puer_linea ='" & Cmb_linea.BoundText & "'")
            sht_est.Cells(linea, 3).Value = devuelve_desc("docu_tamano_equip", "tama_equip_codigo", "tama_equip_descripcion", REG_det("cont_tam_equip"), " tama_equip_linea ='" & Cmb_linea.BoundText & "'")
            sht_est.Cells(linea, 4).Value = devuelve_desc("docu_cond_contenedor", "cond_cont_codigo", "cond_cont_descripcion", REG_det("cont_condicion"), " cond_cont_linea ='" & Cmb_linea.BoundText & "' and cond_cont_status ='A'")
            sht_est.Cells(linea, 5).Value = REG_det("cont_sello1")
            sht_est.Cells(linea, 6).Value = REG_det("cont_sello2")
            sht_est.Cells(linea, 7).Value = REG_det("cont_sello3")
            sht_est.Cells(linea, 8).Value = REG_det("cont_sello4")
            REG_det.MoveNext
            linea = linea + 1
       Loop
       
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       'TL aumento de total de contenedor al reporte
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       linea = linea + 2
       sht_est.Cells(linea, 2).Value = "Total de Contenedores:"
       sht_est.Cells(linea, 4).Value = (linea - 12)
      '*************************************************************

      'If Dir(Grabar.filename) <> "" Then
       If Dir(Gstr_Ruta_BlWord, vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord)
       If Dir(Gstr_Ruta_BlWord & Trim(Cmb_buque.Text), vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord & Trim(Cmb_buque.Text))
       If Dir(Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text), vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text))
       
       NomArchivo = Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text) & "\" & "ContImpoBL.XLS"
       If Dir(NomArchivo, vbArchive) <> "" Then Kill (NomArchivo)
       
       wb_est.ActiveWorkbook.SaveAs FileName:= _
        NomArchivo, FileFormat:= _
        xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
        , CreateBackup:=False
       
       'wb_est.SaveAs Trim$(NomArchivo)
       wb_est.Application.Visible = True ' visible la aplicación
       'Set wb_est = Nothing
Me.MousePointer = 0
End Function


Private Sub Form_Unload(Cancel As Integer)
    Set sht_est = Nothing
    Set wb_est = Nothing
End Sub


Sub Imprime_bls(autoridad As String)
'
' Macro1 Macro
' Macro grabada el 10/21/2006 por careyes
'
    Dim REG_det As New ADODB.Recordset
    Dim NomArchivo  As String
    Dim no_reg As Integer
    Dim no_linea As Integer
    
    
        ' reemplaza el BUQUE
    oWord.Selection.Find.ClearFormatting
    oWord.Selection.Find.Replacement.ClearFormatting
    With oWord.Selection.Find
        .Text = "<FECHA>"
        .Replacement.Text = Format(Date, "dd MMMM /YYYY")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    oWord.Selection.Find.Execute Replace:=wdReplaceAll
    
    ' reemplaza el BUQUE
    oWord.Selection.Find.ClearFormatting
    oWord.Selection.Find.Replacement.ClearFormatting
    With oWord.Selection.Find
        .Text = "<COD_BUQUE>"
        .Replacement.Text = Me.Cmb_buque.Text
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    oWord.Selection.Find.Execute Replace:=wdReplaceAll
    
    'reemplaza el viaje
    oWord.Selection.Find.ClearFormatting
    oWord.Selection.Find.Replacement.ClearFormatting
    With oWord.Selection.Find
        .Text = "<COD_VIAJE>"
        .Replacement.Text = Me.Cmb_viaje.Text
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    oWord.Selection.Find.Execute Replace:=wdReplaceAll
    
    sql = "Select * " & _
               " from docu_buque_viaje " & _
               " Where buqv_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and buqv_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and buqv_viaje       ='" & Cmb_viaje.BoundText & "'"
    REG_det.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    If Not REG_det.EOF Then
        'reemplaza el ETA
        oWord.Selection.Find.ClearFormatting
        oWord.Selection.Find.Replacement.ClearFormatting
        With oWord.Selection.Find
            .Text = "<FECHA_ETA>"
            .Replacement.Text = REG_det("buqv_fecha_arribo")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        oWord.Selection.Find.Execute Replace:=wdReplaceAll
    End If
    REG_det.Close


    sql = "Select * " & _
               " from docu_manifiesto_exp " & _
               " Where mani_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and mani_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and mani_viaje       ='" & Cmb_viaje.BoundText & "'"
    REG_det.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    If Not REG_det.EOF Then
        'reemplaza el no manifiesto
        oWord.Selection.Find.ClearFormatting
        oWord.Selection.Find.Replacement.ClearFormatting
        With oWord.Selection.Find
            .Text = "<NO_MANIFIESTO>"
            .Replacement.Text = REG_det("mani_adua_ing") & "-" & REG_det("mani_ano_car") & "-" & REG_det("mani_tip_mani") & "-" & REG_det("mani_mani_car") & "-" & REG_det("mani_digi_ver")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        oWord.Selection.Find.Execute Replace:=wdReplaceAll
    End If
    REG_det.Close
    
    
    'ActiveWindow.ActivePane.VerticalPercentScrolled = 0
    oWord.Selection.Find.ClearFormatting
    With oWord.Selection.Find
        .Text = "<TABLA>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    oWord.Selection.Find.Execute

oWord.Selection.TypeText Text:=" "

    sql = "Select * " & _
               " from docu_blcab_exp " & _
               " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'" & _
               " Order by blcab_cono_emb"
    REG_det.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        
    no_cols = 6
    
    If Not REG_det.EOF Then
      no_reg = CInt(REG_det.RecordCount / 6)
      If (REG_det.RecordCount / 6) > no_reg Then
         no_reg = no_reg + 1
      End If
    End If
    If no_reg = 0 Then Exit Sub
    
    If no_reg <= 5 Then no_reg = 10
    
    oWord.ActiveDocument.Tables.Add Range:=oWord.Selection.Range, NumRows:=no_reg, NumColumns _
        :=6, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    no_linea = 1
    Do While Not REG_det.EOF
        If no_linea > no_reg Then
            no_linea = 1
            'no_col = 2
            oWord.Selection.MoveUp Unit:=wdLine, Count:=no_reg
            oWord.Selection.EndKey Unit:=wdLine
            oWord.Selection.MoveRight Unit:=wdCharacter, Count:=1
            'oWord.Selection.MoveRight Unit:=wdColumn, Count:=2
            
        End If
        oWord.Selection.TypeText Text:=CStr(REG_det("blcab_cono_emb"))
        oWord.Selection.MoveDown Unit:=wdLine, Count:=1
        no_linea = no_linea + 1
        REG_det.MoveNext
    Loop
    oWord.Selection.MoveUp Unit:=wdLine, Count:=no_linea
    oWord.Selection.MoveDown Unit:=wdLine, Count:=1
    oWord.Selection.HomeKey Unit:=wdLine
      
    oWord.Selection.MoveDown Unit:=wdLine, Count:=no_reg - 1, Extend:=wdExtend
    oWord.Selection.MoveRight Unit:=wdCharacter, Count:=6, Extend:=wdExtend
    oWord.Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    oWord.Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    oWord.Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    oWord.Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    oWord.Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
    oWord.Selection.Borders(wdBorderVertical).LineStyle = wdLineStyleNone
    oWord.Selection.Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
    oWord.Selection.Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    oWord.Selection.Font.Name = "Arial"
    oWord.Selection.Font.Name = "Arial"
    oWord.Selection.Font.Size = 8
    
    oWord.Selection.MoveDown Unit:=wdLine, Count:=1
    
    If Dir(Gstr_Ruta_BlWord, vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord)
    If Dir(Gstr_Ruta_BlWord & Trim(Cmb_buque.Text), vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord & Trim(Cmb_buque.Text))
    If Dir(Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text), vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text))
    NomArchivo = Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text) & "\" & "Carta de Exportacion" & autoridad & ".doc"
    If Dir(NomArchivo, vbArchive) <> "" Then Kill (NomArchivo)
    oWord.ActiveDocument.SaveAs FileName:=NomArchivo
    'Set oWord = Nothing


End Sub

