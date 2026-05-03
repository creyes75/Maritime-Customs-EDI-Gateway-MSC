VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form doc_repo_flete_expo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fletes de Exportacion"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   1830
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4680
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   345
      Top             =   2820
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3915
      Picture         =   "doc_repo_flete_exp.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2700
      Width           =   645
   End
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
      Height          =   645
      Left            =   3240
      Picture         =   "doc_repo_flete_exp.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2700
      Width           =   645
   End
   Begin VB.Frame Frame1 
      Height          =   2505
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   4425
      Begin MSAdodcLib.Adodc data_viaje 
         Height          =   330
         Left            =   3165
         Top             =   1935
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
         Left            =   3060
         Top             =   1215
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
         Top             =   510
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
         Top             =   525
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
         Left            =   150
         TabIndex        =   3
         Top             =   1185
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo Cmb_viaje 
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   1875
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Viaje"
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
         Left            =   180
         TabIndex        =   6
         Top             =   1635
         Width           =   435
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Buque"
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
         Left            =   165
         TabIndex        =   4
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "Linea"
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
         Height          =   210
         Left            =   165
         TabIndex        =   2
         Top             =   270
         Width           =   825
      End
   End
End
Attribute VB_Name = "doc_repo_flete_expo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmb_linea_Change()
carga_datacombo data_buque, Cmb_buque, "Select * from docu_buque Where buqu_linea = '" & Cmb_linea.BoundText & "' and buqu_status ='A'", "buqu_codigo", "buqu_descripcion"
End Sub

Private Sub Command1_Click()
    Set wb_est = CreateObject("excel.Application")
    DoEvents
    Me.MousePointer = 11
    Call LISTADO_EXCEL_FLETE
    Me.MousePointer = 0

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
If Not data_linea.Recordset.EOF Then data_linea.Recordset.MoveFirst
    
End Sub

Public Function LISTADO_EXCEL_FLETE()
Dim REG_det     As New ADODB.Recordset
Dim REG_viaje   As New ADODB.Recordset
Dim i           As Integer
Dim linea       As Integer
Dim puerto_ant  As String
Dim fila_inicial As Integer
Dim total1 As String
Dim total2 As String
Dim total3 As String
Dim total4 As String
Dim total5 As String
Dim total6 As String

Me.MousePointer = 11
      wb_est.Workbooks.Add
      ' aumenta un libro dentro de la aplicacion excel
      '*** crea nuevas hojas dentro del mismo libro *******
      'For i = 1 To 2  'hasta el numero de reportes
      '   wb_est.Sheets.Add
      '   DoEvents
      'Next i
      pagina% = 1
         'Set sht_est = wb_est.ActiveWorkbook.Sheets("sheet1")   'setea la hoja actual
         'Set sht_est = wb_est.ActiveWorkbook.Sheets("hoja1")   'setea la hoja actual
         Set sht_est = wb_est.ActiveWorkbook.Sheets(1)   'setea la hoja actual
         sht_est.Name = "Freight Report"         'asigna un nombre a la hoja
         sht_est.PageSetup.Orientation = 2
         sht_est.PageSetup.PaperSize = xlPaperA4
         sht_est.PageSetup.BottomMargin = 40
         sht_est.PageSetup.TopMargin = 40
         sht_est.PageSetup.FooterMargin = 20
         sht_est.PageSetup.HeaderMargin = 20
         sht_est.Activate
      
   sht_est.Cells(5, 3).Value = "EXPORT FREIGHT REPORT"
   sht_est.Cells(2, 1).Value = "MV." & Cmb_buque.Text & " V " & Cmb_viaje.Text
      
   sht_est.Cells(9, 1).Value = "Consignee"
   sht_est.Cells(9, 2).Value = "Discharge"
   sht_est.Cells(9, 3).Value = "B/L nbr"
   sht_est.Cells(9, 4).Value = "20"
   sht_est.Cells(9, 5).Value = "40"
   sht_est.Cells(9, 6).Value = "N.Freight"
   sht_est.Cells(9, 7).Value = "G.Freight"
   sht_est.Cells(9, 8).Value = "PREPAID"
   sht_est.Cells(9, 9).Value = "COLLECT"
   
   Call Excel_setea_fuente_fila(sht_est, 2, 12, "moderm", True, False, 8000000)
   Call Excel_setea_fuente_fila(sht_est, 5, 16, "moderm", True, False, 8000000)
   Call Excel_setea_fuente_fila(sht_est, 9, 12, "moderm", True, False, 8000000)

   
   DoEvents
   
   Call Excel_setea_ancho_columna(sht_est, 9, 1, 9, 1, 30) ' consignee
   Call Excel_setea_ancho_columna(sht_est, 9, 2, 9, 2, 12)  ' load
   Call Excel_setea_ancho_columna(sht_est, 9, 3, 9, 3, 15) ' bl
   Call Excel_setea_ancho_columna(sht_est, 9, 4, 9, 4, 5) ' 20
   Call Excel_setea_ancho_columna(sht_est, 9, 5, 9, 5, 5) ' 40
   Call Excel_setea_ancho_columna(sht_est, 9, 6, 9, 6, 12) ' n freight
   Call Excel_setea_ancho_columna(sht_est, 9, 7, 9, 7, 12) ' g freight
   Call Excel_setea_ancho_columna(sht_est, 9, 8, 9, 8, 12) ' prepaid
   Call Excel_setea_ancho_columna(sht_est, 9, 9, 9, 9, 12) ' collect
          
 
   Call Excel_setea_borde(sht_est, 9, 1, 9, 9, 1, False)
   sht_est.Columns("F:F").Style = "Currency"
   sht_est.Columns("G:G").Style = "Currency"
   sht_est.Columns("H:H").Style = "Currency"
   sht_est.Columns("I:I").Style = "Currency"
   
   '*****************datos de la fecha de arribo *************************
    sql = "Select * from docu_buque_viaje " & _
          " where buqv_linea = '" & Cmb_linea.BoundText & "'" & _
          "   and buqv_buque       ='" & Cmb_buque.BoundText & "'" & _
          "   and buqv_viaje       ='" & Cmb_viaje.BoundText & "'"
    REG_viaje.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   
    If Not REG_viaje.EOF Then sht_est.Cells(5, 1).Value = "ETS: " & REG_viaje("buqv_fecha_arribo")
    sht_est.Cells(5, 1).Font.Size = 9
    '**********************************************************************
    
      sql = "Select blcab_consig, puer_descripcion, blcab_cono_emb, blcab_flete, blcab_metod_pago, " & _
               "  sum(case cont_tara when 2200 then 1 else 0 end) as no20,  sum(case cont_tara when 4400 then 1 else 0 end) as no40   " & _
               " from docu_blcab_exp, docu_bldet_exp, docu_puerto, docu_cont_exp " & _
               " Where blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
               "   and blcab_buque       ='" & Cmb_buque.BoundText & "'" & _
               "   and blcab_viaje       ='" & Cmb_viaje.BoundText & "'" & _
               "   and blcab_linea  = bldet_linea " & _
               "   and blcab_buque  = bldet_buque " & _
               "   and blcab_viaje  = bldet_viaje " & _
               "   and blcab_puer_emb = bldet_puer_emb " & _
               "   and blcab_puer_des = bldet_puer_des " & _
               "   and blcab_emp_resp = bldet_emp_resp " & _
               "   and blcab_cono_emb = bldet_cono_emb " & _
               "   and bldet_linea    = cont_linea " & _
               "   and bldet_buque    = cont_buque " & _
               "   and bldet_viaje    = cont_viaje " & _
               "   and bldet_puer_emb = cont_puer_emb " & _
               "   and bldet_puer_des = cont_puer_des " & _
               "   and bldet_contenedor=cont_contenedor" & _
               "   and blcab_linea       = puer_linea " & _
               "   and blcab_puer_des    = puer_codigo " & _
               " group by puer_descripcion, blcab_consig, blcab_cono_emb, blcab_flete, blcab_metod_pago"

        ' "ORDER By blcab_linea, blcab_buque, blcab_viaje, blcab_puer_emb, blcab_consig"

       REG_det.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
           
       linea = 10
       fila_inicial = linea
       total1 = "=SUM(": total2 = "=SUM(": total3 = "=SUM("
       total4 = "=SUM(": total5 = "=SUM(": total6 = "=SUM("
       If Not REG_det.EOF Then puerto_ant = REG_det("puer_descripcion")
       
       Do While Not REG_det.EOF
            If puerto_ant <> REG_det("puer_descripcion") Then
                Call Excel_setea_fuente_fila(sht_est, linea, 8, "moderm", True, False, 8000000)
                sht_est.Cells(linea, 1).Value = "TOTAL"
                sht_est.Cells(linea, 4).Value = "= SUM(D" & fila_inicial & ":D" & linea - 1 & ")": total1 = total1 + "D" & linea & "+"
                sht_est.Cells(linea, 5).Value = "= SUM(E" & fila_inicial & ":E" & linea - 1 & ")": total2 = total2 + "E" & linea & "+"
                sht_est.Cells(linea, 6).Value = "= SUM(F" & fila_inicial & ":F" & linea - 1 & ")": total3 = total3 + "F" & linea & "+"
                sht_est.Cells(linea, 7).Value = "= SUM(G" & fila_inicial & ":G" & linea - 1 & ")": total4 = total4 + "G" & linea & "+"
                sht_est.Cells(linea, 8).Value = "= SUM(H" & fila_inicial & ":H" & linea - 1 & ")": total5 = total5 + "H" & linea & "+"
                sht_est.Cells(linea, 9).Value = "= SUM(I" & fila_inicial & ":I" & linea - 1 & ")": total6 = total6 + "I" & linea & "+"
                linea = linea + 2
                fila_inicial = linea
                puerto_ant = REG_det("puer_descripcion")
            End If
            sht_est.Rows(linea).Font.Size = 7
            sht_est.Cells(linea, 1).Value = REG_det("blcab_consig")
            sht_est.Cells(linea, 2).Value = REG_det("puer_descripcion")
            sht_est.Cells(linea, 3).Value = REG_det("blcab_cono_emb")
            sht_est.Cells(linea, 4).Value = REG_det("no20")
            sht_est.Cells(linea, 5).Value = REG_det("no40")
            sht_est.Cells(linea, 6).Value = ""
            sht_est.Cells(linea, 7).Value = REG_det("blcab_flete")
            If REG_det("blcab_metod_pago") = "PP" Then sht_est.Cells(linea, 8).Value = REG_det("blcab_flete")
            If REG_det("blcab_metod_pago") = "CC" Then sht_est.Cells(linea, 9).Value = REG_det("blcab_flete")
            REG_det.MoveNext
            linea = linea + 1
       Loop
       
       '''''''''para el ultimo de los puertos '''''''''''''''''''''''''''''
        Call Excel_setea_fuente_fila(sht_est, linea, 8, "moderm", True, False, 8000000)
        sht_est.Cells(linea, 1).Value = "TOTAL"
        sht_est.Cells(linea, 4).Value = "= SUM(D" & fila_inicial & ":D" & linea - 1 & ")": total1 = total1 + "D" & linea & "+"
        sht_est.Cells(linea, 5).Value = "= SUM(E" & fila_inicial & ":E" & linea - 1 & ")": total2 = total2 + "E" & linea & "+"
        sht_est.Cells(linea, 6).Value = "= SUM(F" & fila_inicial & ":F" & linea - 1 & ")": total3 = total3 + "F" & linea & "+"
        sht_est.Cells(linea, 7).Value = "= SUM(G" & fila_inicial & ":G" & linea - 1 & ")": total4 = total4 + "G" & linea & "+"
        sht_est.Cells(linea, 8).Value = "= SUM(H" & fila_inicial & ":H" & linea - 1 & ")": total5 = total5 + "H" & linea & "+"
        sht_est.Cells(linea, 9).Value = "= SUM(I" & fila_inicial & ":I" & linea - 1 & ")": total6 = total6 + "I" & linea & "+"
        linea = linea + 2
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '''''''''''''''''''''' GRAN TOTAL ''''''''''''''''''''''''''''''''''
        linea = linea + 2
        Call Excel_setea_fuente_fila(sht_est, linea, 10, "moderm", True, False, 8000000)
        sht_est.Cells(linea, 2).Value = "GRAN TOTAL"
        sht_est.Cells(linea, 4).Value = Mid(total1, 1, Len(total1) - 1) + ")"
        sht_est.Cells(linea, 5).Value = Mid(total2, 1, Len(total2) - 1) + ")"
        sht_est.Cells(linea, 6).Value = Mid(total3, 1, Len(total3) - 1) + ")"
        sht_est.Cells(linea, 7).Value = Mid(total4, 1, Len(total4) - 1) + ")"
        sht_est.Cells(linea, 8).Value = Mid(total5, 1, Len(total5) - 1) + ")"
        sht_est.Cells(linea, 9).Value = Mid(total6, 1, Len(total6) - 1) + ")"
      '*************************************************************
      Call Excel_setea_borde(sht_est, 10, 1, linea + 1, 9, 1, True)

      'If Dir(Grabar.filename) <> "" Then
       If Dir(Gstr_Ruta_BlWord, vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord)
       If Dir(Gstr_Ruta_BlWord & Trim(Cmb_buque.Text), vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord & Trim(Cmb_buque.Text))
       If Dir(Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text), vbDirectory) = "" Then MkDir (Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text))
       
       NomArchivo = Gstr_Ruta_BlWord & Trim(Cmb_buque.Text) & "\" & Trim(Cmb_viaje.Text) & "\" & "FleteImpo.XLS"
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
