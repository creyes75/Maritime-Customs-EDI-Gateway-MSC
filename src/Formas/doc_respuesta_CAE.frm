VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form doc_respuesta_CAE 
   Caption         =   "Archivo de Respuesta de la CAE"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   8580
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cmd_buscar 
      Height          =   525
      Left            =   5805
      Picture         =   "doc_respuesta_CAE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Consulta Codigos de Error CAE"
      Top             =   75
      Width           =   555
   End
   Begin MSGrid.Grid Grid_error 
      Height          =   1730
      Left            =   120
      TabIndex        =   21
      Top             =   3120
      Width           =   8295
      _Version        =   65536
      _ExtentX        =   14631
      _ExtentY        =   3052
      _StockProps     =   77
      BackColor       =   16777215
      FixedCols       =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   " Control de Respuesta "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2160
      Left            =   105
      TabIndex        =   5
      Top             =   630
      Width           =   6360
      Begin VB.TextBox Txt_tot_doc_error 
         Height          =   285
         Left            =   4950
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1710
         Width           =   1290
      End
      Begin VB.TextBox Txt_tot_doc_OK 
         Height          =   285
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1710
         Width           =   1290
      End
      Begin VB.TextBox Txt_tipo_procesamiento 
         Height          =   285
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1410
         Width           =   4095
      End
      Begin VB.TextBox Txt_operador 
         Height          =   285
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   510
         Width           =   4770
      End
      Begin VB.TextBox Txt_tipo_operador 
         Height          =   285
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   195
         Width           =   4770
      End
      Begin VB.Frame Frame1 
         Caption         =   " Proceso Fecha / Hora "
         ForeColor       =   &H00000080&
         Height          =   570
         Left            =   150
         TabIndex        =   8
         Top             =   795
         Width           =   6105
         Begin VB.TextBox txt_fecha_hora_fin 
            Height          =   285
            Left            =   4155
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   195
            Width           =   1740
         End
         Begin VB.TextBox txt_fecha_hora_inicio 
            Height          =   285
            Left            =   750
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   195
            Width           =   1740
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Inicio"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   150
            TabIndex        =   10
            Top             =   225
            Width           =   375
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fin "
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   3735
            TabIndex        =   9
            Top             =   225
            Width           =   255
         End
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total doc. Error"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3675
         TabIndex        =   13
         Top             =   1755
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total doc. OK"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   165
         TabIndex        =   12
         Top             =   1755
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Procesamiento :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   165
         TabIndex        =   11
         Top             =   1470
         Width           =   1725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Operador :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   165
         TabIndex        =   7
         Top             =   495
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Operdor :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   165
         TabIndex        =   6
         Top             =   225
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmd_file 
      Caption         =   "Abrir"
      Height          =   525
      Left            =   3960
      Picture         =   "doc_respuesta_CAE.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   75
      Width           =   555
   End
   Begin VB.TextBox Txt_archivo 
      Height          =   285
      Left            =   825
      TabIndex        =   2
      Top             =   210
      Width           =   3015
   End
   Begin Threed.SSCommand cmd_Proceso 
      Height          =   525
      Left            =   4560
      TabIndex        =   0
      Top             =   75
      Width           =   555
      _Version        =   65536
      _ExtentX        =   979
      _ExtentY        =   926
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
      AutoSize        =   1
      Picture         =   "doc_respuesta_CAE.frx":1DE4
   End
   Begin Threed.SSCommand cmd_salir 
      Height          =   525
      Left            =   5175
      TabIndex        =   1
      Top             =   75
      Width           =   555
      _Version        =   65536
      _ExtentX        =   979
      _ExtentY        =   926
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
      Picture         =   "doc_respuesta_CAE.frx":2236
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   3420
      Top             =   -30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSGrid.Grid Grid_cert 
      Height          =   1590
      Left            =   135
      TabIndex        =   23
      Top             =   5145
      Width           =   8295
      _Version        =   65536
      _ExtentX        =   14631
      _ExtentY        =   2805
      _StockProps     =   77
      BackColor       =   16777215
      FixedCols       =   0
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Respuesta para certificacion digital"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   105
      TabIndex        =   24
      Top             =   4890
      Width           =   3030
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Respuesta a los mensajes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   135
      TabIndex        =   22
      Top             =   2850
      Width           =   2220
   End
   Begin VB.Label Label9 
      Caption         =   "Archivo :"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   150
      TabIndex        =   4
      Top             =   210
      Width           =   1140
   End
End
Attribute VB_Name = "doc_respuesta_CAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_buscar_Click()
    doc_aux_cons_Error.Show
End Sub

Private Sub cmd_file_Click()
    CDialog.InitDir = Gstr_Ruta_files
    CDialog.DialogTitle = "Selección archivo Plano"
    CDialog.Filter = "(*.ZIP)|*.zip"
    
    CDialog.ShowOpen
    Txt_archivo.Text = CDialog.FileName
End Sub

Private Sub cmd_Proceso_Click()
If Dir(Txt_archivo.Text, vbArchive) = "" Then
   MsgBox "La ruta del archivo no es válida", vbInformation
   Exit Sub
End If
Call procesa_respuesta

End Sub

Public Sub procesa_respuesta()
On Error GoTo Errores
Dim cadena As String

Me.MousePointer = 11
carga_parametros "CAE"
Call Form_Load

sep = Gstr_param.separador '","
del = Gstr_param.delimitador ' "|"
     'If Dir(Gstr_Ruta_files & "\" & buque, vbDirectory) = "" Then MkDir (Gstr_Ruta_files & "\" & buque)
     'If Dir(Gstr_Ruta_files & "\" & buque & "\" & viaje, vbDirectory) = "" Then MkDir (Gstr_Ruta_files & "\" & buque & "\" & viaje)
     If Dir(Gstr_Ruta_files, vbDirectory) = "" Then MkDir (Gstr_Ruta_files)
     If Dir(Gstr_Ruta_files & "\TMP", vbDirectory) = "" Then MkDir (Gstr_Ruta_files & "\TMP")
     
     If Dir(Gstr_Ruta_files & "TMP/RESCTROL.txt", vbArchive) <> "" Then Kill Gstr_Ruta_files & "TMP/RESCTROL.txt"
     If Dir(Gstr_Ruta_files & "TMP/RESMENSJ.txt", vbArchive) <> "" Then Kill Gstr_Ruta_files & "TMP/RESMENSJ.txt"
     If Dir(Gstr_Ruta_files & "TMP/RESACEPT.txt", vbArchive) <> "" Then Kill Gstr_Ruta_files & "TMP/RESACEPT.txt"
     
     'If Dir(Gstr_Ruta_files & "\" & buque & "\" & viaje & "\CARRES.ZIP", vbArchive) <> "" Then Kill (Gstr_Ruta_files & "\" & buque & "\" & viaje & "\CARRES.ZIP")
     'FileCopy Txt_archivo.Text, Gstr_Ruta_files & "\" & buque & "\" & viaje & "\CARRES.ZIP"
     'For i = 1 To 10000
     '   DoEvents
     'Next i
     i = Shell(Gstr_Ruta_files & "pkunzip " & Txt_archivo.Text & " " & Gstr_Ruta_files & "TMP\", vbMaximizedFocus)
     'Shell (Gstr_Ruta_files & "pkunzip " & Gstr_Ruta_files & buque & "\" & viaje & "\CARRES.ZIP " & Gstr_Ruta_files & "TMP\")
     For i = 1 To 40000
        DoEvents
     Next i
     Sleep (1000)
     
     If Dir(Gstr_Ruta_files & "TMP\RESCTROL.txt", vbArchive) = "" Then MsgBox "No se encuentra archivo " & Gstr_Ruta_files & "TMP/RESCTROL.txt" & " Posible error al desempaquetar archivo o archivo incorrecto", vbInformation
     If Dir(Gstr_Ruta_files & "TMP\RESMENSJ.txt", vbArchive) = "" Then MsgBox "No se encuentra archivo " & Gstr_Ruta_files & "TMP/RESMENSJ.txt" & " Posible error al desempaquetar archivo o archivo incorrecto", vbInformation
     If Dir(Gstr_Ruta_files & "TMP\RESACEPT.txt", vbArchive) = "" Then MsgBox "No se encuentra archivo " & Gstr_Ruta_files & "TMP/RESACEPT.txt" & " Posible error al desempaquetar archivo o archivo incorrecto", vbInformation
     
     If Dir(Gstr_Ruta_files & "TMP\RESCTROL.txt", vbArchive) <> "" Then
        x = FreeFile
        Open Gstr_Ruta_files & "TMP\RESCTROL.txt" For Input As #x
        Do While Not EOF(x)
            Input #x, tipo_oper, operador, fecha_ini, hora_ini, fecha_fin, hora_fin, tipo_proc, doc_ok, doc_error
            
            If Mid(tipo_oper, 2, Len(tipo_oper) - 2) = Gstr_param.tipo_oper Then
               Txt_tipo_operador.Text = "Naviero"
            Else
               Txt_tipo_operador.Text = "Linea"
            End If
            Txt_operador.Text = devuelve_desc("docu_emp_resp", "emp_resp_codigo", "emp_resp_descripcion", Mid(operador, 2, Len(operador) - 2), "emp_resp_linea ='CAE'")
            txt_fecha_hora_inicio.Text = Format(Mid(fecha_ini, 1, 4) & "/" & Mid(fecha_ini, 5, 2) & "/" & Mid(fecha_ini, 7), "DD/MM/YYYY")
            txt_fecha_hora_inicio.Text = txt_fecha_hora_inicio.Text & " " & Mid(hora_ini, 2, Len(hora_ini) - 2)
            
            txt_fecha_hora_fin.Text = Format(Mid(fecha_fin, 1, 4) & "/" & Mid(fecha_fin, 5, 2) & "/" & Mid(fecha_fin, 7), "DD/MM/YYYY")
            txt_fecha_hora_fin.Text = txt_fecha_hora_fin.Text & " " & Mid(hora_fin, 2, Len(hora_fin) - 2)
            
            Select Case Mid(tipo_proc, 2, Len(tipo_proc) - 2)
              Case "01", "1": Txt_tipo_procesamiento.Text = "MENSAJE ACEPTADO TOTALMENTE"
              Case "17": Txt_tipo_procesamiento.Text = "MENSAJE RECIBIDO"
              Case "14": Txt_tipo_procesamiento.Text = "MENSAJE CON ERROR"
              Case "67": Txt_tipo_procesamiento.Text = "MENSAJE ACEPTADO PARCIALMENTE"
            End Select
                        
            Txt_tot_doc_OK.Text = doc_ok
            Txt_tot_doc_error.Text = doc_error
        Loop
        Close #x
     End If
     
     
     If Dir(Gstr_Ruta_files & "TMP\RESMENSJ.txt", vbArchive) <> "" Then
        x = FreeFile
        Open Gstr_Ruta_files & "TMP\RESMENSJ.txt" For Input As #x
        Do While Not EOF(x)
            Grid_error.Rows = Grid_error.Rows + 1
            adu_ing = "": año_doc = "": num_doc = "": tipo_mani = "": pto_emb = "": emp_resp = "": tipo_error = "": Tipo_cont = "": desp1 = "": desp2 = "": bl = "": cod_error = "": desc_error = "": tipo_proc = ""
            Input #x, adu_ing, año_doc, num_doc, tipo_mani, pto_emb, emp_resp, tipo_error, Tipo_cont, desp1, desp2, bl, cod_error, desc_error, tipo_proc
            Do While Right(desc_error, 1) <> "|"
                desc_error = desc_error & tipo_proc
                Input #x, tipo_proc
            Loop
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'TL Aumento para presentar solo los mensajes que sean errores de rechazo'
            '   y no los warnings (los warnings son mayores a 9000)                 '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            val_error = CDbl(Mid(cod_error, 2, Len(cod_error) - 2))
          '  If val_error < 9000 Then
              Grid_error.Row = Grid_error.Rows - 2
              Grid_error.Col = 0 ' adua ingreso
              Grid_error.Text = Trim(devuelve_desc("docu_adua_ing", "adua_ing_codigo", "adua_ing_descripcion", Mid(adu_ing, 2, Len(adu_ing) - 2), "adua_ing_linea ='CAE'"))
              Grid_error.Col = 1 ' año documento
              Grid_error.Text = Mid(año_doc, 2, Len(año_doc) - 2)
              Grid_error.Col = 2 ' num documento
              Grid_error.Text = Mid(num_doc, 2, Len(num_doc) - 2)
              Grid_error.Col = 3 'tipo manifiesto
              Select Case Mid(tipo_mani, 2, Len(tipo_mani) - 2)
                   Case "01": Grid_error.Text = "MARITIMO"
                   Case "02": Grid_error.Text = "AEREO"
                   Case "03": Grid_error.Text = "TERRESTRE"
              End Select
              Grid_error.Col = 4 ' pto emb
              Grid_error.Text = Trim(devuelve_desc("docu_puerto", "puer_codigo", "puer_descripcion", Mid(pto_emb, 2, Len(pto_emb) - 2), "puer_linea ='CAE'"))
              Grid_error.Col = 5 ' emp resp
              Grid_error.Text = devuelve_desc("docu_emp_resp", "emp_resp_codigo", "emp_resp_descripcion", Mid(emp_resp, 2, Len(emp_resp) - 2), "emp_resp_linea ='CAE'")
              Grid_error.Col = 6 'tipo error
              Select Case Trim(Mid(tipo_error, 2, Len(tipo_error) - 2))
                   Case "":         Grid_error.Text = "NO HAY ERROR"
                   Case "01", "1":  Grid_error.Text = "ERROR EN MANIFIESTO"
                   Case "02", "2":  Grid_error.Text = "ERROR EN BL CAB/DET"
                   Case "03", "3":  Grid_error.Text = "ERROR EN CONTENEDOR"
              End Select
              Grid_error.Col = 7 'tipo GUIA/CONT
              Select Case Mid(Tipo_cont, 2, Len(Tipo_cont) - 2)
                   Case "703": Grid_error.Text = "BL HOUSE"
                   Case "740": Grid_error.Text = "GUIA AEREA HOUSE"
                   Case "704": Grid_error.Text = "Bl MASTER"
                   Case "741": Grid_error.Text = "GUIA AEREA MASTER"
                   Case "CN": Grid_error.Text = "CONTENEDOR"
                   Case "PA": Grid_error.Text = "PALLET AEREO"
                   Case "PL": Grid_error.Text = "PLATAFORMA"
                   Case "TE": Grid_error.Text = "TRAILER"
              End Select
              pos = pos2
              Grid_error.Col = 8 ' NO. DOC TRANSPORTE / CONTENEDOR
              Grid_error.Text = Mid(bl, 2, Len(bl) - 2)
              Grid_error.Col = 9 ' CODIGO DEL ERROR
              Grid_error.Text = Mid(cod_error, 2, Len(cod_error) - 2)
              Grid_error.Col = 10 ' DESCRIPCION DEL ERROR
              Grid_error.Text = Mid(desc_error, 2, Len(desc_error) - 2)
              Grid_error.Col = 11 ' DESCRIPCION DEL ERROR tabla
              Grid_error.Text = devuelve_desc("docu_error", "error_codigo", "error_descripcion", CStr(Val(Mid(cod_error, 2, Len(cod_error) - 2))))
              Grid_error.Col = 12 ' TIPO DE PROCESAMIENTO
              Select Case Mid(tipo_proc, 2, Len(tipo_proc) - 2)
                   Case "01", "1": Grid_error.Text = "ACEPTADO"
                   Case "02", "2": Grid_error.Text = "ACEPTADO CONDICIONALMENTE"
                   Case "08", "8": Grid_error.Text = "RECHAZADO(CON ERROR)"
                   Case Else:      Grid_error.Text = "Warning"
              End Select
           'End If
        Loop
        Close #x
     End If
     
     If Dir(Gstr_Ruta_files & "TMP\RESACEPT.txt", vbArchive) <> "" Then
        x = FreeFile
        Open Gstr_Ruta_files & "TMP\RESACEPT.txt" For Input As #x
        Do While Not EOF(x)
            Grid_cert.Rows = Grid_cert.Rows + 1
            adu_ing = "": año_doc = "": num_doc = "": tipo_mani = "": tipo_tran = "": dig_veri = "": sec_tran = "": año_ord = "": num_ord = "": resu_ver = ""
            Input #x, adu_ing, año_doc, tipo_mani, num_doc, tipo_tran, dig_veri, sec_tran, año_ord, num_ord, resu_ver
            Grid_cert.Row = Grid_cert.Rows - 2
            Grid_cert.Col = 0 ' adua ingreso
                Grid_cert.Text = Trim(devuelve_desc("docu_adua_ing", "adua_ing_codigo", "adua_ing_descripcion", Mid(adu_ing, 2, Len(adu_ing) - 2), "adua_ing_linea ='CAE'"))
            Grid_cert.Col = 1 ' año documento
                Grid_cert.Text = Mid(año_doc, 2, Len(año_doc) - 2)
            Grid_cert.Col = 2 'tipo manifiesto
                Select Case Mid(tipo_mani, 2, Len(tipo_mani) - 2)
                   Case "01": Grid_cert.Text = "MARITIMO"
                   Case "02": Grid_cert.Text = "AEREO"
                   Case "03": Grid_cert.Text = "TERRESTRE"
                End Select
            Grid_cert.Col = 3 ' num documento
                Grid_cert.Text = Mid(num_doc, 2, Len(num_doc) - 2)
            Grid_cert.Col = 4 ' tipo transaccion
                If Len(tipo_tran) > 1 Then
                Grid_cert.Text = devuelve_desc("docu_tipo_transaccion", "tipo_tran_codigo", "tipo_tran_descripcion", Mid(tipo_tran, 2, Len(tipo_tran) - 2))
                Else
                Grid_cert.Text = devuelve_desc("docu_tipo_transaccion", "tipo_tran_codigo", "tipo_tran_descripcion", Str(tipo_tran))
                End If
            Grid_cert.Col = 5 ' digito verificador
                Grid_cert.Text = dig_veri
            Grid_cert.Col = 6 ' sec documentos
                Grid_cert.Text = Mid(sec_tran, 2, Len(sec_tran) - 2)
            Grid_cert.Col = 7 'año orden
                Grid_cert.Text = Mid(año_ord, 2, Len(año_ord) - 2)
            Grid_cert.Col = 8 ' num orden
                Grid_cert.Text = Mid(num_ord, 2, Len(num_ord) - 2)
           Grid_cert.Col = 9 ' resumen verificacion
                Grid_cert.Text = Mid(resu_ver, 2, Len(resu_ver) - 2)
        Loop
        Close #x
     End If
     Me.MousePointer = 0
Exit Sub
Errores:
    Me.MousePointer = 0
    MsgBox "Ocurrio el Error " & Err.Description, vbInformation
    Close
    Exit Sub
    Resume 0
End Sub

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Grid_error.Cols = 13
   Grid_error.Rows = 2
   Grid_error.Row = 0
   Grid_error.Col = 0: Grid_error.Text = "Adu Ing":         Grid_error.ColWidth(0) = 800
   Grid_error.Col = 1: Grid_error.Text = "Año doc":         Grid_error.ColWidth(1) = 800
   Grid_error.Col = 2: Grid_error.Text = "Num manif":       Grid_error.ColWidth(2) = 900
   Grid_error.Col = 3: Grid_error.Text = "Tipo mani":       Grid_error.ColWidth(3) = 800
   Grid_error.Col = 4: Grid_error.Text = "Pto emb":         Grid_error.ColWidth(4) = 800
   Grid_error.Col = 5: Grid_error.Text = "Emp resp":        Grid_error.ColWidth(5) = 1000
   Grid_error.Col = 6: Grid_error.Text = "Tipo error":      Grid_error.ColWidth(6) = 1000
   Grid_error.Col = 7: Grid_error.Text = "Tipo cont":       Grid_error.ColWidth(7) = 900
   Grid_error.Col = 8: Grid_error.Text = "Doc Transp/Cont": Grid_error.ColWidth(8) = 1000
   Grid_error.Col = 9: Grid_error.Text = "Codigo error":    Grid_error.ColWidth(9) = 800
   Grid_error.Col = 10: Grid_error.Text = "Desc error":     Grid_error.ColWidth(10) = 5000
   Grid_error.Col = 11: Grid_error.Text = "Desc error Tabla": Grid_error.ColWidth(11) = 12000
   Grid_error.Col = 12: Grid_error.Text = "Tipo Proces":    Grid_error.ColWidth(12) = 1000
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'TL pongo en comentario porque voy a maximizar la forma
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'Me.Width = 8700
   'Me.Height = 7230
   
   
   Grid_cert.Cols = 10
   Grid_cert.Rows = 2
   Grid_cert.Row = 0
   Grid_cert.Col = 0: Grid_cert.Text = "Adu Ing":         Grid_cert.ColWidth(0) = 1000
   Grid_cert.Col = 1: Grid_cert.Text = "Año doc":         Grid_cert.ColWidth(1) = 800
   Grid_cert.Col = 2: Grid_cert.Text = "Tipo mani":       Grid_cert.ColWidth(2) = 900
   Grid_cert.Col = 3: Grid_cert.Text = "Num manif":       Grid_cert.ColWidth(3) = 900
   Grid_cert.Col = 4: Grid_cert.Text = "Tipo trans":      Grid_cert.ColWidth(4) = 2000
   Grid_cert.Col = 5: Grid_cert.Text = "Dig verif":       Grid_cert.ColWidth(5) = 1000
   Grid_cert.Col = 6: Grid_cert.Text = "Sec Docum":       Grid_cert.ColWidth(6) = 1000
   Grid_cert.Col = 7: Grid_cert.Text = "Año Orden Int":   Grid_cert.ColWidth(7) = 1000
   Grid_cert.Col = 8: Grid_cert.Text = "Num Orden Int":   Grid_cert.ColWidth(8) = 1000
   Grid_cert.Col = 9: Grid_cert.Text = "Resumen verif":    Grid_cert.ColWidth(9) = 2500
      
End Sub

Private Sub Form_Resize()

' 7230
'1730
'1590

'3910


Grid_error.Width = Me.Width - 400
Grid_error.Height = (Me.Height - 4010) / 2

Grid_cert.Width = Me.Width - 400
Grid_cert.Height = (Me.Height - 4010) / 2

Label10.Top = 3250 + (Me.Height - 4010) / 2
Grid_cert.Top = 3500 + (Me.Height - 4010) / 2

End Sub
