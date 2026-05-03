VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form doc_aux_busq_bl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda por Bl o Contendor de Importación"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   1830
   ClientWidth     =   5385
   Icon            =   "doc_aux_busq_bl.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   5385
   Begin VB.CommandButton Cmd_modificar 
      Caption         =   "Modificar"
      Height          =   660
      Left            =   4440
      Picture         =   "doc_aux_busq_bl.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Modificar Bl"
      Top             =   2265
      Width           =   750
   End
   Begin VB.CommandButton Cmd_Imprimir 
      Caption         =   "Imprimir"
      Height          =   660
      Left            =   4440
      Picture         =   "doc_aux_busq_bl.frx":213C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprimir Bl"
      Top             =   2985
      Width           =   750
   End
   Begin VB.Frame Frame1 
      Height          =   2310
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   5235
      Begin VB.CommandButton Cmd_busq 
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
         Left            =   4410
         Picture         =   "doc_aux_busq_bl.frx":223E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Buscar"
         Top             =   750
         Width           =   735
      End
      Begin VB.CommandButton Cmd_cancel 
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
         Left            =   4410
         Picture         =   "doc_aux_busq_bl.frx":3F38
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir"
         Top             =   1470
         Width           =   735
      End
      Begin VB.TextBox Txt_contenedor 
         Height          =   285
         Left            =   420
         MaxLength       =   11
         TabIndex        =   6
         Top             =   1725
         Width           =   1440
      End
      Begin VB.TextBox Txt_bl 
         Height          =   285
         Left            =   450
         MaxLength       =   25
         TabIndex        =   4
         Top             =   1155
         Width           =   2385
      End
      Begin MSDataListLib.DataCombo Cmb_linea 
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Top             =   525
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc data_linea 
         Height          =   330
         Left            =   3645
         Top             =   285
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
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "Contenedor"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   420
         TabIndex        =   7
         Top             =   1470
         Width           =   960
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Bill of lading"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   435
         TabIndex        =   5
         Top             =   945
         Width           =   840
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
         Left            =   510
         TabIndex        =   2
         Top             =   270
         Width           =   825
      End
   End
   Begin MSDataGridLib.DataGrid Grid_bl 
      Height          =   2145
      Left            =   75
      TabIndex        =   3
      Top             =   2295
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   3784
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
   Begin MSAdodcLib.Adodc RDC_grid_bl 
      Height          =   330
      Left            =   3315
      Top             =   3735
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
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
End
Attribute VB_Name = "doc_aux_busq_bl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_busq_Click()
            
            If Trim(Txt_bl.Text) = "" And Trim$(Txt_contenedor.Text) = "" Then
                MsgBox "No existe ningún criterio para realizar la búsqueda!", vbInformation
                Exit Sub
            End If
            Me.MousePointer = 11
             sql = "SELECT blcab_buque          as Buque,        blcab_viaje      as Viaje,     " & _
                   "       blcab_puer_emb       as Pto_embarque, blcab_cono_emb   as Bl,           " & _
                   "       blcab_puer_origen    as Pto_origen, blcab_puer_des as Pto_descarga,  " & _
                   "       blcab_pto_dest_final as Pto_final,  blcab_consig   as Consignatario, " & _
                   "       blcab_embarcador     as Embarcador "
                   
             If Trim(Txt_bl.Text) <> "" And Trim$(Txt_contenedor.Text) = "" Then
                  sql = sql & " FROM docu_blcab " & _
                              " WHERE blcab_linea       ='" & Cmb_linea.BoundText & "'" & _
                              "   And blcab_cono_emb    ='" & Txt_bl.Text & "'"
             Else 'If Trim(txt_bl.Text) = "" And Trim$(Txt_contenedor.Text) <> "" Then
                  sql = sql & " FROM docu_blcab, docu_bldet, docu_cont " & _
                              " WHERE blcab_linea       ='" & Cmb_linea.BoundText & "'"
                  If Trim(Txt_contenedor) <> "" Then
                      sql = sql & "   And cont_contenedor  ='" & Txt_contenedor.Text & "'"
                  End If
                  If Trim(Txt_bl.Text) <> "" Then
                      sql = sql & "   And blcab_cono_emb    ='" & Txt_bl.Text & "'"
                  End If
                  sql = sql & "   And blcab_linea = bldet_linea       " & _
                              "   And blcab_buque = bldet_buque       " & _
                              "   And blcab_viaje = bldet_viaje       " & _
                              "   And blcab_puer_emb = bldet_puer_emb " & _
                              "   And blcab_puer_des = bldet_puer_des " & _
                              "   And blcab_emp_resp = bldet_emp_resp " & _
                              "   And blcab_cono_emb = bldet_cono_emb " & _
                              "   And bldet_linea    = cont_linea     " & _
                              "   and bldet_buque    = cont_buque     " & _
                              "   and bldet_viaje    = cont_viaje     " & _
                              "   and bldet_puer_emb = cont_puer_emb  " & _
                              "   and bldet_puer_des = cont_puer_des  " & _
                              "   and bldet_contenedor= cont_contenedor "
             End If

                          
            RDC_grid_bl.ConnectionString = Gstr_conex_docu
            RDC_grid_bl.RecordSource = sql
            RDC_grid_bl.Refresh
            Set Grid_bl.DataSource = RDC_grid_bl
            Grid_bl.Refresh
            Me.MousePointer = 0
            
End Sub

Private Sub Cmd_cancel_Click()
Unload Me
End Sub


Private Sub Cmd_Imprimir_Click()
If Not RDC_grid_bl.Recordset.EOF Then
    doc_repo_bl.Show
    doc_repo_bl.Cmb_linea.BoundText = Cmb_linea.BoundText
    doc_repo_bl.Cmb_buque.BoundText = RDC_grid_bl.Recordset("Buque")
    doc_repo_bl.Cmb_viaje.BoundText = RDC_grid_bl.Recordset("Viaje")
    doc_repo_bl.Cmb_pto_emb.BoundText = RDC_grid_bl.Recordset("Pto_embarque")
    doc_repo_bl.Txt_pto_descarga.Text = RDC_grid_bl.Recordset("Pto_descarga")
    doc_repo_bl.Txt_bl.Text = RDC_grid_bl.Recordset("bl")
    doc_repo_bl.txt_bl_LostFocus
End If
End Sub

Private Sub cmd_modificar_Click()
Dim frm_bl_mod As New doc_bl



If Not RDC_grid_bl.Recordset.EOF Then
    frm_bl_mod.Show
    frm_bl_mod.Gstr_opc_bl = "M"
    Call frm_bl_mod.carga_bl(Cmb_linea.BoundText, RDC_grid_bl.Recordset("Buque"), RDC_grid_bl.Recordset("viaje"), RDC_grid_bl.Recordset("Pto_embarque"), RDC_grid_bl.Recordset("Pto_descarga"), RDC_grid_bl.Recordset("bl"))
    frm_bl_mod.Caption = "Bl Impo: " & RDC_grid_bl.Recordset("bl")
End If

End Sub

Private Sub Form_Load()
carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
If Not data_linea.Recordset.EOF Then data_linea.Recordset.MoveFirst
    
End Sub

Private Sub Grid_bl_HeadClick(ByVal ColIndex As Integer)
    Me.MousePointer = 11
    Me.RDC_grid_bl.Recordset.Sort = Grid_bl.Columns(ColIndex).DataField
    Grid_bl.Refresh
    Me.MousePointer = 0
End Sub
