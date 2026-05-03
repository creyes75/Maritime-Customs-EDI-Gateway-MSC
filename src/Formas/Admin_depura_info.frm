VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Admin_depura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Depuración de Información"
   ClientHeight    =   2385
   ClientLeft      =   4230
   ClientTop       =   2640
   ClientWidth     =   6015
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2385
   ScaleWidth      =   6015
   Begin VB.Frame Fra_linea 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00000080&
      Height          =   540
      Left            =   105
      TabIndex        =   10
      Top             =   60
      Width           =   2820
      Begin MSDataListLib.DataCombo Cmb_linea 
         Height          =   315
         Left            =   165
         TabIndex        =   0
         Top             =   210
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
         TabIndex        =   11
         Top             =   0
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000080&
      Height          =   990
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   660
      Width           =   5760
      Begin VB.Frame Fra_viaje 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   3150
         TabIndex        =   6
         Top             =   210
         Width           =   1965
         Begin MSDataListLib.DataCombo Cmb_viaje 
            Height          =   315
            Left            =   135
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
            TabIndex        =   7
            Top             =   0
            Width           =   345
         End
      End
      Begin VB.Frame Fra_buque 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   120
         TabIndex        =   8
         Top             =   210
         Width           =   2970
         Begin MSDataListLib.DataCombo Cmb_buque 
            Height          =   315
            Left            =   120
            TabIndex        =   1
            Top             =   225
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
            Left            =   1830
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
            TabIndex        =   9
            Top             =   0
            Width           =   465
         End
      End
   End
   Begin Threed.SSCommand btn_sgte 
      Height          =   495
      Left            =   4830
      TabIndex        =   3
      ToolTipText     =   "Continuar"
      Top             =   1800
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   78
      Picture         =   "Admin_depura_info.frx":0000
   End
   Begin Threed.SSCommand btn_elim 
      Height          =   495
      Left            =   5370
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   1800
      Width           =   495
      _Version        =   65536
      _ExtentX        =   868
      _ExtentY        =   868
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Admin_depura_info.frx":0452
   End
End
Attribute VB_Name = "Admin_depura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_elim_Click()
    Unload Me
End Sub

'Procedimiento:  Elimina informaciòn de la base de datos (Import y Export) de buques que no existen, o son de prueba
'Fecha        :  Domingo 21 de Marzo del 2004
'===================================================================================
Private Sub btn_sgte_Click()
On Error GoTo Errores
Dim sql         As String
Dim linea       As String
Dim buque       As String
Dim viaje       As String
Dim puerto      As String
Dim puerto_des  As String
Dim bl          As String

Dim no_cab_del  As Integer
Dim no_det_del  As Integer
Dim no_cont_del As Integer
Dim no_flete_del As Integer
Dim no_mani_del As Integer

Dim no_cab_del_expo  As Integer
Dim no_det_del_expo  As Integer
Dim no_cont_del_expo As Integer
Dim no_flete_del_expo As Integer
Dim no_mani_del_expo As Integer

Dim no_buqv_del As Integer

' variables de las tablas

If Not Cmb_buque.MatchedWithList Or Not Cmb_viaje.MatchedWithList Then
   MsgBox "No ha escogido el buque para el proceso de depuración", vbInformation
   Exit Sub
End If


linea = Cmb_linea.BoundText
buque = Cmb_buque.BoundText
viaje = Cmb_viaje.BoundText

   Gcon_docu.BeginTrans
        sql = "Delete " & _
                    "FROM docu_cont " & _
                    " WHERE cont_linea ='" & linea & "'" & _
                    "   and cont_buque = '" & buque & "'" & _
                    "   and cont_viaje = '" & viaje & "'"
        Gcon_docu.Execute sql, no_cont_del
        
        sql = "Delete " & _
                    "FROM docu_blflete " & _
                    " WHERE flete_linea ='" & linea & "'" & _
                    "   and flete_buque = '" & buque & "'" & _
                    "   and flete_viaje = '" & viaje & "'"
        Gcon_docu.Execute sql, no_flete_del
        
        sql = "Delete FROM docu_bldet " & _
                    " WHERE bldet_linea ='" & linea & "'" & _
                    "   and bldet_buque    = '" & buque & "'" & _
                    "   and bldet_viaje    = '" & viaje & "'"
        Gcon_docu.Execute sql, no_det_del
        
        sql = "Delete From docu_blcab " & _
                    " WHERE blcab_linea       ='" & linea & "'" & _
                    "   and blcab_buque       ='" & buque & "'" & _
                    "   and blcab_viaje       ='" & viaje & "'"
        Gcon_docu.Execute sql, no_cab_del
        
        sql = "Delete From docu_MANIFIESTO "
        sql = sql & " WHERE mani_linea       ='" & linea & "'" & _
                "   and mani_buque       ='" & buque & "'" & _
                "   and mani_viaje       ='" & viaje & "'"
        Gcon_docu.Execute sql, no_mani_del
        
        '//////////// informaciòn de exportacion////////////
        sql = "Delete " & _
                    "FROM docu_cont_exp " & _
                    " WHERE cont_linea ='" & linea & "'" & _
                    "   and cont_buque = '" & buque & "'" & _
                    "   and cont_viaje = '" & viaje & "'"
        Gcon_docu.Execute sql, no_cont_del_expo
        
        sql = "Delete " & _
                    "FROM docu_blflete_exp " & _
                    " WHERE flete_linea ='" & linea & "'" & _
                    "   and flete_buque = '" & buque & "'" & _
                    "   and flete_viaje = '" & viaje & "'"
        Gcon_docu.Execute sql, no_flete_del_expo
        
        sql = "Delete FROM docu_bldet_exp " & _
                    " WHERE bldet_linea ='" & linea & "'" & _
                    "   and bldet_buque    = '" & buque & "'" & _
                    "   and bldet_viaje    = '" & viaje & "'"
        Gcon_docu.Execute sql, no_det_del_expo
        
        sql = "Delete From docu_blcab_exp " & _
                    " WHERE blcab_linea       ='" & linea & "'" & _
                    "   and blcab_buque       ='" & buque & "'" & _
                    "   and blcab_viaje       ='" & viaje & "'"
        Gcon_docu.Execute sql, no_cab_del_expo
        
         sql = "Delete From docu_MANIFIESTO_exp "
        sql = sql & " WHERE mani_linea       ='" & linea & "'" & _
                "   and mani_buque       ='" & buque & "'" & _
                "   and mani_viaje       ='" & viaje & "'"
        Gcon_docu.Execute sql, no_mani_del_expo
        
        ' ////////// se elimina el buque viaje /////////
        sql = "Delete " & _
                    "FROM docu_buque_viaje " & _
                    " WHERE buqv_linea ='" & linea & "'" & _
                    "   and buqv_buque = '" & buque & "'" & _
                    "   and buqv_viaje = '" & viaje & "'"
        Gcon_docu.Execute sql, no_buqv_del

   Gcon_docu.CommitTrans
   msg = " Se Eliminaron : " & Chr(13) & _
         " - Manifietsos Import = " & no_mani_del & Chr(13) & _
         " - Bls         Import = " & no_cab_del & Chr(13) & _
         " - detalles    Import = " & no_det_del & Chr(13) & _
         " - contenedores Import = " & no_cont_del & Chr(13) & _
         " " & Chr(13) & _
         " - Manifietsos Export = " & no_mani_del_expo & Chr(13) & _
         " - Bls         Export = " & no_cab_del_expo & Chr(13) & _
         " - detalles    Export = " & no_det_del_expo & Chr(13) & _
         " - contenedores Export = " & no_cont_del_expo & Chr(13) & _
         " - buque viaje " & no_buqv_del_expo
         

   MsgBox msg, vbInformation
   data_viaje.Refresh
   data_buque.Refresh
   Cmb_viaje.Refresh
   Cmb_buque.Refresh
Exit Sub
Errores:
    Gcon_docu.RollbackTrans
    MsgBox "ERROR " & Err.Description
    Exit Sub
    Resume 0
End Sub



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


Private Sub Cmb_linea_Change()
carga_datacombo data_buque, Cmb_buque, "Select * from docu_buque Where buqu_linea = '" & Cmb_linea.BoundText & "' and buqu_status ='A'", "buqu_codigo", "buqu_descripcion"
carga_parametros Cmb_linea.BoundText

End Sub

Private Sub Cmb_linea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Cmb_viaje_Change()
   Dim sql As String
   sql = "Select * from docu_manifiesto " & _
         " Where mani_linea ='" & Cmb_linea.BoundText & "'" & _
         "   And mani_buque ='" & Cmb_buque.BoundText & "'" & _
         "   And mani_viaje ='" & Cmb_viaje.BoundText & "'"
End Sub

Private Sub Cmb_viaje_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub Form_Load()
    carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
    If Not data_linea.Recordset.EOF Then data_linea.Recordset.MoveFirst
End Sub
