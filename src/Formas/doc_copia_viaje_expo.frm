VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form doc_expo_copia_viaje 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copia de Bl's por Viaje"
   ClientHeight    =   4275
   ClientLeft      =   4230
   ClientTop       =   2640
   ClientWidth     =   6060
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4275
   ScaleWidth      =   6060
   Begin VB.CheckBox Chk_borra_bl 
      Caption         =   "Borrar Bl Original"
      Height          =   195
      Left            =   135
      TabIndex        =   19
      Top             =   3645
      Width           =   2025
   End
   Begin VB.Frame Fra_linea 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00000080&
      Height          =   540
      Left            =   105
      TabIndex        =   13
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
         TabIndex        =   14
         Top             =   0
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Bl Destino "
      ForeColor       =   &H00000080&
      Height          =   1395
      Left            =   120
      TabIndex        =   12
      Top             =   2145
      Width           =   5790
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   3210
         TabIndex        =   17
         Top             =   180
         Width           =   1965
         Begin MSAdodcLib.Adodc data_viaje_new 
            Height          =   330
            Left            =   705
            Top             =   225
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
         Begin MSDataListLib.DataCombo Cmb_viaje_new 
            Height          =   315
            Left            =   105
            TabIndex        =   4
            Top             =   240
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Viaje"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   135
            TabIndex        =   18
            Top             =   0
            Width           =   345
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   180
         TabIndex        =   15
         Top             =   180
         Width           =   2970
         Begin MSDataListLib.DataCombo Cmb_buque_new 
            Height          =   315
            Left            =   105
            TabIndex        =   3
            Top             =   240
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc data_buque_new 
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Buque"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   0
            Width           =   465
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Bl Origen "
      ForeColor       =   &H00000080&
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   660
      Width           =   5760
      Begin VB.Frame Fra_viaje 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   3150
         TabIndex        =   8
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
            TabIndex        =   9
            Top             =   0
            Width           =   345
         End
      End
      Begin VB.Frame Fra_buque 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   120
         TabIndex        =   10
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
            TabIndex        =   11
            Top             =   0
            Width           =   465
         End
      End
   End
   Begin Threed.SSCommand btn_sgte 
      Height          =   495
      Left            =   4860
      TabIndex        =   5
      ToolTipText     =   "Continuar"
      Top             =   3750
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   78
      Picture         =   "doc_copia_viaje_expo.frx":0000
   End
   Begin Threed.SSCommand btn_elim 
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   3750
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
      Picture         =   "doc_copia_viaje_expo.frx":0452
   End
End
Attribute VB_Name = "doc_expo_copia_viaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_elim_Click()
    Unload Me
End Sub

'Procedimiento:  Copia de un Bl a otro Puerto o al mismo con otro número de BL
'Fecha        :  Domingo 15 de Diciembre del 2002
'Proceso      :  Validar que la linea,buque,viaje,puerto,bl no exista ya en la tabla
'                docu_blcab, docu_bldet.
'                copiar el Bl en el linea,buque,viaje seleccionado con el mismo
'                o distinto Número de BL
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

Dim buque_new   As String
Dim viaje_new   As String
Dim puerto_new  As String
Dim bl_new      As String

Dim reg         As New ADODB.Recordset
Dim reg2        As New ADODB.Recordset

Dim no_cab      As Integer
Dim no_det      As Integer
Dim no_cont     As Integer

Dim no_cab_del  As Integer
Dim no_det_del  As Integer
Dim no_cont_del As Integer
' variables de las tablas

If Not Cmb_buque.MatchedWithList Or Not Cmb_viaje.MatchedWithList Then
   MsgBox "No ha escogido los dato origen de la copia", vbInformation
   Exit Sub
End If

If Not Cmb_buque_new.MatchedWithList Or Not Cmb_viaje_new.MatchedWithList Then
   MsgBox "No ha escogido los dato destino de la copia", vbInformation
   Exit Sub
End If

linea = Cmb_linea.BoundText
buque = Cmb_buque.BoundText
viaje = Cmb_viaje.BoundText

'bl = Trim$(Cmb_Bl.BoundText)

buque_new = Cmb_buque_new.BoundText
viaje_new = Cmb_viaje_new.BoundText

'para validar si existe el bl origen
sql = "SELECT * FROM docu_blcab_exp " & _
      " WHERE blcab_linea       ='" & linea & "'" & _
      "   and blcab_buque       ='" & buque & "'" & _
      "   and blcab_viaje       ='" & viaje & "'" '& _
      "   and blcab_puer_emb    ='" & puerto & "'" & _
      "   and blcab_cono_emb    ='" & bl & "'"

reg2.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
If Not reg2.EOF Then
   Gcon_docu.BeginTrans
   
   puerto_des = reg2("blcab_puer_des")
   
   sql = "Insert Into docu_manifiesto_exp select mani_linea, '" & buque_new & "', '" & viaje_new & "', " & _
                                     " mani_puer_emb,       mani_puer_des,      mani_adua_ing, " & _
                                     " mani_ano_car,        mani_tip_mani,      mani_mani_car, " & _
                                     " mani_digi_ver,       mani_ano_ord,       mani_num_ord, " & _
                                     " mani_tot_doc,        mani_tot_peso,      mani_tot_bult," & _
                                     " mani_tot_vol,        mani_tot_con,       mani_jus_men, " & _
                                     " mani_emp_prop,       mani_emp_nave,      mani_fecha_esti_lle, " & _
                                     " mani_capitan,        mani_puer_ini,      mani_fecha_zarp_ini, " & _
                                     " mani_fecha_envio,    mani_hora_envio    "
    sql = sql & " From docu_MANIFIESTO_exp "
    sql = sql & " WHERE mani_linea       ='" & linea & "'" & _
                "   and mani_buque       ='" & buque & "'" & _
                "   and mani_viaje       ='" & viaje & "'"
   Gcon_docu.Execute sql, no_mani
   
   sql = "INSERT INTO docu_blcab_exp SELECT blcab_linea,   '" & buque_new & "', '" & viaje_new & "', " & _
                                     " blcab_puer_emb ,     blcab_puer_des,     blcab_emp_resp, " & _
                                     " blcab_cono_emb,      blcab_emp_anav,     blcab_tipo_doc, " & _
                                     " blcab_cono_master,   blcab_peso_tot,     blcab_bultos_tot, " & _
                                     " blcab_tipo_ide_cons, blcab_ide_cons,     blcab_consig, " & _
                                     " blcab_direc_consig,  blcab_nacio_consig, blcab_tipo_ide_notif," & _
                                     " blcab_ide_notif,     blcab_notif,        blcab_direc_notif, " & _
                                     " blcab_tipo_ide_emb,  blcab_ide_emb,      blcab_embarcador,    " & _
                                     " blcab_direc_embarc,  blcab_puer_origen,  blcab_puer_trasb,  blcab_fecha_trasb," & _
                                     " blcab_almac_temp,    blcab_tipo_depos,   blcab_deposito_adua, " & _
                                     " blcab_regimen_adu,   blcab_consolidad,   blcab_fecha_trans ,  " & _
                                     " blcab_hora_trans,    blcab_tipo_salida,  blcab_docum_aut, " & _
                                     " blcab_pesob_mal_est, blcab_bultos_mal_est, blcab_notif2," & _
                                     " blcab_direc_notif2,  blcab_no_cont,      blcab_volumen_tot," & _
                                     " blcab_pto_dest_final,blcab_metod_pago,   blcab_moneda, " & _
                                     " blcab_flete,         blcab_otros_gastos, blcab_moneda_otros, " & _
                                     " blcab_seguro,        blcab_moneda_seguro,blcab_metod_pago_otros," & _
                                     " blcab_fecha_emb,     blcab_user_ing,     blcab_fecha_ing, " & _
                                     " blcab_hora_ing,      blcab_user_mod,     blcab_fecha_mod, " & _
                                     " blcab_hora_mod"
    sql = sql & " From docu_blcab_exp "
    sql = sql & " WHERE blcab_linea       ='" & linea & "'" & _
                "   and blcab_buque       ='" & buque & "'" & _
                "   and blcab_viaje       ='" & viaje & "'" '& _
                "   and blcab_puer_emb    ='" & puerto & "'" '& _
                "   and blcab_cono_emb    ='" & bl & "'"
   Gcon_docu.Execute sql, no_cab

   '**********************************************************************
   '   llena docu_bldet
   '***********************************************************************
   
   sql = "insert into docu_bldet_exp " & _
         "SELECT bldet_linea,   '" & buque_new & "', '" & viaje_new & "', " & _
               " bldet_puer_emb,    bldet_puer_des,      bldet_emp_resp, " & _
               " bldet_cono_emb,     bldet_secuencia,     bldet_tipo_carga, " & _
               " bldet_peso,         bldet_bultos,        bldet_embalaje, " & _
               " bldet_marcas_num,   bldet_desc_carg,     bldet_tipo_prod, bldet_volumen, " & _
               " bldet_contenedor,   bldet_just_cambio,   bldet_IMO, " & _
               " bldet_pelig_ONU,    bldet_sello,         bldet_tam_equi, " & _
               " bldet_tip_equi,     bldet_cond_cont,     bldet_danger_cargo, bldet_un_number, bldet_imdg_page " & _
         " FROM docu_bldet_exp WHERE bldet_linea ='" & linea & "'"
   sql = sql + " and bldet_buque    = '" & buque & "'"
   sql = sql + " and bldet_viaje    = '" & viaje & "'"
   'sql = sql + " and bldet_puer_emb = '" & puerto & "'"
   'sql = sql + " and bldet_puer_des = '" & puerto_des & "'"
   'sql = sql + " and bldet_cono_emb = '" & bl & "'"
   Gcon_docu.Execute sql, no_det
   
   '**********************************************************************
   '   llena docu_cont
   '**********************************************************************
   
   sql = "INSERT INTO docu_cont_exp  " & _
         "SELECT cont_linea,    '" & buque_new & "', '" & viaje_new & "', " & _
                "cont_puer_emb ,     cont_puer_des,      cont_contenedor,  " & _
                "cont_tam_equip,     cont_condicion,     cont_tara,        " & _
                "cont_tipo_equip,    cont_proved_equip,  cont_estad_equip, " & _
                "cont_indic_llen_vac,cont_cond_contrato, cont_reg_serv_transp, " & _
                "cont_natu_carg,     cont_peso,          cont_tipo_bien,   " & _
                "cont_sello1,        cont_sello2,        cont_sello3,      " & _
                "cont_sello4,        cont_temp_dur_almac,cont_temp_dur_trans, " & _
                "cont_temp_dur_manip,cont_temp_min_almac,cont_temp_max_almac, " & _
                "cont_temp_min_trans,cont_temp_max_trans,cont_temp_min_manip, " & _
                "cont_temp_max_manip "
   sql = sql & "  FROM docu_cont_exp " & _
               " WHERE cont_linea ='" & linea & "'" & _
               "   and cont_buque = '" & buque & "'" & _
               "   and cont_viaje = '" & viaje & "'" & _
               "   and cont_contenedor in (Select " & _
               "   distinct(bldet_contenedor) from docu_bldet_exp, docu_cont_exp where " & _
               "   cont_linea ='" & linea & "'" & _
               "   and cont_buque = '" & buque & "'" & _
               "   and cont_viaje = '" & viaje & "'" & _
               "   and cont_linea    = bldet_linea " & _
               "   and cont_buque    = bldet_buque " & _
               "   and cont_viaje    = bldet_viaje " & _
               "   and cont_puer_emb = bldet_puer_emb " & _
               "   and cont_puer_des = bldet_puer_des )" & _
               "   " & _
               "   and cont_contenedor not in " & _
               " (Select cont_contenedor from docu_cont_exp " & _
               "      Where cont_linea ='" & linea & "'" & _
               "        and cont_buque = '" & buque_new & "'" & _
               "        and cont_viaje = '" & viaje_new & "')" ' & _
               "        and cont_puer_emb = '" & puerto_new & "')" ' & _
               "        and cont_puer_des = '" & puerto_des & "')"
   Gcon_docu.Execute sql, no_cont

   'Call actualiza_totales_manifiesto(Cmb_linea.BoundText, Cmb_buque_new.BoundText, Cmb_viaje_new.BoundText, cmb_pto_emb_new.BoundText, puerto_des, reg2("blcab_emp_resp"))

   ' se procesa la parte de borrado del bl original
   If Chk_borra_bl.Value = 1 Then
        sql = "Delete " & _
                    "FROM docu_cont_exp " & _
                    " WHERE cont_linea ='" & linea & "'" & _
                    "   and cont_buque = '" & buque & "'" & _
                    "   and cont_viaje = '" & viaje & "'" & _
                    "   " & _
                    "   and cont_puer_des = '" & puerto_des & "'" & _
                    "   and cont_contenedor in (" & _
                    "       Select bldet_contenedor FROM docu_bldet_exp " & _
                    "       WHERE bldet_linea      = '" & linea & "'" & _
                    "           and bldet_buque    = '" & buque & "'" & _
                    "           and bldet_viaje    = '" & viaje & "'" & _
                    "           " & _
                    "           and bldet_puer_des = '" & puerto_des & "')" '& _
                    "           " ')" '& _
                    "   and cont_contenedor not in (" & _
                    "       Select bldet_contenedor FROM docu_bldet_exp " & _
                    "       WHERE bldet_linea      = '" & linea & "'" & _
                    "           and bldet_buque    = '" & buque & "'" & _
                    "           and bldet_viaje    = '" & viaje & "'" & _
                    "           and bldet_puer_emb = '" & puerto & "'" & _
                    "           and bldet_puer_des = '" & puerto_des & "'" & _
                    "           and bldet_cono_emb <>'" & bl & "')"
        Gcon_docu.Execute sql, no_cont_del
        
        sql = "Delete FROM docu_bldet_exp " & _
                    " WHERE bldet_linea ='" & linea & "'" & _
                    "   and bldet_buque    = '" & buque & "'" & _
                    "   and bldet_viaje    = '" & viaje & "'" '& _
                    "   and bldet_puer_emb = '" & puerto & "'" '& _
                    "   and bldet_puer_des = '" & puerto_des & "'" & _
                    "   and bldet_cono_emb = '" & bl & "'"
        Gcon_docu.Execute sql, no_det_del
        
        sql = "Delete From docu_blcab_exp " & _
                    " WHERE blcab_linea       ='" & linea & "'" & _
                    "   and blcab_buque       ='" & buque & "'" & _
                    "   and blcab_viaje       ='" & viaje & "'" '& _
                    "   and blcab_puer_emb    ='" & puerto & "'" '& _
                    "   and blcab_cono_emb    ='" & bl & "'"
        Gcon_docu.Execute sql, no_cab_del
        
         sql = sql & "Delete From docu_MANIFIESTO_exp "
        sql = sql & " WHERE mani_linea       ='" & linea & "'" & _
                "   and mani_buque       ='" & buque & "'" & _
                "   and mani_viaje       ='" & viaje & "'" '&
        Gcon_docu.Execute sql, no_mani_del
        
        'Call actualiza_totales_manifiesto(Cmb_linea.BoundText, Cmb_buque.BoundText, Cmb_viaje.BoundText, Cmb_pto_emb.BoundText, puerto_des, reg2("blcab_emp_resp"))
   End If ' fin de borrado de bl original
   Gcon_docu.CommitTrans
   msg = "Se Copiaron exitosamente: " & Chr(13) & _
               " - Manifiestos  = " & no_mani & Chr(13) & _
               " - Bls          = " & no_cab & Chr(13) & _
               " - detalles     = " & no_det & Chr(13) & _
               " - contenedores = " & no_cont
   If Chk_borra_bl.Value = 1 Then
        msg = msg & Chr(13) & " y se Eliminaron : " & Chr(13) & _
                              " - Manifietsos  = " & no_mani_del & Chr(13) & _
                              " - Bls          = " & no_cab_del & Chr(13) & _
                              " - detalles     = " & no_det_del & Chr(13) & _
                              " - contenedores = " & no_cont_del
   End If
   MsgBox msg, vbInformation
Else
    MsgBox "La información Original No existe. Por favor verifique", vbInformation
End If

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

Private Sub Cmb_buque_new_Change()
   Dim sql As String
   sql = "Select * from docu_buque_viaje " & _
         " Where buqv_linea ='" & Cmb_linea.BoundText & "'" & _
         "   And buqv_buque ='" & Cmb_buque_new.BoundText & "'" & _
         "   And buqv_status = 'A'"
   
   carga_datacombo data_viaje_new, Cmb_viaje_new, sql, "buqv_viaje", "buqv_viaje"
End Sub

Private Sub Cmb_buque_new_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmb_linea_Change()
carga_datacombo data_buque, Cmb_buque, "Select * from docu_buque Where buqu_linea = '" & Cmb_linea.BoundText & "' and buqu_status ='A'", "buqu_codigo", "buqu_descripcion"
carga_datacombo data_buque_new, Cmb_buque_new, "Select * from docu_buque Where buqu_linea = '" & Cmb_linea.BoundText & "' and buqu_status ='A'", "buqu_codigo", "buqu_descripcion"
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

Private Sub Cmb_viaje_new_Change()
   Dim sql As String
   sql = "Select * from docu_manifiesto " & _
         " Where mani_linea ='" & Cmb_linea.BoundText & "'" & _
         "   And mani_buque ='" & Cmb_buque_new.BoundText & "'" & _
         "   And mani_viaje ='" & Cmb_viaje_new.BoundText & "'"
   
End Sub

Private Sub Cmb_viaje_new_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
    If Not data_linea.Recordset.EOF Then data_linea.Recordset.MoveFirst
End Sub
