VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm Doc_menu 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de Documentación Naviero"
   ClientHeight    =   6315
   ClientLeft      =   1230
   ClientTop       =   2610
   ClientWidth     =   9165
   Icon            =   "Doc_Menu.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "buque"
            Object.ToolTipText     =   "Nuevo Buque"
            Object.Tag             =   ""
            ImageKey        =   "buque"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "viaje"
            Object.ToolTipText     =   "Nuevo Buque Viaje"
            Object.Tag             =   ""
            ImageKey        =   "viaje"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "manifiesto"
            Object.ToolTipText     =   "Nuevo Manifiesto"
            Object.Tag             =   ""
            ImageKey        =   "manifiesto"
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "bl_status"
            Object.ToolTipText     =   "Bl Status"
            Object.Tag             =   ""
            ImageKey        =   "bl_status"
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "bl"
            Object.ToolTipText     =   "Bill of Lading"
            Object.Tag             =   ""
            ImageKey        =   "bl"
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "sep1"
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "DTX"
            Object.ToolTipText     =   "Procesar archivo DTX"
            Object.Tag             =   ""
            ImageKey        =   "DTX"
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CAE"
            Object.ToolTipText     =   "Generar Informacion CAE"
            Object.Tag             =   ""
            ImageKey        =   "CAE"
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "sep2"
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del sistema"
            Object.Tag             =   ""
            ImageKey        =   "salir"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   6045
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            TextSave        =   ""
            Key             =   "general"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            TextSave        =   "NUM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "version"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "base: documento"
            TextSave        =   "base: documento"
            Key             =   "base"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "sa"
            TextSave        =   "sa"
            Key             =   "user"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "23:59"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doc_Menu.frx":030A
            Key             =   "bl"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doc_Menu.frx":0624
            Key             =   "CAE"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doc_Menu.frx":093E
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doc_Menu.frx":0C58
            Key             =   "bl_status"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doc_Menu.frx":0F72
            Key             =   "historia"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doc_Menu.frx":128C
            Key             =   "DTX"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doc_Menu.frx":15A6
            Key             =   "buque"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doc_Menu.frx":18C0
            Key             =   "viaje"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doc_Menu.frx":1BDA
            Key             =   "manifiesto"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_mant 
      Caption         =   "Mantenimientos"
      Begin VB.Menu mnu_mant_linea 
         Caption         =   "Linea"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_puerto 
         Caption         =   "Puerto"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_buque 
         Caption         =   "Buque"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_separa1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_mant_almacen 
         Caption         =   "Almacen"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_cond_cont 
         Caption         =   "Condición Contenedor"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_cond 
         Caption         =   "Condición Contrato"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_cons 
         Caption         =   "Consolidadora"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_depo 
         Caption         =   "Deposito"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_emba 
         Caption         =   "Embalaje"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_emp_anav 
         Caption         =   "Empresa Naviera"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_emp_resp 
         Caption         =   "Empresa Responsable"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_impexp 
         Caption         =   "Importadores/ Exportadores"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_indic_llen_vac 
         Caption         =   "Indicador Lleno Vacío"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_meto_pago 
         Caption         =   "Método Pago"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_moneda 
         Caption         =   "Moneda"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_nacio 
         Caption         =   "Nacionalidad"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_natura_carga 
         Caption         =   "Naturaleza Carga"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_imo 
         Caption         =   "Peligrosidad (IMO)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_imo_onu 
         Caption         =   "Peligrosidad (IMO ONU)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_prov_equip 
         Caption         =   "Proveedor Equipamiento"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_regimen 
         Caption         =   "Régimen Aduanero"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_requer_serv_trans 
         Caption         =   "Requer. Serv. Transp."
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_tamano_equip 
         Caption         =   "Tamaño Equipo(Tipo cont)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_tipo_carga 
         Caption         =   "Tipo Carga"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_tipo_deposito 
         Caption         =   "Tipo Deposito"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_tipo_doc 
         Caption         =   "Tipo Documento"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_tipo_envio 
         Caption         =   "Tipo Envio"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_tipo_Equipo 
         Caption         =   "Tipo Equipo"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_mant_tipo_salida 
         Caption         =   "Tipo Salida"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_proceso 
      Caption         =   "Procesos"
      Begin VB.Menu mnu_proc_viaje 
         Caption         =   "Buque_viaje"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_proc_bl_status 
         Caption         =   "Bl Status"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_proc_impo 
         Caption         =   "Importación"
         Visible         =   0   'False
         Begin VB.Menu mnu_proc_manifiesto 
            Caption         =   "Manifiesto"
         End
         Begin VB.Menu mnu_proc_bl 
            Caption         =   "Bl"
         End
         Begin VB.Menu mnu_proc_contenedor 
            Caption         =   "Contenedores"
         End
         Begin VB.Menu mnu_proc_busq_bl 
            Caption         =   "Búsqueda de Bl"
         End
         Begin VB.Menu copia_bl 
            Caption         =   "Copia Bl"
         End
         Begin VB.Menu mnu_proc_copia_pto 
            Caption         =   "Copia de Bls por Puerto"
         End
         Begin VB.Menu mnu_proc_copia_viaje 
            Caption         =   "Copia de Bls por Viaje"
         End
      End
      Begin VB.Menu mnu_proc_expo 
         Caption         =   "Exportación"
         Visible         =   0   'False
         Begin VB.Menu mnu_proc_manifiesto_expo 
            Caption         =   "Manifiesto de Exportación"
         End
         Begin VB.Menu mnu_proc_bl_expo 
            Caption         =   "Bl de Exportación"
         End
         Begin VB.Menu mnu_proc_contenedor_expo 
            Caption         =   "Contenedores de Exportación"
         End
         Begin VB.Menu mnu_proc_busq_bl_expo 
            Caption         =   "Búsqueda de Bl"
         End
         Begin VB.Menu mnu_proc_copia_bl_expo 
            Caption         =   "Copia Bl"
         End
         Begin VB.Menu mnu_proc_copia_bl_x_pto_expo 
            Caption         =   "Copia de Bl por Puerto"
         End
         Begin VB.Menu mnu_proc_copia_bl_x_viaje_expo 
            Caption         =   "Copia de Bl por Viaje"
         End
         Begin VB.Menu mnu_proc_gene_dtx 
            Caption         =   "Generación DTX"
         End
      End
      Begin VB.Menu mnu_proc_separa 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_proc_archivo 
         Caption         =   "Archivo CAE"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_proc_respuesta 
         Caption         =   "Respuesta CAE"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_proc_cons_error 
         Caption         =   "Consulta Errores CAE"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_proc_dtx 
         Caption         =   "Proceso DTX"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_reportes 
      Caption         =   "Reportes"
      Begin VB.Menu mnu_repo_manifiesto 
         Caption         =   "Manifiesto"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_repo_blWord 
         Caption         =   "BL Word"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_repo_carta_impo 
         Caption         =   "Carta de Importación"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_repo_correc 
         Caption         =   "Carta de Corrección de Bls"
      End
      Begin VB.Menu contimpo 
         Caption         =   "Contenedores Importación"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_repo_cont_sin_bl 
         Caption         =   "Contenedores Importación sin bl"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_repo_aviso 
         Caption         =   "Aviso Llegada"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_repo_flete_impo 
         Caption         =   "Reporte de Fletes de Importación"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_repo_separa 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_repo_bl_export 
         Caption         =   "Bl Exportación"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_repo_blWord_expo 
         Caption         =   "Bl Word Exportación"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_repo_mani_expo 
         Caption         =   "Manifiesto Exportación"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_repo_carta_expo 
         Caption         =   "Carta de Exportación"
      End
      Begin VB.Menu mnu_repo_cont_expo 
         Caption         =   "Contenedores Exportación"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_repo_cont_sin_bl_expo 
         Caption         =   "Contenedores Exportación sin bl"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_repo_flete_Expo 
         Caption         =   "Reporte de Fletes de Exportación"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_ventana 
      Caption         =   "Ventana"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnu_acerca_de 
      Caption         =   "Acerca De..."
   End
   Begin VB.Menu mnu_salir 
      Caption         =   "Salir"
      Begin VB.Menu mnu_salir_si 
         Caption         =   "Si"
      End
      Begin VB.Menu mnu_salir_no 
         Caption         =   "No"
      End
   End
   Begin VB.Menu mnu_adm 
      Caption         =   "Administración"
      Visible         =   0   'False
      Begin VB.Menu mnu_adm_plano_param 
         Caption         =   "Plano Parámetros"
      End
      Begin VB.Menu mnu_adm_docu_param 
         Caption         =   "Documentos Parámetros"
      End
      Begin VB.Menu mnu_adm_dtx_param 
         Caption         =   "Dtx Parámetros"
      End
      Begin VB.Menu mnu_adm_menu 
         Caption         =   "Permisos Menú"
      End
      Begin VB.Menu mnu_adm_depura 
         Caption         =   "Depura Información"
      End
      Begin VB.Menu mnu_adm_historica 
         Caption         =   "Cambia a base Historica"
      End
   End
End
Attribute VB_Name = "Doc_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub contimpo_Click()
doc_repo_cont.Show
End Sub

Private Sub copia_bl_Click()
    Frm_copia.Show
End Sub

Private Sub MDIForm_Load()
    Call Carga_Menus(Gstr_user)
    
    Doc_menu.StBar.Panels(4).Text = "Versión: " & App.Major & "." & App.Minor & " r" & App.Revision
    Me.StBar.Panels(5).Text = "Base: " & Gstr_base_doc
    Me.StBar.Panels(6).Text = Gstr_user
    
    Dim reg As New ADODB.Recordset
    reg.Open "Select * from master..sysdatabases where name like '" & Mid(Gstr_base_doc, 1, 4) & "%'", Gcon_docu
    i = 0
    Do While Not reg.EOF
        i = i + 1
        If i > 1 Then Load doc_menu_aux.mnu_base(i)
        If i > 1 Then doc_menu_aux.mnu_base(i).Caption = reg("name") Else doc_menu_aux.mnu_base(i - 1).Caption = reg("name")
        reg.MoveNext
    Loop
    'carga_datacombo data_linea, Cmb_linea, "Select * from master..sysdatabases where name like '" & Mid(Gstr_base_doc, 1, 4) & "%'", "name", "name"
    
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnu_acerca_de_Click()
    Doc_About.Show
End Sub

Private Sub mnu_adm_depura_Click()
    Admin_depura.Show
End Sub

Private Sub mnu_adm_docu_param_Click()
    Admin_mant_docu_param.Show
End Sub

Private Sub mnu_adm_dtx_param_Click()
    Admin_mant_dtx_param.Show
End Sub

Private Sub mnu_adm_historica_Click()
    Admin_cambia_historico.Show
End Sub

Private Sub mnu_adm_menu_Click()
    Admin_mant_menu.Show
End Sub

Private Sub mnu_adm_plano_param_Click()
    Admin_mant_plano_param.Show
End Sub

Private Sub mnu_mant_almacen_Click()
    doc_mant_almacen.Show
End Sub

Private Sub mnu_mant_buque_Click()
    doc_mant_buque.Show
End Sub

Private Sub mnu_mant_cond_Click()
    doc_mant_cond.Show
End Sub

Private Sub mnu_mant_cond_cont_Click()
    doc_mant_cond_cont.Show
End Sub

Private Sub mnu_mant_cons_Click()
    doc_mant_cons.Show
End Sub

Private Sub mnu_mant_depo_Click()
    doc_mant_depo.Show
End Sub

Private Sub mnu_mant_emba_Click()
    doc_mant_emba.Show
End Sub

Private Sub mnu_mant_emp_anav_Click()
    doc_mant_emp_anav.Show
End Sub

Private Sub mnu_mant_emp_resp_Click()
    doc_mant_emp_resp.Show
End Sub

Private Sub Mnu_mant_IMO_Click()
    doc_mant_imo.Show
End Sub

Private Sub mnu_mant_imo_onu_Click()
    doc_mant_imo_onu.Show
End Sub

Private Sub mnu_mant_impexp_Click()
    doc_mant_consig.Show
End Sub

Private Sub mnu_mant_indic_llen_vac_Click()
    doc_mant_indic_llen_vac.Show
End Sub

Private Sub mnu_mant_linea_Click()
    doc_mant_linea.Show
End Sub

Private Sub mnu_mant_meto_pago_Click()
    doc_mant_metodo_pago.Show
End Sub

Private Sub mnu_mant_moneda_Click()
    doc_mant_moneda.Show
End Sub

Private Sub mnu_mant_nacio_Click()
    doc_mant_nacionalidad.Show
End Sub

Private Sub mnu_mant_natura_carga_Click()
    doc_mant_natura_carga.Show
End Sub

Private Sub mnu_mant_prov_equip_Click()
    doc_mant_prov_equip.Show
End Sub

Private Sub mnu_mant_puerto_Click()
    doc_mant_puerto.Show
End Sub

Private Sub mnu_mant_regimen_Click()
    doc_mant_regimen.Show
End Sub

Private Sub mnu_mant_requer_serv_trans_Click()
    doc_mant_requer_serv_trans.Show
End Sub

Private Sub mnu_mant_tamano_equip_Click()
    doc_mant_tamano_equip.Show
End Sub

Private Sub mnu_mant_tipo_bien_Click()
    doc_mant_tipo_bien.Show
End Sub

Private Sub mnu_mant_tipo_carga_Click()
    doc_mant_tipo_carga.Show
End Sub

Private Sub mnu_mant_tipo_deposito_Click()
    doc_mant_tipo_deposito.Show
End Sub

Private Sub mnu_mant_tipo_doc_Click()
    doc_mant_tipo_documento.Show
End Sub

Private Sub mnu_mant_tipo_envio_Click()
    doc_mant_tipo_envio.Show
End Sub

Private Sub mnu_mant_tipo_Equipo_Click()
    doc_mant_tipo_equipo.Show
End Sub

Private Sub mnu_mant_tipo_salida_Click()
    doc_mant_tipo_salida.Show
End Sub

Private Sub mnu_proc_archivo_Click()
    doc_archivo_CAE.Show
End Sub

Public Sub mnu_proc_bl_Click()
Dim forma As New doc_bl
   forma.Show
   forma.Gstr_opc_bl = "A"
   'Gstr_opc_bl = "A"
   'doc_bl.Show
End Sub

Public Sub mnu_proc_bl_expo_Click()
Dim forma As New doc_expo_bl
   forma.Show
   forma.Gstr_opc_bl = "A"
End Sub

Private Sub mnu_proc_bl_status_Click()
    Doc_bl_status.Show
End Sub

Private Sub mnu_proc_busq_bl_Click()
    doc_aux_busq_bl.Show
End Sub

Private Sub mnu_proc_busq_bl_expo_Click()
    doc_aux_busq_bl_expo.Show
End Sub

Private Sub mnu_proc_cons_error_Click()
    doc_aux_cons_Error.Show
End Sub

Private Sub mnu_proc_contenedor_Click()
    Dim frm_cont As New doc_contenedor
    frm_cont.Show
    frm_cont.Gbol_CONT_desde_menu = True
End Sub

Private Sub mnu_proc_contenedor_expo_Click()
    Dim frm_cont As New doc_expo_contenedor
    frm_cont.Show
    frm_cont.Gbol_CONT_desde_menu = True
End Sub

Private Sub mnu_proc_copia_bl_expo_Click()
    doc_expo_copia.Show
End Sub

Private Sub mnu_proc_copia_bl_x_pto_expo_Click()
    doc_expo_copia_pto.Show
End Sub

Private Sub mnu_proc_copia_bl_x_viaje_expo_Click()
    doc_expo_copia_viaje.Show
End Sub

Private Sub mnu_proc_copia_pto_Click()
    doc_copia_pto.Show
End Sub

Private Sub mnu_proc_copia_viaje_Click()
    doc_copia_viaje.Show
End Sub

Private Sub mnu_proc_dtx_Click()
    DTX_PROCESO.Show
End Sub

Private Sub mnu_proc_gene_dtx_Click()
    doc_archivo_Expo.Show
End Sub

Private Sub mnu_proc_manifiesto_expo_Click()
    doc_expo_manifiesto.Show
End Sub

Private Sub mnu_proc_plano_Click()
    doc_archivo_plano.Show
End Sub

Private Sub mnu_proc_viaje_Click()
   doc_buque_viaje.Show
End Sub

Private Sub mnu_proc_manifiesto_Click()
    doc_manifiesto.Show
End Sub

Private Sub mnu_proc_respuesta_Click()
    doc_respuesta_CAE.Show
End Sub

Private Sub mnu_repo_aviso_Click()
    doc_repo_aviso_llegada.Show
End Sub

Private Sub mnu_repo_bl_export_Click()
    doc_repo_expo.Show
End Sub

Private Sub mnu_repo_blWord_Click()
   doc_repo_bl.Show
End Sub

Private Sub mnu_repo_blWord_expo_Click()
    doc_repo_expo_bl.Show
End Sub

Private Sub mnu_repo_carta_expo_Click()
    doc_repo_carta_expo.Show
End Sub

Private Sub mnu_repo_carta_impo_Click()
    doc_repo_carta_impo.Show
End Sub

Private Sub mnu_repo_cont_expo_Click()
    doc_repo_expo_cont.Show
End Sub

Private Sub mnu_repo_cont_sin_bl_Click()
    doc_repo_cont_sin_bl.Show
End Sub

Private Sub mnu_repo_cont_sin_bl_expo_Click()
    doc_repo_expo_cont_sin_bl.Show
End Sub

Private Sub mnu_repo_correc_Click()
    doc_repo_correc.Show
End Sub

Private Sub mnu_repo_flete_Expo_Click()
    doc_repo_flete_expo.Show
End Sub

Private Sub mnu_repo_flete_impo_Click()
    doc_repo_flete_impo.Show
End Sub

Private Sub mnu_repo_mani_expo_Click()
    doc_repo_expo_mani.Show
End Sub

Private Sub mnu_repo_manifiesto_Click()
    doc_repo_mani.Show
End Sub

Private Sub mnu_salir_si_Click()
Dim reg As New ADODB.Recordset

'sql = "select * from dtx_puerto, docu_puerto where linea ='KHS' and puer_linea =linea and puer_descripcion = descripcion"
'sql = "select * from dtx_puerto where linea ='KHS'"
'reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
'
'        Do While Not reg.EOF
'            sql = "Update docu_puerto set puer_codigo_alt = '" & Trim(reg("codigo")) & "' Where " & _
'                  " puer_linea ='KHS' and puer_descripcion ='" & Trim(reg("descripcion")) & "'"
'            Gcon_docu.Execute sql, i
'            If i > 0 Then cont = cont + 1
'            reg.MoveNext
'        Loop
    
    End
End Sub

Private Sub StBar_PanelClick(ByVal Panel As ComctlLib.Panel)
If Panel.Key = "base" Then
    PopupMenu doc_menu_aux.mnu_bases
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   Select Case LCase(Button.Key)
        Case "buque":       If mnu_mant_buque.Visible Then Call mnu_mant_buque_Click
        Case "viaje":       If mnu_proc_viaje.Visible Then Call mnu_proc_viaje_Click
        Case "manifiesto":  If mnu_proc_manifiesto.Visible Then Call mnu_proc_manifiesto_Click
        Case "bl_status":   Call mnu_proc_bl_status_Click
        Case "bl":          If mnu_proc_bl.Visible Then Call mnu_proc_bl_Click
        Case "dtx":         If mnu_proc_dtx.Visible Then Call mnu_proc_dtx_Click
        Case "cae":         If mnu_proc_archivo.Visible Then Call mnu_proc_archivo_Click
        Case "salir":       Call mnu_salir_si_Click
   End Select
End Sub
