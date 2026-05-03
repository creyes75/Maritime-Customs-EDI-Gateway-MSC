VERSION 5.00
Begin VB.Form doc_menu_aux 
   Caption         =   "Form1"
   ClientHeight    =   705
   ClientLeft      =   60
   ClientTop       =   915
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   705
   ScaleWidth      =   4785
   Visible         =   0   'False
   Begin VB.Menu mnu_buque 
      Caption         =   "buque"
      Begin VB.Menu mnu_buque_viaje 
         Caption         =   "Nuevo Viaje"
      End
      Begin VB.Menu mnu_buque_modif 
         Caption         =   "Modificar Buque"
      End
      Begin VB.Menu mnu_buque_elim 
         Caption         =   "Eliminar Buque"
      End
   End
   Begin VB.Menu mnu_viaje 
      Caption         =   "viaje"
      Begin VB.Menu mnu_viaje_modif 
         Caption         =   "Modificar Viaje"
      End
      Begin VB.Menu mnu_viaje_elim 
         Caption         =   "Eliminar Viaje"
      End
      Begin VB.Menu mnu_viaje_mani 
         Caption         =   "Nuevo Manifiesto Impo"
      End
      Begin VB.Menu mnu_viaje_mani_expo 
         Caption         =   "Nuevo Manifiesto Expo"
      End
      Begin VB.Menu mnu_buque_mani_cae 
         Caption         =   "Asignar No. Manifiesto CAE (Impo)"
      End
      Begin VB.Menu mnu_buque_mani_cae_expo 
         Caption         =   "Asignar No. Manifiesto CAE (Expo)"
      End
   End
   Begin VB.Menu mnu_mani 
      Caption         =   "manifiestos"
      Begin VB.Menu mnu_mani_modif 
         Caption         =   "Modificar Manifiesto"
      End
      Begin VB.Menu mnu_mani_elim 
         Caption         =   "Eliminar Manifiesto"
      End
      Begin VB.Menu mnu_mani_bl 
         Caption         =   "Nuevo Bl"
      End
   End
   Begin VB.Menu mnu_bl 
      Caption         =   "BL"
      Begin VB.Menu mnu_bl_copia 
         Caption         =   "Copiar Bl"
      End
      Begin VB.Menu mnu_bl_modif 
         Caption         =   "Modificar Bl"
      End
      Begin VB.Menu mnu_bl_elim 
         Caption         =   "Eliminar Bl"
      End
   End
   Begin VB.Menu mnu_mani_expo 
      Caption         =   "Manifiesto Export"
      Begin VB.Menu mnu_mani_expo_modif 
         Caption         =   "Modificar Manifiesto"
      End
      Begin VB.Menu mnu_mani_expo_elim 
         Caption         =   "Eliminar Manifiesto"
      End
      Begin VB.Menu mnu_mani_expo_bl 
         Caption         =   "Nuevo_bl"
      End
   End
   Begin VB.Menu mnu_bl_export 
      Caption         =   "BL Export"
      Begin VB.Menu mnu_bl_export_copia 
         Caption         =   "Copiar Bl"
      End
      Begin VB.Menu mnu_bl_export_modif 
         Caption         =   "Modificar Bl"
      End
      Begin VB.Menu mnu_bl_export_elim 
         Caption         =   "Elimar Bl"
      End
   End
   Begin VB.Menu mnu_bases 
      Caption         =   "Bases"
      Begin VB.Menu mnu_base 
         Caption         =   "mu_base"
         Index           =   0
      End
   End
End
Attribute VB_Name = "doc_menu_aux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnu_base_Click(Index As Integer)
    If mnu_base(Index).Caption <> "" Then
       Call cambio_base(mnu_base(Index).Caption)
    End If
End Sub

Private Sub mnu_bl_copia_Click()
Frm_copia.Show
Frm_copia.Cmb_linea.BoundText = Doc_bl_status.Cmb_linea.BoundText
Frm_copia.Cmb_buque.BoundText = Doc_bl_status.Trv_buque_cod
Frm_copia.Cmb_viaje.BoundText = Doc_bl_status.Trv_viaje
Frm_copia.Cmb_pto_emb.BoundText = Doc_bl_status.Trv_puer_emb
Frm_copia.Cmb_Bl.BoundText = Doc_bl_status.RDC_grid_bl.Recordset("Bl")

End Sub

Private Sub mnu_bl_elim_Click()
Dim frm_bl_mod As New doc_bl

If Not Doc_bl_status.RDC_grid_bl.Recordset.EOF Then
    frm_bl_mod.Show
    frm_bl_mod.Gstr_opc_bl = "M"
    Call frm_bl_mod.carga_bl(Doc_bl_status.Cmb_linea.BoundText, Doc_bl_status.Trv_buque_cod, Doc_bl_status.Trv_viaje, Doc_bl_status.RDC_grid_bl.Recordset("Pto_embarque"), Doc_bl_status.RDC_grid_bl.Recordset("Pto_descarga"), Doc_bl_status.RDC_grid_bl.Recordset("bl"))
    frm_bl_mod.Caption = Doc_bl_status.RDC_grid_bl.Recordset("bl")
    frm_bl_mod.cmd_eliminar_Click
End If
    
'    Call Doc_bl_status.Grid_bl_DblClick
'    doc_bl.cmd_eliminar_Click
End Sub

Private Sub mnu_bl_export_copia_Click()
    doc_expo_copia.Show
    doc_expo_copia.Cmb_linea.BoundText = Doc_bl_status.Cmb_linea.BoundText
    doc_expo_copia.Cmb_buque.BoundText = Doc_bl_status.Trv_buque_cod
    doc_expo_copia.Cmb_viaje.BoundText = Doc_bl_status.Trv_viaje
    doc_expo_copia.Cmb_pto_des.BoundText = Doc_bl_status.Trv_puer_emb
    doc_expo_copia.Cmb_Bl.BoundText = Doc_bl_status.RDC_grid_bl.Recordset("Bl")
End Sub

Private Sub mnu_bl_export_elim_Click()
Dim frm_bl_mod_exp As New doc_expo_bl

If Not Doc_bl_status.RDC_grid_bl.Recordset.EOF Then
    frm_bl_mod_exp.Show
    frm_bl_mod_exp.Gstr_opc_bl = "M"
    Call frm_bl_mod_exp.carga_bl(Doc_bl_status.Cmb_linea.BoundText, Doc_bl_status.Trv_buque_cod, Doc_bl_status.Trv_viaje, Doc_bl_status.RDC_grid_bl.Recordset("Pto_embarque"), Doc_bl_status.RDC_grid_bl.Recordset("Pto_descarga"), Doc_bl_status.RDC_grid_bl.Recordset("bl"))
    frm_bl_mod_exp.Caption = Doc_bl_status.RDC_grid_bl.Recordset("bl")
    frm_bl_mod_exp.cmd_eliminar_Click
End If
End Sub

Private Sub mnu_bl_export_modif_Click()
    Call Doc_bl_status.Grid_bl_DblClick
End Sub

Private Sub mnu_bl_modif_Click()
    Call Doc_bl_status.Grid_bl_DblClick
End Sub

Private Sub mnu_buque_elim_Click()
    doc_mant_buque.Show
    doc_mant_buque.Cmb_linea.BoundText = Doc_bl_status.Cmb_linea.BoundText
    doc_mant_buque.txt_codigo.Text = Doc_bl_status.Trv_buque_cod
    doc_mant_buque.txt_codigo_LostFocus
    DoEvents
    doc_mant_buque.cmd_eliminar_Click
End Sub

Private Sub mnu_buque_Mani_CAE_Click()
    doc_asigna_mani.Show
    doc_asigna_mani.Cmb_linea.BoundText = Doc_bl_status.Cmb_linea.BoundText
    doc_asigna_mani.Cmb_buque.BoundText = Doc_bl_status.Trv_buque_cod ' son variables public en la forma de bl_status
    doc_asigna_mani.Cmb_viaje.BoundText = Doc_bl_status.Trv_viaje     ' son variables public en la forma de bl_status
    doc_asigna_mani.Msk_manifiesto.SetFocus
End Sub

Private Sub mnu_buque_mani_cae_expo_Click()
    doc_expo_asigna_mani.Show
    doc_expo_asigna_mani.Cmb_linea.BoundText = Doc_bl_status.Cmb_linea.BoundText
    doc_expo_asigna_mani.Cmb_buque.BoundText = Doc_bl_status.Trv_buque_cod ' son variables public en la forma de bl_status
    doc_expo_asigna_mani.Cmb_viaje.BoundText = Doc_bl_status.Trv_viaje     ' son variables public en la forma de bl_status
    doc_expo_asigna_mani.Msk_manifiesto.SetFocus
End Sub

Private Sub mnu_buque_modif_Click()
    doc_mant_buque.Show
    doc_mant_buque.Cmb_linea.BoundText = Doc_bl_status.Cmb_linea.BoundText
    doc_mant_buque.txt_codigo.Text = Doc_bl_status.Trv_buque_cod
    doc_mant_buque.txt_codigo_LostFocus
    doc_mant_buque.cmd_modificar_Click
    
End Sub

Private Sub mnu_buque_viaje_Click()
    doc_buque_viaje.Show
    doc_buque_viaje.cmd_agregar_Click
    doc_buque_viaje.Cmb_linea.BoundText = Doc_bl_status.Cmb_linea.BoundText
    doc_buque_viaje.Cmb_vessel.BoundText = Doc_bl_status.Trv_buque_cod
    doc_buque_viaje.Txt_voyage.SetFocus
End Sub

Private Sub mnu_mani_expo_bl_Click()
    Doc_menu.mnu_proc_bl_expo_Click
End Sub

Private Sub mnu_mani_bl_Click()
    Doc_menu.mnu_proc_bl_Click
End Sub

Private Sub mnu_mani_elim_Click()
    doc_manifiesto.Show
    doc_manifiesto.Cmb_linea.BoundText = Doc_bl_status.Cmb_linea.BoundText
    doc_manifiesto.Cmb_buque.BoundText = Doc_bl_status.Trv_buque_cod
    doc_manifiesto.Cmb_viaje.BoundText = Doc_bl_status.Trv_viaje
    doc_manifiesto.Txt_pto_emb.Text = Doc_bl_status.Trv_puer_emb
    doc_manifiesto.Txt_pto_descarga.Text = Gstr_param.pto_descarga
    doc_manifiesto.Txt_pto_emb_LostFocus
    doc_manifiesto.cmd_eliminar_Click
End Sub

Private Sub mnu_mani_expo_elim_Click()
    doc_expo_manifiesto.Show
    doc_expo_manifiesto.Cmb_linea.BoundText = Doc_bl_status.Cmb_linea.BoundText
    doc_expo_manifiesto.Cmb_buque.BoundText = Doc_bl_status.Trv_buque_cod
    doc_expo_manifiesto.Cmb_viaje.BoundText = Doc_bl_status.Trv_viaje
    doc_expo_manifiesto.Txt_pto_emb.Text = Gstr_param.pto_descarga
    doc_expo_manifiesto.Txt_pto_descarga.Text = Doc_bl_status.Trv_puer_emb
    doc_expo_manifiesto.Txt_pto_descarga_LostFocus
    doc_expo_manifiesto.cmd_eliminar_Click
End Sub

Private Sub mnu_mani_expo_modif_Click()
    doc_expo_manifiesto.Show
    doc_expo_manifiesto.Cmb_linea.BoundText = Doc_bl_status.Cmb_linea.BoundText
    doc_expo_manifiesto.Cmb_buque.BoundText = Doc_bl_status.Trv_buque_cod
    doc_expo_manifiesto.Cmb_viaje.BoundText = Doc_bl_status.Trv_viaje
    doc_expo_manifiesto.Txt_pto_emb.Text = Gstr_param.pto_descarga
    doc_expo_manifiesto.Txt_pto_descarga.Text = Doc_bl_status.Trv_puer_emb
    doc_expo_manifiesto.Txt_pto_descarga_LostFocus
    doc_expo_manifiesto.cmd_modificar_Click
End Sub

Private Sub mnu_mani_modif_Click()
    doc_manifiesto.Show
    doc_manifiesto.Cmb_linea.BoundText = Doc_bl_status.Cmb_linea.BoundText
    doc_manifiesto.Cmb_buque.BoundText = Doc_bl_status.Trv_buque_cod
    doc_manifiesto.Cmb_viaje.BoundText = Doc_bl_status.Trv_viaje
    doc_manifiesto.Txt_pto_emb.Text = Doc_bl_status.Trv_puer_emb
    doc_manifiesto.Txt_pto_descarga.Text = Gstr_param.pto_descarga
    doc_manifiesto.Txt_pto_emb_LostFocus
    doc_manifiesto.cmd_modificar_Click
End Sub

Private Sub mnu_viaje_elim_Click()
    doc_buque_viaje.Show
    doc_buque_viaje.Cmb_linea.BoundText = Doc_bl_status.Cmb_linea.BoundText
    doc_buque_viaje.Cmb_vessel.BoundText = Doc_bl_status.Trv_buque_cod
    doc_buque_viaje.Txt_voyage.Text = Doc_bl_status.Trv_viaje
    doc_buque_viaje.Txt_voyage_LostFocus
    DoEvents
    doc_buque_viaje.cmd_eliminar_Click
End Sub

Private Sub mnu_viaje_mani_Click()
    doc_manifiesto.Show
    doc_manifiesto.cmd_agregar_Click
    doc_manifiesto.Cmb_linea.BoundText = Doc_bl_status.Cmb_linea.BoundText
    doc_manifiesto.Cmb_buque.BoundText = Doc_bl_status.Trv_buque_cod
    doc_manifiesto.Cmb_viaje.BoundText = Doc_bl_status.Trv_viaje
    doc_manifiesto.Txt_pto_emb.SetFocus
End Sub

Private Sub mnu_viaje_mani_expo_Click()
    doc_expo_manifiesto.Show
    doc_expo_manifiesto.cmd_agregar_Click
    doc_expo_manifiesto.Cmb_linea.BoundText = Doc_bl_status.Cmb_linea.BoundText
    doc_expo_manifiesto.Cmb_buque.BoundText = Doc_bl_status.Trv_buque_cod
    doc_expo_manifiesto.Cmb_viaje.BoundText = Doc_bl_status.Trv_viaje
    doc_expo_manifiesto.Txt_pto_descarga.SetFocus
End Sub

Private Sub mnu_viaje_modif_Click()
    doc_buque_viaje.Show
    doc_buque_viaje.Cmb_linea.BoundText = Doc_bl_status.Cmb_linea.BoundText
    doc_buque_viaje.Cmb_vessel.BoundText = Doc_bl_status.Trv_buque_cod
    doc_buque_viaje.Txt_voyage.Text = Doc_bl_status.Trv_viaje
    doc_buque_viaje.Txt_voyage_LostFocus
    doc_buque_viaje.cmd_modificar_Click
End Sub
