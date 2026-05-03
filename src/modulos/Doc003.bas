Attribute VB_Name = "DOC_DTX_KHS"
Dim Gstr_unit_embalaje   As String      ' codigo de embalaje del KHS

' el proceso es asi :
' por cada linea VSL se inicializan todas las estructuras,
' por cada puerto que se inicializan todas las estructuras de blimpo y blimpodt
' y asi sucesivamente
' ESTE ES EL PROCESO PRINCIPAL DE TODO EL PROGRAMA DE INTERFASE DE DTX's
' *****************************************************************************
Public Sub procesar_arch_KHS(Archivo As String)
Dim x
Dim linea       As String
Dim cod_linea   As String
Dim sql         As String
Dim reg_tipo    As New ADODB.Recordset
Dim reg_detalle As New ADODB.Recordset
Dim Tot_lineas  As Integer
Dim avance      As Double

'------------------------------------------
Tot_lineas = Calcula_Lineas_Archivo(Archivo)
avance = 2625 / Tot_lineas
DTX_PROCESO.Lbl_avan.ForeColor = &H404040
DTX_PROCESO.Lbl_avan.Caption = "Procesando ..."
DTX_PROCESO.Lbl_prog.Width = 0
DTX_PROCESO.Lbl_prog.Visible = True
x = FreeFile
Open Archivo For Input As #x
Do While Not EOF(x)
    Line Input #x, linea
    cod_linea = Mid$(linea, 1, 2)
    If DTX_PROCESO.Lbl_prog.Width < 2625 Then
       DTX_PROCESO.Lbl_prog.Width = DTX_PROCESO.Lbl_prog.Width + CDbl(Format(avance, "#####0.000"))
       DTX_PROCESO.Lbl_avan.Caption = Format(DTX_PROCESO.Lbl_prog.Width * 100 / 2625, "##0") & " %"
    End If
    If DTX_PROCESO.Lbl_prog.Width >= 2625 / 2 Then ' para que cambie de color al llegar a la mitad
       DTX_PROCESO.Lbl_avan.ForeColor = &HFFFFFF
    End If
    
    DoEvents
    Select Case cod_linea
       Case "V0":  ' viaje
           Call inicializa_registro(True):              'Call inicializa_KHS_Viaje(True)
           Call inicializa_manifiesto(True)
           Call inicializa_impocont(True)
           Call inicializa_blimpo(True):                'Call inicializa_KHS_Bl(True)
           Call inicializa_blimpodt(True):              'Call inicializa_KHS_Bl_det(True)
           Call Procesa_V0(cod_linea, linea):           'Call Procesa_KHS_V0(cod_linea, linea)   'nuevo
           Call graba_registro:                         'Call graba_KHS_Viaje                  'nuevo
       Case "A0": ' no bl
           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
           Call graba_blimpodt
           'Call graba_KHS_bl_det ' NUEVO
           'If Gstr_Tipo = "I" Then Call actualiza_cabecera Else Call actualiza_cabecera_expo
           Call actualiza_cabecera
           'If Gstr_Tipo = "I" Then Call actualiza_manifiesto
           Call actualiza_manifiesto
           Call inicializa_blimpo(True):                'Call inicializa_KHS_Bl(True)
           Call inicializa_blimpodt(True):              'Call inicializa_KHS_Bl_det(True)
           Call procesa_A0(cod_linea, linea):           'Call Procesa_KHS_A0(cod_linea, linea) ' nuevo
           'If Gstr_Tipo = "I" Then Call graba_manifiesto
           Call graba_manifiesto
       Case "A1": ' shipper name
           Call procesa_A1(cod_linea, linea):           'Call Procesa_KHS_A1(cod_linea, linea) ' nuevo
       Case "A2": ' shipper direccion
           Call procesa_A2(cod_linea, linea):           'Call procesa_KHS_A2(cod_linea, linea) ' nuevo
       Case "A3": ' consignee name
           Call procesa_A3(cod_linea, linea):           'Call procesa_KHS_A3(cod_linea, linea) ' nuevo
       Case "A4": ' consignee address
           Call procesa_A4(cod_linea, linea):           'Call procesa_KHS_A4(cod_linea, linea) ' nuevo
       Case "A5": ' notify name
           Call procesa_A5(cod_linea, linea):           'Call procesa_KHS_A5(cod_linea, linea) ' nuevo
       Case "A6": ' notify addess
           Call procesa_A6(cod_linea, linea):           'Call procesa_KHS_A6(cod_linea, linea) ' nuevo
       Case "F1"  ' costo del flete
           Call procesa_F1(cod_linea, linea):           'Call procesa_KHS_F1(cod_linea, linea) 'nuevo
       Case "F2": ' FREIGHT REMARK   'NUEVO
                                                        'Call procesa_KHS_F2(cod_linea, linea) 'nuevo
       Case "B1"  ' unidad/desc_carga
           Call procesa_B1(cod_linea, linea):           'Call procesa_KHS_B1(cod_linea, linea) 'nuevo
       Case "B2"  ' NUMEROS Y MARCAS
           Call procesa_B2(cod_linea, linea)
       Case "B3"  ' Datos Del Contenedor/DETALLES DEL BL
           r_blimpodt.embalaje = Gstr_unit_embalaje
           Call procesa_B3(cod_linea, linea):           'Call procesa_KHS_B3(cod_linea, linea) 'nuevo
           'If Gstr_Tipo = "I" Then Call graba_impocont
           Call graba_impocont
           'If Gstr_Tipo = "I" Then Call graba_blimpo Else Call graba_blexpo
           Call graba_blimpo
           Call inicializa_blimpo(False):               'Call inicializa_KHS_Bl(False) ' nuevo
           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
           Call graba_blimpodt
           Call inicializa_blimpodt(False):             'Call inicializa_KHS_Bl_det(False)
       Case "DE1": ' detalle de carga
           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
           Call graba_blimpodt
           Call procesa_DE1(cod_linea, linea)
           'If Gstr_Tipo = "I" Then Call graba_blimpo Else Call graba_blexpo
           Call graba_blimpo
           Call inicializa_blimpo(False)
       Case "DE2":
           Call procesa_DE2(cod_linea, linea)
       Case "DE3":
           Call procesa_DE3(cod_linea, linea)
       Case "ZZ":
           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
           Call graba_blimpodt
           'Call graba_KHS_bl_det
           'If Gstr_Tipo = "I" Then Call actualiza_cabecera Else Call actualiza_cabecera_expo
           Call actualiza_cabecera
           'If Gstr_Tipo = "I" Then Call actualiza_manifiesto
           actualiza_manifiesto
    End Select
Loop

DTX_PROCESO.Lbl_prog.Width = 2625
DTX_PROCESO.Lbl_avan.Caption = "100 % "

Close #x

End Sub


Public Sub Procesa_V0(cod_linea As String, linea As String)
  Dim sql           As String
  Dim reg           As New ADODB.Recordset
  Dim tipo_cambio   As Double
  Dim MONEDA        As String

' SE DROPEAN LAS TAZAS DE CAMBIO DE LAS MONEDAS QUE LUEGO SE LLENAN EN EL PCY
   sql = "DELETE FROM DTX_CONVERSION"
   Gcon_docu.Execute sql
   '----------------------------------------------------------------------------

   r_registro.buque = DTX_PROCESO.Cmb_buque.BoundText
   r_registro.viaje = DTX_PROCESO.Txt_viaje.Text
   r_registro.fecha_ing = Format$(Date, GSTR_FORMATO_FECHA)
   'r_registro.Linea_Nav = Gstr_Linea_Nav
   r_registro.atraque = DTX_PROCESO.Msk_atraque.Text
   r_registro.naviera = Gstr_naviera
   r_registro.user_ing = Gstr_user
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'REGISTRO'"
   
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         'Case "VIAJE":
         '   r_registro.viaje = Mid$(linea, reg("desde"), reg("longitud"))
      End Select
      reg.MoveNext
   Loop
   reg.Close
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'MANIFIESTO' and campo_tabla ='FECHA_ZARPE'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "FECHA_ZARPE":
            r_manifiesto.fecha_zarpe = Aplica_Formato(cod_linea, reg("CAMPO"), Mid$(linea, reg("desde"), reg("longitud")))
            r_blimpo.fecha_embarque = r_manifiesto.fecha_zarpe
      End Select
      reg.MoveNext
   Loop
   reg.Close

   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'CONVERSION'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Select Case reg("CAMPO_TABLA")
      Case "TIPO_CAMBIO"
          tipo_cambio = CDbl(Aplica_Formato(cod_linea, reg("CAMPO"), Mid$(linea, reg("desde"), reg("longitud"))))
   End Select
   MONEDA = "DEM"
   
   sql = "Insert into dtx_conversion Values ('" & MONEDA & "'," & tipo_cambio & ", '')"
   Gcon_docu.Execute sql
'---------------------------------
      
   ' graba en las otras tablas los datos comunes
   r_manifiesto.buque = DTX_PROCESO.Cmb_buque.BoundText
   r_manifiesto.viaje = DTX_PROCESO.Txt_viaje.Text
         
   r_impocont.buque = DTX_PROCESO.Cmb_buque.BoundText
   r_impocont.viaje = DTX_PROCESO.Txt_viaje.Text
   
End Sub
' este procedimiento ademas de grabar en la tabla manifiesto,
'graba ciertos datos en la blimpo y blimpodt
'
Public Sub procesa_A0(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset

On Error GoTo Errores
      
   r_manifiesto.no_mani = "1"
   r_manifiesto.No_bls = 0
   r_manifiesto.peso_kls = 0
   r_manifiesto.volumen = 0
   r_manifiesto.no_bultos = 0
   r_manifiesto.status = "A"
   ' debido a una estandarización del numero de bl las dos primeras letras del mismo son
   ' el codigo del puerto
       
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "COD_BL":
            r_manifiesto.puerto = Codigo_equivalente_DTX("DTX_PUERTO", Mid$(Mid$(linea, reg("desde"), reg("longitud")), 1, 2), "APG", "CODIGO")
            r_blimpo.puerto = r_manifiesto.puerto
            r_blimpo.pto_origen = r_manifiesto.puerto
            r_blimpodt.puerto = r_manifiesto.puerto
            r_blimpo.pto_descarga = Codigo_equivalente_DTX("DTX_PUERTO", Mid$(Mid$(linea, reg("desde"), reg("longitud")), 3, 2), "APG", "CODIGO")
            r_blimpo.pto_destino = r_blimpo.pto_descarga
            r_blimpo.COD_BL = Mid$(linea, reg("desde"), reg("longitud"))
            r_blimpodt.COD_BL = Mid$(linea, reg("desde"), reg("longitud"))
            r_blimpodt.pto_descarga = r_blimpo.pto_descarga
            r_impocont.pto_des = r_blimpo.pto_descarga
            r_impocont.puerto = r_blimpo.puerto
            r_blimpo.fecha_embarque = r_manifiesto.fecha_zarpe
            r_manifiesto.pto_des = r_blimpo.pto_descarga
      End Select
      reg.MoveNext
   Loop
   
   
' ----------------------------------------------------------------------
' se toman datos de la linea de puerto para el archivo blimpo y blimpodt
' ----------------------------------------------------------------------
   r_blimpo.buque = DTX_PROCESO.Cmb_buque.BoundText
   r_blimpo.viaje = DTX_PROCESO.Txt_viaje.Text
   r_blimpodt.buque = DTX_PROCESO.Cmb_buque.BoundText
   r_blimpodt.viaje = DTX_PROCESO.Txt_viaje.Text
Exit Sub

Errores:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0

End Sub

Public Sub procesa_A1(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "EMBARCADOR"
            r_blimpo.embarcador = r_blimpo.embarcador & Trim$(Mid$(linea, reg("desde"), reg("longitud")))
      End Select
      reg.MoveNext
   Loop

End Sub

Public Sub procesa_A2(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "DIREC_EXP"
            r_blimpo.direc_exp = r_blimpo.direc_exp & Trim$(Mid$(linea, reg("desde"), reg("longitud")))
      End Select
      reg.MoveNext
   Loop
End Sub

Public Sub procesa_A3(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "CONSIGNATARIO"
            r_blimpo.tipo_id = Gstr_Tipo_ID
            r_blimpo.impexp = Gstr_Impexp
            r_blimpo.nacio_consig = Gstr_param.nacionalidad
            r_blimpo.consignatario = r_blimpo.consignatario & Trim$(Mid$(linea, reg("desde"), reg("longitud")))
      End Select
      reg.MoveNext
   Loop
End Sub

Public Sub procesa_A4(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "DIREC_IMP"
            r_blimpo.direc_imp = r_blimpo.direc_imp & Trim$(Mid$(linea, reg("desde"), reg("longitud")))
      End Select
      reg.MoveNext
   Loop
End Sub

Public Sub procesa_A5(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "NO_NOTYFIER"
            r_blimpo.no_notyfier = r_blimpo.no_notyfier & Trim$(Mid$(linea, reg("desde"), reg("longitud")))
      End Select
      reg.MoveNext
   Loop
   r_blimpo.tipo_id_notif = Gstr_Tipo_ID
   
End Sub

Public Sub procesa_A6(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "DIREC_NOTF"
            r_blimpo.direc_notf = r_blimpo.direc_notf & Trim$(Mid$(linea, reg("desde"), reg("longitud")))
      End Select
      reg.MoveNext
   Loop
End Sub
'aqui se encuentran los codigos de las monedas y sus tipos de cambio con relación al dolar
' -----------------------------------------------------------------------------------
Public Sub procesa_PCY(cod_linea As String, linea As String)
Dim sql         As String
Dim reg         As New ADODB.Recordset
Dim MONEDA      As String
Dim tipo_cambio As String

On Error GoTo error_graba
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'CONVERSION'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "MONEDA"
            MONEDA = Mid$(linea, reg("desde"), reg("longitud"))
         Case "TIPO_CAMBIO"
            tipo_cambio = CDbl(Aplica_Formato(cod_linea, reg("CAMPO"), Mid$(linea, reg("desde"), reg("longitud"))))
      End Select
      reg.MoveNext
   Loop
   
   sql = "Insert into dtx_conversion Values ('" & MONEDA & "'," & tipo_cambio & ")"
   Gcon_docu.Execute sql

Exit Sub
error_graba:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0

End Sub

Public Sub procesa_F1(cod_linea As String, linea As String)
Dim sql    As String
Dim sql1   As String
Dim reg    As New ADODB.Recordset
Dim reg1   As New ADODB.Recordset
Dim MONEDA As String
Dim SIGNO  As String
Dim Costo_Flete As Double

On Error GoTo Errores

   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO' ORDER BY CAMPO asc"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "MONEDA"
            MONEDA = Mid$(linea, reg("desde"), reg("longitud"))
            r_blimpo.MONEDA = Gstr_moneda
         Case "COSTO_FLETE"
            If MONEDA = "USD" Or MONEDA = "DEM" Then
               If MONEDA = "DEM" Then
                  sql1 = "select * from dtx_conversion where moneda = '" & MONEDA & "'"
                  reg1.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                  If reg1.EOF Then
                     MsgBox "No se encontro TIPO/CAMBIO para la moneda: " & MONEDA & " su valor no se tomara en cuenta en el costo del flete", vbInformation, App.Title
                  Else
                     Costo_Flete = CDbl(Aplica_Formato(cod_linea, reg("Campo"), Mid$(linea, reg("desde"), reg("longitud"))))
                     r_blimpo.Costo_Flete = r_blimpo.Costo_Flete + (Costo_Flete * reg1.Fields("valor"))
                  End If
               Else
                  Costo_Flete = CDbl(Aplica_Formato(cod_linea, reg("Campo"), Mid$(linea, reg("desde"), reg("longitud"))))
                  r_blimpo.Costo_Flete = r_blimpo.Costo_Flete + (Costo_Flete)
               End If
               r_blimpo.metodo_pago = Gstr_param.metodo_pago
            End If
      End Select
      reg.MoveNext
   Loop
   
Exit Sub

Errores:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0

End Sub
Public Sub procesa_DE1(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset

   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPODT'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "NO_BULTOS"
            r_blimpodt.no_bultos = Aplica_Formato(cod_linea, reg("CAMPO"), Mid$(linea, reg("desde"), reg("longitud")))
         Case "EMBALAJE"
            r_blimpodt.embalaje = Codigo_equivalente_DTX("DTX_EMBALAJE", Mid$(linea, reg("desde"), reg("longitud")), "APG", "CODIGO", "")
         Case "PESO"
            r_blimpodt.peso = Aplica_Formato(cod_linea, reg("CAMPO"), Mid$(linea, reg("desde"), reg("longitud")))
         Case "VOLUMEN"
            r_blimpodt.volumen = Aplica_Formato(cod_linea, reg("CAMPO"), Mid$(linea, reg("desde"), reg("longitud")))
      End Select
      reg.MoveNext
   Loop
End Sub
Public Sub procesa_B1(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset

   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPODT'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "EMBALAJE"
            'Los chinos envian en el campo de embalaje el tipo de contenedor entonces
            'vamos a enviar por default el embalaje contenerizado para esos casos
            If Left(chk_str(Mid$(linea, reg("desde"), reg("longitud"))), 1) = "2" Or _
               Left(chk_str(Mid$(linea, reg("desde"), reg("longitud"))), 1) = "4" Then
               Gstr_unit_embalaje = Gstr_EMBAL_CONTEN '"CNT"
            ElseIf Trim(Mid$(linea, reg("desde"), reg("longitud"))) <> "" Then
               Gstr_unit_embalaje = Codigo_equivalente_DTX("DTX_EMBALAJE", chk_str(Mid$(linea, reg("desde"), reg("longitud"))), "APG", "CODIGO", "")
            End If
            If Trim$(Gstr_unit_embalaje) = "" Then Gstr_unit_embalaje = Gstr_Embalaje '28 SEPT 2002 SE PONE YA QUE A VECES ESTE CAMPO VIENE VACIO
            r_blimpodt.embalaje = Gstr_unit_embalaje
         Case "DESCRIP_CARGA"
            r_blimpodt.desc_carga = r_blimpodt.desc_carga & Trim$(Mid$(linea, reg("desde"), reg("longitud")))
      End Select
      reg.MoveNext
   Loop

   
End Sub
Public Sub procesa_B2(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset

   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPODT'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "MARCA_NUM"
            r_blimpodt.marca_num = r_blimpodt.marca_num & Trim$(Mid$(linea, reg("desde"), reg("longitud")))
      End Select
      reg.MoveNext
   Loop
End Sub
Public Sub procesa_DE3(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset

   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPODT'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "DANGER"
            r_blimpodt.danger = Codigo_equivalente_DTX("DTX_IMO", Mid$(linea, reg("desde"), reg("longitud")), "APG", "CODIGO", "")
      End Select
      reg.MoveNext
   Loop
End Sub

Public Sub procesa_B3(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
Dim TIPO_CONTEINER As String
Dim SIZE_CONTEINER As String
Dim ORIGEN         As String
Dim DESTINO        As String
   
   r_impocont.buque = r_registro.buque
   r_impocont.viaje = r_registro.viaje
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And codigo ='" & cod_linea & "' and (nombre_tabla = 'IMPOCONT' or nombre_tabla = 'BLIMPODT')"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "NO_BULTOS"
             If IsNumeric(Mid$(linea, reg("desde"), reg("longitud"))) Then
                r_blimpodt.no_bultos = CDbl(Mid$(linea, reg("desde"), reg("longitud")))
             Else
                r_blimpodt.no_bultos = 0
             End If
         Case "CONTEINER"
            r_impocont.CONTEINER = Mid$(linea, reg("desde"), reg("longitud"))
            r_blimpodt.CONTEINER = r_impocont.CONTEINER
         Case "TYPE_CONTEINER"
            TIPO_CONTEINER = Mid$(linea, reg("desde"), reg("longitud"))
         Case "SIZE_CONTEINER"
            SIZE_CONTEINER = Mid$(linea, reg("desde"), reg("longitud"))
'            'If Trim$(SIZE_CONTEINER) = "" And Gstr_Linea = "KHS" Then ' ESTA VALIDACION PARA EL CASO UNICO EN EL QUE EL DTX DE EXPO DE KHS TIENE UN PROBLEMA CON ESTA FILA YA QUE A VECES GENERA UN ESPACIO MAS EN ESTA FILA DESPUES DEL NUMERO DEL CONTENEDOR
'            If Trim$(Gstr_Tipo) = "E" And Gstr_Linea = "KHS" Then ' ESTA VALIDACION PARA EL CASO UNICO EN EL QUE EL DTX DE EXPO DE KHS TIENE UN PROBLEMA CON ESTA FILA YA QUE A VECES GENERA UN ESPACIO MAS EN ESTA FILA DESPUES DEL NUMERO DEL CONTENEDOR
'               linea = Mid(linea, 1, reg("desde") - 1) & Mid(linea, reg("desde") + 1)
'            End If                 ' se comentario esto ya que la aplicacion KHS ya esta bien y genera completa la linea
            SIZE_CONTEINER = Mid$(linea, reg("desde"), reg("longitud"))
            Gint_size_contein = CInt(SIZE_CONTEINER) ' ESTE DATO ES PARA PODER EVALUAR DESPUES EL VOLUMEN POR DEFECTO DE UN CONTENEDOR SEGUN SU TAMAÑO(SI NO EXISTE VOLUMEN)
         Case "TARA"
            r_impocont.tara = IIf(IsNumeric(Mid$(linea, reg("desde"), reg("longitud"))), Mid$(linea, reg("desde"), reg("longitud")), 0)
         Case "SELLO1"
            r_impocont.sello1 = Mid$(linea, reg("desde"), reg("longitud"))
         Case "COND_CONTEIN1":
             ORIGEN = Mid$(linea, reg("desde"), reg("longitud"))
         Case "COND_CONTEIN2":
             DESTINO = Mid$(linea, reg("desde"), reg("longitud"))
         Case "PESO":
             r_blimpodt.peso = CDbl(Aplica_Formato(cod_linea, reg("CAMPO"), Mid$(linea, reg("desde"), reg("longitud"))))
         Case "VOLUMEN"
             r_blimpodt.volumen = CDbl(Aplica_Formato(cod_linea, reg("CAMPO"), Mid$(linea, reg("desde"), reg("longitud"))))
      End Select
      reg.MoveNext
   Loop
            
   If ORIGEN <> "" And DESTINO <> "" Then
      condicion = Codigo_equivalente_DTX("DTX_COND_CONTEIN", ORIGEN, "APG", "origen", DESTINO, " destino ")
      'r_impocont.cond_contein = condicion
      ' se pidio poner fijo el valor de la condicion del contenedor, esto mas en KHS ya que en MSC el valor viene bien
      r_impocont.cond_contein = Gstr_param.condicion_cont
   End If
   
   r_impocont.tamano_equipo = Codigo_equivalente_DTX("DTX_TIPO_CONTEIN", TIPO_CONTEINER, "APG", "TYPE", SIZE_CONTEINER, " SIZE ")
   If r_impocont.tamano_equipo = "NO" Then
      MsgBox "ERROR LA EQUIVALENCIA DE CODIGO DE TIPO DE CONTENEDOR " & SIZE_CONTEINER & "  " & TIPO_CONTEINER
      Exit Sub
   End If
   r_blimpodt.tipo_carga = Gstr_TIPO_CARGA_CONT '"40"
   r_blimpodt.CONTEINER = r_impocont.CONTEINER
   r_impocont.ind_lleno_vacio = "5" 'PONER EN PARAMETRO
   
End Sub

