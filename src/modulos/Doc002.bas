Attribute VB_Name = "DOC_DTX_MSC"

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'//todos estos salen de los parametros de DTX */
Public Tipo_rela_APG        As String
Public Gstr_naviera         As String
Public Gstr_Linea           As String
Public Gstr_Nombre_naviera  As String
Public Gstr_user_DTX        As String
Public Gstr_moneda          As String
Public Gstr_tipo_alma       As String
Public Gstr_bodega          As String
Public Gstr_Proposito       As String
Public Gstr_Proposito_expo  As String
Public Gstr_Tipo_ID         As String
Public Gstr_Impexp          As String
'Public Gstr_nav_cons_dui As String
Public Gstr_capitan         As String
'Public Gstr_carga          As String
'Public Gstr_Med_Peso       As String
'Public Gstr_Med_Volumen  As String
Public Gstr_Cond_Contein    As String
Public Gstr_Embalaje        As String
Public Gstr_Puerto          As String
Public Gstr_IMO             As String
Public Gstr_Tipo_Contein    As String
Public Gstr_Tipo_Contein_40 As String
'Public Gstr_Linea_Nav      As String         ' identifica el código para la linea naviera 'MSC o KHS
Public Gstr_Cargo           As String
Public Gstr_Debtor          As String
Public Gstr_trasbordo       As String

Public Gstr_Linea_Nav_APG As String     ' identifica el código APG para la linea naviera '  8390 para mediteranean shipping   ' 9993 para KHS

Public Gstr_TIPO_CARGA_CONT_VACIO   As String
Public Gstr_TIPO_CARGA_CONT         As String
Public Gstr_TIPO_CARGA_GRAL         As String
Public Gstr_EMBAL_CONTEN            As String
Public Gstr_EMBAL_BULTOS            As String

Public Gint_size_contein As Integer   ' definida para almacenar el tamaño del contenedor
Public Gint_volu_20      As Integer   ' y dependiendo de eso, si el volumen viene con cero
Public Gint_volu_40      As Integer   ' poner la variable Gint_volu_20 o Gint_volu_40, que
                                      ' tienen los valores default tomados de la tabla de parametros
Public Gdbl_peso_vacio_20 As Double   ' valor por defecto para el peso en los contenedores vacíos de 20 (esto esta parametrizado en la tabla)
Public Gdbl_peso_vacio_40 As Double   ' valor por defecto para el peso en los contenedores vacíos de 40 (esto esta parametrizado en la tabla)
Public Gstr_Cond_Contein_vacio As String ' valor de la condidiciion del contenedor vacio
Public Gstr_sello_Contein As String      ' valor del  sello de los contenedores si viene vacio
Public Gdbl_T_Cambio_DEM As Double    ' VALOR DEL TIPO DE CAMBIO DE LOS "DEM" (MARCOS ALEMANES)
Public Gstr_Tipo         As String    ' identifica si el proceso de carga es de Importación o Exportación
Public Gint_no_Bls       As Integer   ' Número de Bls procesados
Public Gint_no_Bls_det   As Integer   ' Número de Detalles de Bls procesados
Public Gstr_Pto_desc     As String
Public Gstr_Version_MSC  As String    ' contiene la version del archivo que se esta procesando, esto es importante para poder rechazar archivos que sean de una version previa
Public Gbol_bl_detallado_MSC As Boolean  ' se setea a TRUE cuando el proceso de DTX para MSC graba el detalle de cada contenedor por cada tipo de mercaderia; o a FALSE cuando graba un solo detalle de bl por todo el contenedor(AUNQUE EXISTAN VARIOS DETALLES )
Public Gbol_solo_GYE      As Boolean  'para saber si proceso solo la informacion de guayaquil
Public Gbol_relacion_fija As Boolean  ' se usa para establecer si la relacion de codigos de la CAE con los códigos MSC se setean por defecto si no existen o si se pregunta para establecer la relacion y que presente la pantalla, (Este valor se setea en la ventana de proceso DTX
Public Gstr_TIPO_PRODUCTO As String   ' se usa para mantener el valor por defecto de producto a cargar en el DTXb          '49000
Public Gstr_TIPO_PRODUCTO_VACIO As String ' se usa para mantener el valor por defecto de producto vacio a cargar en el DTX '49500
'usados para Proceso de DTX '   Julio / 2008
Public Gstr_Desc_carga      As String     ' se incluyo para poner un texto fijo en la descripcion de la mercaderia  'Julio/2008
Public Gstr_Marca_num       As String     ' se incluyo para poner un texto fijo en las Marcas y numeros             'Julio/2008


Public P_RutaDeMIEIni  As String
Public P_RUTA_MIECLTES As String
Public P_TipoInst      As String

Type REGISTRO
   naviera   As String * 3
   'Linea_Nav As String
   buque     As String
   viaje     As String
   direcc    As String
   atraque   As Date
   fecha_ing As Date
   user_ing  As String
   fecha_mod As String
   user_mod  As String
End Type

Type manifiesto
   buque     As String
   viaje     As String
   puerto      As String
   pto_des     As String
   no_mani     As String
   fecha_zarpe As Date
   No_bls      As Double
   peso_kls    As Double
   volumen     As Double
   no_bultos   As Double
   no_cont     As Double
   capitan     As String
   pto_inicio  As String
   status      As String
End Type

Type impocont
   buque     As String
   viaje     As String
   puerto      As String
   pto_des     As String
   CONTEINER   As String
   tamano_equipo As String
   'tipo_contein As String
   cond_contein As String
   'bls_Asociados As String
   tara          As Double
'   tipo_equipo   As String
'   proveedor_equipo As String
'   estado_equipo As String
   ind_lleno_vacio As String
'   cond_contrato As String
'   serv_transp   As String
'   natu_carga    As String
   peso          As String
   tipo_bien     As String
   sello1        As String
   sello2        As String
   sello3        As String
   sello4        As String
   temp_dur_almac As Integer
   temp_dur_trans As Integer
   temp_dur_manip As Integer
   temp_min_almac As Integer
   temp_max_almac As Integer
   temp_min_trans As Integer
   temp_max_trans As Integer
   temp_min_manip As Integer
   temp_max_manip As Integer
End Type

Type blimpo
   buque          As String
   viaje          As String
   puerto         As String
   pto_descarga   As String
   'emp_resp       As String
   COD_BL         As String
   'emp_anav       As String
   'tipo_doc       As String
   cono_master    As String
   peso_kls       As Double
   no_bultos      As Double
   tipo_id        As String * 1
   Impexp         As String
   consignatario  As String
   direc_imp      As String
   nacio_consig   As String
   tipo_id_notif  As String * 1
   id_notifier    As String
   no_notyfier    As String
   direc_notf     As String
   embarcador     As String
   direc_exp      As String
   pto_origen     As String
   almac_temp     As String
   tipo_deposito  As String
   deposito_adua  As String
   regimen_adua   As String
   consolidadora  As String
   tipo_salida    As String
   docum_autor    As String
   peso_mal_estado As Double
   bultos_mal_estado As Double
   no_notyfier2   As String
   direc_notf2    As String
   no_cont        As Double
   volumen_mt3    As Double
   pto_destino    As String
   metodo_pago    As String
   MONEDA         As String
   Costo_Flete    As Double
   otros_gastos   As Double
   MONEDA_otros   As String
   seguro         As Double
   MONEDA_seguro  As String
   metodo_pago_otros As String
   fecha_embarque As Date
   
'   nav_cons_dui   As String * 1
'   cod_bl_padre   As String
'   desc_carga     As String
'
'   no_cont20      As Double
'   no_cont40      As Double
'
'   pto_descarga   As String
'   comentario     As String
'
'   st_padre       As String * 1
'   status         As String * 1
'   fecha_ing      As Date
'   user_ing       As String
'   fecha_mod      As Date
'   user_mod       As String
End Type

Type blimpodt
   buque          As String
   viaje          As String
   puerto         As String
   pto_descarga   As String
   'emp_resp       As String
   COD_BL         As String
   sec_carga      As Integer
   tipo_carga     As String
   peso           As Double
   no_bultos      As Double
   embalaje       As String
   marca_num      As String
   desc_carga     As String
   volumen        As Double
   CONTEINER      As String
   danger         As String
   IMO_ONU        As String
   producto       As String
End Type

Public r_registro    As REGISTRO
Public r_manifiesto  As manifiesto
Public r_impocont    As impocont
Public r_blimpo      As blimpo
Public r_blimpodt    As blimpodt
Public GLinea_act_Grid As Integer ' usada para determinar la linea actual del Spread

' el proceso es asi :
' por cada linea VSL se inicializan todas las estructuras,
' por cada puerto que se inicializan todas las estructuras de blimpo y blimpodt
' y asi sucesivamente
' ESTE ES EL PROCESO PRINCIPAL DE TODO EL PROGRAMA DE INTERFASE DE DTX's
' *****************************************************************************
' 31 marzo 2004 se hacen modificaciones para recuperar lo perdido en daño de comput. mas de un sello.
'
'
Public Sub procesar_arch(Archivo As String)
Dim x
Dim linea       As String
Dim cod_linea   As String
Dim sql         As String
Dim reg_tipo    As New ADODB.Recordset
Dim reg_detalle As New ADODB.Recordset
Dim Tot_lineas  As Integer
Dim avance      As Double

'------------------------------------------
Gbol_solo_GYE = False
'If Gstr_Tipo = "I" Then
'    If MsgBox("Desea procesar solamente la carga para Guayaquil?", vbYesNo + vbQuestion) = vbYes Then
'       Gbol_solo_GYE = True
'    Else
       Gbol_solo_GYE = False
'    End If
'End If

'carga_parametros cmb_linea.BoundText

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
    cod_linea = Mid$(linea, 1, 3)
    If DTX_PROCESO.Lbl_prog.Width < 2625 Then
       DTX_PROCESO.Lbl_prog.Width = DTX_PROCESO.Lbl_prog.Width + CDbl(Format(avance, "#####0.000"))
       DTX_PROCESO.Lbl_avan.Caption = Format(DTX_PROCESO.Lbl_prog.Width * 100 / 2625, "##0") & " %"
    End If
    If DTX_PROCESO.Lbl_prog.Width >= 2625 / 2 Then ' para que cambie de color al llegar a la mitad
       DTX_PROCESO.Lbl_avan.ForeColor = &HFFFFFF
    End If
    
    DoEvents
    Select Case cod_linea
       Case "VSL":
           Call inicializa_registro(True)
           Call inicializa_manifiesto(True)
           Call inicializa_impocont(True)
           Call inicializa_blimpo(True)
           Call inicializa_blimpodt(True)
           Call Procesa_VSL(cod_linea, linea)
           
           If Gstr_Version_MSC <> "53" And Gstr_Version_MSC <> "54" Then
              If MsgBox("La versión del archivo a ser procesado no corresponde con la versión de este programa, pueden existir problemas en el proceso, " & Chr(13) & "    ¿Desea continuar de todas maneras (S/N)? ", vbYesNo, App.Title) = vbNo Then
                 Exit Sub
              End If
           End If
           Call graba_registro
       Case "POR":
           'If Gstr_Tipo = "I" Then Call graba_fletes
           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
           'If Gstr_Tipo = "I" Then Call actualiza_cabecera Else Call actualiza_cabecera_expo
           'If Gstr_Tipo = "I" Then Call actualiza_manifiesto
           Call graba_fletes
           Call graba_impocont              'MARZO 31 2004
           Call graba_blimpodt
           Call actualiza_cabecera
           Call actualiza_manifiesto
           Call inicializa_blimpo(True)
           Call inicializa_blimpodt(True)
           Call inicializa_impocont(True)  'MARZO 31 2004
           Call inicializa_fletes
           Call procesa_POR(cod_linea, linea)
           'If Gstr_Tipo = "I" Then Call graba_manifiesto
           Call graba_manifiesto
       Case "PCY":
           Call procesa_PCY(cod_linea, linea)
           'AQUI FALTA PONER LA PANTALLA QUE PIDA CONFIRMAR LOS VALORES O CAMBIARLOS PARA EL TIPO DE CAMBIO
       Case "BL1":
           Call Confirma_Conversion
           'If Gstr_Tipo = "I" Then Call graba_fletes
           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
           'If Gstr_Tipo = "I" Then Call actualiza_cabecera Else actualiza_cabecera_expo
           Call graba_fletes
           Call graba_blimpodt
           Call actualiza_cabecera
           Call inicializa_blimpodt(False)
           Call inicializa_fletes
           Call procesa_BL1(cod_linea, linea)
           'AQUI TALVEZ SE DEBA PREGUNTAR POR EL CAMPO TIPO DE BILL OF LADING (CONSIGNADO A ALGUIEN O "A LA ORDEN")
       Case "BS1":
           Call procesa_BS1(cod_linea, linea)
       Case "BS3":
           Call procesa_BS3(cod_linea, linea)
       Case "BC1":
           Call procesa_BC1(cod_linea, linea)
       Case "BC3":
           Call procesa_BC3(cod_linea, linea)
       Case "BN1":
           Call procesa_BN1(cod_linea, linea)
       Case "BN3":
           Call procesa_BN3(cod_linea, linea)
       Case "CHG"
           Call procesa_CHG(cod_linea, linea)
           ' AQUI ESTO TAMBIEN CAMBIA PERO DEBE CAMBIAR JUNTO CON PCY (COSTO FLETE)
       Case "CTR"  ' datos del contenedor
           'If Gstr_Tipo = "I" Then Call graba_blimpo Else Call graba_blexpo
             Call graba_blimpo
           Call inicializa_blimpo(False)
           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
           Call graba_impocont              'MARZO 31 2004
           Call inicializa_impocont(False)  'MARZO 31 2004
           Call graba_blimpodt
           Call inicializa_blimpodt(False)
           Call procesa_CTR(cod_linea, linea)
           'ESTA LINEA SE PUSO SOLO PARA HACER SUBIR LOS CONTENEDORES DE LA VERSION 52, YA QUE EN ESTA VERSION NO VIENE LA LINEA SEL
           'If Gstr_Tipo = "I" And Gstr_Version_MSC = "52" Then Call graba_impocont
           If Gstr_Version_MSC = "52" Then Call graba_impocont
       Case "SEL"
           Call procesa_SEL(cod_linea, linea)
           'If Gstr_Tipo = "I" Then Call graba_impocont
           'Call graba_impocont MARZO 31 2004

           'AQUI falta EVALUAR QUE PASA SI VIENE MAS DE UN SELLO POR CADA CONTEINER
       Case "CT1" ' SE GRABA UN SOLO DETALLE POR CADA CONTENEDOR
           If Not Gbol_bl_detallado_MSC Then
                'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
                Call graba_blimpodt
                Call procesa_CT1(cod_linea, linea)
                'If Gstr_Tipo = "I" Then Call graba_blimpo Else Call graba_blexpo
                Call graba_blimpo
                Call inicializa_blimpo(False)
           End If
           ' aqui este es una linea nueva, se usa para no grabar cada detalle sino uno solo, mas aplicado a cnts vacios
       Case "DE1": ' detalle de carga
           If Gbol_bl_detallado_MSC Then ' SE PARAMETRIZÓ GRABAR DETALLADO LOS ARTICULOS DE CADA CONTENEDOR
                'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
                Call graba_blimpodt
                Call procesa_DE1(cod_linea, linea)
                'If Gstr_Tipo = "I" Then Call graba_blimpo Else Call graba_blexpo
                Call graba_blimpo
                Call inicializa_blimpo(False)
           End If
       Case "DE2":
           Call procesa_DE2(cod_linea, linea)
       Case "DE3":
           Call procesa_DE3(cod_linea, linea)
       Case "CHK":
           'If Gstr_Tipo = "I" Then Call graba_fletes
           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
           'If Gstr_Tipo = "I" Then Call actualiza_cabecera Else Call actualiza_cabecera_expo
           'If Gstr_Tipo = "I" Then Call actualiza_manifiesto
           Call graba_fletes
           Call graba_impocont              'MARZO 31 2004
           Call graba_blimpodt
           Call actualiza_cabecera
           Call actualiza_manifiesto
    End Select
Loop

DTX_PROCESO.Lbl_prog.Width = 2625
DTX_PROCESO.Lbl_avan.Caption = "100 % "

Close #x

End Sub

'*****************************************************************************
' este proceso es igual al procesar_arch pero este es unicamente para grabar
' los contenedores vacíos, ya que debido al modo en que procesa el procedim.
' anterior, se necesita volver a verificar el archivo.
' jueves 4 de enero del 2001 (petion de Pilar Flores)
'-----------------------------------------------------------------------------
' el proceso es asi :
' por cada linea VSL se inicializan todas las estructuras,
' por cada puerto que se inicializan todas las estructuras de blimpo y blimpodt
' y asi sucesivamente
' ESTE ES EL PROCESO PRINCIPAL DE TODO EL PROGRAMA DE INTERFASE DE DTX's
' *****************************************************************************
Public Sub procesar_arch_MSC_cont_vacios(Archivo As String)
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
DTX_PROCESO.Lbl_avan.Caption = "Procesando contened. vacíos..."
DTX_PROCESO.Lbl_prog.Width = 0
DTX_PROCESO.Lbl_prog.Visible = True
x = FreeFile
Open Archivo For Input As #x
Do While Not EOF(x)
    Line Input #x, linea
    cod_linea = Mid$(linea, 1, 3)
    If DTX_PROCESO.Lbl_prog.Width < 2625 Then
       DTX_PROCESO.Lbl_prog.Width = DTX_PROCESO.Lbl_prog.Width + CDbl(Format(avance, "#####0.000"))
       DTX_PROCESO.Lbl_avan.Caption = Format(DTX_PROCESO.Lbl_prog.Width * 100 / 2625, "##0") & " %"
    End If
    If DTX_PROCESO.Lbl_prog.Width >= 2625 / 2 Then ' para que cambie de color al llegar a la mitad
       DTX_PROCESO.Lbl_avan.ForeColor = &HFFFFFF
    End If
    
    DoEvents
    Select Case cod_linea
       Case "VSL":
           Call inicializa_registro(True)
           Call inicializa_manifiesto(True)
           Call inicializa_impocont(True)
           Call inicializa_blimpo(True)
           Call inicializa_blimpodt(True)
           Call Procesa_VSL(cod_linea, linea)
           'Call graba_registro '  vacio
       Case "POR":
           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt                ' vacios
           'If Gstr_Tipo = "I" Then Call actualiza_cabecera Else Call actualiza_cabecera_expo   ' vacios
           'If Gstr_Tipo = "I" Then Call actualiza_manifiesto                                   ' vacios
           Call inicializa_blimpo(True)
           Call inicializa_blimpodt(True)
           Call procesa_POR(cod_linea, linea)
           'If Gstr_Tipo = "I" Then Call graba_manifiesto       ' vacios
       Case "BL1":
           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt    '   vacios
           'If Gstr_Tipo = "I" Then Call actualiza_cabecera Else actualiza_cabecera_expo    '   vacios
           Call inicializa_blimpodt(False)
           Call procesa_BL1(cod_linea, linea)
       Case "PCY":
           Call procesa_PCY(cod_linea, linea)
       Case "BS1":
           'Call procesa_BS1(cod_linea, linea)
       Case "BS2":
           'Call procesa_BS2(cod_linea, linea)
       Case "BC1":
           'Call procesa_BC1(cod_linea, linea)
       Case "BC2":
           'Call procesa_BC2(cod_linea, linea)
       Case "BN1":
           'Call procesa_BN1(cod_linea, linea)
       Case "BN2":
           'Call procesa_BN2(cod_linea, linea)
       ' aqui se debe poner la verificacion de si no existe agregarlo como vacio
       '************************************************************************
       Case "CTR"  ' datos del contenedor
           'If Gstr_Tipo = "I" Then Call graba_blimpo Else Call graba_blexpo
           Call inicializa_blimpo(False)
           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
           Call inicializa_blimpodt(False)
           Call procesa_CTR(cod_linea, linea)
       Case "SEL"
           Call procesa_SEL(cod_linea, linea)
           'If Gstr_Tipo = "I" Then Call graba_impocont
           'AQUI falta EVALUAR QUE PASA SI VIENE MAS DE UN SELLO POR CADA CONTEINER
       Case "CT1" ' SE GRABA UN SOLO DETALLE POR CADA CONTENEDOR
           'If Not Gbol_bl_detallado_MSC Then
                'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
                Call procesa_CT1(cod_linea, linea)
                'If Gstr_Tipo = "I" Then Call graba_impo_vacios Else graba_expo_vacios ' vacios
                'If Gstr_Tipo = "I" Then Call actualiza_cabecera Else Call actualiza_cabecera_expo
                'If Gstr_Tipo = "I" Then Call actualiza_manifiesto
                Call graba_impo_vacios
                Call actualiza_cabecera
                Call actualiza_manifiesto
'                Call inicializa_blimpo(False)
'                Call inicializa_blimpodt(False)
           'End If
           
           ' aqui este es una linea nueva, se usa para no grabar cada detalle sino uno solo, mas aplicado a cnts vacios
'       Case "DE1": ' detalle de carga
'           If Gbol_bl_detallado_MSC Then ' SE PARAMETRIZÓ GRABAR DETALLADO LOS ARTICULOS DE CADA CONTENEDOR
'                If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
'                Call procesa_DE1(cod_linea, linea)
'                If Gstr_Tipo = "I" Then Call graba_blimpo Else Call graba_blexpo
'                Call inicializa_blimpo(False)
'           End If
       '**********************************************************************
'       Case "CTR"  ' datos del contenedor
'           'If Gstr_Tipo = "I" Then Call graba_blimpo Else Call graba_blexpo
'           Call inicializa_blimpo(False)
'           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
'           Call inicializa_blimpodt(False)
'           Call procesa_CTR(cod_linea, linea)
'           If Gstr_Tipo = "I" Then Call graba_impo_vacios Else graba_expo_vacios ' vacios
'           If Gstr_Tipo = "I" Then Call actualiza_cabecera Else Call actualiza_cabecera_expo
'           If Gstr_Tipo = "I" Then Call actualiza_manifiesto
       '*************************************************************************
       Case "CHG"
           'Call procesa_CHG(cod_linea, linea)
       Case "DE1": ' detalle de carga
           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
           'Call procesa_DE1(cod_linea, linea)
           'If Gstr_Tipo = "I" Then Call graba_blimpo Else Call graba_blexpo
           Call inicializa_blimpo(False)
       Case "DE2":
           'Call procesa_DE2(cod_linea, linea)
       Case "DE3":
           'Call procesa_DE3(cod_linea, linea)
       Case "CHK":
           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt                '   vacios
           'If Gstr_Tipo = "I" Then Call actualiza_cabecera Else Call actualiza_cabecera_expo   '   vacíos
           'If Gstr_Tipo = "I" Then Call actualiza_manifiesto                                   '   vacíos
    End Select
Loop

DTX_PROCESO.Lbl_prog.Width = 2625
DTX_PROCESO.Lbl_avan.Caption = "100 % "

Close #x

End Sub

Public Function verifica_version_DTX(Archivo As String)
Dim x
Dim sql   As String
Dim reg   As New ADODB.Recordset

verifica_version_DTX = "NN"
x = FreeFile
Open Archivo For Input As #x
Do While Not EOF(x)
    Line Input #x, linea
    cod_linea = Mid$(linea, 1, 3)
    Select Case cod_linea
       Case "HDR":
          sql = "Select * from DTX_detalle " & _
                " Where linea = '" & Gstr_Linea & "'" & _
                " And codigo  ='HDR'" & _
                " And campo   = 'HDR02'" & _
                " And version = '55' "
          reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
          If Not reg.EOF Then
             verifica_version_DTX = Mid$(linea, reg("desde"), reg("longitud"))
             Gstr_Version_MSC = Mid$(linea, reg("desde"), reg("longitud"))
          Else
             Gstr_Version_MSC = "54"
          End If
          Exit Do
    End Select
Loop
If Trim(Gstr_Version_MSC) = "" Then Gstr_Version_MSC = "54"
Close
End Function


' el proceso es asi :
' por cada linea VSL se inicializan todas las estructuras,
' por cada puerto que se inicializan todas las estructuras de blimpo y blimpodt
' y asi sucesivamente
' ESTE ES EL PROCESO PRINCIPAL DE TODO EL PROGRAMA DE INTERFASE DE DTX's
' *****************************************************************************
' 31 marzo 2004 se hacen modificaciones para recuperar lo perdido en daño de comput. mas de un sello.
' 29 Sept 2006  se realiza front end para pedir confirmacion de lo que se esta cargando.
'
Public Sub procesar_arch_v55(Archivo As String)
Dim x
Dim linea       As String
Dim cod_linea   As String
Dim sql         As String
Dim reg_tipo    As New ADODB.Recordset
Dim reg_detalle As New ADODB.Recordset
Dim Tot_lineas  As Integer
Dim avance      As Double

Screen.MousePointer = 11

'------------------------------------------
Gbol_solo_GYE = False
'If Gstr_Tipo = "I" Then
'    If MsgBox("Desea procesar solamente la carga para Guayaquil?", vbYesNo + vbQuestion) = vbYes Then
'       Gbol_solo_GYE = True
'    Else
       Gbol_solo_GYE = False
'    End If

'End If

'carga_parametros cmb_linea.BoundText

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
    cod_linea = Mid$(linea, 1, 3)
    If DTX_PROCESO.Lbl_prog.Width < 2625 Then
       DTX_PROCESO.Lbl_prog.Width = DTX_PROCESO.Lbl_prog.Width + CDbl(Format(avance, "#####0.000"))
       DTX_PROCESO.Lbl_avan.Caption = Format(DTX_PROCESO.Lbl_prog.Width * 100 / 2625, "##0") & " %"
    End If
    If DTX_PROCESO.Lbl_prog.Width >= 2625 / 2 Then ' para que cambie de color al llegar a la mitad
       DTX_PROCESO.Lbl_avan.ForeColor = &HFFFFFF
    End If
    
    DoEvents
    Select Case cod_linea
       Case "VSL":
           Call inicializa_registro(True)
           Call inicializa_manifiesto(True)
           Call inicializa_impocont(True)
           Call inicializa_blimpo(True)
           Call inicializa_blimpodt(True)
           Call Procesa_VSL(cod_linea, linea)
                      
           Call graba_registro
       Case "POR":
           'If Gstr_Tipo = "I" Then Call graba_fletes
           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
           'If Gstr_Tipo = "I" Then Call actualiza_cabecera Else Call actualiza_cabecera_expo
           'If Gstr_Tipo = "I" Then Call actualiza_manifiesto
           Call graba_fletes
           Call graba_impocont              'MARZO 31 2004
           Call graba_blimpodt
           Call actualiza_cabecera
           Call actualiza_manifiesto
           Call inicializa_blimpo(True)
           Call inicializa_blimpodt(True)
           Call inicializa_impocont(True)  'MARZO 31 2004
           Call inicializa_fletes
           Call procesa_POR(cod_linea, linea)
           'If Gstr_Tipo = "I" Then Call graba_manifiesto
           Call graba_manifiesto
       Case "PCY":
           Call procesa_PCY(cod_linea, linea)
           'AQUI FALTA PONER LA PANTALLA QUE PIDA CONFIRMAR LOS VALORES O CAMBIARLOS PARA EL TIPO DE CAMBIO
       Case "BL1":
           If Not Gbol_relacion_fija Then Call Confirma_Conversion
           'If Gstr_Tipo = "I" Then Call graba_fletes
           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
           'If Gstr_Tipo = "I" Then Call actualiza_cabecera Else actualiza_cabecera_expo
           Call graba_fletes
           Call graba_blimpodt
           Call actualiza_cabecera
           Call inicializa_blimpodt(False)
           Call inicializa_fletes
           Call procesa_BL1(cod_linea, linea)
           'AQUI TALVEZ SE DEBA PREGUNTAR POR EL CAMPO TIPO DE BILL OF LADING (CONSIGNADO A ALGUIEN O "A LA ORDEN")
       Case "BA1": 'solo version 55
           Call procesa_BA1(cod_linea, linea)
       Case "BA3": 'solo version 55
           Call procesa_BA3(cod_linea, linea)
       Case "CHB" ' SOLO VERSION 55
           Call procesa_CHB(cod_linea, linea)
       Case "CTR"  ' datos del contenedor
           'If Gstr_Tipo = "I" Then Call graba_blimpo Else Call graba_blexpo
             Call graba_blimpo
           Call inicializa_blimpo(False)
           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
           Call graba_impocont              'MARZO 31 2004
           Call inicializa_impocont(False)  'MARZO 31 2004
           Call graba_blimpodt
           Call inicializa_blimpodt(False)
           Call procesa_CTR(cod_linea, linea)
           'ESTA LINEA SE PUSO SOLO PARA HACER SUBIR LOS CONTENEDORES DE LA VERSION 52, YA QUE EN ESTA VERSION NO VIENE LA LINEA SEL
           'If Gstr_Tipo = "I" And Gstr_Version_MSC = "52" Then Call graba_impocont
           If Gstr_Version_MSC = "52" Then Call graba_impocont
       Case "RFR"
            Call procesa_RFR(cod_linea, linea)
       Case "SEL"
           Call procesa_SEL(cod_linea, linea)
       Case "CT1" ' SE GRABA UN SOLO DETALLE POR CADA CONTENEDOR
           If Not Gbol_bl_detallado_MSC Then
                'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
                Call graba_blimpodt
                Call procesa_CT1(cod_linea, linea)
                'If Gstr_Tipo = "I" Then Call graba_blimpo Else Call graba_blexpo
                Call graba_blimpo
                Call inicializa_blimpo(False)
           End If
           ' aqui este es una linea nueva, se usa para no grabar cada detalle sino uno solo, mas aplicado a cnts vacios
       Case "DE1": ' detalle de carga
           If Gbol_bl_detallado_MSC Then ' SE PARAMETRIZÓ GRABAR DETALLADO LOS ARTICULOS DE CADA CONTENEDOR
                'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
                Call graba_blimpodt
                Call procesa_DE1(cod_linea, linea)
                'If Gstr_Tipo = "I" Then Call graba_blimpo Else Call graba_blexpo
                Call graba_blimpo
                Call inicializa_blimpo(False)
           End If
       Case "DE2":
           Call procesa_DE2(cod_linea, linea)
       Case "DE3":
           Call procesa_DE3(cod_linea, linea)
       Case "CHK":
           'If Gstr_Tipo = "I" Then Call graba_fletes
           'If Gstr_Tipo = "I" Then Call graba_blimpodt Else Call graba_blexpodt
           'If Gstr_Tipo = "I" Then Call actualiza_cabecera Else Call actualiza_cabecera_expo
           'If Gstr_Tipo = "I" Then Call actualiza_manifiesto
           Call graba_fletes
           Call graba_impocont              'MARZO 31 2004
           Call graba_blimpodt
           Call actualiza_cabecera
           Call actualiza_manifiesto
    End Select
Loop

DTX_PROCESO.Lbl_prog.Width = 2625
DTX_PROCESO.Lbl_avan.Caption = "100 % "

Close #x
Screen.MousePointer = 0
End Sub


'*****************************************************************************
' este proceso es igual al procesar_arch pero este es unicamente para grabar
' los contenedores vacíos, ya que debido al modo en que procesa el procedim.
' anterior, se necesita volver a verificar el archivo.
' jueves 4 de enero del 2001 (petion de Pilar Flores)
'-----------------------------------------------------------------------------
' el proceso es asi :
' por cada linea VSL se inicializan todas las estructuras,
' por cada puerto que se inicializan todas las estructuras de blimpo y blimpodt
' y asi sucesivamente
' ESTE ES EL PROCESO PRINCIPAL DE TODO EL PROGRAMA DE INTERFASE DE DTX's
' *****************************************************************************
Public Sub procesar_arch_MSC_cont_vacios_V55(Archivo As String)
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
DTX_PROCESO.Lbl_avan.Caption = "Procesando contened. vacíos..."
DTX_PROCESO.Lbl_prog.Width = 0
DTX_PROCESO.Lbl_prog.Visible = True
x = FreeFile
GLinea_act_Grid = 0
Open Archivo For Input As #x
Do While Not EOF(x)
    Line Input #x, linea
    cod_linea = Mid$(linea, 1, 3)
    If DTX_PROCESO.Lbl_prog.Width < 2625 Then
       DTX_PROCESO.Lbl_prog.Width = DTX_PROCESO.Lbl_prog.Width + CDbl(Format(avance, "#####0.000"))
       DTX_PROCESO.Lbl_avan.Caption = Format(DTX_PROCESO.Lbl_prog.Width * 100 / 2625, "##0") & " %"
    End If
    If DTX_PROCESO.Lbl_prog.Width >= 2625 / 2 Then ' para que cambie de color al llegar a la mitad
       DTX_PROCESO.Lbl_avan.ForeColor = &HFFFFFF
    End If
    
    DoEvents
    Select Case cod_linea
       Case "VSL":
           Call inicializa_registro(True)
           Call inicializa_manifiesto(True)
           Call inicializa_impocont(True)
           Call inicializa_blimpo(True)
           Call inicializa_blimpodt(True)
           Call Procesa_VSL(cod_linea, linea)
       Case "POR":
           Call inicializa_manifiesto(True)
           Call inicializa_blimpo(True)
           Call inicializa_blimpodt(True)
           Call procesa_POR(cod_linea, linea)
       Case "BL1":
           Call inicializa_blimpodt(False)
           Call procesa_BL1(cod_linea, linea)
       Case "PCY":
           Call procesa_PCY(cod_linea, linea)
       ' aqui se debe poner la verificacion de si no existe agregarlo como vacio
       '************************************************************************
       Case "CTR"  ' datos del contenedor
           'Call inicializa_blimpo(False)
           'Call inicializa_blimpodt(False)
           Call procesa_CTR(cod_linea, linea)
           Call graba_impo_vacios
           Call actualiza_cabecera
           Call actualiza_manifiesto
       Case "SEL"
           Call procesa_SEL(cod_linea, linea)
       Case "CT1" ' SE GRABA UN SOLO DETALLE POR CADA CONTENEDOR
                Call procesa_CT1(cod_linea, linea)
                Call graba_impo_vacios
                Call actualiza_cabecera
                Call actualiza_manifiesto
       Case "DE1": ' detalle de carga
           'Call inicializa_blimpo(False)
       Case "CHK":
    End Select
Loop

DTX_PROCESO.Lbl_prog.Width = 2625
DTX_PROCESO.Lbl_avan.Caption = "100 % "

Close #x

End Sub

Public Sub inicializa_registro(opc As Boolean)
    If opc Then
        r_registro.buque = ""
        r_registro.viaje = ""
        ReDim Preserve Gtyp_flete_DTX(1)
    End If
    r_registro.naviera = ""
    r_registro.direcc = ""
'    r_registro.Linea_Nav = ""
    r_registro.atraque = Empty
    'r_registro.zarpe = Empty
    r_registro.fecha_ing = Empty
    r_registro.user_ing = ""
    r_registro.fecha_mod = Empty
    r_registro.user_mod = ""
End Sub
' el parametro es para inicializar toda la estructura o
' solo los campos que no son comunes
' -----------------------------------------------------
Public Sub inicializa_manifiesto(opc As Boolean)
   If opc Then
      r_manifiesto.buque = ""
      r_manifiesto.viaje = ""
      r_manifiesto.puerto = ""
   End If
   r_manifiesto.capitan = ""
   r_manifiesto.fecha_zarpe = Empty
   r_manifiesto.No_bls = 0
   r_manifiesto.no_bultos = 0
   r_manifiesto.no_cont = 0
   r_manifiesto.no_mani = ""
   r_manifiesto.peso_kls = 0
   r_manifiesto.pto_des = ""
   r_manifiesto.pto_inicio = ""
   r_manifiesto.status = ""
   r_manifiesto.volumen = 0
      
End Sub
' el parametro es para inicializar toda la estructura o
' solo los campos que no son comunes
' -----------------------------------------------------
Public Sub inicializa_impocont(opc As Boolean)
   If opc Then
      r_impocont.buque = ""
      r_impocont.viaje = ""
      r_impocont.cond_contein = ""
      r_impocont.puerto = ""
      r_impocont.pto_des = ""   ' marzo 31 2004 esto estaba fuera del if
   End If
   r_impocont.CONTEINER = ""
'   r_impocont.cond_contrato = ""
'   r_impocont.estado_equipo = ""
   r_impocont.ind_lleno_vacio = ""
'   r_impocont.natu_carga = ""
   r_impocont.peso = 0
'   r_impocont.proveedor_equipo = ""

   r_impocont.sello1 = ""
   r_impocont.sello2 = ""
   r_impocont.sello3 = ""
   r_impocont.sello4 = ""
'   r_impocont.serv_transp = ""
   r_impocont.tamano_equipo = ""
   r_impocont.tara = 0
   r_impocont.temp_dur_almac = 0
   r_impocont.temp_dur_manip = 0
   r_impocont.temp_dur_trans = 0
   r_impocont.temp_max_almac = 0
   r_impocont.temp_max_manip = 0
   r_impocont.temp_max_trans = 0
   r_impocont.temp_min_almac = 0
   r_impocont.temp_min_manip = 0
   r_impocont.temp_min_trans = 0
   r_impocont.tipo_bien = ""
'   r_impocont.tipo_equipo = ""
End Sub
' el parametro es para inicializar toda la estructura o
' solo los campos que no son comunes
' -----------------------------------------------------
Public Sub inicializa_blimpo(opc As Boolean)
   If opc = True Then
      r_blimpo.buque = ""
      r_blimpo.viaje = ""
      r_blimpo.puerto = ""
      r_blimpo.fecha_embarque = Empty
      r_blimpo.pto_origen = ""
      r_blimpo.pto_destino = ""
      r_blimpo.pto_descarga = ""
   End If
   
   r_blimpo.COD_BL = ""
   r_blimpo.almac_temp = ""
   r_blimpo.bultos_mal_estado = 0
   r_blimpo.cono_master = ""
   r_blimpo.consignatario = ""
   r_blimpo.consolidadora = ""
   r_blimpo.Costo_Flete = 0
   r_blimpo.deposito_adua = ""
   r_blimpo.direc_exp = ""
   r_blimpo.direc_imp = ""
   r_blimpo.direc_notf = ""
   r_blimpo.direc_notf2 = ""
   r_blimpo.docum_autor = ""
   r_blimpo.embarcador = ""
   'r_blimpo.emp_anav = ""
   r_blimpo.fecha_embarque = Empty
   r_blimpo.id_notifier = ""
   r_blimpo.Impexp = ""
   r_blimpo.metodo_pago = ""
   r_blimpo.metodo_pago_otros = ""
   r_blimpo.MONEDA = ""
   r_blimpo.MONEDA_otros = ""
   r_blimpo.MONEDA_seguro = ""
   r_blimpo.nacio_consig = ""
   r_blimpo.no_bultos = 0
   r_blimpo.no_cont = 0
   r_blimpo.no_notyfier = ""
   r_blimpo.no_notyfier2 = ""
   r_blimpo.otros_gastos = 0
   r_blimpo.peso_kls = 0
   r_blimpo.peso_mal_estado = 0
   r_blimpo.regimen_adua = ""
   r_blimpo.seguro = 0
   r_blimpo.tipo_deposito = ""
   'r_blimpo.tipo_doc = ""
   r_blimpo.tipo_id = ""
   r_blimpo.tipo_id_notif = ""
   r_blimpo.tipo_salida = ""
   r_blimpo.volumen_mt3 = 0
   
'   Erase Gtyp_flete_DTX   ' reseta el arreglo de Fletes
'   ReDim Preserve Gtyp_flete_DTX(1)

End Sub

Public Sub inicializa_blimpodt(opc As Boolean)
   If opc Then
      r_blimpodt.buque = ""
      r_blimpodt.viaje = ""
      r_blimpodt.puerto = ""
      r_blimpodt.pto_descarga = ""
      
      r_blimpodt.COD_BL = ""
      r_blimpodt.desc_carga = ""
   End If
   r_blimpodt.sec_carga = 0
   r_blimpodt.CONTEINER = ""
   r_blimpodt.danger = ""
   r_blimpodt.embalaje = ""
   r_blimpodt.IMO_ONU = ""
   r_blimpodt.marca_num = ""
   r_blimpodt.no_bultos = 0
   r_blimpodt.peso = 0
   r_blimpodt.tipo_carga = ""
   r_blimpodt.volumen = 0
         
   If Gstr_Linea <> "KHS" Then
      r_blimpodt.desc_carga = ""
   End If
End Sub
Sub inicializa_fletes()
    Erase Gtyp_flete_DTX
    ReDim Preserve Gtyp_flete_DTX(1)
End Sub

Public Sub graba_registro()
Dim sql As String
Dim reg As New ADODB.Recordset
On Error GoTo error_graba

If r_registro.direcc = "1" Then r_registro.direcc = "A" Else r_registro.direcc = "R"
sql = "select * from docu_buque_viaje " & _
      " where buqv_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
      "   and buqv_buque ='" & r_registro.buque & "'" & _
      "   and buqv_viaje ='" & r_registro.viaje & "'"
reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
If Not reg.EOF Then Exit Sub

   sql = "Insert into docu_buque_viaje " & _
         "(buqv_linea,  buqv_buque,     buqv_viaje, buqv_direccion, buqv_fecha_arribo, " & _
         " buqv_status, buqv_fecha_ing, buqv_user_ing) Values('" & _
       DTX_PROCESO.Cmb_linea.BoundText & "', '" & _
       r_registro.buque & "','" & _
       r_registro.viaje & "','" & _
       r_registro.direcc & "','" & _
       Format(r_registro.atraque, GSTR_FORMATO_FECHA) & "','" & _
       "A" & "', '" & _
       Format(r_registro.fecha_ing, GSTR_FORMATO_FECHA) & "','" & _
       r_registro.user_ing & "')"
   
   Gcon_docu.Execute sql
   
   Exit Sub

error_graba:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
'Proceso_Error r_registro.anio, r_registro.REGISTRO, " ", " ", " ", Err.Description, DTX_PROCESO.Txt_DTX.Text
Exit Sub
Resume 0
End Sub

Public Sub Procesa_VSL(cod_linea As String, linea As String)
  Dim sql As String
  Dim reg As New ADODB.Recordset


   r_registro.buque = DTX_PROCESO.Cmb_buque.BoundText
   r_registro.fecha_ing = Format$(Date, GSTR_FORMATO_FECHA)
   r_registro.atraque = DTX_PROCESO.Msk_atraque.Text
   r_registro.naviera = Gstr_naviera
   r_registro.user_ing = Gstr_user_DTX
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' and codigo ='" & cod_linea & "' and nombre_tabla = 'REGISTRO'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   
   Do While Not reg.EOF
      Select Case Trim$(reg("CAMPO_TABLA"))
         Case "DIRECCION":
            r_registro.direcc = Mid$(linea, reg("desde"), reg("longitud"))
      End Select
      reg.MoveNext
   Loop
      
      
   r_registro.viaje = Trim(DTX_PROCESO.Txt_viaje.Text)
   ' graba en las otras tablas los datos comunes
   r_manifiesto.buque = DTX_PROCESO.Cmb_buque.BoundText
   r_manifiesto.viaje = r_registro.viaje

   r_impocont.buque = DTX_PROCESO.Cmb_buque.BoundText
   r_impocont.viaje = r_registro.viaje
End Sub
' este procedimiento ademas de grabar en la tabla manifiesto,
'graba ciertos datos en la blimpo y blimpodt
' 2 oct 2006 se modifico para que ahora busque del grid de pantalla en lugar de hacer uso de la funcion codigo_equivalente
Public Sub procesa_POR(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
Dim pto_emb  As String
Dim pto_desc As String

On Error GoTo Errores
   
   ' SE DROPEAN LAS TAZAS DE CAMBIO DE LAS MONEDAS QUE LUEGO SE LLENAN EN EL PCY
   sql = "DELETE FROM DTX_CONVERSION"
   Gcon_docu.Execute sql
   '------------------------------------------------------------------------------
   r_manifiesto.no_mani = "1"
   r_manifiesto.No_bls = 0
   r_manifiesto.peso_kls = 0
   r_manifiesto.volumen = 0
   r_manifiesto.no_bultos = 0
   r_manifiesto.status = "A"
   
   'se ubica dentro del Grid la posicion del registro para obtener los valores que el usuario modifico.
   ' se busca el pto de Embarque
       sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'MANIFIESTO' and  CAMPO_TABLA ='PUERTO'"
       reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
       If Not reg.EOF Then pto_emb = Mid$(linea, reg("desde"), reg("longitud"))
       reg.Close
   ' se busca el pto de descarga
       sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO' and CAMPO_TABLA = 'PTO_DESCARGA'"
       reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
       If Not reg.EOF Then pto_desc = Mid$(linea, reg("desde"), reg("longitud"))
       reg.Close
                          
   If Not Grid_confirma_carga("PTO", pto_emb, pto_desc, "") Then Exit Sub
       
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'MANIFIESTO'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "PUERTO":
            r_manifiesto.puerto = Grid_Codigo_equivalente("PTO_EMBA")   'Codigo_equivalente_DTX("DTX_PUERTO", Mid$(linea, reg("desde"), reg("longitud")), "APG", "CODIGO")
            r_blimpo.puerto = r_manifiesto.puerto
            r_blimpodt.puerto = r_manifiesto.puerto
            r_impocont.puerto = r_manifiesto.puerto
         Case "FECHA_ZARPE"
            r_manifiesto.fecha_zarpe = Aplica_Formato(cod_linea, reg("CAMPO"), Mid$(linea, reg("desde"), reg("longitud")))
            r_blimpo.fecha_embarque = r_manifiesto.fecha_zarpe
      End Select
      reg.MoveNext
   Loop
   
' ----------------------------------------------------------------------
' se toman datos de la linea de puerto para el archivo blimpo y blimpodt
' ----------------------------------------------------------------------
   r_manifiesto.buque = DTX_PROCESO.Cmb_buque.BoundText
   r_manifiesto.viaje = r_registro.viaje
   
   r_blimpo.buque = DTX_PROCESO.Cmb_buque.BoundText
   r_blimpo.viaje = r_registro.viaje
   
   r_blimpodt.buque = DTX_PROCESO.Cmb_buque.BoundText
   r_blimpodt.viaje = r_registro.viaje
   reg.Close
   
   'sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
   'reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   
   '
   ' 23sept2005
   'Do While Not reg.EOF
   '   Select Case reg("CAMPO_TABLA")
   '      Case "PTO_ORIGEN":
            r_blimpo.pto_origen = Grid_Codigo_equivalente("PTO_ORIG")   '
   '      Case "PTO_DESCARGA"
            r_blimpo.pto_descarga = Grid_Codigo_equivalente("PTO_DESC")   ' Codigo_equivalente_DTX("DTX_PUERTO", Mid$(linea, reg("desde"), reg("longitud")), "APG", "CODIGO")
            r_blimpo.pto_destino = Grid_Codigo_equivalente("PTO_DEST")   'Codigo_equivalente_DTX("DTX_PUERTO", Mid$(linea, reg("desde"), reg("longitud")), "APG", "CODIGO")
            r_blimpodt.pto_descarga = r_blimpo.pto_descarga
            r_manifiesto.pto_des = r_blimpo.pto_descarga
            r_impocont.pto_des = r_blimpo.pto_descarga
   '   End Select
'      reg.MoveNext
'   Loop
Exit Sub

Errores:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0

End Sub

Public Sub graba_manifiesto()
Dim sql As String
Dim reg As New ADODB.Recordset

On Error GoTo error_graba
'------ esto para filtrar solo inf. de GYE en MSC de import ------
' -------------------------  27/agosto/2001  ---------------------
st_continua = True
If Gstr_Linea = "MSC" And Gstr_Tipo = "I" Then
    If Gbol_solo_GYE Then
        If r_blimpo.pto_descarga <> "ECGYE" Then st_continua = False
    End If
End If
If Not st_continua Then Exit Sub
'-----------------------------------------------------------------
If Gstr_Tipo = "I" Then
    If r_manifiesto.puerto = "" Then Exit Sub
Else
    If r_manifiesto.pto_des = "" Then Exit Sub
End If

sql = "Select * from "
If Gstr_Tipo = "I" Then sql = sql & "docu_manifiesto " Else sql = sql & "docu_manifiesto_exp "
sql = sql & " Where mani_linea    ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
            "   and mani_buque    ='" & r_registro.buque & "'" & _
            "   and mani_viaje    ='" & r_registro.viaje & "'" & _
            "   and mani_puer_emb ='" & r_manifiesto.puerto & "'" & _
            "   and mani_puer_des ='" & r_manifiesto.pto_des & "'"
reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
If Not reg.EOF Then Exit Sub

r_manifiesto.pto_inicio = r_manifiesto.puerto

sql = "Insert into "
If Gstr_Tipo = "I" Then sql = sql & "docu_manifiesto " Else sql = sql & "docu_manifiesto_exp "
sql = sql & "                      (mani_linea,    mani_buque,    mani_viaje, " & _
                                  " mani_puer_emb, mani_puer_des, mani_capitan, " & _
                                  " mani_puer_ini, mani_fecha_zarp_ini , mani_emp_prop, mani_emp_nave) " & _
       " Values('" & _
       DTX_PROCESO.Cmb_linea.BoundText & "','" & _
       r_registro.buque & "','" & _
       r_registro.viaje & "','" & _
       r_manifiesto.puerto & "','" & _
       r_manifiesto.pto_des & "', '" & _
       Gstr_capitan & "', '" & _
       r_manifiesto.pto_inicio & "','" & _
       Format(r_manifiesto.fecha_zarpe, GSTR_FORMATO_FECHA) & "','" & _
       Gstr_param.Linea_Nav & "','" & _
       Gstr_param.naviera & "')"
       
       'Gstr_param.naviera
       
   Gcon_docu.Execute sql
   Exit Sub

error_graba:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0
End Sub

Public Sub graba_blimpo()
Dim sql As String
Dim reg As New ADODB.Recordset
Dim st_continua As Boolean

On Error GoTo error_graba


If Trim$(r_blimpo.COD_BL) = "" Then Exit Sub

'------ esto para filtrar solo inf. de GYE en MSC de import ------
' -------------------------  27/agosto/2001  ---------------------
st_continua = True
If Gstr_Linea = "MSC" And Gstr_Tipo = "I" Then
    If Gbol_solo_GYE Then
        If r_blimpo.pto_descarga <> "ECGYE" Then st_continua = False
    End If
End If
If Not st_continua Then Exit Sub
'-----------------------------------------------------------------
'If r_blimpo.COD_BL = "MSCUR1541680" Then MsgBox "aqui"
   sql = "Select * from "
   If Gstr_Tipo = "I" Then sql = sql & " docu_Blcab " Else sql = sql & " docu_blcab_exp "
   sql = sql & " where blcab_linea    ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
         "   and blcab_buque    ='" & r_registro.buque & "'" & _
         "   and blcab_viaje    ='" & r_registro.viaje & "'" & _
         "   and blcab_puer_emb ='" & r_manifiesto.puerto & "'" & _
         "   and blcab_puer_des ='" & r_manifiesto.pto_des & "'" & _
         "   and blcab_emp_resp ='" & Gstr_param.naviera & "'" & _
         "   and blcab_cono_emb ='" & r_blimpo.COD_BL & "'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   
   If reg.EOF Then
       If r_blimpo.MONEDA = "" Then r_blimpo.MONEDA = Gstr_moneda
       
       r_blimpo.fecha_embarque = r_manifiesto.fecha_zarpe 'OJO CON FECHA DE EMB.
       r_blimpo.regimen_adua = IIf(Gstr_Tipo = "I", Gstr_param.regimen_adua, Gstr_param.regimen_adua_expo)
       If r_blimpo.pto_descarga <> r_blimpo.pto_destino And Mid(r_blimpo.pto_destino, 1, 2) <> "EC" Then 'And Mid(r_blimpo.pto_descarga, 1, 2) <> "EC"
         r_blimpo.regimen_adua = Gstr_param.regimen_adua_trasb
       End If
       r_blimpo.almac_temp = Gstr_param.almacen_temporal
       ' debido a un cambio en la aduana para exportacion, el campo consignatario en realidad debe contener la informacion del embarcador y el campo embarcador debe contener la informacion del consignatario
       sql = "Insert into "
       If Gstr_Tipo = "I" Then sql = sql & " docu_Blcab " Else sql = sql & " docu_blcab_exp "
       sql = sql & " Values('" & _
           DTX_PROCESO.Cmb_linea.BoundText & "','" & _
           r_registro.buque & "','" & _
           r_registro.viaje & "','" & _
           r_blimpo.puerto & "','" & _
           r_manifiesto.pto_des & "','" & _
           Gstr_param.naviera & "','" & _
           r_blimpo.COD_BL & "','" & _
           Gstr_Linea_Nav_APG & "','" & _
           Gstr_param.tipo_documento & "', '" & _
           r_blimpo.cono_master & "', '" & _
           r_blimpo.peso_kls & "','" & _
           r_blimpo.no_bultos & "','" & _
           r_blimpo.tipo_id & "','" & _
           r_blimpo.Impexp & "','"
'       If Gstr_Tipo = "I" Then ' si es exportacion los campos se setean a como debe ser
           sql = sql & _
           Trim(chk_str(r_blimpo.consignatario)) & "','" & _
           Trim(chk_str(r_blimpo.direc_imp)) & "','"
'       Else         ' si es exportacion los campos se invierten
'           sql = sql & _
'           Trim(chk_str(r_blimpo.embarcador)) & "','" & _
'           Trim(chk_str(r_blimpo.direc_exp)) & "','"
'       End If
       sql = sql & _
           r_blimpo.nacio_consig & "','" & _
           r_blimpo.tipo_id_notif & "','" & _
           r_blimpo.id_notifier & "','" & _
           Trim(chk_str(r_blimpo.no_notyfier)) & "','" & _
           Trim(chk_str(r_blimpo.direc_notf)) & "','" & _
           "', '', '"
           ' embarcador
'       If Gstr_Tipo = "E" Then
'           sql = sql & _
'           Trim(chk_str(r_blimpo.consignatario)) & "','" & _
'           Trim(chk_str(r_blimpo.direc_imp)) & "','"
'       Else
           sql = sql & _
           Trim(chk_str(r_blimpo.embarcador)) & "','" & _
           Trim(chk_str(r_blimpo.direc_exp)) & "','"
 '      End If
           'Trim (chk_str(r_blimpo.embarcador)) & "','" & _
           Trim(chk_str(r_blimpo.direc_exp)) & "','"
      sql = sql & r_blimpo.pto_origen & "','" & _
           "', null, '" & _
           r_blimpo.almac_temp & "','" & _
           r_blimpo.tipo_deposito & "','" & _
           r_blimpo.deposito_adua & "','" & _
           r_blimpo.regimen_adua & "','" & _
           r_blimpo.consolidadora & "'," & _
           "Null" & ",'" & _
           "" & "','" & _
           r_blimpo.tipo_salida & "','" & _
           r_blimpo.docum_autor & "','" & _
           r_blimpo.peso_mal_estado & "','" & _
           r_blimpo.bultos_mal_estado & "','" & _
           Trim(r_blimpo.no_notyfier2) & "','" & _
           Trim(r_blimpo.direc_notf2) & "','" & _
           r_blimpo.no_cont & "','" & _
           r_blimpo.volumen_mt3 & "','"
      If r_blimpo.pto_destino = r_blimpo.pto_descarga And Gstr_Tipo = "I" Then ' por un req de aduana este campo de destino final debe estar en blanco
           sql = sql & _
                "" & "','"
      Else
           sql = sql & _
                r_blimpo.pto_destino & "','"
      End If
      sql = sql & _
           Gstr_param.metodo_pago & "','" & _
           r_blimpo.MONEDA & "','" & _
           r_blimpo.Costo_Flete & "','" & _
           r_blimpo.otros_gastos & "','" & _
           r_blimpo.MONEDA_otros & "','" & _
           r_blimpo.seguro & "','" & _
           r_blimpo.MONEDA_seguro & "','"
      sql = sql & r_blimpo.metodo_pago_otros & "','" & _
           Format(r_blimpo.fecha_embarque, GSTR_FORMATO_FECHA) & "','" & _
           Gstr_user_DTX & "','" & _
           Format(Now, GSTR_FORMATO_FECHA) & "', '" & _
           Format(Time, "hh:mm:ss") & "','" & _
           "',Null, '')"
       Gcon_docu.Execute sql
   End If
' ------- se borran los detalles ----------
sql = "Delete from "
If Gstr_Tipo = "I" Then sql = sql & " docu_bldet " Else sql = sql & " docu_bldet_exp "
sql = sql & " Where bldet_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
         "   and bldet_buque    ='" & r_registro.buque & "'" & _
         "   and bldet_viaje    ='" & r_registro.viaje & "'" & _
         "   and bldet_puer_emb ='" & r_manifiesto.puerto & "'" & _
         "   and bldet_puer_des ='" & r_manifiesto.pto_des & "'" & _
         "   and bldet_emp_resp ='" & Gstr_param.naviera & "'" & _
         "   and bldet_cono_emb ='" & r_blimpo.COD_BL & "'"
Gcon_docu.Execute sql
' ----------------------------------------------
' ------- se borran los detalles de Fletes ----------
sql = "Delete from "
If Gstr_Tipo = "I" Then sql = sql & " docu_blflete " Else sql = sql & " docu_blflete_exp "
sql = sql & " where flete_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
         "   and flete_buque    ='" & r_registro.buque & "'" & _
         "   and flete_viaje    ='" & r_registro.viaje & "'" & _
         "   and flete_puer_emb ='" & r_manifiesto.puerto & "'" & _
         "   and flete_puer_des ='" & r_manifiesto.pto_des & "'" & _
         "   and flete_emp_resp ='" & Gstr_param.naviera & "'" & _
         "   and flete_cono_emb ='" & r_blimpo.COD_BL & "'"
Gcon_docu.Execute sql

' ----------------------------------------------
r_blimpo.COD_BL = ""
Exit Sub

error_graba:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0
End Sub

Public Sub graba_blimpodt()
Dim sql As String
Dim reg As New ADODB.Recordset
Dim SECUENCIA As Integer
Dim st_continua As Boolean

On Error GoTo error_graba

If Trim$(r_blimpodt.COD_BL) = "" Then Exit Sub
If Trim$(r_blimpodt.embalaje) = "" Then Exit Sub

'------ esto para filtrar solo inf. de GYE en MSC de import ------
' -------------------------  27/agosto/2001  ---------------------
st_continua = True
If Gstr_Linea = "MSC" And Gstr_Tipo = "I" Then
    If Gbol_solo_GYE Then
        If r_blimpo.pto_descarga <> "ECGYE" Then st_continua = False
    End If
End If
If Not st_continua Then Exit Sub

   
' ************* esto es nuevo al 17/oct/2000 cra. ****************
' aqui se evalua si el contenedor ya existe dentro de ese mismo bl en ese Puerto se acumulan los pesos, volumenes y bultos
   If Trim$(r_blimpodt.CONTEINER) <> "" And Gstr_Linea = "MSC" Then
        If Gstr_Tipo = "I" Then
            sql = "SELECT * FROM docu_Bldet "
        ElseIf Gstr_Tipo = "E" Then
            sql = "SELECT * FROM docu_Bldet_exp "
        End If
        sql = sql & " where bldet_linea       ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
              "   and bldet_buque    ='" & r_registro.buque & "'" & _
              "   and bldet_viaje    ='" & r_registro.viaje & "'" & _
              "   and bldet_puer_emb ='" & r_manifiesto.puerto & "'" & _
              "   and bldet_puer_des ='" & r_manifiesto.pto_des & "'" & _
              "   and bldet_emp_resp ='" & Gstr_param.naviera & "'" & _
              "   and bldet_cono_emb ='" & r_blimpo.COD_BL & "'" & _
              "   and bldet_contenedor='" & r_blimpodt.CONTEINER & "'"
              
        'Set reg = Base_APG.OpenRecordset(sql)
        reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        If Not reg.EOF Then
            If Gstr_Tipo = "I" Then
                sql = "Update docu_Bldet set "
            ElseIf Gstr_Tipo = "E" Then
                sql = "Update docu_Bldet_exp set "
            End If
            sql = sql & " bldet_bultos = " & CDbl(reg("no_bultos") + r_blimpodt.no_bultos) & _
                 ", BLDET_peso =" & CDbl(reg("peso") + r_blimpodt.peso) & _
                 ", bldet_volumen=" & CDbl(reg("volumen") + r_blimpodt.volumen) & _
                 ", bldet_tipo_prod = '" & r_blimpodt.producto & "' " & _
                 " where bldet_linea    ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
                 "   and bldet_buque    ='" & r_registro.buque & "'" & _
                 "   and bldet_viaje    ='" & r_registro.viaje & "'" & _
                 "   and bldet_puer_emb ='" & r_manifiesto.puerto & "'" & _
                 "   and bldet_puer_des ='" & r_manifiesto.pto_des & "'" & _
                 "   and bldet_emp_resp ='" & Gstr_param.naviera & "'" & _
                 "   and bldet_cono_emb ='" & r_blimpo.COD_BL & "'" & _
                 "   and bldet_secuencia= " & reg("bldet_secuencia")
           Gcon_docu.Execute sql
           GoTo fin
        End If
        reg.Close
   End If
' ****************************************************************
' ****************************************************************
   If Gstr_Tipo = "I" Then
        sql = "Select max(bldet_secuencia) from docu_bldet "
   ElseIf Gstr_Tipo = "E" Then
        sql = "Select max(bldet_secuencia) from docu_bldet_exp "
   End If
   sql = sql & " where bldet_linea    ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
         "   and bldet_buque    ='" & r_registro.buque & "'" & _
         "   and bldet_viaje    ='" & r_registro.viaje & "'" & _
         "   and bldet_puer_emb ='" & r_manifiesto.puerto & "'" & _
         "   and bldet_puer_des ='" & r_manifiesto.pto_des & "'" & _
         "   and bldet_emp_resp ='" & Gstr_param.naviera & "'" & _
         "   and bldet_cono_emb ='" & r_blimpodt.COD_BL & "'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   If reg.EOF Then
      SECUENCIA = 1
   Else
      If Not IsNull(reg(0)) Then
         SECUENCIA = reg(0) + 1
      Else
         SECUENCIA = 1
      End If
   End If
   If r_blimpodt.CONTEINER <> "" Then
      r_blimpodt.tipo_carga = Gstr_TIPO_CARGA_CONT '"5"
   Else
      r_blimpodt.tipo_carga = Gstr_TIPO_CARGA_GRAL '"1"
   End If
   If r_blimpodt.danger = "" Then
      r_blimpodt.danger = Gstr_IMO
   End If
   ' Julio/2008 se agrego estos valores para incluir el valor por defecto del texto de numeros y marcas y de descripcion de la carga
   If Trim(r_blimpodt.marca_num) = "" Then
      r_blimpodt.marca_num = Gstr_Marca_num
   End If
   r_blimpodt.desc_carga = Gstr_Desc_carga & Chr(13) & r_blimpodt.desc_carga
   '------------------------- final del cambio del mes de julio/2008 ---------------------------------------------------------------
   
   r_blimpodt.sec_carga = SECUENCIA
   
    ' SE EVALUA SI EL VOLUMEN ESTA EN CERO O VACIO (DESDE EL DTX)
    If r_blimpodt.volumen = 0 Then
       If Gint_size_contein = 20 Or Gint_size_contein = 1 Or Gint_size_contein = 2 Then ' LOS VALORES 1 y 2 SON CONT DE 20' EN LA LINEA KHS
          r_blimpodt.volumen = Gint_volu_20
       ElseIf Gint_size_contein = 40 Or Gint_size_contein = 3 Or Gint_size_contein = 4 Then ' LOS VALORES 3 y 4 SON CONT DE 40' EN LA LINEA KHS
          r_blimpodt.volumen = Gint_volu_40
       End If
    End If
    ' DTX_PROCESO.Cmb_pto_desc.BoundText & "','"
    If Gstr_Tipo = "I" Then
        sql = "Insert into docu_bldet "
    ElseIf Gstr_Tipo = "E" Then
        sql = "Insert into docu_bldet_exp "
    End If
    sql = sql & "                (bldet_linea,      bldet_buque,    bldet_viaje,    bldet_puer_emb,  " & _
                                " bldet_puer_des,   bldet_emp_resp, bldet_cono_emb, bldet_secuencia, " & _
                                " bldet_tipo_carga, bldet_peso,     bldet_bultos,   bldet_embalaje, " & _
                                " bldet_marcas_num, bldet_desc_carg,bldet_tipo_prod,bldet_volumen,  " & _
                                " bldet_contenedor, bldet_IMO,      bldet_pelig_ONU)" & _
       " Values('" & _
       DTX_PROCESO.Cmb_linea.BoundText & "','" & _
       r_registro.buque & "','" & _
       r_registro.viaje & "','" & _
       r_blimpo.puerto & "','" & _
       r_manifiesto.pto_des & "','" & _
       Gstr_param.naviera & "','" & _
       r_blimpodt.COD_BL & "','" & _
       r_blimpodt.sec_carga & "','" & _
       r_blimpodt.tipo_carga & "',"
   sql = sql & CDbl(r_blimpodt.peso) & ", '" & _
       r_blimpodt.no_bultos & "', '" & _
       r_blimpodt.embalaje & "','" & _
       chk_str(r_blimpodt.marca_num) & "','"
        
        
  If r_blimpo.pto_descarga <> r_blimpo.pto_destino And Mid(r_blimpo.pto_destino, 1, 2) <> "EC" Then 'And Mid(r_blimpo.pto_descarga, 1, 2) <> "EC"
        sql = sql & Gstr_trasbordo & " "
  End If

sql = sql & chk_str(r_blimpodt.desc_carga) & "', '" & r_blimpodt.producto & "', " & _
       CDbl(r_blimpodt.volumen) & ",'" & _
       r_blimpodt.CONTEINER & "','" & _
       r_blimpodt.danger & "','" & _
       r_blimpodt.IMO_ONU & "')"
       
       Gcon_docu.Execute sql
       'Debug.Print sql
       
       ' puesto para validar si el peso del contenedor viene vacio se graba el peso del detalle del bl, esto para MSC
       ' 21 septiembre 2002
       If (Not Gbol_bl_detallado_MSC And Gstr_Linea = "MSC") Or Gstr_Linea = "KHS" Then
            If Gstr_Tipo = "I" Then
                sql = "UPDATE docu_cont"
            ElseIf Gstr_Tipo = "E" Then
                sql = "UPDATE docu_cont_exp"
            End If
            sql = sql & " set cont_peso = " & r_blimpodt.peso & _
                  " Where cont_linea   ='" & DTX_PROCESO.Cmb_linea.BoundText & _
                  "'  and cont_buque     = '" & r_registro.buque & _
                  "'  and cont_viaje     = '" & r_registro.viaje & _
                  "'  and cont_puer_emb  = '" & r_blimpo.puerto & _
                  "'  and cont_puer_des  = '" & r_manifiesto.pto_des & _
                  "'  and cont_contenedor= '" & r_blimpodt.CONTEINER & _
                  "'  and cont_peso = 0"
            Gcon_docu.Execute sql, n
       End If

fin:
r_blimpo.COD_BL = ""
Exit Sub

error_graba:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0
End Sub

' 18 de octubre de 2002
' funcion que permite grabar los fletes del bl para poder tener mayor informacion de fletes
Public Sub graba_fletes()
Dim sql   As String
Dim reg   As New ADODB.Recordset
On Error GoTo error_graba
   If Gstr_Tipo = "I" Then
        sql = "Select * from docu_Blcab"
   ElseIf Gstr_Tipo = "E" Then
        sql = "Select * from docu_Blcab_exp"
   End If
   sql = sql & "  Where blcab_linea    ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
                "   and blcab_buque    ='" & r_registro.buque & "'" & _
                "   and blcab_viaje    ='" & r_registro.viaje & "'" & _
                "   and blcab_puer_emb ='" & r_manifiesto.puerto & "'" & _
                "   and blcab_puer_des ='" & r_manifiesto.pto_des & "'" & _
                "   and blcab_emp_resp ='" & Gstr_param.naviera & "'" & _
                "   and blcab_cono_emb ='" & r_blimpodt.COD_BL & "'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   If Not reg.EOF Then  ' solo si el bl ya existe se manda a grabar el/los fletes por el constrain de FK
        If UBound(Gtyp_flete_DTX) > 1 Then ' si existe al menos un flete en el arreglo (1 es el que siempre esta vacio)
            For i = 0 To UBound(Gtyp_flete_DTX) - 2
                If Gstr_Tipo = "I" Then
                    sql = "Insert into docu_blflete "
                ElseIf Gstr_Tipo = "E" Then
                    sql = "Insert into docu_blflete_exp "
                End If
                sql = sql & "                   (flete_linea,    flete_buque,    flete_viaje,    flete_puer_emb, " & _
                                               " flete_puer_des, flete_emp_resp, flete_cono_emb, flete_secuencia," & _
                                               " flete_cargo,    flete_debtor,   flete_valor,    flete_signo,    " & _
                                               " flete_moneda,   flete_marcas,   flete_lugar,    flete_imp_bl)" & _
                      " Values  ('" & DTX_PROCESO.Cmb_linea.BoundText & "', " & _
                                "'" & r_registro.buque & "',    " & _
                                "'" & r_registro.viaje & "',    " & _
                                "'" & r_manifiesto.puerto & "', " & _
                                "'" & r_manifiesto.pto_des & "'," & _
                                "'" & Gstr_param.naviera & "',  " & _
                                "'" & r_blimpodt.COD_BL & "',     " & _
                                " " & i + 1 & ",                " & _
                                "'" & Gtyp_flete_DTX(i).cargo & "',  " & _
                                "'" & Gtyp_flete_DTX(i).debtor & "', " & _
                                " " & Gtyp_flete_DTX(i).valor & ",   " & _
                                "'" & Gtyp_flete_DTX(i).SIGNO & "',  " & _
                                "'" & Gtyp_flete_DTX(i).MONEDA & "', " & _
                                "'" & chk_str(Gtyp_flete_DTX(i).marcas) & "', " & _
                                "'" & Gtyp_flete_DTX(i).lugar & "',  " & _
                                "'" & IIf(Gtyp_flete_DTX(i).imp_bl, "S", "N") & "')"
                Gcon_docu.Execute sql
            Next i
            ' abril 2004 cra
            ' aqui se debe incluir un update para actualizar el valor del flete.
            
            Call inicializa_fletes
        End If
   End If
Exit Sub

error_graba:
    MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
    Exit Sub
    Resume 0

End Sub

Public Sub graba_blexpo()
Dim sql   As String
Dim reg   As New ADODB.Recordset

On Error GoTo error_graba

If Trim$(r_blimpo.COD_BL) = "" Then Exit Sub
   Gint_no_Bls = Gint_no_Bls + 1
  
   sql = "Select * from Blexpo where " & _
         " pto_desc     ='" & Gstr_Pto_desc & "'" & _
         " and anio     ='" & r_blimpo.buque & "'" & _
         " and registro ='" & r_blimpo.viaje & "'" & _
         " and puerto   ='" & r_blimpo.pto_destino & "'" & _
         " and cod_bl   ='" & r_blimpo.COD_BL & "'"
   'Set reg = Base_APG.OpenRecordset(sql)
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   
   ' donde decía consignatario ahora dice embarcador
   If reg.EOF Then
       'r_blimpo.comentario = " "
       'DTX_PROCESO.Cmb_pto_desc.BoundText & "','" & _

       sql = "Insert into Blexpo Values('" & _
           Gstr_Pto_desc & "', '" & _
           r_blimpo.buque & "','" & _
           r_blimpo.viaje & "','" & _
           r_blimpo.pto_destino & "','" & _
           r_blimpo.COD_BL & "','" & _
           Gstr_Linea_Nav_APG & "','" & _
           r_blimpo.tipo_id & "', '" & _
           r_blimpo.Impexp & "','" & _
           chk_str(r_blimpo.embarcador) & "','" & _
           r_blimpo.pto_origen & "','" & _
           r_blimpo.puerto & "','" & _
           r_blimpo.pto_descarga & "','" & _
           r_blimpo.fecha_embarque & "','" & _
           r_blimpo.peso_kls & "','" & _
           r_blimpo.volumen_mt3 & "','" & _
           r_blimpo.no_bultos & "','" & _
           "','" & _
            "',Null,Null)"
       Gcon_docu.Execute sql
   End If
   ' se cambia el puerto de destino al campo puerto que es clave primaria
   ' en el detalle...
   r_blimpodt.puerto = r_blimpo.pto_destino
' ------- se borran los detalles ----------
sql = "Delete from blexpodt where " & _
      " pto_desc     = '" & Gstr_Pto_desc & "'" & _
      " and anio     = '" & r_blimpo.buque & "'" & _
      " and registro = '" & r_blimpo.viaje & "'" & _
      " and puerto   = '" & r_blimpo.pto_destino & "'" & _
      " and cod_bl   = '" & r_blimpo.COD_BL & "'"
Gcon_docu.Execute sql
' ----------------------------------------------

r_blimpo.COD_BL = ""
Exit Sub

error_graba:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0
End Sub

Public Sub graba_blexpodt()
Dim sql As String
Dim reg As New ADODB.Recordset
Dim SECUENCIA As Integer

On Error GoTo error_graba

If Trim$(r_blimpodt.COD_BL) = "" Then Exit Sub
If Trim$(r_blimpodt.embalaje) = "" Then Exit Sub

Gint_no_Bls_det = Gint_no_Bls_det + 1

   sql = "Select max(sec_carga) from Blexpodt " & _
         " where pto_desc ='" & Gstr_Pto_desc & "'" & _
         "   and anio     ='" & r_blimpodt.buque & "'" & _
         "   and registro ='" & r_blimpodt.viaje & "'" & _
         "   and puerto   ='" & r_blimpodt.puerto & "'" & _
         "   and cod_bl   ='" & r_blimpodt.COD_BL & "'"
   'Set reg = Base_APG.OpenRecordset(sql)
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   If reg.EOF Then
      SECUENCIA = 1
   Else
      If Not IsNull(reg(0)) Then
         SECUENCIA = reg(0) + 1
      Else
         SECUENCIA = 1
      End If
   End If
   If r_blimpodt.CONTEINER <> "" Then
      r_blimpodt.tipo_carga = "40"
   Else
      r_blimpodt.tipo_carga = "10"
   End If
   If r_blimpodt.danger = "" Then
      r_blimpodt.danger = Gstr_IMO
   End If
   'r_blimpodt.carga = Gstr_carga
   'r_blimpodt.med_peso = Gstr_Med_Peso
   'r_blimpodt.med_volumen = Gstr_Med_Volumen
   r_blimpodt.sec_carga = SECUENCIA
    ' SE EVALUA SI EL VOLUMEN ESTA EN CERO O VACIO (DESDE EL DTX)
   If r_blimpodt.volumen = 0 Then
      If Gint_size_contein = 20 Then
         r_blimpodt.volumen = Gint_volu_20
      ElseIf Gint_size_contein = 40 Then
         r_blimpodt.volumen = Gint_volu_40
      End If
   End If
   'DTX_PROCESO.Cmb_pto_desc.BoundText & "','" & _

   sql = "Insert into Blexpodt Values('" & _
       Gstr_Pto_desc & "', '" & _
       r_blimpodt.buque & "','" & _
       r_blimpodt.viaje & "','" & _
       r_blimpodt.puerto & "','" & _
       r_blimpodt.COD_BL & "','" & _
       r_blimpodt.sec_carga & "','" & _
       r_blimpodt.tipo_carga & "','" & _
       r_blimpodt.danger & "', '" & _
       r_blimpodt.embalaje & "','" & _
       r_blimpodt.peso & "','" & _
       r_blimpodt.volumen & "','" & _
       r_blimpodt.no_bultos & "','" & _
       r_blimpodt.CONTEINER & "','" & _
       r_impocont.cond_contein & "')"
       Gcon_docu.Execute sql

r_blimpo.COD_BL = ""
Exit Sub

error_graba:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0

End Sub

Public Sub actualiza_cabecera()
Dim sql         As String
Dim reg_impo    As New ADODB.Recordset
Dim reg         As New ADODB.Recordset
Dim reg1        As New ADODB.Recordset
Dim reg_no_det  As New ADODB.Recordset
Dim cont20      As Integer
Dim cont40      As Integer
Dim peso        As Double
Dim volumen     As Double
Dim bultos      As Double
Dim desc_carga  As String

On Error GoTo error_graba

If Trim$(r_blimpodt.COD_BL) = "" Then Exit Sub
   cont20 = 0: cont40 = 0
   peso = 0: volumen = 0: bultos = 0
   
   ' si se trata de la linea KHS se el valor de peso  y volumen esta a nivel de cabecera, no de detalle
   If Gstr_Linea = "KHS" Then
      sql = "Select blcab_peso_tot, blcab_volumen_tot "
      If Gstr_Tipo = "I" Then sql = sql & " from docu_blcab " Else sql = sql & " from docu_blcab_exp "
      sql = sql & " where blcab_linea    ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
            "   and blcab_buque    ='" & r_blimpodt.buque & "'" & _
            "   and blcab_viaje    ='" & r_blimpodt.viaje & "'" & _
            "   and blcab_puer_emb ='" & r_blimpodt.puerto & "'" & _
            "   and blcab_puer_des ='" & r_blimpodt.pto_descarga & "'" & _
            "   and blcab_emp_resp ='" & Gstr_param.naviera & "'" & _
            "   and blcab_cono_emb ='" & r_blimpodt.COD_BL & "'"
      reg_impo.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
      
      sql = "Select count(*) as cuantos from "
      If Gstr_Tipo = "I" Then sql = sql & " docu_bldet " Else sql = sql & " docu_bldet_exp "
      sql = sql & " where bldet_linea       ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
              "   and bldet_buque    ='" & r_blimpodt.buque & "'" & _
              "   and bldet_viaje    ='" & r_blimpodt.viaje & "'" & _
              "   and bldet_puer_emb ='" & r_blimpodt.puerto & "'" & _
              "   and bldet_puer_des ='" & r_blimpodt.pto_descarga & "'" & _
              "   and bldet_emp_resp ='" & Gstr_param.naviera & "'" & _
              "   and bldet_cono_emb ='" & r_blimpodt.COD_BL & "'"
                  
      reg_no_det.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
      If Not reg_no_det.EOF Then
         If reg_no_det!cuantos > 0 Then
            If Gstr_Tipo = "I" Then sql = "Update docu_bldet " Else sql = "Update docu_bldet_exp "
            sql = sql & " set bldet_peso = " & reg_impo("blcab_peso_tot") / reg_no_det("cuantos") & " , bldet_volumen = " & reg_impo("blcab_volumen_tot") / reg_no_det("cuantos") & _
                  " where bldet_linea    ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
                  "   and bldet_buque    ='" & r_blimpodt.buque & "'" & _
                  "   and bldet_viaje    ='" & r_blimpodt.viaje & "'" & _
                  "   and bldet_puer_emb ='" & r_blimpodt.puerto & "'" & _
                  "   and bldet_puer_des ='" & r_blimpodt.pto_descarga & "'" & _
                  "   and bldet_emp_resp ='" & Gstr_param.naviera & "'" & _
                  "   and bldet_cono_emb ='" & r_blimpodt.COD_BL & "'"
           ' Gcon_docu.Execute sql
         End If
      End If
   End If ' hasta aqui la validación solo para KHS
      
   'sumo peso, volumen , no_bultos
   sql = "Select sum(bldet_peso),sum(bldet_volumen),sum(bldet_bultos) From "
   If Gstr_Tipo = "I" Then sql = sql & "docu_bldet " Else sql = sql & "docu_bldet_exp "
   sql = sql & " Where bldet_linea    ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
         "   and bldet_buque    ='" & r_blimpodt.buque & "'" & _
         "   and bldet_viaje    ='" & r_blimpodt.viaje & "'" & _
         "   and bldet_puer_emb ='" & r_blimpodt.puerto & "'" & _
         "   and bldet_puer_des ='" & r_blimpodt.pto_descarga & "'" & _
         "   and bldet_emp_resp ='" & Gstr_param.naviera & "'" & _
         "   and bldet_cono_emb ='" & r_blimpodt.COD_BL & "'"
                  
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   If IsNull(reg(0)) Then
      peso = 0
   Else
      peso = reg(0)
   End If
   If IsNull(reg(1)) Then
      volumen = 0
   Else
      volumen = reg(1)
   End If
   If IsNull(reg(2)) Then
      bultos = 0
   Else
      bultos = reg(2)
   End If
   reg.Close
   
   'descripcion_carga
   sql = "Select bldet_desc_carg from "
   If Gstr_Tipo = "I" Then sql = sql & "docu_bldet " Else sql = sql & "docu_bldet_exp "
   sql = sql & " Where bldet_linea    ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
         "   and bldet_buque    ='" & r_blimpodt.buque & "'" & _
         "   and bldet_viaje    ='" & r_blimpodt.viaje & "'" & _
         "   and bldet_puer_emb ='" & r_blimpodt.puerto & "'" & _
         "   and bldet_puer_des ='" & r_blimpodt.pto_descarga & "'" & _
         "   and bldet_emp_resp ='" & Gstr_param.naviera & "'" & _
         "   and bldet_cono_emb ='" & r_blimpodt.COD_BL & "'"
   
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While IsNull(reg(0)) And Not reg.EOF
      reg.MoveNext
   Loop
   If Not reg.EOF Then
      desc_carga = reg(0)
   Else
      desc_carga = ""
   End If
   
      
   If Gstr_Tipo = "I" Then sql = sql & "Update docu_blcab " Else sql = sql & "Update docu_blcab_exp "
   sql = sql & " Set blcab_peso_tot = " & peso & _
             " , blcab_volumen_tot = " & volumen & " , blcab_bultos_tot = " & bultos & _
            " Where blcab_linea    ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
            "   and blcab_buque    ='" & r_blimpodt.buque & "'" & _
            "   and blcab_viaje    ='" & r_blimpodt.viaje & "'" & _
            "   and blcab_puer_emb ='" & r_blimpodt.puerto & "'" & _
            "   and blcab_puer_des ='" & r_blimpodt.pto_descarga & "'" & _
            "   and blcab_emp_resp ='" & Gstr_param.naviera & "'" & _
            "   and blcab_cono_emb ='" & r_blimpodt.COD_BL & "'"
   Gcon_docu.Execute sql
Exit Sub

error_graba:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0
   
End Sub

Public Sub actualiza_manifiesto()
Dim sql     As String
Dim reg     As New ADODB.Recordset
Dim reg1    As New ADODB.Recordset
Dim No_bls  As Integer
Dim peso    As Double
Dim volumen As Double
Dim bultos  As Double
Dim no_cont As Double

On Error GoTo error_graba
   
   peso = 0: volumen = 0: bultos = 0: No_bls = 0
   If r_blimpo.puerto = "" Then Exit Sub
   
   'sumo peso, volumen , no_bultos, no_bls
   sql = "Select sum(blcab_peso_tot), sum(blcab_volumen_tot), sum(blcab_bultos_tot)," & _
         "     count(blcab_cono_emb) from "
   If Gstr_Tipo = "I" Then sql = sql & "docu_blcab " Else sql = sql & "docu_blcab_exp "
   sql = sql & " Where blcab_linea    ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
         "   and blcab_buque    ='" & r_blimpo.buque & "'" & _
         "   and blcab_viaje    ='" & r_blimpo.viaje & "'" & _
         "   and blcab_puer_emb ='" & r_blimpo.puerto & "'" & _
         "   and blcab_puer_des ='" & r_blimpo.pto_descarga & "'" '& _
         "   and blcab_emp_resp ='" & Gstr_param.naviera & "'" & _
         "   and blcab_cono_emb ='" & r_blimpodt.COD_BL & "'"

   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   If IsNull(reg(0)) Then
      peso = 0
   Else
      peso = reg(0)
   End If
   If IsNull(reg(1)) Then
      volumen = 0
   Else
      volumen = reg(1)
   End If
   If IsNull(reg(2)) Then
      bultos = 0
   Else
      bultos = reg(2)
   End If
   If IsNull(reg(3)) Then
      No_bls = 0
   Else
      No_bls = reg(3)
   End If
   reg.Close
   sql = "Select distinct (bldet_contenedor) from "
   If Gstr_Tipo = "I" Then sql = sql & "docu_bldet, docu_cont " Else sql = sql & "docu_bldet_exp, docu_cont_exp "
   sql = sql & " Where bldet_linea    ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
         "   and bldet_buque    ='" & r_blimpo.buque & "'" & _
         "   and bldet_viaje    ='" & r_blimpo.viaje & "'" & _
         "   and bldet_puer_emb ='" & r_blimpo.puerto & "'" & _
         "   and bldet_puer_des ='" & r_blimpo.pto_descarga & "'" & _
         "   and bldet_emp_resp ='" & Gstr_param.naviera & "'" & _
         "   and bldet_linea    = cont_linea    " & _
         "   and bldet_buque    = cont_buque    " & _
         "   and bldet_viaje    = cont_viaje    " & _
         "   and bldet_puer_emb = cont_puer_emb " & _
         "   and bldet_puer_des = cont_puer_des " & _
         "   and bldet_contenedor=cont_contenedor"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   If reg.EOF Then
      no_cont = 0
   Else
      reg.MoveLast
      no_cont = reg.RecordCount
   End If
   
   sql = "Update "
   If Gstr_Tipo = "I" Then sql = sql & "docu_manifiesto " Else sql = sql & "docu_manifiesto_exp "
   sql = sql & " set mani_tot_peso = " & peso & _
         " , mani_tot_vol  = " & volumen & _
         " , mani_tot_bult = " & bultos & _
         " , mani_tot_doc  = " & No_bls & _
         " , mani_tot_con  = " & no_cont & _
         " Where mani_linea    ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
         "   and mani_buque    ='" & r_blimpo.buque & "'" & _
         "   and mani_viaje    ='" & r_blimpo.viaje & "'" & _
         "   and mani_puer_emb ='" & r_blimpo.puerto & "'" & _
         "   and mani_puer_des ='" & r_blimpo.pto_descarga & "'"
   Gcon_docu.Execute sql
Exit Sub

error_graba:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0
   
End Sub

Public Sub actualiza_cabecera_expo()
Dim sql         As String
Dim reg_expo    As New ADODB.Recordset
Dim reg_no_det  As New ADODB.Recordset
Dim reg         As New ADODB.Recordset
Dim reg1        As New ADODB.Recordset
Dim cont20      As Integer
Dim cont40      As Integer
Dim peso        As Double
Dim volumen     As Double
Dim bultos      As Double
Dim desc_carga  As String

On Error GoTo error_graba

If Trim$(r_blimpodt.COD_BL) = "" Then Exit Sub
   cont20 = 0: cont40 = 0
   peso = 0: volumen = 0: bultos = 0
   ' si se trata de la linea KHS se el valor de peso  y volumen esta a nivel de cabecera, no de detalle
   If Gstr_Linea = "KHS" Then
      sql = "Select peso_kls, volumen_mt3 from Blexpo " & _
            " where pto_desc ='" & Gstr_Pto_desc & "'" & _
            "   and anio     ='" & r_blimpodt.buque & "'" & _
            "   and registro ='" & r_blimpodt.viaje & "'" & _
            "   and puerto   ='" & r_blimpodt.puerto & "'" & _
            "   and cod_bl   ='" & r_blimpodt.COD_BL & "'"
      'Set reg_expo = Base_APG.OpenRecordset(sql)
      reg_expo.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
      
      sql = "Select count(*) as cuantos from Blexpodt " & _
            " where pto_desc ='" & Gstr_Pto_desc & "'" & _
            "   and anio     ='" & r_blimpodtBUQUE & "'" & _
            "   and registro ='" & r_blimpodt.viaje & "'" & _
            "   and puerto   ='" & r_blimpodt.puerto & "'" & _
            "   and cod_bl   ='" & r_blimpodt.COD_BL & "'"
      'Set reg_no_det = Base_APG.OpenRecordset(sql)
      reg_no_det.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
      If Not reg_no_det.EOF Then
         If reg_no_det!cuantos > 0 Then
            sql = "Update blexpodt set peso = " & reg_expo!peso_kls / reg_no_det!cuantos & _
                  " , volumen = " & reg_expo!volumen_mt3 / reg_no_det!cuantos & _
                  " Where pto_desc ='" & Gstr_Pto_desc & "'" & _
                  "   and anio     ='" & r_blimpodt.buque & "'" & _
                  "   and registro ='" & r_blimpodt.viaje & "'" & _
                  "   and puerto   ='" & r_blimpodt.puerto & "'" & _
                  "   and cod_bl   ='" & r_blimpodt.COD_BL & "'"
            Gcon_docu.Execute sql
         End If
      End If
   End If ' hasta aqui la validación solo para KHS
   
   'sumo peso, volumen , no_bultos
   sql = "Select sum(peso),sum(volumen),sum(no_bultos) from Blexpodt " & _
         " where pto_desc ='" & Gstr_Pto_desc & "'" & _
         "   and anio     ='" & r_blimpodt.buque & "'" & _
         "   and registro ='" & r_blimpodt.viaje & "'" & _
         "   and puerto   ='" & r_blimpodt.puerto & "'" & _
         "   and cod_bl   ='" & r_blimpodt.COD_BL & "'"
   'Set reg = Base_APG.OpenRecordset(sql)
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   If IsNull(reg(0)) Then
      peso = 0
   Else
      peso = reg(0)
   End If
   If IsNull(reg(1)) Then
      volumen = 0
   Else
      volumen = reg(1)
   End If
   If IsNull(reg(2)) Then
      bultos = 0
   Else
      bultos = reg(2)
   End If
      
   
   'contenedores 20
   sql = "Select distinct (a.conteiner) from blexpodt a " & _
         " where a.pto_desc ='" & Gstr_Pto_desc & "'" & _
         "   and a.anio     ='" & r_blimpodt.buque & "'" & _
         "   and a.registro ='" & r_blimpodt.viaje & "'" & _
         "   and a.puerto   ='" & r_blimpodt.puerto & "'" & _
         "   and a.cod_bl   ='" & r_blimpodt.COD_BL & "'" & _
         "   and mid$(A.tipo_contein,1,2) = '20'"
   'Set reg1 = Base_APG.OpenRecordset(sql)
   reg1.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   If reg1.EOF Then
      cont20 = 0
   Else
      reg1.MoveLast
      cont20 = reg1.RecordCount
   End If
   
   'contenedores 40
   sql = "Select distinct (a.conteiner) from blexpodt a " & _
         " where a.pto_desc ='" & Gstr_Pto_desc & "'" & _
         "   and a.anio     ='" & r_blimpodt.buque & "'" & _
         "   and a.registro ='" & r_blimpodt.viaje & "'" & _
         "   and a.puerto   ='" & r_blimpodt.puerto & "'" & _
         "   and a.cod_bl   ='" & r_blimpodt.COD_BL & "'" & _
         "   and mid$(A.tipo_contein,1,2) = '40'"
   'Set reg1 = Base_APG.OpenRecordset(sql)
   reg1.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   If reg1.EOF Then
      cont40 = 0
   Else
      reg1.MoveLast
      cont40 = reg1.RecordCount
   End If
   
   sql = "Update blexpo set " & _
         "   peso_kls    = " & peso & _
         " , volumen_mt3 = " & volumen & _
         " , no_bultos = " & bultos & _
         " , no_cont20 = " & cont20 & _
         " , no_cont40 = " & cont40 & _
         " Where pto_desc ='" & Gstr_Pto_desc & "'" & _
         "   and anio     ='" & r_blimpodt.buque & "'" & _
         "   and registro ='" & r_blimpodt.viaje & "'" & _
         "   and puerto   ='" & r_blimpodt.puerto & "'" & _
         "   and cod_bl   ='" & r_blimpodt.COD_BL & "'"
   Gcon_docu.Execute sql
Exit Sub

error_graba:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0
   

End Sub

Public Sub graba_impocont()
Dim sql As String
Dim reg As New ADODB.Recordset
Dim st_continua  As Boolean
On Error GoTo error_graba


st_continua = True
If Gstr_Linea = "MSC" And Gstr_Tipo = "I" Then
    If Gbol_solo_GYE Then
        If r_impocont.pto_des <> "ECGYE" Then st_continua = False
    End If
End If
If Not st_continua Then Exit Sub

If r_impocont.CONTEINER = "" Then Exit Sub
If Gstr_Tipo = "I" Then
    If r_impocont.puerto = "" Then Exit Sub
Else
    If r_impocont.pto_des = "" Then Exit Sub
End If

sql = "Select * from "
If Gstr_Tipo = "I" Then sql = sql & " docu_cont " Else sql = sql & " docu_cont_exp "
sql = sql & "Where cont_linea     ='" & DTX_PROCESO.Cmb_linea.BoundText & _
            "' and cont_buque     = '" & r_impocont.buque & _
            "' and cont_viaje     = '" & r_impocont.viaje & _
            "' and cont_puer_emb  = '" & r_impocont.puerto & _
            "' and cont_puer_des  = '" & r_impocont.pto_des & _
            "' and cont_contenedor= '" & r_impocont.CONTEINER & "'"

reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
If Not reg.EOF Then Exit Sub
   
If Trim$(r_impocont.ind_lleno_vacio) = "" Then r_impocont.ind_lleno_vacio = Gstr_param.indic_llen_vac
If Trim$(r_impocont.cond_contein) = "" Then r_impocont.cond_contein = Gstr_param.condicion_cont
If Trim(r_impocont.sello1) = "" Then r_impocont.sello1 = IIf(Trim(Gstr_sello_Contein) <> "", Gstr_sello_Contein, " ")
' si la tara viene en blanco se le setea un valor fijo
If r_impocont.tara = 0 Then
   If Codigo_equivalente("dtx_tipo_contein", r_impocont.tamano_equipo, "Size", "apg") = "20" Then
        r_impocont.tara = 2200
   Else
        r_impocont.tara = 4400
   End If
End If

sql = "Insert Into "
If Gstr_Tipo = "I" Then sql = sql & " docu_cont " Else sql = sql & " docu_cont_exp "
sql = sql & " Values('" & DTX_PROCESO.Cmb_linea.BoundText & _
                    "','" & r_impocont.buque & _
                    "','" & r_impocont.viaje & _
                    "','" & r_impocont.puerto & _
                    "','" & r_impocont.pto_des & _
                    "','" & r_impocont.CONTEINER & _
                    "','" & r_impocont.tamano_equipo & _
                    "','" & r_impocont.cond_contein & _
                    "','" & r_impocont.tara & _
                    "','" & Gstr_param.tipo_equipo & _
                    "','" & Gstr_param.proveed_eq
                    
If Gstr_Tipo = "I" Then sql = sql & "','" & Gstr_param.estado_equip Else sql = sql & "','" & Gstr_param.estado_equip_expo

sql = sql & "','" & r_impocont.ind_lleno_vacio & _
                    "','" & Gstr_param.cond_contrato & _
                    "','" & Gstr_param.requer_serv & _
                    "','" & Gstr_param.naturaleza_carga & _
                    "','" & r_impocont.peso & _
                    "','" & r_impocont.tipo_bien & _
                    "','" & r_impocont.sello1 & _
                    "','" & r_impocont.sello2 & _
                    "','" & r_impocont.sello3 & _
                    "','" & r_impocont.sello4 & _
                    "','" & r_impocont.temp_dur_almac & _
                    "','" & r_impocont.temp_dur_trans & _
                    "','" & r_impocont.temp_dur_manip
        sql = sql & "','" & r_impocont.temp_min_almac & _
                    "','" & r_impocont.temp_max_almac & _
                    "','" & r_impocont.temp_min_trans & _
                    "','" & r_impocont.temp_max_trans & _
                    "','" & r_impocont.temp_min_manip & _
                    "','" & r_impocont.temp_max_manip & _
                    "')"
                                
Gcon_docu.Execute sql
reg.Close

sql = "Select bldet_peso from"
sql = sql & IIf(Gstr_Tipo = "I", " docu_bldet ", " docu_bldet_exp ")
sql = sql & " Where bldet_linea   ='" & DTX_PROCESO.Cmb_linea.BoundText & _
                  "'  and bldet_buque     = '" & r_impocont.buque & _
                  "'  and bldet_viaje     = '" & r_impocont.viaje & _
                  "'  and bldet_puer_emb  = '" & r_impocont.puerto & _
                  "'  and bldet_puer_des  = '" & r_impocont.pto_des & _
                  "'  and bldet_contenedor= '" & r_impocont.CONTEINER & "'"
reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
 ' puesto para validar si el peso del contenedor viene vacio se graba el peso del detalle del bl, esto para MSC
 ' 21 septiembre 2002 esto esta en graba blimpodt se puso aqui por cambio de proceso por varios sellos mayo 25 2004
       If ((Not Gbol_bl_detallado_MSC And Gstr_Linea = "MSC") Or Gstr_Linea = "KHS") And Not reg.EOF Then
            If Gstr_Tipo = "I" Then
                sql = "UPDATE docu_cont"
            ElseIf Gstr_Tipo = "E" Then
                sql = "UPDATE docu_cont_exp"
            End If
            sql = sql & " set cont_peso = " & reg("bldet_peso") & _
                  " Where cont_linea   ='" & DTX_PROCESO.Cmb_linea.BoundText & _
                  "'  and cont_buque     = '" & r_impocont.buque & _
                  "'  and cont_viaje     = '" & r_impocont.viaje & _
                  "'  and cont_puer_emb  = '" & r_impocont.puerto & _
                  "'  and cont_puer_des  = '" & r_impocont.pto_des & _
                  "'  and cont_contenedor= '" & r_impocont.CONTEINER & _
                  "'  and cont_peso = 0"
            Gcon_docu.Execute sql, n
       End If

Exit Sub

error_graba:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0

End Sub

Public Sub graba_impo_vacios()
Dim sql As String
Dim reg As New ADODB.Recordset
Dim reg2 As New ADODB.Recordset
Dim st_continua  As Boolean

On Error GoTo error_graba


'------ esto para filtrar solo inf. de GYE en MSC de import ------
' -------------------------  27/agosto/2001  ---------------------
st_continua = True
If Gstr_Linea = "MSC" And Gstr_Tipo = "I" Then
    If Gbol_solo_GYE Then
        If r_blimpo.pto_descarga <> "ECGYE" Then st_continua = False
    End If
End If
If Not st_continua Then Exit Sub
If r_blimpo.puerto = "" Then Exit Sub

If r_impocont.tara = 0 Then
   If Codigo_equivalente("dtx_tipo_contein", r_impocont.tamano_equipo, "Size", "apg") = "20" Then
        r_impocont.tara = 2200
   Else
        r_impocont.tara = 4400
   End If
End If
If Trim(r_blimpodt.embalaje) = "" Then r_blimpodt.embalaje = Gstr_EMBAL_CONTEN

'------------------------------------------------------------------

sql = "Select * from "
If Gstr_Tipo = "I" Then sql = sql & " docu_cont " Else sql = sql & " docu_cont_exp "
sql = sql & " Where cont_linea    = '" & DTX_PROCESO.Cmb_linea.BoundText & _
            "' and cont_buque     = '" & r_impocont.buque & _
            "' and cont_viaje     = '" & r_impocont.viaje & _
            "' and cont_puer_emb  = '" & r_impocont.puerto & _
            "' and cont_puer_des  = '" & r_impocont.pto_des & _
            "' and cont_contenedor= '" & r_impocont.CONTEINER & "'"

reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
If reg.EOF Then
    
    
    
    r_impocont.cond_contein = Gstr_Cond_Contein_vacio 'nov 2007
    If Trim(r_impocont.sello1) = "" Then r_impocont.sello1 = IIf(Trim(Gstr_sello_Contein) <> "", Gstr_sello_Contein, " ") 'nov 2007
    If r_blimpodt.peso = 0 Then r_blimpodt.peso = r_impocont.tara    'nov 2007 se incluye el mimsmo valor que la tara al peso para los vacios
    
    sql = "Insert Into "
    If Gstr_Tipo = "I" Then sql = sql & " docu_cont " Else sql = sql & " docu_cont_exp "
    sql = sql & " Values('" & DTX_PROCESO.Cmb_linea.BoundText & _
                        "','" & r_impocont.buque & _
                        "','" & r_impocont.viaje & _
                        "','" & r_impocont.puerto & _
                        "','" & r_impocont.pto_des & _
                        "','" & r_impocont.CONTEINER & _
                        "','" & r_impocont.tamano_equipo & _
                        "','" & r_impocont.cond_contein & _
                        "','" & r_impocont.tara & _
                        "','" & Gstr_param.tipo_equipo & _
                        "','" & Gstr_param.proveed_eq
    If Gstr_Tipo = "I" Then sql = sql & "','" & Gstr_param.estado_equip Else sql = sql & "','" & Gstr_param.estado_equip_expo
    sql = sql & "','" & r_impocont.ind_lleno_vacio & _
                        "','" & Gstr_param.cond_contrato & _
                        "','" & Gstr_param.requer_serv & _
                        "','" & Gstr_param.naturaleza_carga & _
                        "','" & r_impocont.peso & _
                        "','" & r_impocont.tipo_bien & _
                        "','" & r_impocont.sello1 & _
                        "','" & r_impocont.sello2 & _
                        "','" & r_impocont.sello3 & _
                        "','" & r_impocont.sello4 & _
                        "','" & r_impocont.temp_dur_almac & _
                        "','" & r_impocont.temp_dur_trans & _
                        "','" & r_impocont.temp_dur_manip
            sql = sql & "','" & r_impocont.temp_min_almac & _
                        "','" & r_impocont.temp_max_almac & _
                        "','" & r_impocont.temp_min_trans & _
                        "','" & r_impocont.temp_max_trans & _
                        "','" & r_impocont.temp_min_manip & _
                        "','" & r_impocont.temp_max_manip & _
                        "')"
    Gcon_docu.Execute sql
    'Debug.Print "Se inghrso a impocont " & r_impocont.CONTEINER
End If
reg.Close
' hasta aqui se graba los datos del contenedor a la impocont y ahora se graban el detalle del Bl
'puede darse el caso que no  exista una cabecera de Bl y se este tratando de grabar un detalle
   If Trim$(r_blimpodt.CONTEINER) <> "" And Gstr_Linea = "MSC" Then
        sql = "SELECT * FROM "
        If Gstr_Tipo = "I" Then sql = sql & " docu_bldet " Else sql = sql & " docu_bldet_exp "
        sql = sql & " where bldet_linea='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
              "   and bldet_buque      ='" & r_blimpodt.buque & "'" & _
              "   and bldet_viaje      ='" & r_blimpodt.viaje & "'" & _
              "   and bldet_puer_emb   ='" & r_impocont.puerto & "'" & _
              "   and bldet_puer_des   ='" & r_impocont.pto_des & "'" & _
              "   and bldet_emp_resp   ='" & Gstr_param.naviera & "'" & _
              "   and bldet_cono_emb   ='" & r_blimpodt.COD_BL & "'" & _
              "   and bldet_contenedor ='" & r_blimpodt.CONTEINER & "'"
              
        'Set reg = Base_APG.OpenRecordset(sql)
        reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        If reg.EOF Then ' si existe no es un contenedor vacío por lo que se evalua por eof
            If Not Gbol_relacion_fija Then
                If MsgBox("El Contenedor """ & r_blimpodt.CONTEINER & """ no existe dentro de los bls, " & Chr(13) & Chr(10) & "Desea ingresarlo como VACIO en el Bl """ & r_blimpodt.COD_BL & """ del Puerto " & r_blimpodt.puerto & ".", vbInformation + vbYesNo, App.Title) = vbNo Then GoTo fin
            End If
            sql = "Select max(bldet_secuencia) from "
            If Gstr_Tipo = "I" Then sql = sql & " docu_bldet " Else sql = sql & " docu_bldet_exp "
            sql = sql & " where bldet_linea  ='" & DTX_PROCESO.Cmb_linea.BoundText & "'" & _
                    "   and bldet_buque      ='" & r_blimpodt.buque & "'" & _
                    "   and bldet_viaje      ='" & r_blimpodt.viaje & "'" & _
                    "   and bldet_puer_emb   ='" & r_impocont.puerto & "'" & _
                    "   and bldet_puer_des   ='" & r_impocont.pto_des & "'" & _
                    "   and bldet_emp_resp   ='" & Gstr_param.naviera & "'" & _
                    "   and bldet_cono_emb   ='" & r_blimpodt.COD_BL & "'"
            'Set reg = Base_APG.OpenRecordset(sql)
            reg2.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
            If reg2.EOF Then
               SECUENCIA = 1
            Else
               If Not IsNull(reg2(0)) Then
                  SECUENCIA = reg2(0) + 1
               Else
                  SECUENCIA = 1
               End If
            End If
            reg2.Close
            r_blimpodt.tipo_carga = Gstr_TIPO_CARGA_CONT_VACIO '"8" ' tipo de carga de contenedores vacios
            r_blimpodt.producto = Gstr_TIPO_PRODUCTO_VACIO
            If r_blimpodt.danger = "" Then
               r_blimpodt.danger = Gstr_IMO
            End If
            
            r_blimpodt.no_bultos = 1
            r_blimpodt.embalaje = Gstr_EMBAL_CONTEN ' "12"
            
            r_blimpodt.sec_carga = SECUENCIA
            ' SE EVALUA SI EL VOLUMEN ESTA EN CERO O VACIO (DESDE EL DTX)
             If r_blimpodt.volumen = 0 Then
               If Gint_size_contein = 20 Or Gint_size_contein = 1 Or Gint_size_contein = 2 Then ' LOS VALORES 1 y 2 SON CONT DE 20' EN LA LINEA KHS
                  r_blimpodt.volumen = Gint_volu_20
               ElseIf Gint_size_contein = 40 Or Gint_size_contein = 3 Or Gint_size_contein = 4 Then ' LOS VALORES 3 y 4 SON CONT DE 40' EN LA LINEA KHS
                  r_blimpodt.volumen = Gint_volu_40
               End If
             End If
             ' SE EVALUA SI EL PESO ESTA EN CERO O VACIO (DESDE EL DTX) (POSIBLEMENTE CONTENEDOR VACIO)
             If r_blimpodt.peso = 0 Then
               If Gint_size_contein = 20 Or Gint_size_contein = 1 Or Gint_size_contein = 2 Then ' LOS VALORES 1 y 2 SON CONT DE 20' EN LA LINEA KHS
                  r_blimpodt.peso = Gdbl_peso_vacio_20
               ElseIf Gint_size_contein = 40 Or Gint_size_contein = 3 Or Gint_size_contein = 4 Then ' LOS VALORES 3 y 4 SON CONT DE 40' EN LA LINEA KHS
                  r_blimpodt.peso = Gdbl_peso_vacio_40
               End If
             End If
             
             r_blimpodt.desc_carga = "EMPTY CONTEINER"
             ' DTX_PROCESO.Cmb_pto_desc.BoundText & "','" & _
             ' Julio/2008 se agrego estos valores para incluir el valor por defecto del texto de numeros y marcas y de descripcion de la carga
             
             If Trim(r_blimpodt.marca_num) = "" Then
                r_blimpodt.marca_num = Gstr_Marca_num
             End If
             r_blimpodt.desc_carga = Gstr_Desc_carga & Chr(13) & r_blimpodt.desc_carga
             

            sql = "Insert into "
            If Gstr_Tipo = "I" Then sql = sql & " docu_bldet " Else sql = sql & " docu_bldet_exp "
            sql = sql & "                 (bldet_linea,      bldet_buque,    bldet_viaje,    bldet_puer_emb,  " & _
                                         " bldet_puer_des,   bldet_emp_resp, bldet_cono_emb, bldet_secuencia, " & _
                                         " bldet_tipo_carga, bldet_peso,     bldet_bultos,   bldet_embalaje, " & _
                                         " bldet_marcas_num, bldet_desc_carg,bldet_volumen,  bldet_contenedor," & _
                                         " bldet_IMO,        bldet_pelig_ONU,bldet_tipo_prod)" & _
                " Values('" & _
                DTX_PROCESO.Cmb_linea.BoundText & "','" & _
                r_registro.buque & "','" & _
                r_registro.viaje & "','" & _
                r_impocont.puerto & "','" & _
                r_impocont.pto_des & "','" & _
                Gstr_param.naviera & "','" & _
                r_blimpodt.COD_BL & "','" & _
                r_blimpodt.sec_carga & "','" & _
                r_blimpodt.tipo_carga & "','"
            sql = sql & r_blimpodt.peso & "', '" & _
                r_blimpodt.no_bultos & "', '" & _
                r_blimpodt.embalaje & "','" & _
                chk_str(r_blimpodt.marca_num) & "','" & _
                chk_str(r_blimpodt.desc_carga) & "', '" & _
                r_blimpodt.volumen & "','" & _
                r_blimpodt.CONTEINER & "','" & _
                r_blimpodt.danger & "','" & _
                r_blimpodt.IMO_ONU & "','" & r_blimpodt.producto & "')"
            
             Gcon_docu.Execute sql
            ' Debug.Print "Se inghrso a blimpodt " & r_impocont.CONTEINER & " en el bl " & r_blimpodt.COD_BL
        End If
        reg.Close
   End If
' ****************************************************************
' ****************************************************************


sql = "Select bldet_peso from"
sql = sql & IIf(Gstr_Tipo = "I", " docu_bldet ", " docu_bldet_exp ")
sql = sql & " Where bldet_linea   ='" & DTX_PROCESO.Cmb_linea.BoundText & _
                  "'  and bldet_buque     = '" & r_impocont.buque & _
                  "'  and bldet_viaje     = '" & r_impocont.viaje & _
                  "'  and bldet_puer_emb  = '" & r_impocont.puerto & _
                  "'  and bldet_puer_des  = '" & r_impocont.pto_des & _
                  "'  and bldet_contenedor= '" & r_impocont.CONTEINER & "'"
reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
 ' puesto para validar si el peso del contenedor viene vacio se graba el peso del detalle del bl, esto para MSC
 ' 21 septiembre 2002 esto esta en graba blimpodt se puso aqui por cambio de proceso por varios sellos mayo 25 2004
       If ((Not Gbol_bl_detallado_MSC And Gstr_Linea = "MSC") Or Gstr_Linea = "KHS") And Not reg.EOF Then
            If Gstr_Tipo = "I" Then
                sql = "UPDATE docu_cont"
            ElseIf Gstr_Tipo = "E" Then
                sql = "UPDATE docu_cont_exp"
            End If
            sql = sql & " set cont_peso = " & reg("bldet_peso") & _
                  " Where cont_linea   ='" & DTX_PROCESO.Cmb_linea.BoundText & _
                  "'  and cont_buque     = '" & r_impocont.buque & _
                  "'  and cont_viaje     = '" & r_impocont.viaje & _
                  "'  and cont_puer_emb  = '" & r_impocont.puerto & _
                  "'  and cont_puer_des  = '" & r_impocont.pto_des & _
                  "'  and cont_contenedor= '" & r_impocont.CONTEINER & _
                  "'  and cont_peso = 0"
            Gcon_docu.Execute sql, n
       End If




fin:

Exit Sub

error_graba:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0

End Sub

Public Sub graba_expo_vacios()
' ya no se usa este procedimiento
Dim sql As String
Dim reg As New ADODB.Recordset
Dim SECUENCIA As Integer

On Error GoTo error_graba


Gint_no_Bls_det = Gint_no_Bls_det + 1
   ' se cambia el puerto de destino al campo puerto que es clave primaria
   ' en el detalle...
   r_blimpodt.puerto = r_blimpo.pto_destino

   sql = "Select * from Blexpodt " & _
         " where pto_desc ='" & Gstr_Pto_desc & "'" & _
         "   and anio     ='" & r_blimpodt.buque & "'" & _
         "   and registro ='" & r_blimpodt.viaje & "'" & _
         "   and puerto   ='" & r_blimpodt.puerto & "'" & _
         "   and cod_bl   ='" & r_blimpodt.COD_BL & "'" & _
         "   and conteiner='" & r_blimpodt.CONTEINER & "'"
   'Set reg = Base_APG.OpenRecordset(sql)
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
  
   If reg.EOF Then ' este contenedor no existe en este bl por llo que se asume que es vacio
        If MsgBox("El Contenedor """ & r_blimpodt.CONTEINER & """ no existe dentro de los bls " & Chr(13) & Chr(10) & " Desea ingresarlo como VACIO en el Bl """ & r_blimpodt.COD_BL & """ del Puerto " & r_blimpodt.puerto & ".", vbQuestion + vbYesNo, App.Title) = vbNo Then GoTo fin
        sql = "Select max(sec_carga) from Blexpodt " & _
              " where pto_desc ='" & Gstr_Pto_desc & "'" & _
              "   and anio     ='" & r_blimpodt.buque & "'" & _
              "   and registro ='" & r_blimpodt.viaje & "'" & _
              "   and puerto   ='" & r_blimpodt.puerto & "'" & _
              "   and cod_bl   ='" & r_blimpodt.COD_BL & "'"
        'Set reg = Base_APG.OpenRecordset(sql)
        reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        If reg.EOF Then
           SECUENCIA = 1
        Else
           If Not IsNull(reg(0)) Then
              SECUENCIA = reg(0) + 1
           Else
              SECUENCIA = 1
           End If
        End If
        r_blimpodt.tipo_carga = "44"
        
        
        If r_blimpodt.danger = "" Then
           r_blimpodt.danger = Gstr_IMO
        End If
        'r_blimpodt.carga = Gstr_carga
        'r_blimpodt.med_peso = Gstr_Med_Peso
        'r_blimpodt.med_volumen = Gstr_Med_Volumen
        r_blimpodt.embalaje = "CNT"
        r_blimpodt.no_bultos = 1
        r_blimpodt.sec_carga = SECUENCIA
         ' SE EVALUA SI EL VOLUMEN ESTA EN CERO O VACIO (DESDE EL DTX)
        If r_blimpodt.volumen = 0 Then
           If Gint_size_contein = 20 Then
              r_blimpodt.volumen = Gint_volu_20
           ElseIf Gint_size_contein = 40 Then
              r_blimpodt.volumen = Gint_volu_40
           End If
        End If
        ' SE EVALUA SI EL PESO ESTA EN CERO O VACIO (POSIBLEMENTE EN CONTEEDORES VACIOS)
        If r_blimpodt.peso = 0 Then
           r_blimpodt.peso = r_impocont.tara 'nov 2007 se incluye el mimsmo valor que la tara al peso para los vacios
           
'           If Gint_size_contein = 20 Then
'              r_blimpodt.peso = Gdbl_peso_vacio_20
'           ElseIf Gint_size_contein = 40 Then
'              r_blimpodt.peso = Gdbl_peso_vacio_40
'           End If
        End If
        'DTX_PROCESO.Cmb_pto_desc.BoundText & "','" & _

        r_impocont.cond_contein = Gstr_Cond_Contein_vacio 'nov 2007
        If Trim(r_impocont.sello1) = "" Then r_impocont.sello1 = IIf(Trim(Gstr_sello_Contein) <> "", Gstr_sello_Contein, " ") 'nov 2007

        sql = "Insert into Blexpodt Values('" & _
            Gstr_Pto_desc & "', '" & _
            r_blimpodt.buque & "','" & _
            r_blimpodt.viaje & "','" & _
            r_blimpodt.puerto & "','" & _
            r_blimpodt.COD_BL & "','" & _
            r_blimpodt.sec_carga & "','" & _
            r_blimpodt.tipo_carga & "','" & _
            r_blimpodt.danger & "', '" & _
            r_blimpodt.embalaje & "','" & _
            r_blimpodt.peso & "','" & _
            r_blimpodt.volumen & "','" & _
            r_blimpodt.no_bultos & "','" & _
            r_blimpodt.CONTEINER & "','" & _
            r_impocont.cond_contein & "')"
            Gcon_docu.Execute sql
            r_blimpo.COD_BL = ""
   End If

fin:
Exit Sub

error_graba:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0

End Sub

Public Sub procesa_BL1(cod_linea As String, linea As String)
Dim sql             As String
Dim reg             As New ADODB.Recordset
Dim ORIGEN, DESTINO As String
Dim condicion       As String
Dim pto_emb         As String
Dim pto_desc        As String
   
   r_blimpodt.buque = DTX_PROCESO.Cmb_buque.BoundText
   r_blimpodt.viaje = r_blimpo.viaje
   r_blimpodt.puerto = r_blimpo.puerto
       
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "COD_BL":
            Call inicializa_blimpo(False)
            'se ubica el registro del ultimo puerto leido se diferencia entre impo y expo.
            
            pto_emb = Grid_Codigo_equivalente("PTO_EMBA_ORIG")
            pto_desc = Grid_Codigo_equivalente("PTO_DESC_ORIG")
            
            If Not Grid_confirma_carga("BL", pto_emb, pto_desc, Mid$(linea, reg("desde"), reg("longitud"))) Then Exit Sub
            r_blimpo.COD_BL = Mid$(linea, reg("desde"), reg("longitud"))
            Call inicializa_blimpodt(False)
            r_blimpodt.COD_BL = Mid$(linea, reg("desde"), reg("longitud"))
      End Select
      reg.MoveNext
   Loop
   reg.Close
   
   ' se procesan los datos para la condicion del contenedor
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'IMPOCONT'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "COND_CONTEIN1":
             ORIGEN = Mid$(linea, reg("desde"), reg("longitud"))
         Case "COND_CONTEIN2":
             DESTINO = Mid$(linea, reg("desde"), reg("longitud"))
      End Select
      reg.MoveNext
   Loop
   If ORIGEN <> "" And DESTINO <> "" Then
      condicion = Codigo_equivalente_DTX("DTX_COND_CONTEIN", ORIGEN, "APG", "origen", DESTINO, " destino ")
      ' se pidio poner fijo el valor de la condicion del contenedor, esto mas en KHS ya que en MSC el valor viene bien
      r_impocont.cond_contein = condicion
   End If
'    If r_blimpo.COD_BL = "MSCUEC084596" Then
'        MsgBox "aqui esta el bl"
'    End If
   
   DTX_PROCESO.Lbl_proceso_BL.Caption = "Procesando " & "PTO Embarq " & r_blimpo.puerto & " PTO Descarg " & r_blimpo.pto_descarga & " BL No. " & r_blimpo.COD_BL
   DoEvents
End Sub

Public Sub procesa_BA1(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
Dim tipo As String
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
   'Set reg = Gcon_docu.OpenRecordset(sql)
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "TIPO_ID"
            tipo = Trim(Mid$(linea, reg("desde"), reg("longitud")))
         Case "NOMBRE"
            Select Case tipo
               Case "SH"
                    r_blimpo.embarcador = r_blimpo.embarcador & Trim(Mid$(linea, reg("desde"), reg("longitud")))
               Case "CN"
                   r_blimpo.tipo_id = Gstr_Tipo_ID
                   r_blimpo.Impexp = Gstr_Impexp
                   r_blimpo.nacio_consig = Gstr_param.nacionalidad
                   r_blimpo.consignatario = r_blimpo.consignatario & Trim(Mid$(linea, reg("desde"), reg("longitud")))
               Case "N1"
                    r_blimpo.tipo_id_notif = Gstr_Tipo_ID
                    r_blimpo.no_notyfier = r_blimpo.no_notyfier & Trim$(Mid$(linea, reg("desde"), reg("longitud")))
               Case "N2"
                    r_blimpo.no_notyfier2 = r_blimpo.no_notyfier2 & Trim$(Mid$(linea, reg("desde"), reg("longitud")))
            End Select
      End Select
      reg.MoveNext
   Loop

End Sub

Public Sub procesa_BA3(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
Dim tipo As String
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "'  And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
   'Set reg = Gcon_docu.OpenRecordset(sql)
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "TIPO_ID"
            tipo = Trim(Mid$(linea, reg("desde"), reg("longitud")))
         Case "DIRECCION"
            Select Case tipo
               Case "SH"
                    r_blimpo.direc_exp = r_blimpo.direc_exp & Trim(Mid$(linea, reg("desde"), reg("longitud")))
               Case "CN"
                   r_blimpo.direc_imp = r_blimpo.direc_imp & Trim(Mid$(linea, reg("desde"), reg("longitud")))
               Case "N1"
                    r_blimpo.direc_notf = r_blimpo.direc_notf & Trim$(Mid$(linea, reg("desde"), reg("longitud")))
               Case "N2"
                    r_blimpo.direc_notf2 = r_blimpo.direc_notf2 & Trim$(Mid$(linea, reg("desde"), reg("longitud")))
            End Select
      End Select
      reg.MoveNext
   Loop

End Sub




Public Sub procesa_BS1(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
Dim tipo As String
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "'  And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
   'Set reg = Gcon_docu.OpenRecordset(sql)
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
        Case "EMBARCADOR"
            r_blimpo.embarcador = r_blimpo.embarcador & Trim(Mid$(linea, reg("desde"), reg("longitud")))
      End Select
      reg.MoveNext
   Loop

End Sub


Public Sub procesa_BS3(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "'  And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "DIREC_EXP"
            r_blimpo.direc_exp = r_blimpo.direc_exp & Trim(Mid$(linea, reg("desde"), reg("longitud")))
      End Select
      reg.MoveNext
   Loop
End Sub

Public Sub procesa_BC1(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "CONSIGNATARIO"
            r_blimpo.tipo_id = Gstr_Tipo_ID
            r_blimpo.Impexp = Gstr_Impexp
            r_blimpo.nacio_consig = Gstr_param.nacionalidad
            r_blimpo.consignatario = r_blimpo.consignatario & Trim$(Mid$(linea, reg("desde"), reg("longitud")))
      End Select
      reg.MoveNext
   Loop
End Sub

Public Sub procesa_BC3(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
   'Set reg = Gcon_docu.OpenRecordset(sql)
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "DIREC_IMP"
            r_blimpo.direc_imp = r_blimpo.direc_imp & Trim$(Mid$(linea, reg("desde"), reg("longitud")))
      End Select
      reg.MoveNext
   Loop
End Sub

Public Sub procesa_BN1(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
   'Set reg = Gcon_docu.OpenRecordset(sql)
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "NO_NOTYFIER"
            r_blimpo.no_notyfier = r_blimpo.no_notyfier & Trim$(Mid$(linea, reg("desde"), reg("longitud")))
      End Select
      reg.MoveNext
   Loop
End Sub

Public Sub procesa_BN3(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
   'Set reg = Gcon_docu.OpenRecordset(sql)
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "DIREC_NOTF"
            r_blimpo.tipo_id_notif = Gstr_Tipo_ID
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
Dim MONEDA_EQ   As String
Dim tipo_cambio As String

On Error GoTo error_graba
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'CONVERSION'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "MONEDA"
            MONEDA = Mid$(linea, reg("desde"), reg("longitud"))
         Case "TIPO_CAMBIO"
            tipo_cambio = CDbl(Aplica_Formato(cod_linea, reg("CAMPO"), Mid$(linea, reg("desde"), reg("longitud"))))
         Case "MONEDA_EQ" ' ES LA MONEDA A LA CUAL SE CONVIERTE EL VALOR
            MONEDA_EQ = Mid$(linea, reg("desde"), reg("longitud"))
      End Select
      reg.MoveNext
   Loop
   reg.Close
   sql = "Select * from DTX_conversion  where moneda = '" & MONEDA & "'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   If reg.EOF Then
      sql = "Insert into DTX_conversion Values ('" & MONEDA & "'," & tipo_cambio & ", '" & MONEDA_EQ & "')"
      Gcon_docu.Execute sql
   End If
Exit Sub
error_graba:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0
End Sub

Public Sub procesa_CHB(cod_linea As String, linea As String)
Dim sql    As String
Dim sql1   As String
Dim reg    As New ADODB.Recordset
Dim reg1    As New ADODB.Recordset
Dim MONEDA As String
Dim SIGNO  As String
Dim Costo_Flete As Double

On Error GoTo Errores

   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO' ORDER BY CAMPO DESC"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "SIGNO"
            SIGNO = Mid$(linea, reg("desde"), reg("longitud"))
         Case "MONEDA"
            MONEDA = Mid$(linea, reg("desde"), reg("longitud"))
            r_blimpo.MONEDA = Gstr_moneda
         Case "COSTO_FLETE"
            sql1 = "select * from DTX_conversion where moneda = '" & MONEDA & "'"
            'Set reg1 = Gcon_docu.OpenRecordset(sql1)
            reg1.Open sql1, Gcon_docu, adOpenDynamic, adLockOptimistic
            If reg.EOF Then
               MsgBox "No se encontro TIPO/CAMBIO para la moneda: " & MONEDA & " su valor no se tomara en cuenta en el costo del flete", vbInformation, App.Title
            Else
               If SIGNO = "+" And (MONEDA = "USD" Or MONEDA = "DEM") Then
                  Costo_Flete = CDbl(Aplica_Formato(cod_linea, reg("Campo"), Mid$(linea, reg("desde"), reg("longitud"))))
                  If MONEDA = "USD" Then
                     r_blimpo.Costo_Flete = r_blimpo.Costo_Flete + (Costo_Flete)
                  ElseIf MONEDA = "DEM" Then ' EL VALOR DE FLETE EN MONEDA "DEM" SE DEBE MULTIPLICAR POR UN FACTOR DADO POR LA MISMA LINEA (1.5 HASTA OCT/2000)
                     r_blimpo.Costo_Flete = r_blimpo.Costo_Flete + (Costo_Flete * reg1.Fields("valor")) 'Gdbl_T_Cambio_DEM)
                  Else ' ESTA ES LA INSTRUCCION ORIGINAL Y LA QUE ES POR LOGICA LA CORRECTA
                     r_blimpo.Costo_Flete = r_blimpo.Costo_Flete + (Costo_Flete * reg1.Fields("valor"))
                  End If
               End If
               r_blimpo.metodo_pago = Gstr_param.metodo_pago
            End If
      End Select
      reg.MoveNext
   Loop
   '----------------  INICIA EL PROCESO DE CARGAR LA INFORMACION REAL DE FLETES TAL COMO BIENE EN EL DTX  --------
   reg.Close
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' ORDER BY CAMPO DESC"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   
   i = UBound(Gtyp_flete_DTX) - 1 ' SETEA EL INDICE ACTUAL DEL ARREGLO
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "SIGNO"
            Gtyp_flete_DTX(i).SIGNO = Mid$(linea, reg("desde"), reg("longitud"))
         Case "MONEDA"
            Gtyp_flete_DTX(i).MONEDA = Mid$(linea, reg("desde"), reg("longitud"))
            If Trim(Gtyp_flete_DTX(i).MONEDA) <> "" Then Gtyp_flete_DTX(i).MONEDA = Codigo_equivalente_DTX("DTX_moneda", Gtyp_flete_DTX(i).MONEDA, "codigo", "codigo")
         Case "COSTO_FLETE"
            Gtyp_flete_DTX(i).valor = CDbl(Aplica_Formato(cod_linea, reg("Campo"), Mid$(linea, reg("desde"), reg("longitud"))))
         Case "CHARGE"
            Gtyp_flete_DTX(i).cargo = Mid$(linea, reg("desde"), reg("longitud"))
            If Trim(Gtyp_flete_DTX(i).cargo) <> "" Then Gtyp_flete_DTX(i).cargo = Codigo_equivalente_DTX("Docu_cargo", Gtyp_flete_DTX(i).cargo, "cargo_codigo", "cargo_codigo")
         Case "DEBTOR"
            Gtyp_flete_DTX(i).debtor = Mid$(linea, reg("desde"), reg("longitud"))
            If Trim(Gtyp_flete_DTX(i).debtor) <> "" Then Gtyp_flete_DTX(i).debtor = Codigo_equivalente_DTX("Docu_debtor", Gtyp_flete_DTX(i).debtor, "debtor_codigo", "debtor_codigo")
         Case "REMARKS"
            Gtyp_flete_DTX(i).marcas = Mid$(linea, reg("desde"), reg("longitud"))
         Case "LOCATION"
            Gtyp_flete_DTX(i).lugar = Mid$(linea, reg("desde"), reg("longitud"))
         Case "APPEARS"
            Gtyp_flete_DTX(i).imp_bl = IIf(Mid$(linea, reg("desde"), reg("longitud")) = "Y", True, False)
      End Select
      reg.MoveNext
   Loop
   ReDim Preserve Gtyp_flete_DTX(UBound(Gtyp_flete_DTX) + 1)
Exit Sub

Errores:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0
End Sub

Public Sub procesa_CHG(cod_linea As String, linea As String)
Dim sql    As String
Dim sql1   As String
Dim reg    As New ADODB.Recordset
Dim reg1    As New ADODB.Recordset
Dim MONEDA As String
Dim SIGNO  As String
Dim Costo_Flete As Double

On Error GoTo Errores

   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO' ORDER BY CAMPO DESC"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "SIGNO"
            SIGNO = Mid$(linea, reg("desde"), reg("longitud"))
         Case "MONEDA"
            MONEDA = Mid$(linea, reg("desde"), reg("longitud"))
            r_blimpo.MONEDA = Gstr_moneda
         Case "COSTO_FLETE"
            sql1 = "select * from DTX_conversion where moneda = '" & MONEDA & "'"
            'Set reg1 = Gcon_docu.OpenRecordset(sql1)
            reg1.Open sql1, Gcon_docu, adOpenDynamic, adLockOptimistic
            If reg.EOF Then
               MsgBox "No se encontro TIPO/CAMBIO para la moneda: " & MONEDA & " su valor no se tomara en cuenta en el costo del flete", vbInformation, App.Title
            Else
               If SIGNO = "+" And (MONEDA = "USD" Or MONEDA = "DEM") Then
                  Costo_Flete = CDbl(Aplica_Formato(cod_linea, reg("Campo"), Mid$(linea, reg("desde"), reg("longitud"))))
                  If MONEDA = "USD" Then
                     r_blimpo.Costo_Flete = r_blimpo.Costo_Flete + (Costo_Flete)
                  ElseIf MONEDA = "DEM" Then ' EL VALOR DE FLETE EN MONEDA "DEM" SE DEBE MULTIPLICAR POR UN FACTOR DADO POR LA MISMA LINEA (1.5 HASTA OCT/2000)
                     r_blimpo.Costo_Flete = r_blimpo.Costo_Flete + (Costo_Flete * reg1.Fields("valor")) 'Gdbl_T_Cambio_DEM)
                  Else ' ESTA ES LA INSTRUCCION ORIGINAL Y LA QUE ES POR LOGICA LA CORRECTA
                     r_blimpo.Costo_Flete = r_blimpo.Costo_Flete + (Costo_Flete * reg1.Fields("valor"))
                  End If
               End If
               r_blimpo.metodo_pago = Gstr_param.metodo_pago
            End If
      End Select
      reg.MoveNext
   Loop
   '----------------  INICIA EL PROCESO DE CARGAR LA INFORMACION REAL DE FLETES TAL COMO BIENE EN EL DTX  --------
   reg.Close
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' ORDER BY CAMPO DESC"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   
   i = UBound(Gtyp_flete_DTX) - 1 ' SETEA EL INDICE ACTUAL DEL ARREGLO
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "SIGNO"
            Gtyp_flete_DTX(i).SIGNO = Mid$(linea, reg("desde"), reg("longitud"))
         Case "MONEDA"
            Gtyp_flete_DTX(i).MONEDA = Mid$(linea, reg("desde"), reg("longitud"))
            If Trim(Gtyp_flete_DTX(i).MONEDA) <> "" Then Gtyp_flete_DTX(i).MONEDA = Codigo_equivalente_DTX("DTX_moneda", Gtyp_flete_DTX(i).MONEDA, "codigo", "codigo")
         Case "COSTO_FLETE"
            Gtyp_flete_DTX(i).valor = CDbl(Aplica_Formato(cod_linea, reg("Campo"), Mid$(linea, reg("desde"), reg("longitud"))))
         Case "CHARGE"
            Gtyp_flete_DTX(i).cargo = Mid$(linea, reg("desde"), reg("longitud"))
            If Trim(Gtyp_flete_DTX(i).cargo) <> "" Then Gtyp_flete_DTX(i).cargo = Codigo_equivalente_DTX("Docu_cargo", Gtyp_flete_DTX(i).cargo, "cargo_codigo", "cargo_codigo")
         Case "DEBTOR"
            Gtyp_flete_DTX(i).debtor = Mid$(linea, reg("desde"), reg("longitud"))
            If Trim(Gtyp_flete_DTX(i).debtor) <> "" Then Gtyp_flete_DTX(i).debtor = Codigo_equivalente_DTX("Docu_debtor", Gtyp_flete_DTX(i).debtor, "debtor_codigo", "debtor_codigo")
         Case "REMARKS"
            Gtyp_flete_DTX(i).marcas = Mid$(linea, reg("desde"), reg("longitud"))
         Case "LOCATION"
            Gtyp_flete_DTX(i).lugar = Mid$(linea, reg("desde"), reg("longitud"))
         Case "APPEARS"
            Gtyp_flete_DTX(i).imp_bl = IIf(Mid$(linea, reg("desde"), reg("longitud")) = "Y", True, False)
      End Select
      reg.MoveNext
   Loop
   ReDim Preserve Gtyp_flete_DTX(UBound(Gtyp_flete_DTX) + 1)
Exit Sub

Errores:
MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
Exit Sub
Resume 0

End Sub

Public Sub procesa_DE1(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset

   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPODT'"
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
Public Sub procesa_DE2(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset

   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPODT'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "MARCA_NUM"
            r_blimpodt.marca_num = r_blimpodt.marca_num & Trim$(Mid$(linea, reg("desde"), reg("longitud")))
         Case "DESCRIP_CARGA"
            r_blimpodt.desc_carga = r_blimpodt.desc_carga & Trim$(Mid$(linea, reg("desde"), reg("longitud")))
      End Select
      reg.MoveNext
   Loop
End Sub
Public Sub procesa_DE3(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset

   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "'  And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPODT'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "DANGER"
            r_blimpodt.danger = Codigo_equivalente_DTX("DTX_IMO", Mid$(linea, reg("desde"), reg("longitud")), "APG", "CODIGO", "")
      End Select
      reg.MoveNext
   Loop
End Sub

Public Sub procesa_CTR(cod_linea As String, linea As String)
Dim sql            As String
Dim reg            As New ADODB.Recordset
Dim TIPO_CONTEINER As String
Dim SIZE_CONTEINER As String
Dim FULL_EMPTY     As String
   
   r_impocont.buque = r_registro.buque
   r_impocont.viaje = r_registro.viaje
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'IMPOCONT'"
   'Set reg = Gcon_docu.OpenRecordset(sql)
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "CONTEINER"
            r_impocont.CONTEINER = Mid$(linea, reg("desde"), reg("longitud"))
         Case "CONTEINER1" ' SOLO VERSION 55
            r_impocont.CONTEINER = Mid$(linea, reg("desde"), reg("longitud"))
         Case "CONTEINER2" ' SOLO VERSION 55
            r_impocont.CONTEINER = Trim$(r_impocont.CONTEINER) & Trim(Mid$(linea, reg("desde"), reg("longitud")))
         Case "TYPE_CONTEINER"
            TIPO_CONTEINER = Mid$(linea, reg("desde"), reg("longitud"))
         Case "SIZE_CONTEINER"
            SIZE_CONTEINER = Mid$(linea, reg("desde"), reg("longitud"))
            Gint_size_contein = CInt(SIZE_CONTEINER) ' ESTE DATO ES PARA PODER EVALUAR DESPUES EL VOLUMEN POR DEFECTO DE UN CONTENEDOR SEGUN SU TAMAÑO(SI NO EXISTE VOLUMEN)
         Case "TARA"
            r_impocont.tara = IIf(IsNumeric(Mid$(linea, reg("desde"), reg("longitud"))), Mid$(linea, reg("desde"), reg("longitud")), 0)
         Case "FULL/EMPTY"
            FULL_EMPTY = Mid$(linea, reg("desde"), reg("longitud"))
'         Case "SELLO1"
'            r_impocont.sello1 = Mid$(linea, reg("desde"), reg("longitud"))
      End Select
      reg.MoveNext
   Loop
         
   'cra
   r_impocont.tamano_equipo = Codigo_equivalente_DTX("DTX_tipo_contein", TIPO_CONTEINER, "APG", "TYPE", SIZE_CONTEINER, " SIZE ")
   If r_impocont.tamano_equipo = "NO" Then
      MsgBox "ERROR LA EQUIVALENCIA DE CODIGO DE TIPO DE CONTENEDOR " & SIZE_CONTEINER & "  " & TIPO_CONTEINER
      Exit Sub
   End If
   If UCase(FULL_EMPTY) = "E" Then ' SE EVALUA SI EL CONTENEDOR VIENE LLENO O VACIO
      r_blimpodt.tipo_carga = Gstr_TIPO_CARGA_CONT_VACIO '"8"
      r_impocont.ind_lleno_vacio = "4" 'PONER EN PARAMETRO
      If Trim(Gstr_Cond_Contein_vacio) <> "" Then r_impocont.cond_contein = Gstr_Cond_Contein_vacio 'nov 2007
      If Trim(r_impocont.sello1) = "" Then r_impocont.sello1 = Gstr_sello_Contein
      r_blimpodt.peso = r_impocont.tara    'nov 2007 se incluye el mimsmo valor que la tara al peso para los vacios
      r_blimpodt.producto = Gstr_TIPO_PRODUCTO_VACIO '23 nov 2007 se incluye valor para el producto de contenddores vacios
   Else
      r_blimpodt.tipo_carga = Gstr_TIPO_CARGA_CONT '"5"
      r_impocont.ind_lleno_vacio = "5" 'PONER EN PARAMETRO
      r_blimpodt.producto = Gstr_TIPO_PRODUCTO '23 nov 2007 se incluye valor para el producto de contenddores o por defecto
   End If
   r_blimpodt.CONTEINER = r_impocont.CONTEINER
   'ESTA VALIDACION SE PUSO SOLO POR SI ACASO PARA VERSIONES INFERIORES A LA 53
   If Gstr_Version_MSC = "52" Then r_impocont.sello1 = Mid$(linea, 23, 15)
   
   
End Sub


Public Sub procesa_RFR(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
   
   If Mid(linea, 4, 1) = "N" Then Exit Sub ' si es
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'IMPOCONT'"
   'Set reg = Gcon_docu.OpenRecordset(sql)
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "MIN_TEMP_SIG"
                SIGNO = Mid$(linea, reg("desde"), reg("longitud"))
         Case "MIN_TEMP_SET"
                r_impocont.temp_min_trans = CInt(SIGNO & "1") * CInt(Mid$(linea, reg("desde"), reg("longitud")))
         Case "MAX_TEMP_SIG"
                SIGNO = Mid$(linea, reg("desde"), reg("longitud"))
         Case "MAX_TEMP_SET"
                r_impocont.temp_max_trans = CInt(SIGNO & "1") * CInt(Mid$(linea, reg("desde"), reg("longitud")))
      End Select
      reg.MoveNext
   Loop
   
End Sub


Public Sub procesa_SEL(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
   
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'IMPOCONT'"
   'Set reg = Gcon_docu.OpenRecordset(sql)
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "SELLO1"
            If r_impocont.sello1 = "" Then
                r_impocont.sello1 = Mid$(linea, reg("desde"), reg("longitud"))
            ElseIf r_impocont.sello2 = "" Then
                r_impocont.sello2 = Mid$(linea, reg("desde"), reg("longitud"))
            ElseIf r_impocont.sello3 = "" Then
                r_impocont.sello3 = Mid$(linea, reg("desde"), reg("longitud"))
            ElseIf r_impocont.sello4 = "" Then
                r_impocont.sello4 = Mid$(linea, reg("desde"), reg("longitud"))
            End If
      End Select
      reg.MoveNext
   Loop
   
End Sub
Public Sub procesa_CT1(cod_linea As String, linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset
   sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPODT'"
   
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   Do While Not reg.EOF
      Select Case reg("CAMPO_TABLA")
         Case "NO_BULTOS"
            r_blimpodt.no_bultos = Aplica_Formato(cod_linea, reg("CAMPO"), Mid$(linea, reg("desde"), reg("longitud")))
         'Case "EMBALAJE"
         '   r_blimpodt.embalaje = Codigo_equivalente_DTX("EMBALAJE", Mid$(linea, reg("desde"), reg("longitud")), "APG", "CODIGO", "")
         Case "PESO"
            r_blimpodt.peso = Aplica_Formato(cod_linea, reg("CAMPO"), Mid$(linea, reg("desde"), reg("longitud")))
         Case "VOLUMEN"
            r_blimpodt.volumen = Aplica_Formato(cod_linea, reg("CAMPO"), Mid$(linea, reg("desde"), reg("longitud")))
      End Select
      reg.MoveNext
   Loop
   r_blimpodt.embalaje = Gstr_EMBAL_BULTOS
End Sub



'Funcion        : Codigo Equivalente
'Objetivo       : Devolver el código equivalente en APG en base al código
'                 de los archivos DTX
'Fecha Creacion : Sábado 15 de Abril del 2000 - 15:00:00
'Creado por     : Carlos Reyes Atiencia
'Parámetros     : Tabla .- es el nombre de la tabla intermedia que tiene ambos codigos
'                 codigo_buscado .- es el valor en si que se esta buscando
'                 campo_apg .- es el nombre del campo que contiene el código APG
'===============================================================================

Public Function Codigo_equivalente_DTX(tabla As String, ByVal codigo_buscado As String, campo_apg As String, campo_msc As String, Optional codigo_buscado2 As String, Optional campo_msc2 As String, Optional condicion As String) As String

On Error GoTo error_busca_cod_apg
   Dim sql As String
   Dim REGISTRO   As New ADODB.Recordset
   Dim reg        As New ADODB.Recordset
   Dim Response   As Integer
   Dim ENCONTRADO As Boolean
   Dim ST_NUEVO   As Boolean
   
If DTX_PROCESO.Cmb_linea.BoundText = "CAE" Then
BUSCAR:
   sql = "select " & Trim$(campo_apg) & " from " & Trim$(tabla) & _
         " where " & Trim$(campo_msc) & " = '" & Trim$(codigo_buscado) & "'"
   
   If Not IsMissing(codigo_buscado2) And Trim$(codigo_buscado2) <> "" And _
      Not IsMissing(campo_msc2) And Trim$(campo_msc2) <> "" Then
      sql = sql & " and " & Trim$(campo_msc2) & " = '" & Trim$(codigo_buscado2) & "'"
   End If
   
   If Not IsMissing(condicion) And Trim$(condicion) <> "" Then
      sql = sql + " and " & condicion
   End If
   Select Case UCase(tabla)
        Case "DOCU_DEBTOR":  sql = sql & " and debtor_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "'"
        Case "DOCU_CARGO":   sql = sql & " and cargo_linea  ='" & DTX_PROCESO.Cmb_linea.BoundText & "'"
        Case Else:   sql = sql & " and LINEA ='" & Gstr_Linea & "'"
   End Select
   
   REGISTRO.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   
   ENCONTRADO = True:   ST_NUEVO = True
   If REGISTRO.EOF Then
      ENCONTRADO = False
      ST_NUEVO = True
   Else
      If Trim(REGISTRO.Fields(campo_apg)) = "" Or IsNull(REGISTRO.Fields(campo_apg)) Then
         ENCONTRADO = False
         ST_NUEVO = False
      End If
   End If
   
   If Not ENCONTRADO Then
      If UCase(tabla) = "DOCU_CARGO" Then
          Codigo_equivalente_DTX = Gstr_Cargo  'se las pso aqui fijas para que no pregunte nada
          Exit Function
      End If
      If UCase(tabla) = "DOCU_DEBTOR" Then
          Codigo_equivalente_DTX = Gstr_Debtor 'se las pso aqui fijas para que no pregunte nada
          Exit Function
      End If
      If tabla <> "DTX_IMO" Then
         If Not Gbol_relacion_fija Then ' si la relacion entre los codigos de CAE y MSC es fija no pregunta para establecer una nueva relación.
            Response = MsgBox("El código Equivalente " & UCase(tabla) & "." & UCase(campo_msc) & " no tiene un codigo relacionado en CAE, si desea asignar un codigo default presione SI, caso contrario debera establecer la relacion!", vbQuestion + vbYesNo)
         Else
            Response = vbYes
         End If
      Else
         Response = vbYes
      End If
      If Response = vbYes Then
         Select Case UCase(tabla)
            Case "DTX_EMBALAJE":     Codigo_equivalente_DTX = Gstr_Embalaje
            Case "DTX_PUERTO":       Codigo_equivalente_DTX = Gstr_Puerto
            Case "DTX_IMO":          Codigo_equivalente_DTX = Gstr_IMO
            Case "DTX_COND_CONTEIN": Codigo_equivalente_DTX = Gstr_Cond_Contein
            Case "DTX_MONEDA":       Codigo_equivalente_DTX = Gstr_moneda
            Case "DTX_TIPO_CONTEIN": Codigo_equivalente_DTX = Gstr_Tipo_Contein
'            Case "DOCU_CARGO":       Codigo_equivalente_DTX = Gstr_Cargo
'            Case "DOCU_DEBTOR":      Codigo_equivalente_DTX = Gstr_Debtor
         End Select
      ElseIf Response = vbNo Then
         Tipo_rela_APG = tabla
         Load DTX_RELA_APG
         DTX_RELA_APG.Opt_nuevo = ST_NUEVO
         Select Case UCase(tabla) 'forma de mantenimiento
            Case "DTX_EMBALAJE":       DTX_RELA_APG.txt_emba_linea = Gstr_Linea:      DTX_RELA_APG.txt_emba_codi = codigo_buscado:
            Case "DTX_PUERTO":         DTX_RELA_APG.Txt_puer_linea = Gstr_Linea:      DTX_RELA_APG.Txt_puer_codi = codigo_buscado:
            Case "DTX_IMO":            DTX_RELA_APG.Txt_imo_linea = Gstr_Linea:       DTX_RELA_APG.Txt_imo_codi = codigo_buscado:
            Case "DTX_COND_CONTEIN":   DTX_RELA_APG.txt_cond_cont_linea = Gstr_Linea: DTX_RELA_APG.txt_cond_cont_orig = codigo_buscado: DTX_RELA_APG.txt_cond_cont_dest = codigo_buscado2
            Case "DTX_MONEDA":         DTX_RELA_APG.Txt_mone_linea = Gstr_Linea:      DTX_RELA_APG.txt_mone_codi = codigo_buscado:
            Case "DTX_TIPO_CONTEIN":   DTX_RELA_APG.txt_tipo_cont_linea = Gstr_Linea: DTX_RELA_APG.txt_tipo_cont_size = codigo_buscado: DTX_RELA_APG.txt_tipo_cont_type = codigo_buscado2
         End Select
         DTX_RELA_APG.Show 1
         If ST_NUEVO Then 'SE DEBE AGREGAR UN NUEVO REGISTRO A LA TABLA DE CODIGO EQUIVALENTE
         
         Else             'SE DEBE MODIFICAR EL REGISTRO DE LA TABLA DE CODIGO EQUIVALENTE
         
         End If
         REGISTRO.Close
         GoTo BUSCAR
      Else
         Codigo_equivalente_DTX = "NO"
      End If
   Else
      Codigo_equivalente_DTX = Trim(REGISTRO.Fields(campo_apg))
   End If
Else ' si se trata de un dtx que sube a los datos de la misma linea no de la CAE, se usan los mismos codigos
    Select Case UCase(tabla)
        Case "DTX_EMBALAJE":     Codigo_equivalente_DTX = Trim(codigo_buscado) 'Gstr_Embalaje
        Case "DTX_PUERTO":
                 If DTX_PROCESO.Cmb_linea.BoundText = "MSC" Then
                    Codigo_equivalente_DTX = Trim(codigo_buscado) ' Gstr_Puerto
                 Else
                    Codigo_equivalente_DTX = devuelve_desc("docu_puerto", "puer_codigo_alt", "puer_codigo", Trim(codigo_buscado), " puer_linea ='KHS' and puer_status='A'")
                 End If
        Case "DTX_IMO":          Codigo_equivalente_DTX = Trim(codigo_buscado) 'Gstr_IMO
        Case "DTX_COND_CONTEIN": Codigo_equivalente_DTX = codigo_buscado & "-" & codigo_buscado2 'Gstr_Cond_Contein
        Case "DTX_MONEDA":       Codigo_equivalente_DTX = Trim(codigo_buscado) 'Gstr_moneda
        Case "DTX_TIPO_CONTEIN": Codigo_equivalente_DTX = codigo_buscado & "-" & codigo_buscado2  'Gstr_Tipo_Contein
        Case "DOCU_DEBTOR":      Codigo_equivalente_DTX = Trim(codigo_buscado) 'Gstr_debtor
        Case "DOCU_CARGO":       Codigo_equivalente_DTX = Trim(codigo_buscado) 'Gstr_CARGO
    End Select
    'Codigo_equivalente_DTX = Trim(codigo_buscado)
End If
   
   
   ' se busca en la tabla maestra para verificar si existe en la tabla de maestra de su correspondiente
   Select Case UCase(tabla)
            Case "DTX_EMBALAJE":
                  sql = "Select * from docu_embalaje where emba_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "' and emba_codigo ='" & Codigo_equivalente_DTX & "' and emba_status = 'A'"
                  reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                  If reg.EOF Then Codigo_equivalente_DTX = Gstr_Embalaje
            Case "DTX_PUERTO":
                  sql = "Select * from docu_puerto where puer_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "' and puer_codigo ='" & Codigo_equivalente_DTX & "'"
                  reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                  If reg.EOF Then
                        MsgBox "El código de puerto " & codigo_buscado & " NO existe o esta mal relacionado, por favor verifique la relación, por el momento se pondra el codigo por defecto " & Gstr_Puerto, vbCritical
                        Codigo_equivalente_DTX = Gstr_Puerto
                  End If
            Case "DTX_IMO":
                  sql = "Select * from docu_imo where imo_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "' and imo_codigo ='" & Codigo_equivalente_DTX & "'"
                  reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                  If reg.EOF Then Codigo_equivalente_DTX = Gstr_IMO
            Case "DTX_COND_CONTEIN":
                  sql = "Select * from docu_cond_contenedor where cond_cont_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "' and cond_cont_codigo ='" & Codigo_equivalente_DTX & "' and cond_cont_status ='A'"
                  reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                  If reg.EOF Then Codigo_equivalente_DTX = Gstr_Cond_Contein
            Case "DTX_MONEDA":
                  sql = "Select * from docu_moneda where mone_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "' and mone_codigo ='" & Codigo_equivalente_DTX & "'"
                  reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                  If reg.EOF Then Codigo_equivalente_DTX = Gstr_moneda
            Case "DTX_TIPO_CONTEIN":
                  sql = "Select * from docu_tamano_equip where tama_equip_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "' and tama_equip_codigo ='" & Codigo_equivalente_DTX & "' and tama_equip_status ='A'"
                  reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                  If reg.EOF Then
                    If codigo_buscado = "20" Or codigo_buscado = "0" Or codigo_buscado = "1" Or codigo_buscado = "2" Then
                       Codigo_equivalente_DTX = Gstr_Tipo_Contein    ' de 20
                    Else
                       Codigo_equivalente_DTX = Gstr_Tipo_Contein_40 ' de 40
                    End If
                  End If
            Case "DOCU_DEBTOR":
                  sql = "Select * from docu_debtor where debtor_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "' and debtor_codigo ='" & Codigo_equivalente_DTX & "'"
                  reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                  If reg.EOF Then Codigo_equivalente_DTX = Gstr_Debtor
            Case "DOCU_CARGO":
                  sql = "Select * from docu_cargo where cargo_linea ='" & DTX_PROCESO.Cmb_linea.BoundText & "' and cargo_codigo ='" & Codigo_equivalente_DTX & "'"
                  reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                  If reg.EOF Then Codigo_equivalente_DTX = Gstr_Cargo
   End Select
   reg.Close
   
   Exit Function
   
error_busca_cod_apg:
   MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
   Exit Function
   Resume 0
End Function

Public Function Aplica_Formato(ByVal cod_linea As String, ByVal campo As String, ByVal valor As String) As Variant
Dim sql As String
Dim reg As New ADODB.Recordset
Dim val_numerico As Double
Dim fecha As String
Dim mes As String

On Error GoTo Errores
sql = "Select * from DTX_Formato_Detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' and codigo = '" & cod_linea & "' and campo ='" & campo & "'"
reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
If reg.Fields("entero") <> 0 And (reg.Fields("fecha") = "" Or IsNull(reg.Fields("fecha"))) Then
   If IsNumeric(valor) Then
      Aplica_Formato = Mid$(valor, 1, reg.Fields("entero")) & "." & Mid$(valor, (reg.Fields("entero") + 1), reg.Fields("decimal"))
   Else
      Aplica_Formato = 0
   End If
ElseIf reg.Fields("entero") = 0 And reg.Fields("decimal") = 0 And reg.Fields("fecha") <> "" And Not IsNull(reg.Fields("fecha")) Then
       ini1 = 0: ini2 = 0: ini3 = 0
       lon1 = 0: lon2 = 0: lon3 = 0
       For i% = 1 To Len(reg.Fields("fecha"))
           If UCase(Mid$(reg.Fields("fecha"), i%, 1)) = "D" Then
              If ini1 = 0 Then ini1 = i%
              lon1 = lon1 + 1
           ElseIf UCase(Mid$(reg.Fields("fecha"), i%, 1)) = "M" Then
                  If ini2 = 0 Then ini2 = i%
                  lon2 = lon2 + 1
               ElseIf UCase(Mid$(reg.Fields("fecha"), i%, 1)) = "Y" Or _
                      UCase(Mid$(reg.Fields("fecha"), i%, 1)) = "A" Then
                      If ini3 = 0 Then ini3 = i%
                      lon3 = lon3 + 1
           End If
       Next i%
       
       mes = Devuelve_mes(Int(Mid$(valor, ini2, lon2))) 'OBTIENE EL NOMBRE DEL MES
       fecha = Mid$(valor, ini1, lon1) & "/" & mes & "/" & Mid$(valor, ini3, lon3)
       'DEPENDIENDO DEL FORMATO DE LA FECHA SE SETEA EL VALOR DE LA FECHA
       Select Case UCase(GSTR_FORMATO_FECHA)
            Case "DD/MM/YYYY", "DD/MM/AAAA": Aplica_Formato = Mid$(valor, ini1, lon1) & "/" & mes & "/" & Mid$(valor, ini3, lon3)
            Case "DD-MM-YYYY", "DD-MM-AAAA": Aplica_Formato = Mid$(valor, ini1, lon1) & "-" & mes & "-" & Mid$(valor, ini3, lon3)
            Case "MM/DD/YYYY", "MM/DD/AAAA": Aplica_Formato = mes & "/" & Mid$(valor, ini1, lon1) & "/" & Mid$(valor, ini3, lon3)
            Case "MM-DD-YYYY", "MM-DD-AAAA": Aplica_Formato = mes & "-" & Mid$(valor, ini1, lon1) & "-" & Mid$(valor, ini3, lon3)
            Case "YYYY/MM/DD", "AAAA/MM/DD": Aplica_Formato = Mid$(valor, ini3, lon3) & "/" & mes & "/" & Mid$(valor, ini1, lon1)
            Case "YYYY-MM-DD", "AAAA-MM-DD": Aplica_Formato = Mid$(valor, ini3, lon3) & "-" & mes & "-" & Mid$(valor, ini1, lon1)
            Case Else: Aplica_Formato = Format$(fecha, GSTR_FORMATO_FECHA)
       End Select
End If

Exit Function
   
Errores:
   MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
   Exit Function
   Resume 0

End Function

' ESTA FUNCION PERMITE REALIZAR LA OBTENCION DE LOS MES EN TEXTO EVALUANDO SI ESTA LA
' CONFIGURACION EN INGLES O ESTA EN ESPANOL
Public Function Devuelve_mes(no_mes As Integer) As String
Dim mes As String

    If UCase(Format("31/12/2003", "mmm")) = "DIC" Then
         Select Case no_mes
              Case 1:  mes = "Ene"
              Case 2:  mes = "Feb"
              Case 3:  mes = "Mar"
              Case 4:  mes = "Abr"
              Case 5:  mes = "May"
              Case 6:  mes = "Jun"
              Case 7:  mes = "Jul"
              Case 8:  mes = "Ago"
              Case 9:  mes = "Sep"
              Case 10: mes = "Oct"
              Case 11: mes = "Nov"
              Case 12: mes = "Dic"
         End Select
    Else
         Select Case no_mes
              Case 1:  mes = "Jan"
              Case 2:  mes = "Feb"
              Case 3:  mes = "Mar"
              Case 4:  mes = "Apr"
              Case 5:  mes = "May"
              Case 6:  mes = "Jun"
              Case 7:  mes = "Jul"
              Case 8:  mes = "Aug"
              Case 9:  mes = "Sep"
              Case 10: mes = "Oct"
              Case 11: mes = "Nov"
              Case 12: mes = "Dec"
         End Select
    End If
    Devuelve_mes = mes
End Function



' ESTA FUNCION REALIZA LA INVERSA DE LA FUNCION APLICA_FORMATO, ESTA LE DA EL FORMATO DE LA LINEA A UN VALOR
' EN FORMATO NORMAL USADO EN EL SISTEMA,
' ESTA FUNCION SE USA PARA GENERAR LOS ARCHIVOS DTX, PRINCIPALMENTE DE LA LINEA KHS, LA OTRA ES PARA LEERLOS.
' ENERO 2 DE 2003
Public Function Devuelve_Formato(ByVal cod_linea As String, ByVal campo As String, ByVal valor As String) As Variant
Dim sql As String
Dim reg As New ADODB.Recordset
Dim val_numerico As Double
Dim fecha As String

On Error GoTo Errores
sql = "Select * from DTX_Formato_Detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' and codigo = '" & cod_linea & "' and campo ='" & campo & "'"
reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
If reg.Fields("entero") <> 0 And (reg.Fields("fecha") = "" Or IsNull(reg.Fields("fecha"))) Then
   If IsNumeric(valor) Then
      valor_ENTERO = Fix(CDbl(valor))
      valor_DECIMAL = CDbl(valor) - Fix(CDbl(valor))
      valor_DECIMAL = valor_DECIMAL * 10 ^ (Len(Mid(Str(valor_DECIMAL), 3))) ' VALOR DECIMAL SE LO MULTIPLICA POR 10 A LA N (QUE ES LA CANTIDAD DE CARACTERES QUE SE DEBE CORRER LA COMA)
      If valor_DECIMAL = "0" Then valor_DECIMAL = ""
      
      Devuelve_Formato = String(reg.Fields("entero") - Len(Trim(valor_ENTERO)), "0") & Trim(valor_ENTERO) & Trim(valor_DECIMAL) & String(reg.Fields("decimal") - Len(Trim(valor_DECIMAL)), "0")
      
'      If CDbl(valor) - Fix(CDbl(valor)) <> 0 Then ' SI LA CANTIDAD DE DECIMALES ES <> 0
'
'        '                   (CEROS A LA IZQUIERDA DEL VALOR PARA COMPLETAR                      &  PARTE ENTERA     & PARTE DECIMAL                          & CEROS PARA COMPLETAR #DECIMALES - TAMANO DE LA PARTE DECIMAL HECHA TEXTO -1 (PUNTO DECIMAL)
'         Devuelve_Formato = String(reg.Fields("entero") - Len(Trim(Str(Fix(CDbl(valor))))), "0") & Fix(CDbl(valor)) & Mid(CDbl(valor) - Fix(CDbl(valor)), 3) & String(reg.Fields("decimal") - (Len(Trim(Str(CDbl(valor) - Fix(CDbl(valor)))) - 1)), "0")
'      Else 'SI NO TIENE DECIMALES
'         Devuelve_Formato = String(reg.Fields("entero") - Len(Trim(Str(Fix(CDbl(valor))))), "0") & Fix(CDbl(valor)) & String(reg.Fields("decimal"), "0")
'      End If
   Else
      Devuelve_Formato = String(reg.Fields("entero"), "0") & String(reg.Fields("decimal"), "0")
   End If
ElseIf reg.Fields("entero") = 0 And reg.Fields("decimal") = 0 And reg.Fields("fecha") <> "" And Not IsNull(reg.Fields("fecha")) Then
       Devuelve_Formato = Format(valor, reg.Fields("fecha")) ' se espera que el año sea "yy"
End If
Devuelve_Formato = Trim(Devuelve_Formato)

Exit Function
   
Errores:
   MsgBox "Ocurrio el error " & Err.Number & " - " & Err.Description, vbInformation, App.Title
   Exit Function
   Resume 0
End Function

'esta funcion permite agregar espacios al final de una cadena de caracteres hasta completar la longitud deseada
Public Function Completa_Spacios(texto As String, longitud As Integer) As String
Dim tama As Integer
'    tama = Len(Trim(texto))
'    If tama > longitud Then longitud = tama
    Completa_Spacios = Mid(texto, 1, longitud) & Space(longitud - Len(Mid(texto, 1, longitud)))
End Function


' ***********************************************************************************************
'Function LeeArchivoIni(W_Etiqueta As String, W_Campo As String, W_RutaDelIni As String, ByRef W_VariableDestino As String, Optional W_ShowMsg As Variant) As Boolean
'Dim W_Tamano As Integer
'Dim W_Resultado As Integer
'Dim W_CadenaDeRetorno As String
'    'se asume que W_Campo del archivo ini no tiene datos
'    W_VariableDestino = ""
'    LeeArchivoIni = False
'
'    W_CadenaDeRetorno = Space$(255)
'    W_Tamano = Len(W_CadenaDeRetorno)
'    W_Resultado = GetPrivateProfileString(W_Etiqueta, W_Campo, C_CampoNoExiste, W_CadenaDeRetorno, W_Tamano, W_RutaDelIni)
'    If (Left(W_CadenaDeRetorno, W_Resultado) = C_CampoNoExiste) Then
'        If IsMissing(W_ShowMsg) Then
'            MsgBox "No existe el campo:" & W_Campo & " en la etiqueta: " & W_Etiqueta & " del archivo: " & W_RutaDelIni, vbCritical, App.Title
'        Else
'            If W_ShowMsg Then
'                MsgBox "No existe el campo:" & W_Campo & " en la etiqueta: " & W_Etiqueta & " del archivo: " & W_RutaDelIni, vbCritical, App.Title
'            End If
'        End If
'        Exit Function
'    End If
'    W_VariableDestino = Trim(Left(W_CadenaDeRetorno, W_Resultado))
'    LeeArchivoIni = True
'End Function

'Public Function GetWindowsDir() As String
'Dim strBuf As String
'On Error GoTo MSG_ERROR
'    strBuf = Space$(255)
'
'    If GetWindowsDirectory(strBuf, 254) > 0 Then
'        strBuf = StripTerminator$(strBuf)
'        AddDirSep strBuf
'        GetWindowsDir = strBuf
'    Else
'        GetWindowsDir = App.Path
'    End If
'Exit Function
'MSG_ERROR:
'    MsgBox Err.Description
'End Function

'Public Sub CargaComun()
'Dim W_Temporal As String
'
'    W_OKMonedas = False
'    P_RutaDeMIEIni = GetWindowsDir() & "MIE.INI"
'
'    If Not LeeArchivoIni("Directorios", "BaseClientes", P_RutaDeMIEIni, P_RUTA_MIECLTES) Then End
'    If Not LeeArchivoIni("Instalacion", "TIPO", P_RutaDeMIEIni, P_TipoInst) Then End
'
'End Sub
'-----------------------------------------------------------
' FUNCTION: StripTerminator
'
' Returns a string without any zero terminator.  Typically,
' this was a string returned by a Windows API call.
'
' IN: [strString] - String to remove terminator from
'
' Returns: The value of the string passed in minus any
'          terminating zero.
'-----------------------------------------------------------
''
'Function StripTerminator(ByVal strString As String) As String
'Dim intZeroPos As Integer
'
'    intZeroPos = InStr(strString, Chr$(0))
'    If intZeroPos > 0 Then
'        StripTerminator = Left$(strString, intZeroPos - 1)
'    Else
'        StripTerminator = strString
'    End If
'
'End Function
'-----------------------------------------------------------
' SUB: AddDirSep
' Add a trailing directory path separator (back slash) to the
' end of a pathname unless one already exists
'
' IN/OUT: [strPathName] - path to add separator to
'-----------------------------------------------------------

Sub AddDirSep(strPathName As String)
    
    If Right$(RTrim$(strPathName), 1) <> "\" Then
        strPathName = RTrim$(strPathName) & "\"
    End If

End Sub

'Public Function Calcula_Lineas_Archivo(ARCHIVO As String) As Integer
'Dim x
'Dim i As Integer
'
'x = FreeFile
'i = 0
'Open ARCHIVO For Input As #x
'Do While Not EOF(x)
'    Line Input #x, linea
'    i = i + 1
'Loop
'Calcula_Lineas_Archivo = i
'Close #x
'End Function


' la opcion OPC indica que se duplique el caracter y solo se aplica con la comilla
' simple
Public Function chk_str(wstr As String, Optional opc As String) As String
Dim WSTR2 As String
Dim i As Integer
WSTR2 = ""
For i = 1 To Len(wstr)
   If Mid$(wstr, i, 1) = "'" Or Mid$(wstr, i, 1) = "|" Or _
      Mid$(wstr, i, 1) = """" Or Asc(Mid$(wstr, i, 1)) = 13 Or _
      Asc(Mid$(wstr, i, 1)) = 10 Then
      If opc = "S" And Mid$(wstr, i, 1) = "'" Then
        WSTR2 = WSTR2 & "''"
      Else
        WSTR2 = WSTR2 & ""
      End If
   Else
      WSTR2 = WSTR2 & Mid$(wstr, i, 1)
   End If
Next i
chk_str = WSTR2
End Function

Public Function chk_str2(wstr As String) As String
Dim WSTR2 As String
Dim i As Integer
WSTR2 = ""
For i = 1 To Len(wstr)
   If Mid$(wstr, i, 1) = "'" Then
      WSTR2 = WSTR2 & "''"
   Else
      WSTR2 = WSTR2 & Mid$(wstr, i, 1)
   End If
Next i
chk_str2 = WSTR2
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Esta funcion permite llamar a la forma de ayuda.
' Los parámetros pasados a esta funcion son :
'           Tabla       -> El nombre de la tabla de donde se consultaran los datos
'           Código      -> El nombre del campo código de la tabla
'           Descripcion -> El nombre del campo descripcion de la tabla
'           Criterio    -> El criterio de busqueda inicial para la seleccion de la ayuda
'           Nombre      -> El nombre que aparecerá en la cabecera de la forma de ayuda para saber de que se trata
'           pos_codigo  -> La posicion del campo codigo dentro de la tabla
'           pos_desc    -> La posicion del campo descripcion dentro de la tabla
'
' La funcion retorna un valor string que es el código de la tabla a la que se consulta, si se cancela devuelve Nulo ("")
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Ayuda(ByVal tabla As String, ByVal codigo As String, ByVal DESCRIPCION As String, ByVal criterio As String, ByVal nombre As String, ByVal Serv_MSC As Boolean) As String
'
'    Gtyp_ayuda.tabla = tabla
'    Gtyp_ayuda.campo1 = codigo
'    Gtyp_ayuda.campo2 = DESCRIPCION
'    Gtyp_ayuda.criterio = criterio
'    Gtyp_ayuda.nombre = nombre
'    Gtyp_ayuda.Serv_rol = Serv_MSC
'    DTX_HELP.Show 1
'    Ayuda = Gtyp_ayuda.valor_devuelto
'End Function

Public Sub Confirma_Conversion()
   DTX_CONFIRMA_CONVERSION.Show 1

End Sub

'linea_proc es el codigo de la linea la cual se esta subiendo la informacion puede ser CAE, MSC, KHS
' linea o Gstr_Linea es el codigo de la linea a la que pertenece el archivo puede ser unicamente MSC y KHS
Public Sub carga_parametros_DTX(linea As String, linea_proc As String)
Dim sql As String
Dim reg As New ADODB.Recordset
    
    sql = "Select * from DTX_parametro Where LINEA_PROC ='" & linea_proc & "' and  Linea ='" & linea & "'"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   
    Do While Not reg.EOF
      'Debug.Print reg("codigo")
      Select Case Trim(reg("codigo"))
         Case "USUARIO":        Gstr_user_DTX = Trim$(reg("valor"))
         Case "MONEDA":         Gstr_moneda = Trim$(reg("valor"))
         Case "PROPOSITO":      Gstr_Proposito = Trim(reg("valor"))
         Case "PROPOSITO_EXPO": Gstr_Proposito_expo = Trim(reg("valor"))
         Case "TIPO_ID":        Gstr_Tipo_ID = Trim(reg("valor"))
         Case "IMPEXP":         Gstr_Impexp = Trim(reg("valor"))
         'Case "NAV_CONS_DUI":   Gstr_nav_cons_dui = Trim(reg("valor"))
         Case "CAPITAN":        Gstr_capitan = Trim(reg("valor"))
         'Case "CARGA":          Gstr_carga = Trim(reg("valor"))
         'Case "MED_PESO":       Gstr_Med_Peso = Trim(reg("valor"))
         'Case "MED_VOLUMEN":    Gstr_Med_Volumen = Trim(reg("valor"))
         Case "COND_CONTEIN":   Gstr_Cond_Contein = Trim(reg("valor"))
         Case "LINEA_NAV":      Gstr_Linea_Nav_APG = Trim(reg("valor"))
         
         Case "EMBALAJE":       Gstr_Embalaje = Trim(reg("valor"))
         Case "PUERTO":         Gstr_Puerto = Trim(reg("valor"))
         Case "IMO":            Gstr_IMO = Trim(reg("valor"))
         Case "TIPO_CONTEIN":   Gstr_Tipo_Contein = Trim(reg("valor"))
         Case "TIPO_CONTEIN_40": Gstr_Tipo_Contein_40 = Trim(reg("valor"))
         Case "TIPO_ALMACEN":   Gstr_tipo_alma = Trim(reg("valor"))
         Case "BODEGA":         Gstr_bodega = Trim(reg("valor"))
         Case "VOLU_20":        Gint_volu_20 = CInt(reg("valor"))
         Case "VOLU_40":        Gint_volu_40 = CInt(reg("valor"))
         Case "T_CAMBIO_DEM":   Gdbl_T_Cambio_DEM = CDbl(reg("valor"))
         Case "PESO_VACIO_20":  Gdbl_peso_vacio_20 = CDbl(reg("valor"))
         Case "PESO_VACIO_40":  Gdbl_peso_vacio_40 = CDbl(reg("valor"))
         Case "BL_DETALLADO":   Gbol_bl_detallado_MSC = CBool(Trim$(reg("valor")))
         
         Case "DTX_TIPO_CARGA_CONT_VACIO": Gstr_TIPO_CARGA_CONT_VACIO = Trim(reg("valor"))
         Case "DTX_TIPO_CARGA_CONT":  Gstr_TIPO_CARGA_CONT = Trim(reg("valor"))
         Case "DTX_TIPO_CARGA_GRAL":  Gstr_TIPO_CARGA_GRAL = Trim(reg("valor"))
         Case "DTX_EMBAL_CONTEN":     Gstr_EMBAL_CONTEN = Trim(reg("valor"))
         Case "DTX_EMBAL_BULTOS":     Gstr_EMBAL_BULTOS = Trim(reg("valor"))
         Case "DTX_NAVIERA":          Gstr_naviera = Trim(reg("valor"))
         Case "CARGO":                Gstr_Cargo = Trim(reg("valor"))
         Case "DEBTOR":               Gstr_Debtor = Trim(reg("valor"))
         Case "TRASBORDO":            Gstr_trasbordo = Trim(reg("valor"))
         
         Case "COND_CONTEIN_VACIO":   Gstr_Cond_Contein_vacio = Trim(reg("valor"))
         Case "SELLO_CONTEIN_VACIO":  Gstr_sello_Contein = Trim(reg("valor"))
         Case "TIPO_PRODUCTO":        Gstr_TIPO_PRODUCTO = Trim(reg("valor"))            '23/11/2007
         Case "TIPO_PRODUCTO_VACIO":  Gstr_TIPO_PRODUCTO_VACIO = Trim(reg("valor"))      '23/11/2007
         Case "DESC_CARGA":           Gstr_Desc_carga = Trim(reg("valor"))               'Julio /2008
         Case "MARCAS_NUM":           Gstr_Marca_num = Trim(reg("valor"))                'Julio /2008
      End Select
      reg.MoveNext
    Loop
    reg.Close
    If Trim(Gstr_Linea_Nav_APG) = "" Then Gstr_Linea_Nav_APG = "9993"
    
    If Trim(Gstr_Cond_Contein_vacio) = "" Then ' ESTO SE PUSO POR UNA ACTUALIZACION ELIMIANR ESTO LUEGO.
        sql = "Select * from DTX_parametro Where LINEA_PROC ='" & linea_proc & "' and  Linea ='" & linea & "' and codigo = 'COND_CONTEIN_VACIO'"
        reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        If reg.EOF Then
            Gcon_docu.Execute "Insert into DTX_parametro VALUES ('CAE', 'MSC', 'COND_CONTEIN_VACIO',  'Condicion de contenedor Vacio', '6')"
            Gcon_docu.Execute "Insert into DTX_parametro VALUES ('CAE', 'MSC', 'SELLO_CONTEIN_VACIO',  'Sello por defecto para Cont vacios', 'WITHOUT')"
        End If
        reg.Close
    End If
    If Trim(Gstr_TIPO_PRODUCTO) = "" Then ' ESTO SE PUSO POR UNA ACTUALIZACION ELIMIANR ESTO LUEGO.
        sql = "Select * from DTX_parametro Where LINEA_PROC ='" & linea_proc & "' and  Linea ='" & linea & "' and codigo = 'TIPO_PRODUCTO'"
        reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        If reg.EOF Then
            Gcon_docu.Execute "Insert into DTX_parametro VALUES ('CAE', 'MSC', 'TIPO_PRODUCTO',  'Tipo de Producto', '49000')"
            Gcon_docu.Execute "Insert into DTX_parametro VALUES ('CAE', 'MSC', 'TIPO_PRODUCTO_VACIO',  'Tipo de Producto Cont vacios', '49500')"
        End If
        Gstr_TIPO_PRODUCTO = "49000"
        Gstr_TIPO_PRODUCTO_VACIO = "49500"
        reg.Close
    End If
    If Trim(Gstr_Desc_carga) = "" Then ' ESTO SE PUSO POR UNA ACTUALIZACION ELIMIANR ESTO LUEGO.' Julio/2008
        sql = "Select * from DTX_parametro Where LINEA_PROC ='" & linea_proc & "' and  Linea ='" & linea & "' and codigo = 'DESC_CARGA'"
        reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        If reg.EOF Then
            Gcon_docu.Execute "Insert into DTX_parametro VALUES ('CAE', 'MSC', 'DESC_CARGA',  'Descripcion de la Carga por defecto', 'SAID TO CONTAIN')"
            Gcon_docu.Execute "Insert into DTX_parametro VALUES ('CAE', 'MSC', 'MARCAS_NUM',  'Marcas y Numeros', 'S/M')"
        End If
        Gstr_Desc_carga = "SAID TO CONTAIN"
        Gstr_Marca_num = "S/M"
        reg.Close
    End If

End Sub

Sub procesa_archivo_DTX_XML(Archivo As String)

   Dim reg As New ADODB.Recordset
   
   Dim Objxlm          As New MSXML.DOMDocument
  
    ' accesa a cada elemento JOB
    
    
    Dim objXML As New MSXML.DOMDocument
    Dim objElemList As MSXML.IXMLDOMNodeList
    Dim objElem As MSXML.IXMLDOMElement
    Dim objElemBL As MSXML.IXMLDOMElement
    Dim objRes As MSXML.IXMLDOMElement
    Dim puerto As String
    Dim fecha_zarpe
    Dim pto_origen
    Dim pto_descarga
    Dim COD_BL As String
   
    objXML.loadXML "<DTXs></DTXs>"
    
    
'    While Not rs.EOF
'        codigo = Trim(rs("equip_coord"))
'        retObtDatos = ObtDatos(ObjConn, codigo, nombre, idusuario)
'        If retObtDatos <> "ok" Then
'           LeeUsuariosaNotificar = retObtDatos
'           Exit Function
'        End If
'
'        Set objElem = objXML.createElement("Usuario")
'        objElem.setAttribute "IdUsuario", idusuario
'        objElem.setAttribute "codigo", codigo
'        objElem.setAttribute "nombre", nombre
'        objXML.selectSingleNode("/Usuarios").appendChild objElem
'
'        rs.MoveNext
'    Wend
'
    'XmlUsuarios = objXML.xml
  
   
    x = FreeFile
   
    Open Archivo For Input As #x
Do While Not EOF(x)
    Line Input #x, linea
    cod_linea = Mid$(linea, 1, 3)
   
  Select Case cod_linea
    Case "VLS"
       sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' and codigo ='VLS' and nombre_tabla = 'REGISTRO'"
       reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
       
       Do While Not reg.EOF
          Select Case Trim$(reg("CAMPO_TABLA"))
             Case "DIRECCION":
    '            r_registro.direcc = Mid$(linea, reg("desde"), reg("longitud"))
          End Select
          reg.MoveNext
       Loop
       reg.Close
          
    Case "POR"
        sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'MANIFIESTO'"
        
        reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        Do While Not reg.EOF
           Select Case reg("CAMPO_TABLA")
              Case "PUERTO":
                 puerto = Codigo_equivalente_DTX("DTX_PUERTO", Mid$(linea, reg("desde"), reg("longitud")), "APG", "CODIGO")
              Case "FECHA_ZARPE"
                 fecha_zarpe = Aplica_Formato(cod_linea, reg("CAMPO"), Mid$(linea, reg("desde"), reg("longitud")))
           End Select
           reg.MoveNext
        Loop
        reg.Close
        
        
        sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
        reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                
        Do While Not reg.EOF
           Select Case reg("CAMPO_TABLA")
              Case "PTO_ORIGEN":
                 If Trim(Mid$(linea, reg("desde"), reg("longitud"))) <> "" Then ' si el pto de origen viene en blanco
                     pto_origen = Codigo_equivalente_DTX("DTX_PUERTO", Mid$(linea, reg("desde"), reg("longitud")), "APG", "CODIGO")
                 Else
                     pto_origen = puerto
                 End If
              Case "PTO_DESCARGA"
                 pto_descarga = Codigo_equivalente_DTX("DTX_PUERTO", Mid$(linea, reg("desde"), reg("longitud")), "APG", "CODIGO")
           End Select
           reg.MoveNext
        Loop
        reg.Close
        
        Set objElem = objXML.createElement("Puerto")
        objElem.setAttribute "PUERTO", puerto
        objElem.setAttribute "FECHA_ZARPE", fecha_zarpe
        objElem.setAttribute "PTO_ORIGEN", pto_origen
        objElem.setAttribute "PTO_DESCARGA", pto_descarga
        
        objXML.selectSingleNode("/DTXs").appendChild objElem
        
        
    Case "BL1"

        sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
          'Set reg = Gcon_docu.OpenRecordset(sql)
          reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
          Do While Not reg.EOF
             Select Case reg("CAMPO_TABLA")
                Case "COD_BL":
                   COD_BL = Mid$(linea, reg("desde"), reg("longitud"))
             End Select
             reg.MoveNext
          Loop
          reg.Close
          
          
          Set objElemBL = objXML.createElement("BL")
          objElemBL.setAttribute "COD_BL", COD_BL
          'objXML.selectSingleNode("/Puerto").appendChild objElemBL
          objElem.appendChild objElemBL
          
    End Select
    
Loop
objXML.Save ("c:\Dtxxml.xml")
End Sub


Public Sub lee_dtx_xml()
    Dim Objxlm          As New MSXML.DOMDocument
    Dim objElemList     As MSXML.IXMLDOMNodeList
    Dim objElem         As MSXML.IXMLDOMElement


Objxlm.loadXML XmlComponentes
            
    If Trim(XmlComponentes) <> "<DTXs></DTXs>" Then
       Set objElemList = Objxlm.selectNodes("/DTXs/DTX")
       'cada elemento que exista en componente
       For Each objElem In objElemList
          BPACWMaster = objElem.getAttribute("PTO")
          
          
       Next
    End If

End Sub

Sub carga_archivo_DTX_grid(Archivo As String)
On Error GoTo errores_grid
   Dim reg As New ADODB.Recordset
    Dim puerto As String
    Dim fecha_zarpe
    Dim pto_origen
    Dim pto_descarga
    Dim COD_BL As String
   
    x = FreeFile
   
    Open Archivo For Input As #x
    
    DTX_PROCESO.Spr_dtx.MaxRows = 0
Do While Not EOF(x)
    Line Input #x, linea
    cod_linea = Mid$(linea, 1, 3)
   
  Select Case cod_linea
    Case "VLS"
       sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' and codigo ='VLS' and nombre_tabla = 'REGISTRO'"
       reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
       
       Do While Not reg.EOF
          Select Case Trim$(reg("CAMPO_TABLA"))
             Case "DIRECCION":
    '            r_registro.direcc = Mid$(linea, reg("desde"), reg("longitud"))
          End Select
          reg.MoveNext
       Loop
       reg.Close
          
    Case "POR"
        sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'MANIFIESTO'"
        
        reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        Do While Not reg.EOF
           Select Case reg("CAMPO_TABLA")
              Case "PUERTO":
                 puerto = Mid$(linea, reg("desde"), reg("longitud"))
              Case "FECHA_ZARPE"
                 'fecha_zarpe = Aplica_Formato(cod_linea, reg("CAMPO"), Mid$(linea, reg("desde"), reg("longitud")))
           End Select
           reg.MoveNext
        Loop
        reg.Close
        
        
        sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
        reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
                
        Do While Not reg.EOF
           Select Case reg("CAMPO_TABLA")
              Case "PTO_ORIGEN":
                 If Trim(Mid$(linea, reg("desde"), reg("longitud"))) <> "" Then ' si el pto de origen viene en blanco
                     pto_origen = Mid$(linea, reg("desde"), reg("longitud"))
                 Else
                     pto_origen = puerto
                 End If
              Case "PTO_DESCARGA"
                 pto_descarga = Mid$(linea, reg("desde"), reg("longitud"))
           End Select
           reg.MoveNext
        Loop
        reg.Close
        
        DTX_PROCESO.Spr_dtx.MaxRows = DTX_PROCESO.Spr_dtx.MaxRows + 1
        DTX_PROCESO.Spr_dtx.Row = DTX_PROCESO.Spr_dtx.MaxRows
        DTX_PROCESO.Setea_Formato_Spread_pto DTX_PROCESO.Spr_dtx.MaxRows
        DTX_PROCESO.Spr_dtx.Col = 2: DTX_PROCESO.Spr_dtx.Text = puerto
        DTX_PROCESO.Spr_dtx.Col = 3: DTX_PROCESO.Spr_dtx.Text = pto_origen
        DTX_PROCESO.Spr_dtx.Col = 4: DTX_PROCESO.Spr_dtx.Text = pto_descarga
        DTX_PROCESO.Spr_dtx.Col = 5: DTX_PROCESO.Spr_dtx.Text = pto_descarga
        
        DTX_PROCESO.Spr_dtx.Col = 6: DTX_PROCESO.Spr_dtx.Text = Codigo_equivalente_DTX("DTX_PUERTO", puerto, "APG", "CODIGO")
        DTX_PROCESO.Spr_dtx.Col = 7: DTX_PROCESO.Spr_dtx.Text = Codigo_equivalente_DTX("DTX_PUERTO", pto_origen, "APG", "CODIGO")
        DTX_PROCESO.Spr_dtx.Col = 8: DTX_PROCESO.Spr_dtx.Text = Codigo_equivalente_DTX("DTX_PUERTO", pto_descarga, "APG", "CODIGO")
        DTX_PROCESO.Spr_dtx.Col = 9: DTX_PROCESO.Spr_dtx.Text = Codigo_equivalente_DTX("DTX_PUERTO", pto_descarga, "APG", "CODIGO")
        
    Case "BL1"

        sql = "Select * from DTX_detalle Where linea = '" & Gstr_Linea & "' And version = '" & Gstr_Version_MSC & "' And codigo ='" & cod_linea & "' and nombre_tabla = 'BLIMPO'"
        reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
        Do While Not reg.EOF
             Select Case reg("CAMPO_TABLA")
                Case "COD_BL":
                   COD_BL = Mid$(linea, reg("desde"), reg("longitud"))
             End Select
             reg.MoveNext
        Loop
        reg.Close
          
        DTX_PROCESO.Spr_dtx.MaxRows = DTX_PROCESO.Spr_dtx.MaxRows + 1
        DTX_PROCESO.Spr_dtx.Row = DTX_PROCESO.Spr_dtx.MaxRows
        DTX_PROCESO.Setea_Formato_Spread_bl DTX_PROCESO.Spr_dtx.MaxRows
        DTX_PROCESO.Spr_dtx.Col = 10: DTX_PROCESO.Spr_dtx.Text = COD_BL
        
          
    End Select
    
Loop
Exit Sub
errores_grid:
    MsgBox "Ocurrio el error " & Err.Description
    Screen.MousePointer = 0
    Exit Sub
'objXML.Save ("c:\Dtxxml.xml")
End Sub


' esta funcion sirve para confirmar si el pto o el bl que se etsa procesando esta
' seleccionado o no para que se procese, faltaria que tome los codigos equivalentes
Public Function Grid_confirma_carga(tipo As String, pto_emb As String, pto_desc As String, Optional bl As String) As Boolean
Dim i As Integer
Dim spr_pto_emb  As String
Dim spr_pto_desc As String
confirma_carga = False
    If GLinea_act_Grid = 0 Then GLinea_act_Grid = 1
    For i = GLinea_act_Grid To DTX_PROCESO.Spr_dtx.MaxRows
        DTX_PROCESO.Spr_dtx.Row = i
        DTX_PROCESO.Spr_dtx.Col = 2: spr_pto_emb = DTX_PROCESO.Spr_dtx.Text
        DTX_PROCESO.Spr_dtx.Col = 4: spr_pto_desc = DTX_PROCESO.Spr_dtx.Text
        
        
        If Trim(pto_emb) = Trim(spr_pto_emb) And Trim(pto_desc) = Trim(spr_pto_desc) Then
            DTX_PROCESO.Spr_dtx.Col = 1
            If CBool(DTX_PROCESO.Spr_dtx.Value) Then ' si esta marcado entonces si se sube
                confirma_carga_pto = True
            Else
                confirma_carga_pto = False
            End If
            If Mid(tipo, 1, 3) = "PTO" Then
                Grid_confirma_carga = confirma_carga_pto
                GLinea_act_Grid = i
                Exit Function
            End If
        End If
    
        
        DTX_PROCESO.Spr_dtx.Col = 10
        If Trim(bl) = Trim(DTX_PROCESO.Spr_dtx.Text) And bl <> "" Then
            DTX_PROCESO.Spr_dtx.Col = 1
            If CBool(DTX_PROCESO.Spr_dtx.Value) Then ' si esta marcado entonces si se sube
               confirma_carga_bl = True
            Else
                confirma_carga_bl = False
            End If
            If tipo = "BL" Then
                Grid_confirma_carga = confirma_carga_pto And confirma_carga_bl
                Exit Function
            End If
            
        End If
        
    Next i
End Function

'devuelve del grid de pantalla el valor equivalente al codigo que se recibe del archivo DTX.
' es decir toma lo que esta en el grid que el usuario modifico para registrarlo en la Base de datos
Public Function Grid_Codigo_equivalente(tipo As String) As String

DTX_PROCESO.Spr_dtx.Row = GLinea_act_Grid
Select Case tipo
    Case "PTO_EMBA":  DTX_PROCESO.Spr_dtx.Col = 6:
    Case "PTO_ORIG":  DTX_PROCESO.Spr_dtx.Col = 7
    Case "PTO_DESC":  DTX_PROCESO.Spr_dtx.Col = 8
    Case "PTO_DEST":  DTX_PROCESO.Spr_dtx.Col = 9
    Case "PTO_EMBA_ORIG":       DTX_PROCESO.Spr_dtx.Col = 2
    Case "PTO_DESC_ORIG":      DTX_PROCESO.Spr_dtx.Col = 4
End Select
Grid_Codigo_equivalente = DTX_PROCESO.Spr_dtx.Text

End Function
