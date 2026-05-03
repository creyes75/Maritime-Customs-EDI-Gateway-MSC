Attribute VB_Name = "Doc001"
Declare Function WinExec Lib "kernel32" (ByVal lpszCmdLine As String, ByVal fuCmdShow As Integer) As Integer
Declare Function FindWindow Lib "user32" (ByVal lpszClassName As Any, ByVal lpszWindow As Any) As Integer
Declare Function GetActiveWindow Lib "user32" () As Integer
Declare Function IsWindow Lib "user32" (ByVal hwnd As Integer) As Integer
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'para sql
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
          (ByVal hwndParent As Long, ByVal fRequest As Long, _
          ByVal lpszDriver As String, ByVal lpszAttributes As String) _
          As Long
Private Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)
Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer



' tipos definidos por el usuario
Type Ayuda
     campo1            As String
     campo2            As String
     tabla             As String
     criterio          As String
     valor_devuelto    As String
     nombre            As String
     Serv_rol          As Boolean
End Type

Type struct_blimpodt 'estructura para manejar el arreglo de los detalles
   SECUENCIA        As Integer
   tipo_carga       As String * 2
   peso             As Double
   no_bultos        As Double
   embalaje         As String * 3
   marca_num        As String * 255
   desc_carga       As String
   tipo_prod        As String
   volumen          As Double
   CONTEINER        As String * 11
   just_cambio      As String
   Imo              As String * 7
   pelig_ONU        As String * 5
   Sello            As String * 15
   tamano_equip     As String * 2
   tipo_equip       As String * 3
   Tipo_cont        As String * 3
   cond_cont        As String * 3
   elim             As String * 1 'Si tiene "E" significa que ese detalle ha sido eliminado
                                  'y no pasara a la tabla de la base
   danger_cargo     As String * 3 ' para manejar la peligrosidad de KHS
   UN_number        As String * 4
   IMDG_page        As String * 7
End Type

Type parametro
    pto_descarga    As String
    adua_ing        As String
    tipo_transac    As String
    Linea_Nav       As String
    naviera         As String
    tipo_manif      As String
    tipo_manif_expo As String
    delimitador     As String
    separador       As String
    clave           As String
    clave_KHS       As String
    clave_TAYLOR    As String
    tipo_oper       As String
    MONEDA          As String
    tipo_documento  As String
    naturaleza_carga As String
    condicion_cont  As String
    proveed_eq      As String
    tipo_carga      As String
    tipo_id         As String
    nacionalidad    As String
    metodo_pago     As String
    requer_serv     As String
    cond_contrato   As String
    estado_equip    As String
    estado_equip_expo As String
    indic_llen_vac  As String
    regimen_adua    As String
    regimen_adua_expo As String
    regimen_adua_trasb As String
    tipo_equipo     As String
    almacen_temporal As String
End Type

Type struct_flete
    cargo       As String
    cargo_cod   As String
    debtor      As String
    deptor_cod  As String
    valor       As Double
    SIGNO       As String
    MONEDA      As String
    marcas      As String
    lugar       As String
    imp_bl      As Boolean
    valor_unit  As Double
    nivel_cargo As String
    cantidad    As Double
    reg_cancel  As Boolean
End Type

Type Parametro_plano    ' para los parametros del proceso de archivo plano
    embarcador          As String
    consignatario       As String
    PESO_CONT20         As String
    PESO_CONT40         As String
    Direc_Embarcador    As String
    Direc_consignatario As String
    notificador         As String
    Direc_notificador   As String
    embalaje            As String
    Sello               As String
    Volumen20           As Double
    Volumen40           As Double
    Imo                 As String
    
    id_consignatario    As String
    id_embarcador       As String
    id_notificador      As String
    Indic_vacio         As String
    Tipo_carga_cont_vacio As String
    desc_carga          As String
    pto_opt             As String       'puerto opcional OPT
    marcas_num          As String
    cond_cont           As String
End Type

Global fileconfig                   As String
Global Gcon_docu                    As ADODB.Connection  'As rdoConnection
Global Gbol_Cont_OK                 As Boolean
'Global ambiente_base                As rdoEnvironment
Global Gstr_DSN                     As String
Global Gstr_user                    As String ' usuario de auditorio se obtiene el usuario de red o nombre de maquina
Global Gstr_user_conex              As String ' usuario para realizar la conexion
Global Gstr_psw                     As String
Global Gstr_base_doc                As String
Global Gstr_conex_docu              As String
Global Gstr_ruta                    As String
Global Gstr_ruta_reportes           As String
Global Gstr_ruta_windows            As String
Global Gstr_Nombre_Maquina          As String
Global Gstr_Ruta_files              As String
Global Gstr_Ruta_BlWord             As String

Global Gstr_ruta_log                As String
Global Gstr_arch_log                As String
'Global Qry_sp                       As rdoQuery
Global GSTR_FORMATO_FECHA           As String       ' UTILIZADA PARA TENER ALMACENADO EL FORMATO DE LA FECHA PARA EL INSERT.

Global Ver_msg                      As Boolean      ' seteada para presentar o no los mensajes de error

Global Gstr_opc_cont                As String * 1   ' variable para saber si esta en modificacion o adición en la opcion de movimiento de contenedores
Global Gtyp_ayuda      As Ayuda            ' contiene una estructura de la ayuda parametrizada
Global mani_bl_det()   As struct_blimpodt  ' para el manejo de los detalles de los bls.
'Global Gstr_opc_bl     As String           ' sirve para saber si se trata de adicion o modificacion de un bl en particular
Global Gstr_param      As parametro
Global Gtyp_flete      As struct_flete      ' para el manejo de los fletes en la pantalla asociada al bl(con grid)
Global Gtyp_flete_DTX() As struct_flete     ' arreglo de la estructura de fletes, sirve para manejar los valores de fletes que se incluyan en el DTX
Global Gstr_param_plano As Parametro_plano  ' para el proceso de archivos planos

'Global Gbol_CONT_desde_menu As Boolean   ' sirve para saber si la pantalla de contenedores fue llamada desde el mennu o desde la opcion de bls
Public Const Gstr_cia = "MSC Ecuador "

Function Open_Base_doc() As Integer
Dim Connect$
On Error GoTo error_open_base
Gbol_sigma_OK = False

Open_Base_SIGMA = False

Gstr_conex_docu = "DSN=" & Gstr_DSN & ";" & "UID=" & Gstr_user_conex & ";PWD=" & Gstr_psw & ";DATABASE=" & Gstr_base_doc & ";"
Set Gcon_docu = New ADODB.Connection
Gcon_docu.ConnectionString = "PROVIDER=MSDASQL;" & Gstr_conex_docu 'dsn=audiopubli;uid=;pwd=;"
Gcon_docu.Open
Gcon_docu.CursorLocation = adUseClient
'Set ambiente_base = rdoEnvironments(0)
'ambiente_base.CursorDriver = rdUseOdbc
'ambiente_base.LoginTimeout = 120
'Set Gcon_docu = ambiente_base.OpenConnection("", rdDriverNoPrompt, False, Gstr_conex_docu)

Open_Base_doc = True
Gbol_Cont_OK = True
Exit Function

error_open_base:
    
    MsgBox Error, vbCritical, App.Title
    If Err.Number = 3146 Or Err.Number = 3151 Then
        Open_Base_doc = 3
        Exit Function
    Else
        Open_Base_doc = False
        Exit Function
    End If
    Resume 0
End Function

'===================================================================*
' Función: Lee el archivo de configuración Monitor.INI el cual con-     *
'          tiene parámetros generales de la aplicación.             *
'===================================================================*
Function Lectura_INI() As Integer
Dim cc As String * 50
Dim rc As Integer
Dim ws_ini As String
On Error GoTo error_Lectura_INI
    ' Obtención del directorio donde esta instalado Windows
    Lectura_INI = 0
    ws_ini = Space(256)
    GetWindowsDirectory ws_ini, 256
    If InStr(ws_ini, Chr$(0)) Then
        ws_ini = Mid(ws_ini, 1, InStr(ws_ini, Chr$(0)) - 1)
    End If
    If Right(ws_ini, 1) = "\" Then
       Gstr_ruta_windows = ws_ini
       ws_ini = ws_ini & "doc.ini"
    Else
       Gstr_ruta_windows = ws_ini & "\"
       ws_ini = ws_ini & "\doc.ini"
    End If
    dirconfig = "c:\"
    fileconfig = ws_ini
    If Dir(dirconfig, 16) = "" Or Dir(fileconfig, 0) = "" Then
       MsgBox "ERROR : Archivo no existe." & Chr(13) & Chr(10) & "El archivo de configuración " & fileconfig & " no existe." & "Favor llamar a Computación.", 16
       Lectura_INI = -1
       Exit Function
    End If
    'Carga Variables de ambiente a la estructura var_sys_ambiente
    Gstr_DSN = Trim$(leer_parametro_rec("Generales", "DSN"))
    Gstr_base_doc = Trim$(leer_parametro_rec("Generales", "NAMEBASE"))
        
    If InStr(Gstr_DSN, Chr$(0)) Then
        Gstr_DSN = Mid(Gstr_DSN, 1, InStr(Gstr_DSN, Chr$(0)) - 1)
    End If
    If InStr(Gstr_base_doc, Chr$(0)) Then
        Gstr_base_doc = Mid(Gstr_base_doc, 1, InStr(Gstr_base_doc, Chr$(0)) - 1)
    End If
    
    ' se leen los parametros del servidor NT
    Gstr_ruta_log = Trim$(leer_parametro_rec("Errores", "Log_path"))
    Gstr_arch_log = Trim$(leer_parametro_rec("Errores", "log_name"))
    If InStr(Gstr_ruta_log, Chr$(0)) Then
       Gstr_ruta_log = Mid(Gstr_ruta_log, 1, InStr(Gstr_ruta_log, Chr$(0)) - 1)
    End If
    If InStr(Gstr_arch_log, Chr$(0)) Then
       Gstr_arch_log = Mid(Gstr_arch_log, 1, InStr(Gstr_arch_log, Chr$(0)) - 1)
    End If
    If Dir(Gstr_ruta_log, vbDirectory) = "" Then
       MsgBox "La ruta del archivo Log de Errores no existe, Por favor verifique en el archivo doc.ini", vbInformation, App.Title
       End
    End If
    
    Gstr_Nombre_Maquina = Space(256)
    GetComputerName Gstr_Nombre_Maquina, 256
    Gstr_Nombre_Maquina = Mid(Gstr_Nombre_Maquina, 1, InStr(Gstr_Nombre_Maquina, Chr$(0)) - 1)
    Gstr_ruta = App.Path
    If Right(Gstr_ruta, 1) = "\" Then
       Gstr_ruta_reportes = Gstr_ruta & "Reportes\"
    Else
       Gstr_ruta_reportes = Trim$(Gstr_ruta) & "\Reportes\"
       Gstr_ruta = Trim$(Gstr_ruta) & "\"
    End If
    
    Gstr_Ruta_files = Trim$(leer_parametro_rec("Rutas", "files"))
    Gstr_Ruta_files = Mid(Gstr_Ruta_files, 1, InStr(Gstr_Ruta_files, Chr$(0)) - 1)
    If Right(Gstr_Ruta_files, 1) <> "\" Then Gstr_Ruta_files = Gstr_Ruta_files & "\"


    Gstr_Ruta_BlWord = Trim$(leer_parametro_rec("Rutas", "blword"))
    Gstr_Ruta_BlWord = Mid(Gstr_Ruta_BlWord, 1, InStr(Gstr_Ruta_BlWord, Chr$(0)) - 1)
    If Right(Gstr_Ruta_BlWord, 1) <> "\" Then Gstr_Ruta_BlWord = Gstr_Ruta_BlWord & "\"
    ' ******************************************************************
    
    GSTR_FORMATO_FECHA = Trim$(leer_parametro_rec("Generales", "FECHA"))
    GSTR_FORMATO_FECHA = Mid(GSTR_FORMATO_FECHA, 1, InStr(GSTR_FORMATO_FECHA, Chr$(0)) - 1)
    If GSTR_FORMATO_FECHA = "" Then GSTR_FORMATO_FECHA = "dd/mm/yyyy"
    
    Lectura_INI = 1
    Exit Function
error_Lectura_INI:
    MsgBox Err.Description
    Lectura_INI = 0
    Exit Function
    Resume 0
End Function

'Sección: sección del archivo ini
'Parámetro: el valor a recuperar
'-------------------------------------------------------
Function leer_parametro_rec(section As String, parametro As String) As String
Dim valor_parametro As String

  valor_parametro = String$(255, " ")
  lectura% = GetPrivateProfileString(section, parametro, Chr$(0), valor_parametro, Len(valor_parametro), fileconfig)
  If lectura% = 0 Then
    MsgBox "ERROR: PARAMETRO" + NL + NL + "Falta un valor para " & parametro & " en el archivo de configuración", 48
    Exit Function
  End If
leer_parametro_rec = valor_parametro
End Function

'Funcion        : Codigo Equivalente
'Objetivo       : Devolver el código equivalente en APG en base al código
'                 Internacional impuesto por las Agencias Navieras
'Fecha Creacion : Lunes 29 de Junio de 1998 - 21:10:25 (una noche triste)
'Creado por     : Carlos Reyes Atiencia
'Parámetros     : Tabla .- es el nombre de la tabla intermedia que tiene ambos codigos
'                 codigo_buscado .- es el valor en si que se esta buscando
'                 campo_apg .- es el nombre del campo que contiene el código APG
'===============================================================================
' no se esta usando actualmente 18/julio/2002
Public Function Codigo_equivalente(tabla As String, ByVal codigo_buscado As String, campo_apg As String, campo_inter As String, Optional campo_st As String) As String

On Error GoTo error_busca_cod_apg
   Dim sql As String
   Dim REGISTRO As New ADODB.Recordset
'   If tabla = "factcon" And codigo_buscado = "40RE" Then
'       codigo_buscado = "40RF"
'   End If
   sql = "select " & Trim$(campo_apg) & " from " & Trim$(tabla)
   sql = sql + " where " & Trim$(campo_inter) & " = '" & Trim$(codigo_buscado) & "'"
   If Not IsMissing(campo_st) And Trim$(campo_st) <> "" Then
      sql = sql + "and " & campo_st & " = 'A'"
   End If
   'Set registro = Gcon_docu.OpenRecordset(sql)
   REGISTRO.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   If REGISTRO.EOF Then
      Codigo_equivalente = "NO"
   Else
      Codigo_equivalente = Trim(REGISTRO(campo_apg))
   End If
   Exit Function
error_busca_cod_apg:
'   If Err.Number >= 40000 Then
'      Dim er As rdoError
'      For Each er In rdoErrors
'          If Ver_msg Then MsgBox "Ocurrio el error " & er.Description, vbInformation, App.Title
'          Call Graba_log("Codigo_equivalente", er.Number, er.Description)
'      Next er
'      rdoErrors.Clear
'    Else
      If Ver_msg Then MsgBox "Ocurrio el error " + Err.Description
      Call Graba_log("Codigo_equivalente", Err.Number, Err.Description)
'    End If
   Exit Function
   Resume 0
End Function

' procedimiento de grabación de archivo log para errores de caidas en los diferentes
' programas   ( lunes 30 de Octubre /2000 despues que se perdió la información por daño en disco duro)
' ----------------------------------------------------------------------------------------------------
Public Sub Graba_log(ByVal Procedimiento As String, _
                     ByVal no_error As Double, ByVal desc_error As String)
Dim x
On Error GoTo errores_log
  x = FreeFile
  Open Gstr_ruta_log & Gstr_arch_log For Append As #x
  Print #1, App.Title & "|" & Gstr_Nombre_Maquina & "|" & Format$(Now, "dd/mm/yyyy") & "|" & Format$(Time, "hh:mm:ss") & "|" & Procedimiento & "|" & no_error & "|" & Trim$(desc_error)
  Close #x
  Exit Sub
errores_log:
   MsgBox "Ocurrio el error " & Err.Description, vbInformation, App.Title
  
End Sub

'Function LlamadaStoreProcedure(Prmsalida As Byte, indice As Byte, StoreProcedure As String, ParamArray parametro()) As Boolean
''Indice = 0 Ejecuta un Store Procedure sin Parámetros que devuelve datos
''Indice = 1 Ejecuta un Store Procedure con Parámetros que devuelve datos de un SELECT
''Indice = 2 Ejecuta un Store Procedure con Parámetros que no devuelve datos por ejemplo un Update, Delete, Insert
''Indice = 3 Ejecuta una sentencia concatenada, "SELECT" enviado desde Visual
''Indice = 4 Ejecuta una sentencia concatenada, "INSERT" enviado desde Visual
'
'On Error GoTo ERRLlamadaStoreProcedure
'Dim emCntIndice As Integer
'
'LlamadaStoreProcedure = True
'
'    Set Qry_sp = Gcon_doc.CreateQuery("", "")
'
'    Select Case indice
'       Case 0
'            Qry_sp.sql = "{call " & StoreProcedure & " }"
'       Case 1, 2
'            Qry_sp.sql = "{call " & StoreProcedure & " (?"
'            For emCntIndice = 1 To UBound(parametro)
'                Qry_sp.sql = Qry_sp.sql & ",?"
'            Next emCntIndice
'            Qry_sp.sql = Qry_sp.sql & ")}"
'
'            For emCntIndice = 1 To Prmsalida
'                Qry_sp.rdoParameters(emCntIndice - 1).Direction = rdParamInputOutput
'            Next emCntIndice
'
'            For emCntIndice = 0 To UBound(parametro)
'                Qry_sp.rdoParameters(emCntIndice).Value = parametro(emCntIndice)
'            Next emCntIndice
'
'       Case 3, 4
'            Qry_sp.sql = StoreProcedure
'
'    End Select
'    LlamadaStoreProcedure = True
'    Select Case indice
'        Case 0, 1, 3
'                Set ogrdoConsultaTemporal = Nothing
'                Set ogrdoConsultaTemporal = Qry_sp.OpenRecordset(rdOpenDynamic)
'                 If ogrdoConsultaTemporal.RowCount < 1 Then
'                    LlamadaStoreProcedure = False
'                    ogrdoConsultaTemporal.Close
'                    Set ogrdoConsultaTemporal = Nothing
'                    Exit Function
'                End If
'        Case 2, 4
'                Qry_sp.Execute
'                If Qry_sp.rdoParameters(0) = 0 Then ' se valida la respuesta del SP para saber si esta bien
'                   LlamadaStoreProcedure = True
'                Else
'                   LlamadaStoreProcedure = False
'                End If
'    End Select
'
'Exit Function
'ERRLlamadaStoreProcedure:
'    'If ogrdoConexionGeneral.Transactions Then ogrdoConexionGeneral.RollbackTrans
'    'Set ogrdoConsultaTemporal = Nothing
'    'MsgBox Mid(rdoErrors(rdoErrors.Count - 1).Description, InStr(1, rdoErrors(0).Description, "ERROR:") + 6, Len(rdoErrors(0).Description)), vbCritical, "Cuentas Por Pagar"
'    'ControlErrorSQL
'
'    LlamadaStoreProcedure = False
'    If Err.Number >= 40000 Then
'      Dim er As rdoError
'      For Each er In rdoErrors
'          If Ver_msg Then MsgBox "Ocurrio el error " & er.Description, vbInformation, App.Title
'          Call Graba_log("LlamadaStoreProcedure", er.Number, er.Description & " En el Store Procedure " & StoreProcedure)
'      Next er
'      rdoErrors.Clear
'    Else
'      If Ver_msg Then MsgBox "Ocurrio el Error " & Err.Description, vbInformation, App.Title
'      Call Graba_log("LlamadaStoreProcedure", Err.Number, Err.Description & " En el Store Procedure " & StoreProcedure)
'    End If
'    Exit Function
'    Resume 0
'
'End Function


Public Sub Main()
Dim cadena      As String
Dim segur        As String
Dim reg          As New ADODB.Recordset
Dim rc, rc2      As Integer
    On Error GoTo LError
    
    If Trim$(Command) <> "" Then
        cadena = Trim$(Command)
        Gstr_user = Trim$(Mid$(cadena, 1, InStr(1, cadena, " ") - 1))
        Gstr_psw = Trim$(Mid$(cadena, InStr(1, cadena, " ") + 1))
    End If
INICIO:
    If Trim$(Gstr_user) = "" Or Trim$(Gstr_psw) = "" Then
       Seguridad.Show 1
       'MsgBox "Error en el Envío de Parámetros." & Chr(10) & Chr(13) & "Esta Aplicación debe ser iniciada desde el Menu Principal", vbCritical, App.Title
       'Exit Sub
    End If
        
    Screen.MousePointer = vbHourglass
    If Lectura_INI() = 1 Then
         ' se verifica que existta el ODBC y se procede a crear uno si no existe.
         If Not gbolCheckDSN(Gstr_DSN) Then
            'If Not gbolCreateDSN(Gstr_DSN, "Documentacion", Gstr_base_doc) Then
                MsgBox "No Existe el ODBC, " & Gstr_DSN & " comuníquese con su Administrador", vbCritical
                Exit Sub
            'End If
         End If
         ' finaliza la creacion y se inicia la apertura de la base
         Gstr_user_conex = Gstr_user ' se almacena el usuario de la conexion en otra variable
         If Open_Base_doc() = True Then
             Gstr_user = Mid(Gstr_Nombre_Maquina, 1, 8) ' ojo usuario=maquina
             Doc_menu.Show
         Else
             MsgBox "Error con la Base de datos Documentación en el Servidor NT, proceso cancelado.", vbCritical, App.Title
             Screen.MousePointer = vbDefault
             If Trim$(Command) = "" Then
                Gstr_user = "": Gstr_psw = "": Gstr_user_conex = ""
                GoTo INICIO
             End If
             Exit Sub
         End If
    Else
        MsgBox "Error en Carga de Variables de Ambiente de doc.INI"
        Screen.MousePointer = vbDefault
        End
    End If
salir:
    Screen.MousePointer = vbDefault
    Exit Sub
    
LError:
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " " & Err.Description, vbInformation, App.Title
    Exit Sub
    Resume 0
End Sub


Public Function gbolCheckDSN(DSN_NAME As String) As Boolean
    Dim i As Integer
    Dim sDSNItem As String * 1024
    Dim sDRVItem As String * 1024
    Dim sDSN As String
    Dim sDRV As String
    Dim iDSNLen As Integer
    Dim iDRVLen As Integer
    Dim lHenv As Long         'handle to the environment

    On Error Resume Next
    
    gbolCheckDSN = False

    'get the DSNs
    If SQLAllocEnv(lHenv) <> -1 Then
        Do Until i <> SQL_SUCCESS
            sDSNItem = Space$(1024)
            sDRVItem = Space$(1024)
            i = SQLDataSources(lHenv, 1, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
            sDSN = Left$(sDSNItem, iDSNLen)
            sDRV = Left$(sDRVItem, iDRVLen)
                
            If sDSN <> Space(iDSNLen) Then
                If LCase(Trim(sDSN)) = LCase(DSN_NAME) Then
                    gbolCheckDSN = True
                    Exit Function
                End If
            End If
        Loop
    End If
    
End Function

Public Function gbolCreateDSN(DSN_NAME As String, DSN_DESCRIPTION As String, DSN_DATABASE As String) As Boolean

' Inicio - declaracion Dll para crear DSN si no existe. - 082702 JCTB
Dim intRet As Long
Dim strDriver As String
Dim strAttributes As String
gbolCreateDSN = False
vbAPINull = 0
ODBC_ADD_DSN = 1

'Set the driver.
strDriver = "SQL Server"
'Set the attributes delimited by null.
    strAttributes = ""
    '"SERVER=" & mstrDSN_SERVER & Chr$(0)
      strAttributes = strAttributes & "DSN=" & DSN_NAME & "2" & Chr$(0)
      strAttributes = strAttributes & "DESCRIPTION=" & DSN_DESCRIPTION & Chr$(0)
      strAttributes = strAttributes & "DBQ=" & DSN_DATABASE & Chr$(0)
      strAttributes = strAttributes & "SERVER=" & "." & Chr$(0)
      'To show dialog, use Form1.Hwnd instead of vbAPINull.
      intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, strDriver, strAttributes)
      If intRet Then
            gbolCreateDSN = True ' "DSN Created"
            MsgBox "El sistema creo de forma automatica el DSN=" & DSN_NAME & " por favor informe a su soporte Tecnico.", vbInformation
      Else
            gbolCreateDSN = False '"Create Failed"
            MsgBox "El sistema NO pudo crear de forma automatica el DSN=" & DSN_NAME & " por favor informe a su soporte Tecnico.", vbInformation
      End If
' Fin - declaracion Dll para crear DSN si no existe. - 082702 JCTB

End Function


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
Function Ayuda(ByVal tabla As String, ByVal codigo As String, ByVal DESCRIPCION As String, ByVal criterio As String, ByVal nombre As String) As String
    
    Gtyp_ayuda.tabla = tabla
    Gtyp_ayuda.campo1 = codigo
    Gtyp_ayuda.campo2 = DESCRIPCION
    Gtyp_ayuda.criterio = criterio
    Gtyp_ayuda.nombre = nombre
    Gtyp_ayuda.Serv_rol = Serv_ROLES
    doc_HELP.Show 1
    Ayuda = Gtyp_ayuda.valor_devuelto
End Function

Sub Centra_Pantalla(forma As Form)
  forma.Left = (Screen.Width - forma.Width) / 2
  forma.Top = ((Screen.Height - 1300) - forma.Height) / 2
End Sub



Public Function Calcula_Lineas_Archivo(Archivo As String) As Integer
Dim x
Dim i As Integer

x = FreeFile
i = 0
Open Archivo For Input As #x
Do While Not EOF(x)
    Line Input #x, linea
    i = i + 1
Loop
Calcula_Lineas_Archivo = i
Close #x
End Function


Public Sub carga_datacombo(ByRef data_ado As Adodc, ByRef data_combo As DataCombo, _
                           ByVal sql As String, ByVal codigo As String, ByVal DESCRIPCION As String)
On Error GoTo Errores

   data_ado.ConnectionString = Gstr_conex_docu
   data_ado.RecordSource = sql & " order by " & DESCRIPCION
   data_ado.Refresh
   data_combo.BoundColumn = codigo '"move_codigo"
   data_combo.ListField = DESCRIPCION '"move_descripcion"
   Set data_combo.RowSource = data_ado
Exit Sub
Errores:
    MsgBox "Ocurrio el Error " & Err.Description
    Exit Sub
    Resume 0
End Sub

'   Funcion :  Carga Combo Spread, permite cargar los datos de la tabla maestra o sql que se recibe de parametro en variables
'              para ser utilizadas posteriormente en el Spread Sheet, se devuelven valores para la columna del Spread de codigo y de Descripcion
'   Parametros:Sql.- es el query que se ejecuta normalmente es el select a la tabla maestra
'              codigo.- contiene el nombre del campo que sera el codigo
'              descripcion.- contiene el nombre del campo que sera la descripcion
'              ComboDesc .- contiene la variable que se devuelve con la cadena que se utilizara como combo en el Spread con los valores de la descripcion
'              ComboCod  .- contiene la variable que se devuelve con la cadena que se utilizara como combo en el Spread con los valores del codigo
Public Sub carga_combo_Spread(ByVal sql As String, ByVal codigo As String, ByVal DESCRIPCION As String, ByRef comboDesc As String, ByRef comboCod As String)
On Error GoTo Errores
Dim reg As New ADODB.Recordset


   comboDesc = ""
   comboCod = ""
   reg.Open sql, Gcon_docu
   Do While Not reg.EOF
      If comboDesc <> "" Then comboDesc = comboDesc & Chr(9)
      comboDesc = comboDesc & reg.Fields(DESCRIPCION)
      
      If comboCod <> "" Then comboCod = comboCod & Chr(9)
      comboCod = comboCod & reg.Fields(codigo)
      reg.MoveNext
   Loop
   
   Exit Sub
Errores:
    MsgBox "Ocurrio el Error " & Err.Description
    Exit Sub
    Resume 0
End Sub


Public Sub actualiza_totales_manifiesto(linea As String, buque As String, viaje As String, pto_emb As String, pto_descarga As String, oper_carga As String)
On Error GoTo Errores
Dim sql As String
Dim sql_cond As String
Dim sql_act  As String
Dim reg_sum  As New ADODB.Recordset
Dim reg_doc  As New ADODB.Recordset
Dim reg_cont As New ADODB.Recordset

         sql_cond = " from docu_bldet " & _
                    " Where bldet_linea      ='" & linea & "'" & _
                    "   and bldet_buque    ='" & buque & "'" & _
                    "   and bldet_viaje    ='" & viaje & "'" & _
                    "   and bldet_puer_emb ='" & pto_emb & "'" & _
                    "   and bldet_puer_des ='" & pto_descarga & "'" & _
                    "   and bldet_emp_resp ='" & oper_carga & "'" '& _
                    "   and bldet_cono_emb ='" & bl & "'"
         
         sql = "Select sum(bldet_peso) as peso,   sum(bldet_bultos) as bultos, " & _
                     " sum(bldet_volumen) as volumen "
         
         reg_sum.Open sql & sql_cond, Gcon_docu, adOpenDynamic, adLockOptimistic
               
         sql = "Select count(*) as numero  "
         reg_doc.Open sql & sql_cond, Gcon_docu, adOpenDynamic, adLockOptimistic
         reg_cont.Open sql & sql_cond & " and bldet_contenedor <> ''", Gcon_docu, adOpenDynamic, adLockOptimistic
         
         sql_act = "Update docu_manifiesto set mani_tot_peso =" & IIf(Not IsNull(reg_sum("peso")), reg_sum("peso"), "0") & ", " & _
                                             " mani_tot_doc  =" & IIf(Not IsNull(reg_doc("numero")), reg_doc("numero"), "0") & ", " & _
                                             " mani_tot_bult =" & IIf(Not IsNull(reg_sum("bultos")), reg_sum("bultos"), "0") & ", " & _
                                             " mani_tot_vol  =" & IIf(Not IsNull(reg_sum("volumen")), reg_sum("volumen"), "0") & ", " & _
                                             " mani_tot_con  =" & IIf(Not IsNull(reg_cont("numero")), reg_cont("numero"), "0") & _
                    " Where mani_linea    ='" & linea & "'" & _
                    "   and mani_buque    ='" & buque & "'" & _
                    "   and mani_viaje    ='" & viaje & "'" & _
                    "   and mani_puer_emb ='" & pto_emb & "'" & _
                    "   and mani_puer_des ='" & pto_descarga & "'"

         Gcon_docu.Execute sql_act
Exit Sub
Errores:
MsgBox "Ocurio el error " & Err.Description
Exit Sub
Resume 0
End Sub

Public Sub actualiza_totales_manifiesto_expo(linea As String, buque As String, viaje As String, pto_emb As String, pto_descarga As String, oper_carga As String)
On Error GoTo Errores
Dim sql As String
Dim sql_cond As String
Dim sql_act  As String
Dim reg_sum  As New ADODB.Recordset
Dim reg_doc  As New ADODB.Recordset
Dim reg_cont As New ADODB.Recordset

         sql_cond = " from docu_bldet_exp " & _
                    " Where bldet_linea      ='" & linea & "'" & _
                    "   and bldet_buque    ='" & buque & "'" & _
                    "   and bldet_viaje    ='" & viaje & "'" & _
                    "   and bldet_puer_emb ='" & pto_emb & "'" & _
                    "   and bldet_puer_des ='" & pto_descarga & "'" & _
                    "   and bldet_emp_resp ='" & oper_carga & "'" '& _
                    "   and bldet_cono_emb ='" & bl & "'"
         
         sql = "Select sum(bldet_peso) as peso,   sum(bldet_bultos) as bultos, " & _
                     " sum(bldet_volumen) as volumen "
         
         reg_sum.Open sql & sql_cond, Gcon_docu, adOpenDynamic, adLockOptimistic
               
         sql = "Select count(*) as numero  "
         reg_doc.Open sql & sql_cond, Gcon_docu, adOpenDynamic, adLockOptimistic
         reg_cont.Open sql & sql_cond & " and bldet_contenedor <> ''", Gcon_docu, adOpenDynamic, adLockOptimistic
         
         sql_act = "Update docu_manifiesto_exp set mani_tot_peso =" & IIf(Not IsNull(reg_sum("peso")), reg_sum("peso"), "0") & ", " & _
                                             " mani_tot_doc  =" & IIf(Not IsNull(reg_doc("numero")), reg_doc("numero"), "0") & ", " & _
                                             " mani_tot_bult =" & IIf(Not IsNull(reg_sum("bultos")), reg_sum("bultos"), "0") & ", " & _
                                             " mani_tot_vol  =" & IIf(Not IsNull(reg_sum("volumen")), reg_sum("volumen"), "0") & ", " & _
                                             " mani_tot_con  =" & IIf(Not IsNull(reg_cont("numero")), reg_cont("numero"), "0") & _
                    " Where mani_linea    ='" & linea & "'" & _
                    "   and mani_buque    ='" & buque & "'" & _
                    "   and mani_viaje    ='" & viaje & "'" & _
                    "   and mani_puer_emb ='" & pto_emb & "'" & _
                    "   and mani_puer_des ='" & pto_descarga & "'"

         Gcon_docu.Execute sql_act
Exit Sub
Errores:
MsgBox "Ocurio el error " & Err.Description
Exit Sub
Resume 0
End Sub

Public Sub actualiza_totales_bl_expo(linea As String, buque As String, viaje As String, pto_emb As String, pto_descarga As String, oper_carga As String, bl As String)
On Error GoTo Errores
Dim sql As String
Dim sql_cond As String
Dim sql_act  As String
Dim reg_sum  As New ADODB.Recordset
'Dim reg_doc  As New ADODB.Recordset
'Dim reg_cont As New ADODB.Recordset

         sql_cond = " from docu_bldet_exp " & _
                    " Where bldet_linea      ='" & linea & "'" & _
                    "   and bldet_buque    ='" & buque & "'" & _
                    "   and bldet_viaje    ='" & viaje & "'" & _
                    "   and bldet_puer_emb ='" & pto_emb & "'" & _
                    "   and bldet_puer_des ='" & pto_descarga & "'" & _
                    "   and bldet_emp_resp ='" & oper_carga & "'" & _
                    "   and bldet_cono_emb ='" & bl & "'"
         
         sql = "Select sum(bldet_peso) as peso,   sum(bldet_bultos) as bultos, " & _
                     " sum(bldet_volumen) as volumen "
         
         reg_sum.Open sql & sql_cond, Gcon_docu, adOpenDynamic, adLockOptimistic
         
         sql_act = "Update docu_blcab_exp set  blcab_peso_tot    =" & IIf(Not IsNull(reg_sum("peso")), reg_sum("peso"), "0") & ", " & _
                                             " blcab_bultos_tot  =" & IIf(Not IsNull(reg_sum("bultos")), reg_sum("bultos"), "0") & ", " & _
                                             " blcab_volumen_tot =" & IIf(Not IsNull(reg_sum("volumen")), reg_sum("volumen"), "0") & " " & _
                    " Where blcab_linea      ='" & linea & "'" & _
                    "   and blcab_buque    ='" & buque & "'" & _
                    "   and blcab_viaje    ='" & viaje & "'" & _
                    "   and blcab_puer_emb ='" & pto_emb & "'" & _
                    "   and blcab_puer_des ='" & pto_descarga & "'" & _
                    "   and blcab_emp_resp ='" & oper_carga & "'" & _
                    "   and blcab_cono_emb ='" & bl & "'"

         Gcon_docu.Execute sql_act
Exit Sub
Errores:
MsgBox "Ocurio el error " & Err.Description
Exit Sub
Resume 0
End Sub

Public Sub actualiza_totales_bl(linea As String, buque As String, viaje As String, pto_emb As String, pto_descarga As String, oper_carga As String, bl As String)
On Error GoTo Errores
Dim sql As String
Dim sql_cond As String
Dim sql_act  As String
Dim reg_sum  As New ADODB.Recordset
'Dim reg_doc  As New ADODB.Recordset
'Dim reg_cont As New ADODB.Recordset

         sql_cond = " from docu_bldet " & _
                    " Where bldet_linea      ='" & linea & "'" & _
                    "   and bldet_buque    ='" & buque & "'" & _
                    "   and bldet_viaje    ='" & viaje & "'" & _
                    "   and bldet_puer_emb ='" & pto_emb & "'" & _
                    "   and bldet_puer_des ='" & pto_descarga & "'" & _
                    "   and bldet_emp_resp ='" & oper_carga & "'" & _
                    "   and bldet_cono_emb ='" & bl & "'"
         
         sql = "Select sum(bldet_peso) as peso,   sum(bldet_bultos) as bultos, " & _
                     " sum(bldet_volumen) as volumen "
         
         reg_sum.Open sql & sql_cond, Gcon_docu, adOpenDynamic, adLockOptimistic
         
         sql_act = "Update docu_blcab      set blcab_peso_tot    =" & IIf(Not IsNull(reg_sum("peso")), reg_sum("peso"), "0") & ", " & _
                                             " blcab_bultos_tot  =" & IIf(Not IsNull(reg_sum("bultos")), reg_sum("bultos"), "0") & ", " & _
                                             " blcab_volumen_tot =" & IIf(Not IsNull(reg_sum("volumen")), reg_sum("volumen"), "0") & ", " & _
                    " Where blcab_linea      ='" & linea & "'" & _
                    "   and blcab_buque    ='" & buque & "'" & _
                    "   and blcab_viaje    ='" & viaje & "'" & _
                    "   and blcab_puer_emb ='" & pto_emb & "'" & _
                    "   and blcab_puer_des ='" & pto_descarga & "'" & _
                    "   and blcab_emp_resp ='" & oper_carga & "'" & _
                    "   and blcab_cono_emb ='" & bl & "'"

         Gcon_docu.Execute sql_act
Exit Sub
Errores:
MsgBox "Ocurio el error " & Err.Description
Exit Sub
Resume 0
End Sub



Public Sub carga_parametros(linea As String)
Dim sql As String
Dim reg As New ADODB.Recordset

Gstr_param.pto_descarga = ""
Gstr_param.adua_ing = ""
Gstr_param.tipo_transac = ""
Gstr_param.Linea_Nav = ""
Gstr_param.naviera = ""
Gstr_param.tipo_manif = ""
Gstr_param.delimitador = ""
Gstr_param.separador = ""
Gstr_param.clave = ""
Gstr_param.tipo_oper = ""

sql = "Select * from docu_parametro where para_linea ='" & linea & "'"
reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
Do While Not reg.EOF
    Select Case Trim(reg("para_codigo"))
       Case "PTO_DESCARGA":         Gstr_param.pto_descarga = reg("para_valor")
       Case "ADUA_ING":             Gstr_param.adua_ing = reg("para_valor")
       Case "TIPO_TRANSAC":         Gstr_param.tipo_transac = reg("para_valor")
       Case "LINEA_NAV":            Gstr_param.Linea_Nav = reg("para_valor")
       Case "NAVIERA":              Gstr_param.naviera = reg("para_valor")
       Case "TIPO_MANIF":           Gstr_param.tipo_manif = reg("para_valor")
       Case "TIPO_MANIF_EXPO":      Gstr_param.tipo_manif = reg("para_valor")
       Case "DELIMITADOR":          Gstr_param.delimitador = reg("para_valor")
       Case "SEPARADOR":            Gstr_param.separador = reg("para_valor")
       Case "CLAVE":                Gstr_param.clave = reg("para_valor")
       Case "CLAVE_KHS":            Gstr_param.clave_KHS = reg("para_valor")
       Case "CLAVE_TAYLOR":         Gstr_param.clave_TAYLOR = reg("para_valor")
       Case "TIPO_OPERADOR":        Gstr_param.tipo_oper = reg("para_valor")
       Case "MONEDA":               Gstr_param.MONEDA = reg("para_valor")
       Case "TIPO_DOCUMENTO":       Gstr_param.tipo_documento = reg("para_valor")
       Case "NATURALEZA_CARGA":     Gstr_param.naturaleza_carga = reg("para_valor")
       Case "COND_CONTENEDOR":      Gstr_param.condicion_cont = reg("para_valor")
       Case "PROVEEDOR_EQ":         Gstr_param.proveed_eq = reg("para_valor")
       Case "TIPO_CARGA":           Gstr_param.tipo_carga = reg("para_valor")
       Case "TIPO_ID":              Gstr_param.tipo_id = reg("para_valor")
       Case "NACIONALIDAD":         Gstr_param.nacionalidad = reg("para_valor")
       Case "METODO_PAGO":          Gstr_param.metodo_pago = reg("para_valor")
       Case "REQUER_SERV":          Gstr_param.requer_serv = reg("para_valor")
       Case "COND_CONTRATO":        Gstr_param.cond_contrato = reg("para_valor")
       Case "ESTADO_EQ":            Gstr_param.estado_equip = reg("para_valor")
       Case "ESTADO_EQ_EXPO":       Gstr_param.estado_equip_expo = reg("para_valor")
       Case "INDIC_LLE_VAC":        Gstr_param.indic_llen_vac = reg("para_valor")
       Case "REGIMEN_ADUA":         Gstr_param.regimen_adua = IIf(IsNull(reg("para_valor")), "", reg("para_valor"))
       Case "REGIMEN_ADUA_EXPO":    Gstr_param.regimen_adua_expo = IIf(IsNull(reg("para_valor")), "", reg("para_valor"))
       Case "REGIMEN_ADUA_TRASB":   Gstr_param.regimen_adua_trasb = IIf(IsNull(reg("para_valor")), "", reg("para_valor"))
       Case "TIPO_EQUIPO":          Gstr_param.tipo_equipo = reg("para_valor")
       Case "ALMACEN_TEMPORAL":     Gstr_param.almacen_temporal = reg("para_valor")
    End Select
    reg.MoveNext
Loop
If Gstr_param.regimen_adua_trasb = "" Then Gstr_param.regimen_adua_trasb = "81"
End Sub

' linea_proc es el codigo de la linea A la cual se esta subiendo la informacion puede ser CAE, MSC, KHS
' linea es el codigo de la linea a la que pertenece el archivo puede ser unicamente MSC y KHS
Public Sub carga_parametros_plano(linea As String, linea_proc As String)
Dim sql As String
Dim reg As New ADODB.Recordset


sql = "Select * from plano_parametro where para_linea_proc='" & linea_proc & "' and para_linea ='" & linea & "'"
reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
Do While Not reg.EOF
    Select Case Trim(reg("para_codigo"))
       Case "EMBARCADOR":           Gstr_param_plano.embarcador = reg("para_valor")
       Case "CONSIGNATARIO":        Gstr_param_plano.consignatario = IIf(Not IsNull(reg("para_valor")), reg("para_valor"), "")
       Case "PESO_CONT20":          Gstr_param_plano.PESO_CONT20 = reg("para_valor")
       Case "PESO_CONT40":          Gstr_param_plano.PESO_CONT40 = reg("para_valor")
       Case "DIREC_EMBARCADOR":     Gstr_param_plano.Direc_Embarcador = reg("para_valor")
       
       Case "DIREC_CONSIGNATARIO":  Gstr_param_plano.Direc_consignatario = reg("para_valor")
       Case "DIREC_NOTIFICADOR":    Gstr_param_plano.Direc_notificador = reg("para_valor")
       Case "EMBALAJE":             Gstr_param_plano.embalaje = reg("para_valor")
       Case "IMO":                  Gstr_param_plano.Imo = reg("para_valor")
       Case "NOTIFICADOR":          Gstr_param_plano.notificador = reg("para_valor")
       Case "SELLO":                Gstr_param_plano.Sello = reg("para_valor")
       Case "VOLUMEN20":            Gstr_param_plano.Volumen20 = reg("para_valor")
       Case "VOLUMEN40":            Gstr_param_plano.Volumen40 = reg("para_valor")
       Case "ID_CONSIGNATARIO":     Gstr_param_plano.id_consignatario = reg("para_valor")
       Case "ID_EMBARCADOR":        Gstr_param_plano.id_embarcador = reg("para_valor")
       Case "ID_NOTIFICADOR":       Gstr_param_plano.id_notificador = reg("para_valor")
       Case "INDIC_VACIO":          Gstr_param_plano.Indic_vacio = reg("para_valor")
       Case "TIPO_CARGA_CONT_VACIO": Gstr_param_plano.Tipo_carga_cont_vacio = reg("para_valor")
       Case "DESC_CARGA":           Gstr_param_plano.desc_carga = reg("para_valor")
       Case "PTO_OPT":              Gstr_param_plano.pto_opt = reg("para_valor")
       Case "MARCAS_NUM":           Gstr_param_plano.marcas_num = reg("para_valor")
       Case "COND_CONT":            Gstr_param_plano.cond_cont = reg("para_valor")
    End Select
    reg.MoveNext
Loop
End Sub


Public Function devuelve_desc(tabla As String, campo_cod As String, campo_desc As String, codigo As String, Optional condicion As String) As String
Dim reg As New ADODB.Recordset
    sql = "Select " & campo_desc & " from " & tabla & " Where " & _
           campo_cod & " ='" & codigo & "'"
    If Not IsMissing(condicion) And condicion <> "" Then
        sql = sql & " and " & condicion
    End If
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    If Not reg.EOF Then
       devuelve_desc = Trim$(reg(0))
    End If
End Function


Public Function Elimina_manifiesto(linea As String, buque As String, viaje As String, pto_emb As String, pto_descarga As String) As Boolean
On Error GoTo errores_elimina
     sql = "Delete from docu_bldet " & _
              "Where bldet_linea      ='" & linea & "'" & _
               "   and bldet_buque    ='" & buque & "'" & _
               "   and bldet_viaje    ='" & viaje & "'" & _
               "   and bldet_puer_emb ='" & pto_emb & "'" & _
               "   and bldet_puer_des ='" & pto_descarga & "'" '& _
               "   and bldet_emp_resp ='" & oper_carga & "'"
     Gcon_docu.Execute sql
     
     sql = "Delete from docu_blflete " & _
              "Where flete_linea      ='" & linea & "'" & _
               "   and flete_buque    ='" & buque & "'" & _
               "   and flete_viaje    ='" & viaje & "'" & _
               "   and flete_puer_emb ='" & pto_emb & "'" & _
               "   and flete_puer_des ='" & pto_descarga & "'" '& _
               "   and bldet_emp_resp ='" & oper_carga & "'"
     Gcon_docu.Execute sql
     
     sql = "DELETE FROM docu_blcab " & _
                        "  Where blcab_linea   ='" & linea & "'" & _
                        "   and blcab_buque    ='" & buque & "'" & _
                        "   and blcab_viaje    ='" & viaje & "'" & _
                        "   and blcab_puer_emb ='" & pto_emb & "'" & _
                        "   and blcab_puer_des ='" & pto_descarga & "'" '& _
                        "   and blcab_emp_resp ='" & oper_carga & "'"
     Gcon_docu.Execute sql
     
     sql = "DELETE FROM docu_cont" & _
                        "  Where cont_linea   ='" & linea & "'" & _
                        "   and cont_buque    ='" & buque & "'" & _
                        "   and cont_viaje    ='" & viaje & "'" & _
                        "   and cont_puer_emb ='" & pto_emb & "'" & _
                        "   and cont_puer_des ='" & pto_descarga & "'"
     Gcon_docu.Execute sql
     
     sql = "DELETE FROM docu_manifiesto" & _
                        "  Where mani_linea   ='" & linea & "'" & _
                        "   and mani_buque    ='" & buque & "'" & _
                        "   and mani_viaje    ='" & viaje & "'" & _
                        "   and mani_puer_emb ='" & pto_emb & "'" & _
                        "   and mani_puer_des ='" & pto_descarga & "'"
     Gcon_docu.Execute sql
     
     
     Elimina_manifiesto = True
     Exit Function
errores_elimina:
    MsgBox "Ocurrio el Error " & Err.Description, vbCritical
    Elimina_manifiesto = False
End Function

Public Function Elimina_bl(linea As String, buque As String, viaje As String, pto_emb As String, pto_descarga As String, oper_carga As String, bl As String) As Boolean
On Error GoTo errores_elimina
     sql = "Delete from docu_bldet " & _
              "Where bldet_linea      ='" & linea & "'" & _
               "   and bldet_buque    ='" & buque & "'" & _
               "   and bldet_viaje    ='" & viaje & "'" & _
               "   and bldet_puer_emb ='" & pto_emb & "'" & _
               "   and bldet_puer_des ='" & pto_descarga & "'" & _
               "   and bldet_emp_resp ='" & oper_carga & "'" & _
               "   and bldet_cono_emb ='" & bl & "'"
     Gcon_docu.Execute sql
     
     sql = "DELETE FROM docu_blcab "
     sql = sql & "  Where blcab_linea       ='" & linea & "'" & _
                        "   and blcab_buque    ='" & buque & "'" & _
                        "   and blcab_viaje    ='" & viaje & "'" & _
                        "   and blcab_puer_emb ='" & pto_emb & "'" & _
                        "   and blcab_puer_des ='" & pto_descarga & "'" & _
                        "   and blcab_emp_resp ='" & oper_carga & "'" & _
                        "   and blcab_cono_emb ='" & bl & "'"
     Gcon_docu.Execute sql
     Elimina_bl = True
     Exit Function
errores_elimina:
    MsgBox "Ocurrio el Error " & Err.Description, vbCritical
    Elimina_bl = False
End Function

Public Function Elimina_manifiesto_exp(linea As String, buque As String, viaje As String, pto_emb As String, pto_descarga As String) As Boolean
On Error GoTo errores_elimina
     sql = "Delete from docu_bldet_exp " & _
              "Where bldet_linea      ='" & linea & "'" & _
               "   and bldet_buque    ='" & buque & "'" & _
               "   and bldet_viaje    ='" & viaje & "'" & _
               "   and bldet_puer_emb ='" & pto_emb & "'" & _
               "   and bldet_puer_des ='" & pto_descarga & "'" '& _
               "   and bldet_emp_resp ='" & oper_carga & "'"
     Gcon_docu.Execute sql
     
     sql = "Delete from docu_blflete_exp " & _
              "Where flete_linea      ='" & linea & "'" & _
               "   and flete_buque    ='" & buque & "'" & _
               "   and flete_viaje    ='" & viaje & "'" & _
               "   and flete_puer_emb ='" & pto_emb & "'" & _
               "   and flete_puer_des ='" & pto_descarga & "'" '& _
               "   and bldet_emp_resp ='" & oper_carga & "'"
     Gcon_docu.Execute sql
     
     sql = "DELETE FROM docu_blcab_exp " & _
                        "  Where blcab_linea   ='" & linea & "'" & _
                        "   and blcab_buque    ='" & buque & "'" & _
                        "   and blcab_viaje    ='" & viaje & "'" & _
                        "   and blcab_puer_emb ='" & pto_emb & "'" & _
                        "   and blcab_puer_des ='" & pto_descarga & "'" '& _
                        "   and blcab_emp_resp ='" & oper_carga & "'"
     Gcon_docu.Execute sql
     
     sql = "DELETE FROM docu_cont_exp" & _
                        "  Where cont_linea   ='" & linea & "'" & _
                        "   and cont_buque    ='" & buque & "'" & _
                        "   and cont_viaje    ='" & viaje & "'" & _
                        "   and cont_puer_emb ='" & pto_emb & "'" & _
                        "   and cont_puer_des ='" & pto_descarga & "'"
     Gcon_docu.Execute sql
     
     sql = "DELETE FROM docu_manifiesto_exp" & _
                        "  Where mani_linea   ='" & linea & "'" & _
                        "   and mani_buque    ='" & buque & "'" & _
                        "   and mani_viaje    ='" & viaje & "'" & _
                        "   and mani_puer_emb ='" & pto_emb & "'" & _
                        "   and mani_puer_des ='" & pto_descarga & "'"
     Gcon_docu.Execute sql
     
     
     Elimina_manifiesto_exp = True
     Exit Function
errores_elimina:
    MsgBox "Ocurrio el Error " & Err.Description, vbCritical
    Elimina_manifiesto_exp = False
End Function


Public Function Elimina_bl_EXP(linea As String, buque As String, viaje As String, pto_emb As String, pto_descarga As String, oper_carga As String, bl As String) As Boolean
On Error GoTo errores_elimina
     sql = "Delete from docu_bldet_exp " & _
              "Where bldet_linea      ='" & linea & "'" & _
               "   and bldet_buque    ='" & buque & "'" & _
               "   and bldet_viaje    ='" & viaje & "'" & _
               "   and bldet_puer_emb ='" & pto_emb & "'" & _
               "   and bldet_puer_des ='" & pto_descarga & "'" & _
               "   and bldet_emp_resp ='" & oper_carga & "'" & _
               "   and bldet_cono_emb ='" & bl & "'"
     Gcon_docu.Execute sql
     
     sql = "DELETE FROM docu_blcab_exp "
     sql = sql & "  Where blcab_linea       ='" & linea & "'" & _
                        "   and blcab_buque    ='" & buque & "'" & _
                        "   and blcab_viaje    ='" & viaje & "'" & _
                        "   and blcab_puer_emb ='" & pto_emb & "'" & _
                        "   and blcab_puer_des ='" & pto_descarga & "'" & _
                        "   and blcab_emp_resp ='" & oper_carga & "'" & _
                        "   and blcab_cono_emb ='" & bl & "'"
     Gcon_docu.Execute sql
     Elimina_bl_EXP = True
     Exit Function
errores_elimina:
    MsgBox "Ocurrio el Error " & Err.Description, vbCritical
    Elimina_bl_EXP = False
End Function


Public Sub valida_consist_cont(linea As String, buque As String, viaje As String)
', pto_emb As String, pto_descarga As String)
Dim reg As New ADODB.Recordset

On Error GoTo errores_consistencia
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'TL Validacion de cantidad de contenedores entre el detalle del bl y la tabla
    '   de contenedores
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = "select distinct bldet_contenedor as cont from docu_bldet " & _
          "Where bldet_linea     ='" & linea & "'" & _
          "   and bldet_buque    ='" & buque & "'" & _
          "   and bldet_viaje    ='" & viaje & "'" & _
          "   and bldet_contenedor not in " & _
          "   (Select cont_contenedor from docu_cont " & _
          "   where  cont_linea ='" & linea & "'" & _
          "   and cont_buque    ='" & buque & "'" & _
          "   and cont_viaje    ='" & viaje & "')"
'de docu_bldet
'          "   and bldet_puer_emb ='" & pto_emb & "'" & _
'          "   and bldet_puer_des ='" & pto_descarga & "'" & _
'de docu_cont
'          "   and cont_puer_emb ='" & pto_emb & "'" & _
'          "   and cont_puer_des ='" & pto_descarga & "')"
           
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    If Not reg.EOF Then
       sql = "Los siguientes contenedores CONSTAN EN EL DETALLE DE BLS DEL PUERTO " & _
             "PROCESADO, PERO NO EN LA PANTALLA DE CONTENEDORES, POR FAVOR INGRESARLOS: "
       Do While Not reg.EOF
          sql = Trim$(sql) & reg("cont") & ", "
          reg.MoveNext
       Loop
       MsgBox sql, vbCritical
    End If
    
Exit Sub

errores_consistencia:
    MsgBox "Ocurrio el Error " & Err.Description, vbCritical
    Exit Sub
    Resume 0
End Sub

Public Sub valida_consist_cabecera(linea As String, buque As String, viaje As String)
Dim reg As New ADODB.Recordset

On Error GoTo errores_consistencia
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'TL Validacion de bls con cabecera y sin detalle
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = "select blcab_puer_emb as puer_emb,blcab_puer_des as puer_des,blcab_cono_emb as cono " & _
          "from docu_blcab " & _
          "Where  blcab_linea    ='" & linea & "'" & _
          "   and blcab_buque    ='" & buque & "'" & _
          "   and blcab_viaje    ='" & viaje & "'" & _
          "   and blcab_cono_emb not in " & _
          "   (Select distinct bldet_cono_emb from docu_bldet " & _
          "   where  bldet_linea ='" & linea & "'" & _
          "   and bldet_buque    ='" & buque & "'" & _
          "   and bldet_viaje    ='" & viaje & "')"
 
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    If Not reg.EOF Then
       sql = "Los siguientes BLS BAJARON SIN DETALLE POR FAVOR INGRESARLO Y REVISAR " & _
             "EL ARCHIVO!: "
       Do While Not reg.EOF
          sql = Trim$(sql) & "P.EMB: " & reg("puer_emb") & " - P.DES:" & reg("puer_des") & " - BL:" & reg("cono") & Chr(10) & Chr(13)
          reg.MoveNext
       Loop
       MsgBox sql, vbCritical
    End If
    reg.Close
 
Exit Sub

errores_consistencia:
    MsgBox "Ocurrio el Error " & Err.Description, vbCritical
    Exit Sub
    Resume 0
End Sub

' esta funcion realiza una inspeccion previa de los posibles errores antes
' de realizar el envio de los datos a la aduana
Public Function Valida_datos_CAE(linea As String, buque As String, viaje As String, puerto As String) As Boolean
Dim sql As String
Dim reg As New ADODB.Recordset
    Valida_datos_CAE = True
    sql = "Select * from docu_manifiesto " & _
          " Where mani_linea    ='" & linea & "'" & _
          "   And mani_buque    ='" & buque & "'" & _
          "   And mani_viaje    ='" & viaje & "'" & _
          "   And mani_puer_emb ='" & puerto & "'" & _
          "   And (mani_puer_ini='' or mani_puer_ini IS NULL)"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Do While Not reg.EOF
        MsgBox "El puerto de Inicio de Viaje esta en NULL, posible error en la CAE. Por Favor Revise el Manifiesto" & _
                linea & " - " & buque & " - " & viaje & " - " & puerto, vbInformation
        reg.MoveNext
        Valida_datos_CAE = False
    Loop
    reg.Close
           
    sql = "Select * from docu_manifiesto " & _
          " Where mani_linea    ='" & linea & "'" & _
          "   And mani_buque    ='" & buque & "'" & _
          "   And mani_viaje    ='" & viaje & "'" & _
          "   And mani_puer_emb ='" & puerto & "'" & _
          "   And (mani_emp_prop='' or mani_emp_prop IS NULL)"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Do While Not reg.EOF
        MsgBox "La Empresa Propietaria esta en NULL, posible error en la CAE. Por Favor Revise el Manifiesto" & _
                linea & " - " & buque & " - " & viaje & " - " & puerto, vbInformation
        reg.MoveNext
        Valida_datos_CAE = False
    Loop
    reg.Close
    
    
    sql = "Select * from docu_manifiesto " & _
          " Where mani_linea    ='" & linea & "'" & _
          "   And mani_buque    ='" & buque & "'" & _
          "   And mani_viaje    ='" & viaje & "'" & _
          "   And mani_puer_emb ='" & puerto & "'" & _
          "   And (mani_emp_nave='' or mani_emp_nave IS NULL)"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Do While Not reg.EOF
        MsgBox "La Empresa Naviera esta en NULL, posible error en la CAE. Por Favor Revise el Manifiesto" & _
                linea & " - " & buque & " - " & viaje & " - " & puerto, vbInformation
        reg.MoveNext
        Valida_datos_CAE = False
    Loop
    reg.Close
    
    
    sql = "Select * from docu_manifiesto " & _
          " Where mani_linea    ='" & linea & "'" & _
          "   And mani_buque    ='" & buque & "'" & _
          "   And mani_viaje    ='" & viaje & "'" & _
          "   And mani_puer_emb ='" & puerto & "'" & _
          "   And (mani_adua_ing='' or mani_adua_ing IS NULL or" & _
          "        mani_ano_car ='' or mani_ano_car  IS NULL or" & _
          "        mani_tip_mani='' or mani_tip_mani IS NULL or" & _
          "        mani_mani_car='' or mani_mani_car IS NULL or" & _
          "        mani_digi_ver='' or mani_digi_ver IS NULL )"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Do While Not reg.EOF
        MsgBox "El Número de manifiesto(adua - año - tipo - No. - digito) esta Incompleto o no existe, posible error en la CAE. Por Favor Revise el Manifiesto" & _
                linea & " - " & buque & " - " & viaje & " - " & puerto, vbInformation
        reg.MoveNext
        Valida_datos_CAE = False
    Loop
    reg.Close
    
    ' boletin 113 y 114 de la Aduana
    
    sql = "Select * from docu_blcab " & _
          " Where blcab_linea    ='" & linea & "'" & _
          "   And blcab_buque    ='" & buque & "'" & _
          "   And blcab_viaje    ='" & viaje & "'" & _
          "   And blcab_puer_emb ='" & puerto & "'" & _
          "   And (blcab_regimen_adu='' or blcab_regimen_adu IS NULL ) AND " & _
          "        (blcab_almac_temp ='' or blcab_almac_temp IS NULL )"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Do While Not reg.EOF
        MsgBox "No  existen  los campos Regimen Aduanero ni Almacen Temporal, posible error en la CAE. Por Favor Revise el bl " & _
                linea & " - " & buque & " - " & viaje & " - " & puerto & " - " & reg("blcab_cono_emb"), vbInformation
        reg.MoveNext
        Valida_datos_CAE = False
    Loop
    reg.Close
    
    
    sql = "Select * from docu_cont " & _
          " Where cont_linea    ='" & linea & "'" & _
          "   And cont_buque    ='" & buque & "'" & _
          "   And cont_viaje    ='" & viaje & "'" & _
          "   And cont_puer_emb ='" & puerto & "'" & _
          "   And cont_peso     = 0"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Do While Not reg.EOF
        MsgBox "El Peso esta en CERO, posible error en la CAE. Por Favor Revise el Manifiesto" & _
                linea & " - " & buque & " - " & viaje & " - " & puerto & " - " & reg("cont_contenedor"), vbInformation
        reg.MoveNext
        Valida_datos_CAE = False
    Loop
    reg.Close

End Function


' esta funcion realiza una inspeccion previa de los posibles errores antes
' de realizar el envio de los datos a la aduana
Public Function Valida_datos_CAE_expo(linea As String, buque As String, viaje As String, puerto As String) As Boolean
Dim sql As String
Dim reg As New ADODB.Recordset
    Valida_datos_CAE_expo = True
    sql = "Select * from docu_manifiesto_exp " & _
          " Where mani_linea    ='" & linea & "'" & _
          "   And mani_buque    ='" & buque & "'" & _
          "   And mani_viaje    ='" & viaje & "'" & _
          "   And mani_puer_des ='" & puerto & "'" & _
          "   And (mani_puer_ini='' or mani_puer_ini IS NULL)"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Do While Not reg.EOF
        MsgBox "El puerto de Inicio de Viaje esta en NULL, posible error en la CAE. Por Favor Revise el Manifiesto" & _
                linea & " - " & buque & " - " & viaje & " - " & puerto, vbInformation
        reg.MoveNext
        Valida_datos_CAE_expo = False
    Loop
    reg.Close
           
    sql = "Select * from docu_manifiesto_exp " & _
          " Where mani_linea    ='" & linea & "'" & _
          "   And mani_buque    ='" & buque & "'" & _
          "   And mani_viaje    ='" & viaje & "'" & _
          "   And mani_puer_des ='" & puerto & "'" & _
          "   And (mani_emp_prop='' or mani_emp_prop IS NULL)"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Do While Not reg.EOF
        MsgBox "La Empresa Propietaria esta en NULL, posible error en la CAE. Por Favor Revise el Manifiesto" & _
                linea & " - " & buque & " - " & viaje & " - " & puerto, vbInformation
        reg.MoveNext
        Valida_datos_CAE_expo = False
    Loop
    reg.Close
    
    
    sql = "Select * from docu_manifiesto_exp " & _
          " Where mani_linea    ='" & linea & "'" & _
          "   And mani_buque    ='" & buque & "'" & _
          "   And mani_viaje    ='" & viaje & "'" & _
          "   And mani_puer_des ='" & puerto & "'" & _
          "   And (mani_emp_nave='' or mani_emp_nave IS NULL)"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Do While Not reg.EOF
        MsgBox "La Empresa Naviera esta en NULL, posible error en la CAE. Por Favor Revise el Manifiesto" & _
                linea & " - " & buque & " - " & viaje & " - " & puerto, vbInformation
        reg.MoveNext
        Valida_datos_CAE_expo = False
    Loop
    reg.Close
    
    
    sql = "Select * from docu_manifiesto_exp " & _
          " Where mani_linea    ='" & linea & "'" & _
          "   And mani_buque    ='" & buque & "'" & _
          "   And mani_viaje    ='" & viaje & "'" & _
          "   And mani_puer_des ='" & puerto & "'" & _
          "   And (mani_adua_ing='' or mani_adua_ing IS NULL or" & _
          "        mani_ano_car ='' or mani_ano_car  IS NULL or" & _
          "        mani_tip_mani='' or mani_tip_mani IS NULL or" & _
          "        mani_mani_car='' or mani_mani_car IS NULL or" & _
          "        mani_digi_ver='' or mani_digi_ver IS NULL )"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Do While Not reg.EOF
        MsgBox "El Número de manifiesto(adua - año - tipo - No. - digito)para esta Exportación esta Incompleto o no existe, posible error en la CAE. Por Favor Revise el Manifiesto" & _
                linea & " - " & buque & " - " & viaje & " - " & puerto, vbInformation
        reg.MoveNext
        Valida_datos_CAE_expo = False
    Loop
    reg.Close
    
        ' boletin 113 y 114 de la Aduana
    
    sql = "Select * from docu_blcab_exp " & _
          " Where blcab_linea    ='" & linea & "'" & _
          "   And blcab_buque    ='" & buque & "'" & _
          "   And blcab_viaje    ='" & viaje & "'" & _
          "   And blcab_puer_emb ='" & puerto & "'" & _
          "   And (blcab_regimen_adu='' or blcab_regimen_adu IS NULL ) AND " & _
          "        (blcab_almac_temp ='' or blcab_almac_temp IS NULL )"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Do While Not reg.EOF
        MsgBox "No  existen  los campos Regimen Aduanero ni Almacen Temporal, posible error en la CAE. Por Favor Revise el bl " & _
                linea & " - " & buque & " - " & viaje & " - " & puerto & " - " & reg("blcab_cono_emb"), vbInformation
        reg.MoveNext
        Valida_datos_CAE_expo = False
    Loop
    reg.Close

    
    sql = "Select * from docu_cont_exp " & _
          " Where cont_linea    ='" & linea & "'" & _
          "   And cont_buque    ='" & buque & "'" & _
          "   And cont_viaje    ='" & viaje & "'" & _
          "   And cont_puer_des ='" & puerto & "'" & _
          "   And cont_peso     = 0"
    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
    Do While Not reg.EOF
        MsgBox "El Peso esta en CERO, posible error en la CAE. Por Favor Revise el Manifiesto" & _
                linea & " - " & buque & " - " & viaje & " - " & puerto & " - " & reg("cont_contenedor"), vbInformation
        reg.MoveNext
        Valida_datos_CAE_expo = False
    Loop
    reg.Close

End Function


Public Sub Carga_Menus(user As String)
Dim reg As New ADODB.Recordset

'    sql = "SELECT * FROM docu_menu WHERE mnu_usuario='" & Trim(user) & "'"
'    reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
     
'    Do While Not reg.EOF
        For i% = 0 To Doc_menu.Count - 1
            If TypeOf Doc_menu.Controls(i%) Is Menu Then
                opcion$ = Trim(Doc_menu.Controls(i%).Name)
                'If Trim(reg("mnu_menu")) = opcion$ Then
                        Doc_menu.Controls(i%).Enabled = True
                        Doc_menu.Controls(i%).Visible = True
                '   Exit For
                'End If
                'Debug.Print Doc_menu.Controls(i%).Name
            End If
        Next i%
'        reg.MoveNext
'    Loop
'    reg.Close
    
    
    'Doc_menu.migra.Visible = False
  '  Doc_menu.mnu_repo_flete_impo.Visible = False
  '  Doc_menu.mnu_repo_flete_Expo.Visible = False
    
End Sub

'para poner el fuente por rango
Public Sub Excel_setea_fuente_rango(ByRef hoja As Excel.Worksheet, ByRef rango As Excel.Range, fuente As String, negrita As Boolean, italica As Boolean, tamano As Integer, Optional color As Double)

End Sub

'para poner el fuente por FILA
Public Sub Excel_setea_fuente_fila(ByRef hoja As Excel.Worksheet, fila As Integer, tamano As Integer, Optional fuente As String, Optional negrita As Boolean, Optional italica As Boolean, Optional color As Double)
    With hoja.Rows(fila).Font
        .Size = tamano
        If Not IsMissing(fuente) And fuente <> "" Then .Name = fuente
        If Not IsMissing(negrita) Then .Bold = negrita
        If Not IsMissing(italica) Then .Italic = italica
        If Not IsMissing(color) And color <> 0 Then .color = color
    End With
End Sub

'para poner el borde a un rango
Public Sub Excel_setea_borde(ByRef hoja As Excel.Worksheet, fila1 As Integer, columna1 As Integer, fila2 As Integer, columna2 As Integer, Estilo_Linea As Integer, lineas_internas As Boolean)
Dim rng_est As Excel.Range
    
    Set rng_est = hoja.Range(hoja.Cells(fila1, columna1), hoja.Cells(fila2, columna2))
    rng_est.Borders.LineStyle = Estilo_Linea
    If Not lineas_internas Then
        rng_est.Borders(xlInsideVertical).LineStyle = xlNone
        rng_est.Borders(xlInsideHorizontal).LineStyle = xlNone
    End If
    
End Sub
' este procedimiento setea el ancho de una columna en base a un rango de celdas y el valor del ancho
Public Sub Excel_setea_ancho_columna(ByRef hoja As Excel.Worksheet, fila1 As Integer, columna1 As Integer, fila2 As Integer, columna2 As Integer, ancho As Integer)
Dim rng_est As Excel.Range
    
    Set rng_est = hoja.Range(hoja.Cells(fila1, columna1), hoja.Cells(fila2, columna2))
    rng_est.ColumnWidth = ancho
    
End Sub

' este procedimiento genera en formato excel reportes con los registros correspondientes a las
' tablas maestras del sistema, el mismo que recibe como parametro
' 10 mayo /2005
Public Sub Excel_repo_mant(tabla As String)
On Error GoTo Errores
Dim ApExcel As Variant
    
    Set ApExcel = CreateObject("Excel.application")
     ApExcel.Visible = True
     ApExcel.Workbooks.Add
     ApExcel.Worksheets(1).Activate
     ApExcel.Worksheets(1).Cells.ClearContents ' limpia todas las celdas
     
     ApExcel.Worksheets(1).Name = tabla
     ApExcel.Range("A1:AM1").Font.Size = 10
     ApExcel.Range("A1:AM1").Font.Bold = True
Select Case LCase(tabla)
Case "almacen":         cadenasql = "Select * from docu_almacen"
Case "buque":           cadenasql = "Select * from docu_buque"
Case "cond_contrato":   cadenasql = "Select * from docu_cond_contrato"
Case "cond_contenedor": cadenasql = "Select * from docu_cond_contenedor"
Case "consolidadora":   cadenasql = "Select * from docu_consolidadora"
Case "consignatario":   cadenasql = "Select * from docu_consignatario"

Case "deposito":        cadenasql = "Select * from docu_deposito"
Case "embalaje":        cadenasql = "Select * from docu_embalaje"
Case "emp_anav":        cadenasql = "Select * from docu_emp_anav"
Case "emp_resp":        cadenasql = "Select * from docu_emp_resp"
Case "estado_equipo":   cadenasql = "Select * from docu_estado_equipo"
Case "imo":             cadenasql = "Select * from docu_imo"
Case "imo_onu":             cadenasql = "Select * from docu_imo_onu"
Case "indic_lleno_vacio":   cadenasql = "Select * from docu_indicad_llen_vac"
Case "linea":               cadenasql = "Select * from docu_linea_nav"
Case "metodo_pago":         cadenasql = "Select * from docu_metodo_pago"
Case "moneda":              cadenasql = "Select * from docu_moneda"
Case "nacionalidad":        cadenasql = "Select * from docu_nacionalidad"
Case "natura_carga":        cadenasql = "Select * from docu_natura_carga"
Case "prov_equipo":         cadenasql = "Select * from docu_proveedor_equip"
Case "puerto":              cadenasql = "Select * from docu_puerto"
Case "regimen":             cadenasql = "Select * from docu_regimen"
Case "req_serv_transp":     cadenasql = "Select * from docu_requer_serv_trans"
Case "tamano_equipo":       cadenasql = "Select * from docu_tamano_equip"
Case "tipo_bien":           cadenasql = "Select * from docu_tipo_bien"
Case "tipo_carga":          cadenasql = "Select * from docu_tipo_carga"
Case "tipo_deposito":       cadenasql = "Select * from docu_tipo_deposito"
Case "tipo_documento":      cadenasql = "Select * from docu_tipo_documento"
Case "tipo_envio":          cadenasql = "Select * from docu_tipo_envio"
Case "tipo_equipo":         cadenasql = "Select * from docu_tipo_equip"
Case "tipo_salida":         cadenasql = "Select * from docu_tipo_salida"
End Select
With ApExcel.Worksheets(1).QueryTables.Add("ODBC;DSN=" & Gstr_DSN & ";UID=" & Gstr_user_conex & "PWD=" & Gstr_psw & ";APP=Microsoft® Query;DATABASE=" & Gstr_base_doc & ";Trusted_Connection=Yes", ApExcel.Range("A1:A1"), cadenasql)
    .Refresh
End With
Do While ApExcel.Range("B1").Value = ""
    DoEvents
Loop

Exit Sub
Errores:
MsgBox "Ocurrio el Error " & Err.Description
Exit Sub
Resume 0
End Sub

'funcion : Existe Bl
'          Esta funcion permite verificar si el bl que se recibe como parametro existe o no
'          dentro de la base de datos de documentos.
'          Se creo principalmente para la creacion de bl desde los archivos planos.
'parametro:ImpoExpo     .- puede tomar los valores de "I" / "E", para indicar si se trata de infirmacion de importacion o exportacion
'          Linea        .- indica la linea de la cual se esta buscando el bl
'          Buque        .- indica el buque del cual se esta buscando el bl
'          Viaje        .- indica el Viaje del cual se esta buscando el bl
'          pto_emb      .-indica el Puerto de Embarque del cual se esta buscando el bl
'          pto_descarga .-indica el Puerto de Descarga del cual se esta buscando el bl
'          Oper_carga   .-indica el Operador de Carga del cual se esta buscando el bl
'          bl           .-indica el bl que se esta buscando
Public Function Existe_BL(ImpoExpo As String, linea As String, buque As String, viaje As String, pto_emb As String, pto_descarga As String, oper_carga As String, bl As String) As Boolean
Dim sql As String
Dim reg As New ADODB.Recordset
On Error GoTo errores_existe
        sql = "Select * from "
        If ImpoExpo = "I" Then sql = sql & " docu_blcab " Else sql = sql & " docu_blcab_exp "
        sql = sql & " " & _
                        " Where blcab_linea    ='" & linea & "'" & _
                        "   and blcab_buque    ='" & buque & "'" & _
                        "   and blcab_viaje    ='" & viaje & "'" & _
                        "   and blcab_puer_emb ='" & pto_emb & "'" & _
                        "   and blcab_puer_des ='" & pto_descarga & "'" & _
                        "   and blcab_emp_resp ='" & oper_carga & "'" & _
                        "   and blcab_cono_emb ='" & bl & "'"
        reg.Open sql, Gcon_docu
        If reg.EOF Then Existe_BL = False Else Existe_BL = True
        
        

Exit Function
errores_existe:
    MsgBox "Ocurrrio el error " & Err.Description
    Exit Function
    Resume 0
End Function

'funcion : Existe Contenedor
'          Esta funcion permite verificar si el contenedor que se recibe como parametro existe o no
'          dentro de la base de datos de documentos.
'          Se creo principalmente para la creacion de bl desde los archivos planos.
'parametro:ImpoExpo     .- puede tomar los valores de "I" / "E", para indicar si se trata de infirmacion de importacion o exportacion
'          Linea        .- indica la linea de la cual se esta buscando el contenedor
'          Buque        .- indica el buque del cual se esta buscando el contenedor
'          Viaje        .- indica el Viaje del cual se esta buscando el contenedor
'          pto_emb      .-indica el Puerto de Embarque del cual se esta buscando el contenedor
'          pto_descarga .-indica el Puerto de Descarga del cual se esta buscando el contenedor
'          contenedor   .-indica el contenedor que se esta buscando
Public Function Existe_Contenedor(ImpoExpo As String, linea As String, buque As String, viaje As String, pto_emb As String, pto_descarga As String, contenedor As String) As Boolean
Dim sql As String
Dim reg As New ADODB.Recordset
On Error GoTo errores_existe
        sql = "Select * from "
        If ImpoExpo = "I" Then sql = sql & " docu_cont " Else sql = sql & " docu_cont_exp "
        sql = sql & "Where cont_linea     ='" & linea & _
                    "' and cont_buque     = '" & buque & _
                    "' and cont_viaje     = '" & viaje & _
                    "' and cont_puer_emb  = '" & pto_emb & _
                    "' and cont_puer_des  = '" & pto_descarga & _
                    "' and cont_contenedor= '" & contenedor & "'"

        reg.Open sql, Gcon_docu
        If reg.EOF Then Existe_Contenedor = False Else Existe_Contenedor = True
        
Exit Function
errores_existe:
    MsgBox "Ocurrrio el error " & Err.Description
    Exit Function
    Resume 0
End Function

'funcion : Existe Manifiesto
'          Esta funcion permite verificar si el manifiesto que se recibe como parametro existe o no
'          dentro de la base de datos de documentos.
'          Se creo principalmente para la creacion de bl desde los archivos planos.
'parametro:ImpoExpo     .- puede tomar los valores de "I" / "E", para indicar si se trata de infirmacion de importacion o exportacion
'          Linea        .- indica la linea de la cual se esta buscando el manifiesto
'          Buque        .- indica el buque del cual se esta buscando el manifiesto
'          Viaje        .- indica el Viaje del cual se esta buscando el manifiesto
'          pto_emb      .-indica el Puerto de Embarque del cual se esta buscando el manifiesto
'          pto_descarga .-indica el Puerto de Descarga del cual se esta buscando el manifiesto
Public Function Existe_Manifiesto(ImpoExpo As String, linea As String, buque As String, viaje As String, pto_emb As String, pto_descarga As String) As Boolean
Dim sql As String
Dim reg As New ADODB.Recordset
On Error GoTo errores_existe
        sql = "Select * from "
        If ImpoExpo = "I" Then sql = sql & "docu_manifiesto " Else sql = sql & "docu_manifiesto_exp "
        sql = sql & " Where mani_linea    ='" & linea & "'" & _
                    "   and mani_buque    ='" & buque & "'" & _
                    "   and mani_viaje    ='" & viaje & "'" & _
                    "   and mani_puer_emb ='" & pto_emb & "'" & _
                    "   and mani_puer_des ='" & pto_descarga & "'"

        reg.Open sql, Gcon_docu
        If reg.EOF Then Existe_Manifiesto = False Else Existe_Manifiesto = True
        
Exit Function
errores_existe:
    MsgBox "Ocurrrio el error " & Err.Description
    Exit Function
    Resume 0
End Function

'funcion : Existe Bl_det
'          Esta funcion permite verificar si el detalle del bl que se recibe como parametro existe o no
'          dentro de la base de datos de documentos.
'          Se creo principalmente para la creacion de bl desde los archivos planos.
'parametro:ImpoExpo     .- puede tomar los valores de "I" / "E", para indicar si se trata de infirmacion de importacion o exportacion
'          Linea        .- indica la linea de la cual se esta buscando el bl
'          Buque        .- indica el buque del cual se esta buscando el bl
'          Viaje        .- indica el Viaje del cual se esta buscando el bl
'          pto_emb      .-indica el Puerto de Embarque del cual se esta buscando el bl
'          pto_descarga .-indica el Puerto de Descarga del cual se esta buscando el bl
'          Oper_carga   .-indica el Operador de Carga del cual se esta buscando el bl
'          bl           .-indica el bl que se esta buscando
'          contenedor   .-indica el contenedor que se esta buscando
Public Function Existe_BL_det(ImpoExpo As String, linea As String, buque As String, viaje As String, pto_emb As String, pto_descarga As String, oper_carga As String, bl As String, cont As String) As Boolean
Dim sql As String
Dim reg As New ADODB.Recordset
On Error GoTo errores_existe
        sql = "Select * from "
        If ImpoExpo = "I" Then sql = sql & " docu_bldet " Else sql = sql & " docu_bldet_exp "
        sql = sql & " " & _
                        " Where bldet_linea    ='" & linea & "'" & _
                        "   and bldet_buque    ='" & buque & "'" & _
                        "   and bldet_viaje    ='" & viaje & "'" & _
                        "   and bldet_puer_emb ='" & pto_emb & "'" & _
                        "   and bldet_puer_des ='" & pto_descarga & "'" & _
                        "   and bldet_emp_resp ='" & oper_carga & "'" & _
                        "   and bldet_cono_emb ='" & bl & "'" & _
                        "   and bldet_contenedor='" & cont & "'"
        reg.Open sql, Gcon_docu
        If reg.EOF Then Existe_BL_det = False Else Existe_BL_det = True
        
        

Exit Function
errores_existe:
    MsgBox "Ocurrrio el error " & Err.Description
    Exit Function
    Resume 0
End Function

'funcion : Max_Secuencia_BL_det
'          Esta funcion devuelve la siguiente secuencia disponible del bl que se recibe como parametro
'          Se creo principalmente para la creacion de bl desde los archivos planos.
'parametro:ImpoExpo     .- puede tomar los valores de "I" / "E", para indicar si se trata de infirmacion de importacion o exportacion
'          Linea        .- indica la linea de la cual se esta buscando el bl
'          Buque        .- indica el buque del cual se esta buscando el bl
'          Viaje        .- indica el Viaje del cual se esta buscando el bl
'          pto_emb      .-indica el Puerto de Embarque del cual se esta buscando el bl
'          pto_descarga .-indica el Puerto de Descarga del cual se esta buscando el bl
'          Oper_carga   .-indica el Operador de Carga del cual se esta buscando el bl
'          bl           .-indica el bl que se esta buscando
Public Function Max_Secuencia_BL_det(ImpoExpo As String, linea As String, buque As String, viaje As String, pto_emb As String, pto_descarga As String, oper_carga As String, bl As String) As Integer
Dim sql As String
Dim reg As New ADODB.Recordset

   If ImpoExpo = "I" Then
        sql = "Select max(bldet_secuencia) from docu_bldet "
   ElseIf ImpoExpo = "E" Then
        sql = "Select max(bldet_secuencia) from docu_bldet_exp "
   End If
   sql = sql & " where bldet_linea    ='" & linea & "'" & _
         "   and bldet_buque    ='" & buque & "'" & _
         "   and bldet_viaje    ='" & viaje & "'" & _
         "   and bldet_puer_emb ='" & pto_emb & "'" & _
         "   and bldet_puer_des ='" & pto_descarga & "'" & _
         "   and bldet_emp_resp ='" & oper_carga & "'" & _
         "   and bldet_cono_emb ='" & bl & "'"
   reg.Open sql, Gcon_docu, adOpenDynamic, adLockOptimistic
   If reg.EOF Then
      Max_Secuencia_BL_det = 1
   Else
      If Not IsNull(reg(0)) Then
         Max_Secuencia_BL_det = reg(0) + 1
      Else
         Max_Secuencia_BL_det = 1
      End If
   End If
   
Exit Function
errores_existe:
    MsgBox "Ocurrrio el error " & Err.Description
    Exit Function
    Resume 0
End Function


Public Sub cambio_base(nom_base As String)

    If MsgBox("Para realizar el cambio de la conexión a una base histórica, debe cerrar antes todas las opciones!, Desea Continuar?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Gcon_docu.Close
        Gstr_base_doc = nom_base
        Call Open_Base_doc
        'Doc_menu.StBar.Panels(4).Text = "Versión: " & App.Major & "." & App.Minor
        Doc_menu.StBar.Panels(5).Text = "Base: " & Gstr_base_doc
        'Doc_menu.StBar.Panels(6).Text = Gstr_user
    End If

End Sub
