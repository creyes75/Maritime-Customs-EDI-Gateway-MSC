VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form Frm_copia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manifiestos"
   ClientHeight    =   2760
   ClientLeft      =   4230
   ClientTop       =   2640
   ClientWidth     =   5385
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5385
   Begin VB.Frame Frame2 
      Height          =   600
      Left            =   90
      TabIndex        =   5
      Top             =   2100
      Width           =   4020
      Begin VB.TextBox Txt_bl 
         Height          =   315
         Left            =   1950
         MaxLength       =   25
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   180
         Width           =   1920
      End
      Begin VB.Label Label1 
         Caption         =   "Código de BL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   300
         TabIndex        =   6
         Top             =   255
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Copia de BL "
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
      Height          =   2040
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   5250
      Begin MSGrid.Grid Grd_copia 
         Height          =   1620
         Left            =   135
         TabIndex        =   1
         Top             =   300
         Width           =   5010
         _Version        =   65536
         _ExtentX        =   8837
         _ExtentY        =   2857
         _StockProps     =   77
         BackColor       =   16777215
         Cols            =   4
         FixedCols       =   0
      End
   End
   Begin Threed.SSCommand btn_sgte 
      Height          =   495
      Left            =   4260
      TabIndex        =   2
      ToolTipText     =   "Continuar"
      Top             =   2205
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   78
      Picture         =   "Mie_copia.frx":0000
   End
   Begin Threed.SSCommand btn_elim 
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   2205
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
      Picture         =   "Mie_copia.frx":0452
   End
End
Attribute VB_Name = "Frm_copia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_elim_Click()
  P_Salir = True
  Unload Me
End Sub

'Procedimiento:  Copia de un Bl a otro Puerto o al mismo con otro número de BL
'Autor        :  Carlos Reyes Atiencia
'Fecha        :  Sábado 20 y Domingo 21 de Junio de 1998
'Proceso      :  Validar que el anio,registro,puerto,bl no exista ya en la tabla
'                BLIMPO , tampoco en la BLIMPODT.
'                copiar el Bl en el Anio, Registro, Puerto seleccionado con el mismo
'                o distinto Número de BL
'===================================================================================
Private Sub btn_sgte_Click()
On Error GoTo errores
Dim sql As String
Dim Anio, registro, puerto, bl As String
Dim reg, reg2, reg3 As Recordset            ' reg valida que la informacion a copiar no se encuentre ya en las tablas
                                            ' reg2 selecciona lo que se va a copiar
Dim blimpo, blimpodt, impocont As Recordset ' variables de las tablas

Grd_copia.col = 0: Anio = Grd_copia.Text
Grd_copia.col = 1: registro = Grd_copia.Text
Grd_copia.col = 3: puerto = Grd_copia.Text

bl = Trim$(txt_bl.Text)
' para validar si existe ya en la blimpo
sql = "SELECT * FROM BLIMPO " & _
      " WHERE ANIO     ='" & Anio & "'" & _
      "   and registro = '" & registro & "'" & _
      "   and puerto   = '" & puerto & "'" & _
      "   and cod_bl   = '" & bl & "'" & _
      "   and pto_desc = '" & P_Pto_Descarga.codigo & "'"
Set reg = db_manifi.OpenRecordset(sql)

If reg.RecordCount = 0 Then
   Set blimpo = db_manifi.OpenRecordset("Blimpo")
   Set blimpodt = db_manifi.OpenRecordset("Blimpodt")
   Set impocont = db_manifi.OpenRecordset("Impocont")
   '***********************************************************************
   ' llena la blimpo
   '***********************************************************************
   sql = "SELECT * FROM BLIMPO " & _
         " WHERE ANIO     = '" & MAN_IMPO_Bls.t_anio & "'" & _
         "   and registro = '" & MAN_IMPO_Bls.t_reg & "'" & _
         "   and puerto   = '" & MAN_IMPO_Bls.Cmb_pto & "'" & _
         "   and cod_bl   = '" & MAN_IMPO_Bls.txt_bl & "'" & _
         "   and pto_desc = '" & P_Pto_Descarga.codigo & "'"
   Set reg2 = db_manifi.OpenRecordset(sql)
   blimpo.AddNew
   With blimpo
      !Anio = Anio
      !pto_desc = reg2!pto_desc
      !registro = registro
      !puerto = puerto
      !cod_bl = bl
      !proposito = reg2!proposito
      !tipo_id = reg2!tipo_id
      !impexp = reg2!impexp
      !consignatario = reg2!consignatario
      If IsNull(reg2!direc_imp) Then
         !direc_imp = " "
      Else
         !direc_imp = reg2!direc_imp
      End If
      !fecha_embarque = reg2!fecha_embarque
      If IsNull(reg2!nav_cons_dui) Then
         !nav_cons_dui = " "
      Else
         !nav_cons_dui = reg2!nav_cons_dui
      End If
      If IsNull(reg2!consolidadora) Then
         !consolidadora = " "
      Else
         !consolidadora = reg2!consolidadora
      End If
      If IsNull(reg2!cod_bl_padre) Then
         !cod_bl_padre = " "
      Else
         !cod_bl_padre = reg2!cod_bl_padre
      End If
      !embarcador = reg2!embarcador
      If IsNull(reg2!direc_exp) Then
        !direc_exp = " "
      Else
        !direc_exp = reg2!direc_exp
      End If
      !id_notyfier = IIf(Not IsNull(reg2!id_notyfier), reg2!id_notyfier, "")
      !no_notyfier = reg2!no_notyfier
      !direc_notf = reg2!direc_notf
      If IsNull(reg2!desc_carga) Then
         !desc_carga = " "
      Else
         !desc_carga = reg2!desc_carga
      End If
      !moneda = reg2!moneda
      !costo_flete = reg2!costo_flete
      !no_bultos = reg2!no_bultos
      !no_cont20 = reg2!no_cont20
      !no_cont40 = reg2!no_cont40
      !pto_origen = reg2!pto_origen
      !pto_destino = reg2!pto_destino
      !pto_descarga = reg2!pto_descarga
      !linea_nav = IIf(Not IsNull(reg2!linea_nav), reg2!linea_nav, "9993")
      If IsNull(reg2!comentario) Then
         !comentario = " "
      Else
         !comentario = reg2!comentario
      End If
      !peso_kls = reg2!peso_kls
      !volumen_mt3 = reg2!volumen_mt3
      !st_padre = reg2!st_padre
      If IsNull(reg2!Status) Then
         !Status = " "
      Else
         !Status = reg2!Status
      End If
      !FECHA_ING = reg2!FECHA_ING
      !USER_ING = reg2!USER_ING
      If IsNull(reg2!fecha_mod) Then
         '!fecha_mod = " "
         !USER_mod = " "
      Else
         !fecha_mod = reg2!fecha_mod
         !USER_mod = reg2!USER_mod
      End If
   End With
   blimpo.Update
   reg2.Close
   blimpo.Close
'**********************************************************************
'   llena blimpodt
'***********************************************************************
   sql = "SELECT * FROM BLIMPODT WHERE ANIO ='" & MAN_IMPO_Bls.t_anio & "'"
   sql = sql + " and registro = '" & MAN_IMPO_Bls.t_reg & "'"
   sql = sql + " and puerto = '" & MAN_IMPO_Bls.Cmb_pto & "'"
   sql = sql + " and cod_bl = '" & MAN_IMPO_Bls.txt_bl & "'"
   sql = sql + " and pto_desc = '" & P_Pto_Descarga.codigo & "'"
   Set reg2 = db_manifi.OpenRecordset(sql)
   Do While Not reg2.EOF
        blimpodt.AddNew
        With blimpodt
           !Anio = Anio:                          !registro = registro
           !pto_desc = reg2!pto_desc
           !puerto = puerto:                      !cod_bl = bl
           !sec_carga = reg2!sec_carga:           !tipo_carga = reg2!tipo_carga
           !carga = reg2!carga
           If IsNull(reg2!danger) Then
              !danger = " "
           Else
              !danger = reg2!danger
           End If
           !embalaje = reg2!embalaje:             !no_bultos = reg2!no_bultos
           If IsNull(reg2!marca_num) Then
              !marca_num = " "
           Else
              !marca_num = reg2!marca_num
           End If
           If IsNull(reg2!conteiner) Then
              !conteiner = " "
           Else
              !conteiner = reg2!conteiner
           End If
           !med_peso = reg2!med_peso:           !peso = reg2!peso
           !med_volumen = reg2!med_volumen:     !volumen = reg2!volumen
           !tipo_almacena = reg2!tipo_almacena
           If IsNull(reg2!bodega) Then
               !bodega = " "
           Else
              !bodega = reg2!bodega
           End If
           If IsNull(reg2!almacen) Then
              !almacen = " "
           Else
              !almacen = reg2!almacen
           End If
           If IsNull(reg2!deposito) Then
               !deposito = " "
           Else
               !deposito = reg2!deposito
           End If
           If IsNull(reg2!st_shipconv) Then
             !st_shipconv = " "
           Else
              !st_shipconv = reg2!st_shipconv
           End If
           If IsNull(reg2!st_prorrateo) Then
             !st_prorrateo = " "
           Else
              !st_prorrateo = reg2!st_prorrateo
           End If
           If IsNull(reg2!desc_carga) Then
              !desc_carga = " "
           Else
              !desc_carga = reg2!desc_carga
           End If
        End With
        blimpodt.Update
   reg2.MoveNext
   Loop
   reg2.Close
   blimpodt.Close
'**********************************************************************
'   llena impocont
'***********************************************************************
   sql = "SELECT * FROM IMPOCONT WHERE ANIO ='" & MAN_IMPO_Bls.t_anio & "'"
   sql = sql + " and registro = '" & MAN_IMPO_Bls.t_reg & "'"
   sql = sql + " and pto_desc = '" & P_Pto_Descarga.codigo & "'"
   Set reg2 = db_manifi.OpenRecordset(sql)
   Do While Not reg2.EOF
        ' para validar si existe ya en la impocont
        sql = "SELECT * FROM IMPOCONT WHERE ANIO ='" & Anio & "'"
        sql = sql + " and registro = '" & registro & "'"
        sql = sql + " and conteiner = '" & reg2!conteiner & "'"
        sql = sql + " and pto_desc = '" & P_Pto_Descarga.codigo & "'"
        Set reg3 = db_manifi.OpenRecordset(sql)
        If reg3.RecordCount <= 0 Then
            impocont.AddNew
            With impocont
                !Anio = Anio
                !registro = registro
                !pto_desc = reg2!pto_desc
                !conteiner = reg2!conteiner
                !tipo_contein = reg2!tipo_contein
                !cond_contein = reg2!cond_contein
                !bls_asociados = reg2!bls_asociados
                !tara = reg2!tara
                !sello1 = reg2!sello1
                If IsNull(reg2!sello2) Then
                    !sello2 = " "
                Else
                    !sello2 = reg2!sello2
                End If
                If IsNull(reg2!sello3) Then
                    !sello3 = " "
                Else
                    !sello3 = reg2!sello3
                End If
                If IsNull(reg2!sello4) Then
                    !sello4 = " "
                Else
                    !sello4 = reg2!sello4
                End If
                !FECHA_ING = reg2!FECHA_ING
                !USER_ING = reg2!USER_ING
                If IsNull(reg2!fecha_mod) Then
                    !USER_mod = " "
                Else
                    !fecha_mod = reg2!fecha_mod
                    !USER_mod = reg2!USER_mod
                End If
                
            End With
            impocont.Update
        End If
        reg3.Close
        reg2.MoveNext
    Loop
   reg2.Close
   impocont.Close
   MsgBox "La copia se Realizó con Exito. Por Seguridad Revise sus Datos. ", vbInformation, "Manifiestos"
   Unload Me
Else
  MsgBox "La copia no se pudo realizar, El bl ya existe en ese Puerto", vbInformation, "Manifiestos"
  Exit Sub
End If
Exit Sub
errores:
  MsgBox " Ha ocurrido un Error" & Err.Description, vbCritical, "manifiestos"
  Exit Sub
  Resume 0
End Sub

Private Sub Form_Load()
Dim sql As String
Dim i As Integer
Dim registro As Recordset
'****************************************
Grd_copia.Cols = 4
Grd_copia.Row = 0
Grd_copia.col = 0
Grd_copia.Text = "Año"
Grd_copia.ColWidth(0) = 600

Grd_copia.col = 1
Grd_copia.Text = "Registro"
Grd_copia.ColWidth(1) = 1000

Grd_copia.col = 2
Grd_copia.Text = "Viaje"
Grd_copia.ColWidth(2) = 2000

Grd_copia.col = 3
Grd_copia.Text = "Puerto"
Grd_copia.ColWidth(3) = 2000

txt_bl.Text = MAN_IMPO_Bls.txt_bl

sql = "Select a.anio, a.registro, d.viaje, a.puerto, c.descripcion "
sql = sql + "from registro d, manifiesto a, puerto c "
sql = sql + "where a.anio = d.anio and "
sql = sql + "a.registro = d.registro and "
sql = sql + "a.puerto = c.puerto "

Set registro = db_manifi.OpenRecordset(sql)
i = 1
If registro.RecordCount > 0 Then
   Do While Not registro.EOF
       Grd_copia.Row = i
       Grd_copia.col = 0
       Grd_copia.Text = registro!Anio
       Grd_copia.col = 1
       Grd_copia.Text = registro!registro
       Grd_copia.col = 2
       Grd_copia.Text = registro!viaje
       Grd_copia.col = 3
       Grd_copia.Text = registro!puerto
       Grd_copia.Rows = Grd_copia.Rows + 1
       registro.MoveNext
       i = i + 1
   Loop
Else
   MsgBox "No existen Registros ", vbInformation, "Manifiestos"
   Exit Sub
End If
End Sub


