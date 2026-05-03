VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form docu_ext_actualiza 
   Caption         =   "Actualiza Catálogos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Opt_embalaje 
      Caption         =   "Embalaje"
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   2040
      Width           =   1575
   End
   Begin VB.OptionButton Opt_producto 
      Caption         =   "Productos"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Readme"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   2760
      Width           =   855
   End
   Begin VB.OptionButton Opt_deposito 
      Caption         =   "Depositos"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.OptionButton Opt_Almac 
      Caption         =   "Almacenes"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Cmd_OK 
      Caption         =   "Procersar"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox Txt_DTX 
      Height          =   285
      Left            =   1545
      TabIndex        =   1
      Top             =   240
      Width           =   2715
   End
   Begin VB.CommandButton Cmd_DTX 
      Caption         =   "..."
      Height          =   285
      Left            =   4320
      TabIndex        =   0
      ToolTipText     =   "Seleccionar Archivo a Procesar"
      Top             =   240
      Width           =   285
   End
   Begin MSComDlg.CommonDialog Dlg_DTX 
      Left            =   4560
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Abrir Archivo DTX"
   End
   Begin VB.Label Label6 
      Caption         =   "Archivo EXCEL :"
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
      Height          =   180
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   1365
   End
End
Attribute VB_Name = "docu_ext_actualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_DTX_Click()
   Dlg_DTX.Filter = "Todos los Archivos (*.*)|*.*|Archivos Excel (*.xls)|*.xls"
   Dlg_DTX.ShowOpen
   Txt_DTX.Text = Dlg_DTX.FileName
End Sub

Private Sub Command1_Click()
MsgBox "Importante: " & Chr(13) & " El Archivo a Procesar debe ser EXCEL (.xls), y los datos deben estar sin titulos, ni encabezados y con todas las columnas que tenga la tabla asociada."
End Sub

Private Sub Cmd_OK_Click()
Dim mensaje As String
Dim ApExcel As Variant
Dim reg As New ADODB.Recordset
Dim opc As String

If Opt_Almac.Value Then opc = "ALMACEN"
If Opt_deposito.Value Then opc = "DEPOSITO"
If Opt_producto.Value Then opc = "PRODUCTO"
If Opt_embalaje.Value Then opc = "EMBALAJE"

'***************   abre el excel/setea titulo  *****************
     Set ApExcel = CreateObject("Excel.application")
     ApExcel.Visible = True
     ApExcel.Workbooks.Open (Me.Txt_DTX.Text)   'usa un archivo de excel existente
     'ApExcel.Workbooks.Add
     ApExcel.Worksheets(1).Activate
     'ApExcel.Worksheets(1).Cells.ClearContents ' limpia todas las celdas
     'ApExcel.Worksheets(1).Name = "Productos"
'     ApExcel.Range("A1:AM1").Font.Size = 10
 '    ApExcel.Range("A1:AM1").Font.Bold = True
     i = 1
     Select Case UCase(opc)
            Case "ALMACEN":
                 Gcon_docu.Execute "update docu_almacen set almac_status = 'I' "
            Case "DEPOSITO"
                 Gcon_docu.Execute "update docu_deposito set depo_status = 'I' "
            Case "PRODUCTO"
                 Gcon_docu.Execute "update docu_tipo_producto set tipo_prod_status = 'I' where tipo_prod_linea = 'CAE'"
            Case "EMBALAJE"
                 Gcon_docu.Execute "update docu_embalaje set emba_status = 'I' where emba_linea = 'CAE'"
                 
     End Select
     Do While 1 = 1
     
        campo1 = Trim(ApExcel.Worksheets(1).Cells(i, 1).Value)
        campo2 = Trim(ApExcel.Worksheets(1).Cells(i, 2).Value)
        campo3 = Trim(ApExcel.Worksheets(1).Cells(i, 3).Value)
        campo4 = Trim(ApExcel.Worksheets(1).Cells(i, 4).Value)
        campo5 = Trim(ApExcel.Worksheets(1).Cells(i, 5).Value)
        campo6 = Trim(ApExcel.Worksheets(1).Cells(i, 6).Value)
        campo7 = Trim(ApExcel.Worksheets(1).Cells(i, 7).Value)
        campo8 = Trim(ApExcel.Worksheets(1).Cells(i, 8).Value)
        If campo1 = "" Then Exit Do
        
        Select Case UCase(opc)
            Case "ALMACEN"
                sql = "Select * from docu_almacen where almac_linea = 'CAE' and almac_codigo = '" & campo1 & "'"
                reg.Open sql, Gcon_docu
                If reg.EOF Then
                    sql = "Insert into docu_almacen values ('CAE', '" & campo1 & "', '" & campo2 & "','A')"
                Else
                    sql = "Update docu_almacen set almac_descripcion = '" & campo2 & "', almac_status ='A'  where almac_linea = 'CAE' and almac_codigo = '" & campo1 & "'"
                End If
                reg.Close
            Case "DEPOSITO"
                Select Case campo1
                Case "1": campo3 = "DEP COM - " & Trim(campo3)
                Case "2": campo3 = "DEP IND - " & Trim(campo3)
                Case "3": campo3 = "ZONA FRAN - " & Trim(campo3)
                Case "4": campo3 = "ALM - " & Trim(campo3)
                Case "5": campo3 = "MAQUI - " & Trim(campo3)
                End Select
                
                sql = "Select * from docu_deposito where depo_linea = 'CAE' and depo_tipo_depo = '" & campo1 & "' and depo_codigo = '" & campo2 & "'"
                reg.Open sql, Gcon_docu
                If reg.EOF Then
                    sql = "Insert into docu_deposito values ('CAE', '" & campo1 & "', '" & campo2 & "','" & campo3 & "', 'A')"
                Else
                    sql = "Update docu_deposito set depo_descripcion = '" & campo3 & "', depo_status ='A'  where depo_linea = 'CAE' and depo_tipo_depo = '" & campo1 & "' and depo_codigo = '" & campo2 & "'"
                End If
                reg.Close
            Case "PRODUCTO"
                sql = "Select * from docu_tipo_producto where tipo_prod_linea = 'CAE' and tipo_prod_codigo = '" & campo1 & "'"
                reg.Open sql, Gcon_docu
                If reg.EOF Then
                    sql = "Insert into docu_tipo_producto values ('CAE', '" & campo1 & "', '" & campo2 & "', 'A')"
                Else
                    sql = "Update docu_tipo_producto set tipo_prod_descripcion = '" & campo2 & "', tipo_prod_status ='A'  where tipo_prod_linea = 'CAE' and tipo_prod_codigo = '" & campo1 & "'"
                End If
                reg.Close
            Case "EMBALAJE"
                sql = "Select * from docu_embalaje where emba_linea = 'CAE' and emba_codigo = '" & campo1 & "'"
                reg.Open sql, Gcon_docu
                If reg.EOF Then
                    sql = "Insert into docu_embalaje values ('CAE', '" & campo1 & "', '" & campo2 & "', 'A')"
                Else
                    sql = "Update docu_embalaje set emba_descripcion = '" & campo2 & "', emba_status ='A'  where emba_linea = 'CAE' and emba_codigo = '" & campo1 & "'"
                End If
                reg.Close
        End Select
        Gcon_docu.Execute sql
        i = i + 1
     Loop
     

'cadenasql = "select prod_id as codigo, prod_descripcion as descripcion, prod_stock as cantidad, prod_valor_unit_publico as PVP,prod_valor_unit_emp as PVE, prod_aplic_IVA as IVA, prod_precio_ult_compra as Precio_ult_compra, prod_fech_ult_compra as Fecha_ult_compra,  prod_presentacion as Presentacion from producto order by prod_descripcion"
ApExcel.Workbooks.Close

Set ApExcel = Nothing
MsgBox "Proceso terminado se procesaron " & i - 1 & " Registros", vbInformation

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Opt_Almac.Value = True
End Sub

Private Sub Label1_Click()

End Sub
