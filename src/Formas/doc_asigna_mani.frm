VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form doc_asigna_mani 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asigna Manifiesto CAE Importación"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   Icon            =   "doc_asigna_mani.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4860
   Begin VB.Frame Frame1 
      Height          =   1980
      Left            =   60
      TabIndex        =   2
      Top             =   45
      Width           =   4740
      Begin MSMask.MaskEdBox Msk_manifiesto 
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Top             =   1575
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         Mask            =   "###-##-##-######-#"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo Cmb_buque 
         Height          =   315
         Left            =   150
         TabIndex        =   4
         Top             =   1035
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc data_buque 
         Height          =   330
         Left            =   1890
         Top             =   1035
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
      Begin MSDataListLib.DataCombo cmb_linea 
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   435
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc data_linea 
         Height          =   330
         Left            =   3105
         Top             =   420
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
      Begin MSDataListLib.DataCombo Cmb_viaje 
         Height          =   315
         Left            =   3300
         TabIndex        =   6
         Top             =   1035
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc data_viaje 
         Height          =   330
         Left            =   3480
         Top             =   1020
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buque :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   810
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Viaje :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   3300
         TabIndex        =   9
         Top             =   810
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Línea :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   8
         Top             =   195
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. Manifiesto"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   7
         Top             =   1365
         Width           =   1020
      End
   End
   Begin VB.CommandButton Cmd_acep 
      Cancel          =   -1  'True
      Height          =   510
      Left            =   3690
      Picture         =   "doc_asigna_mani.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Aceptar"
      Top             =   2100
      Width           =   510
   End
   Begin VB.CommandButton Cmd_canc 
      Height          =   510
      Left            =   4275
      Picture         =   "doc_asigna_mani.frx":5E94
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cancelar"
      Top             =   2100
      Width           =   510
   End
End
Attribute VB_Name = "doc_asigna_mani"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmb_buque_Change()
   Dim sql As String
   'If Not Cmb_buque.MatchedWithList Then Exit Sub
   sql = "Select * from docu_buque_viaje " & _
         " Where buqv_linea ='" & Cmb_linea.BoundText & "'" & _
         "   And buqv_buque ='" & Cmb_buque.BoundText & "'" & _
         "   And buqv_status = 'A'"
   
   carga_datacombo data_viaje, Cmb_viaje, sql, "buqv_viaje", "buqv_viaje"

End Sub

Private Sub cmb_linea_Change()
    If Not Cmb_linea.MatchedWithList Then Exit Sub
    carga_datacombo data_buque, Cmb_buque, "Select * from docu_buque Where buqu_linea = '" & Cmb_linea.BoundText & "' and buqu_status ='A'", "buqu_codigo", "buqu_descripcion"

End Sub

Private Sub Cmd_acep_Click()
Dim adua_ing    As String
Dim año_carga   As String
Dim tipo_mani   As String
Dim no_mani     As String
Dim digito_v    As String
On Error GoTo Errores
    If Not Cmb_linea.MatchedWithList Then
        MsgBox "Seleccione una línea naviera Válida.", vbInformation:    Exit Sub
    End If
    If Not Cmb_buque.MatchedWithList Then
        MsgBox "Seleccione un Buque Válido.", vbInformation:    Exit Sub
    End If
    If Not Cmb_viaje.MatchedWithList Then
        MsgBox "Seleccione un Viaje Válido.", vbInformation:    Exit Sub
    End If
    
    If Len(Trim$(Msk_manifiesto.Text)) = 14 Then
       adua_ing = Mid(Msk_manifiesto.Text, 1, 3)
       año_carga = Mid(Msk_manifiesto.Text, 4, 2)
       tipo_mani = Mid(Msk_manifiesto.Text, 6, 2)
       no_mani = Mid(Msk_manifiesto.Text, 8, 6)
       digito_v = Mid(Msk_manifiesto.Text, 14, 1)
       If Len(año_carga) = 2 Then
          año_carga = "20" & Trim$(año_carga)
       End If
       sql = "UPDATE docu_manifiesto SET mani_adua_ing = '" & adua_ing & "', " & _
                                        "mani_ano_car  = '" & año_carga & "'," & _
                                        "mani_tip_mani  = '" & tipo_mani & "', " & _
                                        "mani_mani_car  = '" & no_mani & "', " & _
                                        "mani_digi_ver  = '" & digito_v & "' " & _
             "Where mani_linea = '" & Cmb_linea.BoundText & "' and mani_buque = '" & Cmb_buque.BoundText & "' and mani_viaje = '" & Cmb_viaje.BoundText & "'"
       Gcon_docu.Execute sql, no
       MsgBox "Se actualizaron " & no & " manifiestos", vbInformation
    Else
       MsgBox "El No. de manifiesto no tiene el tamaño adecuado", vbInformation
    End If
    Exit Sub
Errores:
    MsgBox "Ocurrio el Error " & Err.Description, vbInformation
    Exit Sub
    Resume 0
End Sub

Private Sub Cmd_canc_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    carga_datacombo data_linea, Cmb_linea, "Select * from docu_linea_nav Where linea_status ='A'", "linea_codigo", "linea_descripcion"
    If Not data_linea.Recordset.EOF Then data_linea.Recordset.MoveFirst
End Sub

Private Sub Msk_manifiesto_LostFocus()
    If Len(Trim$(Msk_manifiesto.Text)) <> 14 Then
        MsgBox "El No. de manifiesto no tiene el tamaño adecuado", vbInformation
    End If
End Sub
