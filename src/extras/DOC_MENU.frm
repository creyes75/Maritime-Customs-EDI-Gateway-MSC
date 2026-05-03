VERSION 5.00
Begin VB.Form DOC_MENU 
   Caption         =   "Extras"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd_salir 
      Caption         =   "Salir"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton Cmd_actualiza 
      Caption         =   "Actualiza Catálogos"
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
End
Attribute VB_Name = "DOC_MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_actualiza_Click()
    docu_ext_actualiza.Show
End Sub

Private Sub Cmd_salir_Click()
End
End Sub
