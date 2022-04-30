VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} atualizar_dados 
   Caption         =   "Mensagem do Sistema"
   ClientHeight    =   1275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "atualizar_dados.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "atualizar_dados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
        ActiveWorkbook.RefreshAll
    Unload atualizar_dados
End Sub

Private Sub CommandButton2_Click()
    Unload atualizar_dados
End Sub
