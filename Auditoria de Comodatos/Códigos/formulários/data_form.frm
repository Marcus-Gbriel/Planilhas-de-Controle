VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} data_form 
   Caption         =   "Selecinar Data"
   ClientHeight    =   3045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3165
   OleObjectBlob   =   "data_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "data_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    Range("G7").Select
    ActiveCell.FormulaR1C1 = data_anterior.Value
Unload data_form

End Sub

Private Sub CommandButton2_Click()

Unload data_form

End Sub
