Sub limpar_depara()
    Sheets("TRANSFERENCIA (DE-PARA))").Select
    Columns("O:O").Select
    Selection.ClearContents
        Range("B5").Select
        Selection.ClearContents
            Range("B46").Select
            Selection.ClearContents
                Range("B5").Select
End Sub

Sub limpar_auditoria()
    Sheets("AUDITORIA COMODATOS").Select
    Range("B5").Select
    Selection.ClearContents
End Sub