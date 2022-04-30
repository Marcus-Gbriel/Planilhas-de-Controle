Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    
    data_form.Show
    
End Sub

Sub exportar_pdf()

    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "U:\Cadastro\Auditoria Comodatos\PDF Exportado\PDF.pdf", Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        True

End Sub

Sub imprimir_comodadto()

    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False

End Sub