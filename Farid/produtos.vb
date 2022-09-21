Sub ProdFarid()
'
' ProdFarid Macro
' Desenvolvido por Marcus Gabriel no dia 05/08/2022 para enviar automaticamente a indicação dos produtos no palmtop para o Farid.
'
' Atalho do teclado: Ctrl+i
'

'Apagar as cêlulas que não serão utilizadas pelo Faridão
    Cells.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.EntireColumn.AutoFit
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:O").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:AH").Select
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
'Fim

'Formatar as cêlulas existentes como uma tabela
        Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$D$14661"), , xlYes).Name _
        = "Tabela1"
    Range("Tabela1[#All]").Select
    ActiveSheet.ListObjects("Tabela1").TableStyle = "TableStyleLight8"
'Fim

'Aplicar filtros à tabela criada
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=3, Criteria1:= _
        "<>"
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=4, Criteria1:= _
        "="
'Fim

'Input de informações na Coluna "D" (ind palmto)
    Selection.Replace What:=" ", Replacement:="N", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
'Fim

'Reaplicar filtro Final
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=4
'Fim

'Exportar a planilha na U:\cpd\gestao\Produtos Farid\produtos.xlsx (subistituindo se haver outra planilha)
    ChDir "U:\cpd\gestao\Produtos Farid"
    ActiveWorkbook.SaveAs Filename:="U:\cpd\gestao\Produtos Farid\produtos.xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
'Fim

'Se der tudo certo apresentar a mensagem de sucesso
    export_prod.Show
'Fim
'Código Finalizado
End Sub