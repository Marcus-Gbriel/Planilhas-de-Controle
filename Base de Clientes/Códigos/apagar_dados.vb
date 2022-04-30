Sub apagar()
'Criado por Marcus Gabriel dia 13/04/2022.
'Esse Script apaga toda a base de dados dos clientes puxados pelo relatório.
'No final ele puxa os dados apagados, atualizando todos eles atraves dos arquivos gerados.
'Caso tenha alguma dúvida acesse https://marcusgabriel.space e entre em contato comigo.
    'Apagar todos os dados puxados anteriormente.
        Sheets("Base Completa").Select
        Rows("4:4").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Delete Shift:=xlUp
        Range("Tabela5[NB]").Select
        Selection.ClearContents
        Range("Tabela5[RAZÃO SOCIAL]").Select
        Selection.ClearContents
        Range("Tabela5[CATEGORIA]").Select
        Selection.ClearContents
        Range("Tabela5[INSCR EST]").Select
        Selection.ClearContents
        Range("Tabela5[NOME FANTASIA]").Select
        Selection.ClearContents
        Range("Tabela5[CODIGO]").Select
        Selection.ClearContents
        Range("Tabela5[DESCRICAO DE COMPRA]").Select
        Selection.ClearContents
        Range("Tabela5[E-MAIL]").Select
        Selection.ClearContents
        Range("Tabela5[SETOR SEG]").Select
        Selection.ClearContents
        Range("Tabela5[CIDADE]").Select
        Selection.ClearContents
        Range("Tabela5[BAIRRO]").Select
        Selection.ClearContents
        Range("Tabela5[CEP]").Select
        Selection.ClearContents
        Range("Tabela5[CPF/CNPJ]").Select
        Selection.ClearContents
        Range("Tabela5[TELEFONE]").Select
        Selection.ClearContents
        Range("Tabela5[DATA CADASTRAMENTO]").Select
        Selection.ClearContents
        Range("R3").Select
        Sheets("base 806").Select
        ActiveWindow.SmallScroll Down:=-99
        Columns("G:H").Select
        Selection.Delete Shift:=xlToLeft
        Sheets("Calculos de Data").Select
        Rows("4:4").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Delete Shift:=xlUp
        Range("Tabela4[SEG/]").Select
        Selection.ClearContents
        Range("Tabela4[TER/]").Select
        Selection.ClearContents
        Range("Tabela4[QUAR/]").Select
        Selection.ClearContents
        Range("Tabela4[QUIN/]").Select
        Selection.ClearContents
        Range("Tabela4[SEX/]").Select
        Selection.ClearContents
        Range("Tabela4[SAB/]").Select
        Selection.ClearContents
        Range("Tabela4[DOM/]").Select
        Selection.ClearContents
        Range("Tabela4[SEG2/]").Select
        Selection.ClearContents
        Range("Tabela4[TER2/]").Select
        Selection.ClearContents
        Range("Tabela4[QUAR2/]").Select
        Selection.ClearContents
        Range("Tabela4[QUIN2/]").Select
        Selection.ClearContents
        Range("Tabela4[SEX2/]").Select
        Selection.ClearContents
        Range("Tabela4[SAB2/]").Select
        Selection.ClearContents
        Range("Tabela4[DOM2/]").Select
        Selection.ClearContents
        Sheets("Base Completa").Select
        Range("A2").Select

    'Apagar Novas Tabelas
        Sheets("Base de Clientes Extra").Select
        Rows("3:3").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Delete Shift:=xlUp
        Range("Tabela6[NB]").Select
        Selection.ClearContents
        Range("Tabela6[RAZÃO]").Select
        Selection.ClearContents
        Range("Tabela6[CATEGORIA]").Select
        Selection.ClearContents
        Range("Tabela6[INSCR EST]").Select
        Selection.ClearContents
        Range("Tabela6[FANTASIA]").Select
        Selection.ClearContents
        Range("Tabela6[Código]").Select
        Selection.ClearContents
        Range("Tabela6[DESCRIÇÃO]").Select
        Selection.ClearContents
        Range("Tabela6[E-MAIL]").Select
        Selection.ClearContents
        Range("Tabela6[ÁREA SEG]").Select
        Selection.ClearContents
        Range("Tabela6[CIDADE]").Select
        Selection.ClearContents
        Range("Tabela6[BAIRRO]").Select
        Selection.ClearContents
        Range("Tabela6[CEP]").Select
        Selection.ClearContents
        Range("Tabela6[CPF/CGC]").Select
        Selection.ClearContents
        Range("Tabela6[Tel]").Select
        Selection.ClearContents
        Range("Tabela6[NB]").Select
        Sheets("base 806").Select
        Columns("G:G").Select
        Selection.Delete Shift:=xlToLeft
        Columns("G:G").Select
        Selection.Delete Shift:=xlToLeft
        Sheets("Menu").Select

    'Atualizar Dados
        ActiveWorkbook.RefreshAll
        Sheets("Menu").Select

    'Caixa de Mensagem
        MsgBox "Dados apagados com Sucesso!", vbInformation, "Mensagem do Sistema"
    'Fim

End Sub