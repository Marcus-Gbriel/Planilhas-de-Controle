'Módulo 1

    Sub Puxar()
    '
    ' Puxar Macro
    '

    'Puxar Informações
        'NB Clientes
            Sheets("21 04 01").Select
            Range("M2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Atualizações").Select
            Range("Tabela4[NB]").Select
            ActiveSheet.Paste

        'Movimentação Cadastro (Motivo)
            Sheets("21 04 01").Select
            Range("BH2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Atualizações").Select
            Range("Tabela4[MOTIVO]").Select
            ActiveSheet.Paste

        'CPF/CNPJ
            Sheets("21 04 01").Select
            Range("N2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Atualizações").Select
            Range("Tabela4[CPF/CNPJ]").Select
            ActiveSheet.Paste

        'Nome Fantasia
            Sheets("21 04 01").Select
            Range("P2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Atualizações").Select
            Range("Tabela4[NOME FANTASIA]").Select
            ActiveSheet.Paste
        
        'Situação Cliente
            Sheets("21 04 01").Select
            Range("BI2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Atualizações").Select
            Range("Tabela4[SITUAÇÂO]").Select
            ActiveSheet.Paste

        'GV
            Sheets("21 04 01").Select
            Range("BE2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Atualizações").Select
            Range("Tabela4[GV]").Select
            ActiveSheet.Paste
            Range("A1").Select

    End Sub

'formulário

    Private Sub CommandButton1_Click()
            ActiveWorkbook.RefreshAll
        Unload atualizar_dados
    End Sub

    Private Sub CommandButton2_Click()
        Unload atualizar_dados
    End Sub