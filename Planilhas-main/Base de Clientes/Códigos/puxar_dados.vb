Sub atualizar_dados()
'Criado por MArcus Gabriel no dia 13/04/2022.
'Esse Script atualiza e calcula todas as informações recebebidas pelos arquivos gerados.
'A cada etapa a um comentário para expecificar tudo corretamente.
'Caso tenha alguma dúvida acesse https://marcusgabriel.space e entre em contato comigo.
'Atualização de dados
    'Caixa de Mensagem Inicio
            MsgBox "Atualizando dados, Aguarde...", vbInformation, "Mensagem do Sistema"

    'Puxar e Separar FREQ de Visitas:
        'FREQ SEGMENTADA
            Sheets("base 804").Select
            Range("D2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Calculos de Data").Select
            Range("A3").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            Selection.TextToColumns Destination:=Range("A3"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
                :="/", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
                1), Array(6, 1), Array(7, 1)), TrailingMinusNumbers:=True

        'FREQ TELEVENDAS
            Sheets("base 806").Select
            Range("D2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Calculos de Data").Select
            Range("H3").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            Selection.TextToColumns Destination:=Range("H3"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
                :="/", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
                1), Array(6, 1), Array(7, 1)), TrailingMinusNumbers:=True
        'FINALIZADO

    'Compilar dados de Clientes das Bases 804, 805 e 806
        'NB Cliente
            Sheets("base 804").Select
            Range("A2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base Completa").Select
            Range("Tabela5[NB]").Select
            ActiveSheet.Paste

        'Razão Social
            Sheets("base 804").Select
            Range("B2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base Completa").Select
            Range("Tabela5[RAZÃO SOCIAL]").Select
            ActiveSheet.Paste

        'Categoria
            Sheets("base 804").Select
            Range("C2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base Completa").Select
            Range("Tabela5[CATEGORIA]").Select
            ActiveSheet.Paste

        'Incrição Estadual
            Sheets("base 804").Select
            Range("E2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base Completa").Select
            Range("Tabela5[INSCR EST]").Select
            ActiveSheet.Paste

        'Nome Fantasia
            Sheets("base 804").Select
            Range("F2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base Completa").Select
            Range("Tabela5[NOME FANTASIA]").Select
            ActiveSheet.Paste

        'Código Forma de Compra
            Sheets("base 804").Select
            Range("G2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base Completa").Select
            Range("Tabela5[CODIGO]").Select
            ActiveSheet.Paste

        'DESCRIÇÃO DE COMPRA
            Sheets("base 804").Select
            Range("H2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base Completa").Select
            Range("Tabela5[DESCRICAO DE COMPRA]").Select
            ActiveSheet.Paste

        'E-Mail
            Sheets("base 805").Select
            Range("A2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base Completa").Select
            Range("Tabela5[E-MAIL]").Select
            ActiveSheet.Paste

        'Setor Segmentado
            Sheets("base 805").Select
            Range("B2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base Completa").Select
            Range("Tabela5[SETOR SEG]").Select
            ActiveSheet.Paste

        'Cidade
            Sheets("base 805").Select
            Range("E2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base Completa").Select
            Range("Tabela5[CIDADE]").Select
            ActiveSheet.Paste

        'Bairro
            Sheets("base 805").Select
            Range("F2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base Completa").Select
            Range("Tabela5[BAIRRO]").Select
            ActiveSheet.Paste

        'CEP
            Sheets("base 805").Select
            Range("G2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base Completa").Select
            Range("Tabela5[CEP]").Select
            ActiveSheet.Paste

        'CPF/CNPJ
            Sheets("base 806").Select
            Range("B2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base Completa").Select
            Range("Tabela5[CPF/CNPJ]").Select
            ActiveSheet.Paste

        'Telefone (Contato)
            Sheets("base 806").Select
            Range("C2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base Completa").Select
            Range("Tabela5[TELEFONE]").Select
            ActiveSheet.Paste

        'Data Cadastramento
            Sheets("base 806").Select
            Columns("F:F").Select
            Selection.TextToColumns Destination:=Range("base_806[[#Headers],[Column6]]") _
                , DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                ConsecutiveDelimiter:=True, Tab:=False, Semicolon:=False, Comma:=False _
                , Space:=False, Other:=True, OtherChar:=" ", FieldInfo:=Array(Array(1, 1 _
                ), Array(2, 1)), TrailingMinusNumbers:=True
            Range("F2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base Completa").Select
            Range("Tabela5[DATA CADASTRAMENTO]").Select
            ActiveSheet.Paste

        'Compilar Dados de SEGMENTO
            'Sheets("base 806").Select
            'Range("I2").Select
            'Application.CutCopyMode = False
            'ActiveCell.FormulaR1C1 = "=CONCAT(base_806[@[Column7]:[Column8]])"
            'Range("I3").Select

        'Formatação de Dados
            Columns("P:P").Select
            Selection.NumberFormat = "(00)00000-0000"
            Columns("N:N").Select
            Selection.NumberFormat = "00000-000"
            Columns("Q:Q").Select
            Selection.NumberFormat = "00""/""00""/""0000"
            Range("A2").Select
            Sheets("Menu").Select
            Range("A1").Select
        'Fim
    'Puxar dados para Base de Clientes Extra, para exportação
        'NB Clientes
            Sheets("Base Completa").Select
            Range("A3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base de Clientes Extra").Select
            Range("Tabela6[NB]").Select
            ActiveSheet.Paste

        'Razão Social
            Sheets("Base Completa").Select
            Range("B3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base de Clientes Extra").Select
            Range("Tabela6[RAZÃO]").Select
            ActiveSheet.Paste

        'Categoria do Cliente
            Sheets("Base Completa").Select
            Range("C3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base de Clientes Extra").Select
            Range("Tabela6[CATEGORIA]").Select
            ActiveSheet.Paste

        'Dia de Visita
            Sheets("Base Completa").Select
            Range("E3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base de Clientes Extra").Select
            Range("Tabela6[DV]").Select
            ActiveSheet.Paste

        'Incrição Estadual
            Sheets("Base Completa").Select
            Range("D3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base de Clientes Extra").Select
            Range("Tabela6[INSCR EST]").Select
            ActiveSheet.Paste

        'Nome Fantasia
            Sheets("base 804").Select
            Range("F2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base de Clientes Extra").Select
            Range("Tabela6[FANTASIA]").Select
            ActiveSheet.Paste

        'Código de Tipo de Compra
            Sheets("Base Completa").Select
            Range("G3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base de Clientes Extra").Select
            Range("Tabela6[Código]").Select
            ActiveSheet.Paste

        'Descrição de Compra
            Sheets("Base Completa").Select
            Range("H3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base de Clientes Extra").Select
            Range("Tabela6[DESCRIÇÃO]").Select
            ActiveSheet.Paste

        'E-mail
            Sheets("Base Completa").Select
            Range("I3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base de Clientes Extra").Select
            Range("Tabela6[E-MAIL]").Select
            ActiveSheet.Paste

        'GV SEG
            'Calculo GV
                'NB Cliente
                    Sheets("base 804").Select
                    Range("A2").Select
                    Range(Selection, Selection.End(xlDown)).Select
                    Selection.Copy
                    Sheets("Calc").Select
                    Range("Tabela8[NB Cliente]").Select
                    ActiveSheet.Paste
                'Setor
                    Sheets("Base Completa").Select
                    Range("J3").Select
                    Range(Selection, Selection.End(xlDown)).Select
                    Selection.Copy
                    Sheets("Calc").Select
                    Range("Tabela8[SETOR]").Select
                    ActiveSheet.Paste
                'FIM
        'Area SEG
            Sheets("Base Completa").Select
            Range("J3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base de Clientes Extra").Select
            Range("Tabela6[ÁREA SEG]").Select
            ActiveSheet.Paste

        'Cidade
            Sheets("Base Completa").Select
            Range("L3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base de Clientes Extra").Select
            Range("Tabela6[CIDADE]").Select
            ActiveSheet.Paste

        'Bairro
            Sheets("Base Completa").Select
            Range("M3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base de Clientes Extra").Select
            Range("Tabela6[BAIRRO]").Select
            ActiveSheet.Paste

        'CEP
            Sheets("Base Completa").Select
            Range("N3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base de Clientes Extra").Select
            Range("Tabela6[CEP]").Select
            ActiveSheet.Paste

        'Cidade
            Sheets("Base Completa").Select
            Range("L3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base de Clientes Extra").Select
            Range("Tabela6[CIDADE]").Select
            ActiveSheet.Paste

        'CPF/CNPJ
            Sheets("Base Completa").Select
            Range("O3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base de Clientes Extra").Select
            Range("Tabela6[CPF/CGC]").Select
            ActiveSheet.Paste

        'Telefone
            Sheets("Base Completa").Select
            Range("P3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base de Clientes Extra").Select
            Range("Tabela6[Tel]").Select
            ActiveSheet.Paste

        'Cidade
            Sheets("Base Completa").Select
            Range("L3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Base de Clientes Extra").Select
            Range("Tabela6[CIDADE]").Select
            ActiveSheet.Paste
    
        'Calcular Segmento de Clientes
            Sheets("base 806").Select
            Range("base_806[Column7]").Select
            Selection.Copy
            Sheets("Calc").Select
            Range("Tabela9[Coluna1]").Select
            ActiveSheet.Paste
            Sheets("base 806").Select
            Range("base_806[Column8]").Select
            Selection.Copy
            Sheets("Calc").Select
            Range("Tabela9[Coluna2]").Select
            ActiveSheet.Paste
        'Finalizado
    '
    'Atualização de planilhas a partir daqui!
    '
    '
    '
    '
    '
    'Alinhas Planilhas
        Sheets("Pesquisa").Select
        Range("A8").Select
        Sheets("Base Completa").Select
        Range("A3").Select
        Sheets("Menu").Select
        Range("A1").Select
    'Finalizado
    'Caixa de Mensagem Final
            Sheets("Menu").Select
            MsgBox "Dados atualizado com Sucesso!", vbInformation, "Mensagem do Sistema"
    'Finalizado
End Sub