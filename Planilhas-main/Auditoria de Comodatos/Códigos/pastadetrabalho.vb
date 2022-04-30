Private Sub Workbook_Open()
'Criado por Marcus Gabriel no dia 19/04/2022

    'Mensagens de Boas Vindas
        Dim usuario As String
        
        usuario = Environ("USERNAME")
        usuario = UCase(usuario)

        If Time >= "06:00:00" And Time <= "11:59:59" Then
            MsgBox "Bom dia, " & usuario + vbCrLf + "Planilha Criada por Marcus Gabriel, Hoje é " & Date, vbInformation
        ElseIf Time >= "12:00:00" And Time <= "17:59:59" Then
            MsgBox "Boa tarde, " & usuario + vbCrLf + "Planilha Criada por Marcus Gabriel, Hoje é" & Date, vbInformation
        ElseIf Time >= "18:00:00" And Time <= "23:59:59" Then
            MsgBox "Boa noite, " & usuario + vbCrLf + "Planilha Criada por Marcus Gabriel, Hoje é" & Date, vbInformation
        End If
        'Fim
    'Formatar Planilha
        'Etapa 1: Limpar dados na AUDITORIA
           Sheets("AUDITORIA COMODATOS").Select
           Range("B5").Select
           Selection.ClearContents
           Range("B5").Select
           'Fim
        'Etapa 1: Limpar dados na DePara
            Sheets("TRANSFERENCIA (DE-PARA))").Select
            Columns("O:O").Select
            Selection.ClearContents
            Range("B5").Select
            Selection.ClearContents
            Range("B46").Select
            Selection.ClearContents
                Range("B5").Select
                'Fim
        'Ajeitar para não ficar feio.
            Sheets("MENU").Select
            Range("A1").Select
            'Fim
    'Caixa de Atualização
        atualizar_dados.Show
        'Fim
    'Levar para o MENU
        Sheets("MENU").Select
        Range("A1").Select
        'Fim
End Sub