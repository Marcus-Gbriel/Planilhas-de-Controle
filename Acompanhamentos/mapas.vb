'atualização da planilha

    Sub Atualizar()
        'Desenvolvido por Marcus Gabriel no dia 02/09/2022
        'Apagar todos os dados
            Sheets("MAPA").Select
            Rows("2:2").Select
            Range("B2").Activate
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Delete Shift:=xlUp
            Sheets("CALC").Select
            Columns("D:D").Select
            Selection.Delete Shift:=xlToLeft
            Sheets("MAPA").Select
            Range("TAB_MAP[MAPA]").Select
            Range("A2").Select
            Selection.ClearContents
        'Atulizar dados
            ActiveWorkbook.RefreshAll
        'Aplicar filtro 01200105
                Sheets("01 20 01 05").Select
                ActiveSheet.ListObjects("_01_20_01_05").Range.AutoFilter Field:=1, Criteria1 _
                    :=Array("AC", "AL", "AM", "AP", "BA", "CE", "Cod", "DAM", "DF", "ES", "GO", "MA", "MS", _
                    "MT", "PA", "PB", "PE", "PI", "PP0", "PR", "RJ", "RN", "RO", "RR", "RS", "SC", "SE", "SP", _
                    "TO", "="), Operator:=xlFilterValues
                Rows("2:2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Delete Shift:=xlUp
                ActiveSheet.ListObjects("_01_20_01_05").Range.AutoFilter Field:=1
            'transformar números
                    Columns("B:B").Select
                Selection.TextToColumns Destination:=Range( _
                    "_01_20_01_05[[#Headers],[COD]]"), DataType:=xlDelimited, TextQualifier _
                    :=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, Semicolon:= _
                    False, Comma:=False, Space:=True, Other:=False, FieldInfo:=Array(1, 1), _
                    TrailingMinusNumbers:=True
                Selection.NumberFormat = "0"
            'apagar cêlulas vazias
                    Columns("B:B").Select
                Selection.TextToColumns Destination:=Range( _
                    "_01_20_01_05[[#Headers],[COD]]"), DataType:=xlDelimited, TextQualifier _
                    :=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, Semicolon:= _
                    False, Comma:=False, Space:=True, Other:=False, FieldInfo:=Array(1, 1), _
                    TrailingMinusNumbers:=True
                Selection.NumberFormat = "0"
                    ActiveSheet.ListObjects("_01_20_01_05").Range.AutoFilter Field:=1, Criteria1 _
                    :="="
                    ActiveSheet.ListObjects("_01_20_01_05").Range.AutoFilter Field:=1
                Sheets("MAPA").Select
        'Puxar dados atualizados
            Sheets("03 05 17").Select
            Range("D2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("CALC").Select
            Range("D2").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            Columns("D:D").Select
            ActiveSheet.Range("$D$1:$D$96622").RemoveDuplicates Columns:=1, Header:= _
                xlNo
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Range("D2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("MAPA").Select
            Range("TAB_MAP[MAPA]").Select
            ActiveSheet.Paste
    End Sub