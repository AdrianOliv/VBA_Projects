Function UsuarioRede() As String
''''''''''''''''' ESSA FUNCAO PESQUISA O NOME DE USUARIO DA MAQUINA'''''''''''
    Dim ObjNetwork
    Set ObjNetwork = CreateObject("WScript.Network")
    UsuarioRede = ObjNetwork.UserName
End Function

Sub l___VBA__Faturamento_Loja()

''''''''''''''''''' FORMATAR CELULAS '''''''''''''''''
    Cells.Select
    Selection.Borders.LineStyle = xlNone
    With Selection.Font
        .Name = "Calibri"
        .Size = 10
        .ColorIndex = xlAutomatic
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .WrapText = True
        .WrapText = False
        .MergeCells = False
    End With
    Selection.Font.Underline = xlUnderlineStyleSingle
    Selection.Font.Underline = xlUnderlineStyleNone
    
'''''''''''''' RETIRA ÚLTIMAS LINHAS ''''''''''''''''
    If IsEmpty(Cells.Find("Total Geral").Offset(-1, 0).Value) Then
        Cells.Find(What:="Total Geral").EntireRow.Select
        Selection.Resize(Selection.Rows.Count + 20).Delete
    Else
        Cells.Find(What:="Total Geral").Offset(-1, 0).EntireRow.Select
        Selection.Resize(Selection.Rows.Count + 20).Delete
    End If
            
'''''''''''''' TIRA CABEÇALHOS ''''''''''''
    While Cells.Find("Data Doc").Row <> 1
        Cells.Find("Data Doc").Offset(-1, 0).EntireRow.Delete
    Wend

'''''''''''''   RETIRA COLUNA EMP E SERIE '''''''''''''''
    Cells.Find("Emp.").EntireColumn.Delete
    Cells.Find("Sér.").EntireColumn.Delete

''''''''''''''''''''AUTOFIT''''''''''''''''''''''''''''
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit

'''''''''''''''RETIRA LINHAS EM BRANCO ''''''''''''''
    Range("A700").End(xlUp).End(xlUp).Select
    While ActiveCell.Row <> 1
        ActiveCell.Offset(-1, 0).EntireRow.Delete
        ActiveCell.End(xlUp).Select
    Wend

''''''''''''' FORMATA BORDAS '''''''''''''''
    Range("A1").CurrentRegion.Select
    For k = 5 To 12
        If k > 6 Then
            With Selection.Borders.Item(k)
            .LineStyle = xlContinuos
            .ColorIndex = 0
            .Weight = xlThin
            End With
        End If
    Next k
    
''''''''' COPIA NATUREZA   ''''''''''''
linha = 1
linha_fim = Range("A1").End(xlDown).Row
Range("D2:D" & linha_fim).Copy
Range("J1").PasteSpecial
Application.CutCopyMode = False
        
'''''''''''REMOVE DUPLICADAS'''''''''''
ActiveSheet.Range("$J$1:$J$" & linha_fim).RemoveDuplicates Columns:=1, Header:= _
    xlNo
    
linha_fim = Range("J1").End(xlDown).Row

''''''''''''''' FORMULA CFOP''''''''''
Range("I1").FormulaR1C1 = "=INDEX(C[-6],MATCH(RC[1],C[-5],0))"
Range("I1").Copy
Range("I2:I" & linha_fim).PasteSpecial
Range("I1:I" & linha_fim).Copy
Range("I1:I" & linha_fim).PasteSpecial (xlPasteValues)

''''''''''''' RESUME NOME NATUREZA'''''''''''
    Range("D:J").Replace What:="(S) - ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("D:J").Replace What:="(E) - ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("D:J").Replace What:="/", Replacement:=".", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("D:J").Replace What:="DE FATURAMENTO", Replacement:="DE FAT", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

''''''''''''''CRIA NOVAS PLANILHAS'''''''''''''
While linha <= linha_fim
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = Sheets(1).Cells(linha, 10)
    
    Sheets(1).Range("A1:H1").Copy
    ActiveSheet.Range("A1").PasteSpecial
    
    linha = linha + 1
Wend

'''''''''''''''''' INICIA INSERÇÃO DOS DADOS '''''''''''''''''''
linha = 2

While Sheets(1).Cells(linha, 1) <> ""
    Sheets(1).Range("A" & linha & ":H" & linha).Copy
    natureza = Sheets(1).Cells(linha, 4)
    Sheets(natureza).Select
    Range("A300").End(xlUp).Offset(1, 0).PasteSpecial
    Application.CutCopyMode = False
    
    Sheets(1).Select

    linha = linha + 1
Wend

i = 1
folha = 2

'''''''''''''''' NOMEIA NOVAMENTE AS PLANILHAS '''''''''''''''
While i <= linha_fim
    Sheets(folha).Select
    ActiveSheet.Name = Sheets(1).Cells(i, 9)

''''''''''''''''' SOMA '''''''''''''''''''
    Columns("G:G").Select
    Selection.Style = "Currency"
    
    linha_calc = Range("G1").End(xlDown).Offset(1, 0).Row - 1
    Range("G1").End(xlDown).Offset(1, 0).Value = WorksheetFunction.Sum(Range("G2:G" & linha_calc))
    Range("G1").End(xlDown).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 9621584
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    
''''''''''''''''' FIM SOMA '''''''''''''''''

'''''''''''''''''' ADICIONA NA PRIMEIRA PLANILHA O NUMERO DE NOTAS DE CADA NATUREZA'''''''''''''''''''
    Total = Range("A1").End(xlDown).Row
    Sheets(1).Range("K" & i) = Total - 1
    i = i + 1
    folha = folha + 1
Wend

'''''''''''''RENOMEIA PRIMEIRA PLANILHA ''''''''''''''''''
Sheets(1).Select
ActiveSheet.Name = "Todos"
Sheets("Todos").Range("I:J").Clear

''''''''''''''''SOMAR NOTAS''''''''''''''''''''''''
    linha_calc = Range("K1").End(xlDown).Row
    Range("K1").End(xlDown).Offset(1, 0).Value = WorksheetFunction.Sum(Range("K1:K" & linha_calc)) & " NFs"
    Range("K1:K" & linha_calc).Delete shift:=xlUp
    Range("K1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 9621584
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

''''''''''''''''' AUTO AJUSTAR'''''''''''''''''
For Each aba In Worksheets
    aba.Activate
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
Next aba

''''''''''''''''''''''' SALVAR SEM MACRO ''''''''''''''''''''''
    ActiveWorkbook.SaveAs Filename:="C:\Users\" & UsuarioRede & "\Downloads\" & "Faturamento Loja " & Month(Range("A2")) & "_" & Format(Year(Range("A2"))), FileFormat:=xlWorkbookDefault

End Sub