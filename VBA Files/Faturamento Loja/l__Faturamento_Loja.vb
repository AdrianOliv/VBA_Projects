Sub l__Faturamento_Loja()

    formata_celula
    limpa_fundo
    cabecalho
    colunas
    linhas_brancas
    formata_bordas
    criar_planilhas
    insercao_dados
    formata_planilhas
    formata_planilha1
    ajustar
    salvar

End Sub



Function ajustar()
    ' ajustar planilhas para salvar
    
    For Each aba In Worksheets
        aba.Activate
        
        '' ADICIONA TOTAL DE NOTAS
        total = Range("A1").End(xlDown).Row
        Range("K1").Value = total - 1 & " Notas Fiscais"
        
        Range("K1").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 9621584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        ActiveWindow.DisplayGridlines = False
        Cells.EntireColumn.AutoFit
        Cells.EntireRow.AutoFit
        Range("A1").Select
    Next aba
End Function


Function cabecalho()
    ' retira o cabecalho desnecessario
    
    While Cells.Find("Data Doc").Row <> 1
        Cells.Find("Data Doc").Offset(-1, 0).EntireRow.Delete
    Wend
End Function


Function colunas()
    ' retira colunas Emp. e Série
    
    Cells.Find("Emp.").EntireColumn.Delete
    Cells.Find("Sér.").EntireColumn.Delete
End Function


Function criar_planilhas()
    ' formata as naturezas de operacao para a criacao das planilhas
    ' o programa identifica as naturezas para então dividir as linhas por planilhas
    ' cada planilha corresponde a uma natureza de operacao
    
    '' COPIA NATUREZA ''
    linha = 1
    linha_fim = Range("A1").End(xlDown).Row
    Range("D2:D" & linha_fim).Copy
    Range("J1").PasteSpecial
    Application.CutCopyMode = False
        
    '' REMOVE DUPLICADAS
    ActiveSheet.Range("$J$1:$J$" & linha_fim).RemoveDuplicates Columns:=1, Header:= _
    xlNo
    
    linha_fim = Range("J1").End(xlDown).Row

    '' FORMULA CFOP
    Range("I1").FormulaR1C1 = "=INDEX(C[-6],MATCH(RC[1],C[-5],0))"
    Range("I1").Copy
    Range("I2:I" & linha_fim).PasteSpecial
    Range("I1:I" & linha_fim).Copy
    Range("I1:I" & linha_fim).PasteSpecial (xlPasteValues)

    '' RESUME NOME NAT OPERACAO
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

    '' CRIA NOVAS PLANILHAS
    While linha <= linha_fim
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = Sheets(1).Cells(linha, 10)
    
        Sheets(1).Range("A1:H1").Copy
        ActiveSheet.Range("A1").PasteSpecial
    
        linha = linha + 1
    Wend
End Function


Function formata_bordas()
    ' formatar bordas
    
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
End Function


Function formata_celula()
    ' formatar celulas
    
    Cells.Select
    Selection.Borders.LineStyle = xlNone
    With Selection.Font
        .Name = "Calibri"
        .Size = 10
        .ColorIndex = xlAutomatic
        .Underline = xlUnderlineStyleNone
    End With
    
    With Selection
        .HorizontalAlignment = xlLeft
        .WrapText = False
        .MergeCells = False
    End With
End Function


Function formata_planilhas()
    ' formatar as planilhas
    
    i = 1
    folha = 2
    linha_fim = Range("J1").End(xlDown).Row
    While i <= linha_fim
        Sheets(folha).Select
        ActiveSheet.Name = Sheets(1).Cells(i, 9)

        '' SOMA DAS NFS
        Columns("G:G").Select
        Selection.Style = "Currency"
    
        '' FORMATA CELULA SOMA
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
        
        i = i + 1
        folha = folha + 1
    Wend
End Function


Function formata_planilha1()
    ' formatar primeira planilha
    
    '' RENOMEAR
    Sheets(1).Select
    ActiveSheet.Name = "Todos"
    Sheets("Todos").Range("I:J").Clear
End Function


Function insercao_dados()
    ' insere as notas fiscais em cada planilha conforme natureza
    ' o mesmo possui barra de progresso

    Dim total As Long
    Dim contador As Long
    Dim largura As Long
    Dim percent As Double
    
    Application.ScreenUpdating = False
    linha = 2
    
    '' BARRA DE PROGRESSO
    total = Sheets(1).Range("A1").End(xlDown).Row
    Progresso.Show
    largura = Progresso.Evolucao.Width

    While Sheets(1).Cells(linha, 1) <> ""
        Sheets(1).Range("A" & linha & ":H" & linha).Copy
        natureza = Sheets(1).Cells(linha, 4)
        Sheets(natureza).Range("A300").End(xlUp).Offset(1, 0).PasteSpecial
        Application.CutCopyMode = False
        
        Sheets(1).Select
        linha = linha + 1

        '' BARRA DE PROGRESSO
        DoEvents
        contador = linha
        percent = contador / total
        Progresso.Evolucao.Width = percent * largura
        Progresso.Valor = Round(percent * 100, 1) & "%"
    Wend
    
    Application.ScreenUpdating = True
    Unload Progresso
End Function


Function linhas_brancas()
    ' retira linhas em branco
    
    Range("A700").End(xlUp).End(xlUp).Select
    While ActiveCell.Row <> 1
        ActiveCell.Offset(-1, 0).EntireRow.Delete
        ActiveCell.End(xlUp).Select
    Wend
End Function


Function limpa_fundo()
    ' limpar fundo da planilha
    
    Cells.Find("Total Geral").Offset(-1, 0).EntireRow.Select
    Selection.Resize(Selection.Rows.Count + 20).Delete
End Function


Function salvar()
    ' salvar arquivo
    Sheets(1).Select
    
    MkDir ("C:\Users\" & UsuarioRede & "\Downloads\" & Month(Range("A2")) _
    & " " & StrConv(MonthName(Month(Range("A2"))), vbProperCase))
    
    pasta = ("C:\Users\" & UsuarioRede & "\Downloads\" _
    & Month(Range("A2")) & " " & StrConv(MonthName(Month(Range("A2"))), vbProperCase))
    
    ActiveWorkbook.SaveAs Filename:=pasta & "\" & "Faturamento Loja " _
    & Month(Range("A2")) & "_" & Format(Year(Range("A2"))), FileFormat:=xlWorkbookDefault
End Function


Function UsuarioRede() As String
    ' Encontra o nome do usuário
    
    Dim ObjNetwork
    Set ObjNetwork = CreateObject("WScript.Network")
    UsuarioRede = ObjNetwork.UserName
End Function