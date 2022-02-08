Sub l__SaldoProjetado()
    
    formatacao
    limpa_colunas
    formata_planilha
    ajustar
    filtro

End Sub



Function formatacao()
    '' formatar celulas
    
    Range("A1").CurrentRegion.Select
    Selection.Replace What:="", Replacement:="-", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
        
    Selection.Borders.LineStyle = xlNone
    With Selection.Font
        .Name = "Calibri"
        .Size = 10
        .ColorIndex = xlAutomatic
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .WrapText = False
        .MergeCells = False
    End With
    Selection.Font.Underline = xlUnderlineStyleNone
    Range("A1").Select
End Function


Function limpa_colunas()
    '' limpar colunas desnecessarias
    
    Cells.Find("Cliente").EntireColumn.Delete
    Cells.Find("Rua").EntireColumn.Delete
    Cells.Find("Complemento").EntireColumn.Delete
    Cells.Find("Número").EntireColumn.Delete
    Cells.Find("Bairro").EntireColumn.Delete
    Cells.Find("Cidade").EntireColumn.Delete
    Cells.Find("Pais").EntireColumn.Delete
    Cells.Find("Estado").EntireColumn.Delete
    Cells.Find("CEP").EntireColumn.Delete
    Cells.Find("Quant. Cancelada").EntireColumn.Delete
    Cells.Find("Vendedor").EntireColumn.Delete
    Cells.Find("O.C. Cliente").EntireColumn.Delete
    Cells.Find("Und.").EntireColumn.Delete
    Cells.Find("Retenção de Credito").EntireColumn.Delete
    Cells.Find("Outras Retenções").EntireColumn.Delete
    Cells.Find("Data Minima  Faturamento").EntireColumn.Delete
    Cells.Find("Faturamento Parcial ").EntireColumn.Delete
    Cells.Find("Data Criação Pedido").EntireColumn.Delete
    Cells.Find("Data Criação Linha").EntireColumn.Delete
    Cells.Find("Transportadora").EntireColumn.Delete
    Cells.Find("Distribuição").EntireColumn.Delete
    Cells.Find("Percurso").EntireColumn.Delete
    Cells.Find("Planejado").EntireColumn.Delete
    Cells.Find("Data Percurso").EntireColumn.Delete
    Cells.Find("Peso Liquido").EntireColumn.Delete
    Cells.Find("Peso Bruto").EntireColumn.Delete
    Cells.Find("Oc cliente linha").EntireColumn.Delete
    Cells.Find("Filtro CP").EntireColumn.Delete
    Cells.Find("Nota Fiscal").EntireColumn.Delete
    Cells.Find("Data Nota Fiscal").EntireColumn.Delete
    Cells.Find("Prazo Médio ").EntireColumn.Delete
    Cells.Find("Cidade loja").EntireColumn.Delete
    Cells.Find("Data Desejo Cliente").EntireColumn.Delete
    Cells.Find("Canal de Vendas").EntireColumn.Delete
    Cells.Find("Linha postergada atp").EntireColumn.Delete
    Cells.Find("Ordem postergada atp").EntireColumn.Delete
    Cells.Find("Data ressuprimento").EntireColumn.Delete
    Cells.Find("Cliente").EntireColumn.Delete
    Cells.Find("Tipo pedido").EntireColumn.Delete
End Function


Function formata_planilha()
    '' eliminar linhas desnecessarias
    
    rg = ""
    Do 'METODO DE REPETICAO'
        Set rg = Sheets(1).Cells.Find("Reservada") 'PROCURA A PALAVRA'
        If rg Is Nothing Then 'DECISAO'
            Cells.Find("Ton Cal").EntireColumn.Delete
            Do
                Set rg = Sheets(1).Cells.Find("-") 'PROCURA A PALAVRA'
                If rg Is Nothing Then 'DECISAO'
                Else
                    rg.EntireRow.Delete
                End If
            Loop Until rg Is Nothing   'LOOP'
        Else
            rg.EntireRow.Delete
        End If
    Loop Until rg Is Nothing   'LOOP'
End Function


Function ajustar()
    '' Autofit
    
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
End Function


Function filtro()
    '' Filtro AZ
    
    Cells.Find("Linha").Select
    ActiveCell.CurrentRegion.Sort Key1:=Cells.Find("Linha"), Order1:=xlAscending, Header:=xlYes
    Cells.Find("Nº Pedido").Select
    ActiveCell.CurrentRegion.Sort Key1:=Cells.Find("Nº Pedido"), Order1:=xlAscending, Header:=xlYes
End Function
