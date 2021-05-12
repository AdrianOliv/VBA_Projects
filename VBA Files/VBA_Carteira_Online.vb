Sub l___VBA__CarteiraOnline()

''''''''''''''''''' FORMATAR CELULAS ''''''''''''''''''
    Range("A1").CurrentRegion.Select
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

'''''''''''''''   RETIRA COLUNAS DESNECESSARIAS '''''''''''''''
    Cells.Find("Espelho").EntireColumn.Delete
    Cells.Find("Carteira").EntireColumn.Delete
    Cells.Find("Data Mínima Faturamento ").EntireColumn.Delete
    Cells.Find("Canal de Vendas").EntireColumn.Delete
    Cells.Find("Cidade Loja").EntireColumn.Delete
    Cells.Find("Data Criação Linha").EntireColumn.Delete
    Cells.Find("Unid.").EntireColumn.Delete
    Cells.Find("Retenção de Credito").EntireColumn.Delete
    Cells.Find("Outras Retenções ").EntireColumn.Delete
    Cells.Find("Faturamento Parcial").EntireColumn.Delete
    Cells.Find("O.C. Linha").EntireColumn.Delete
    Cells.Find("O.C. Cliente").EntireColumn.Delete
    Cells.Find("Data Desejada Entrega").EntireColumn.Delete
    Cells.Find("Transportadora").EntireColumn.Delete
    Cells.Find("Embarque Expedição").EntireColumn.Delete
    Cells.Find("Data percurso").EntireColumn.Delete
    Cells.Find("Data prev embarque").EntireColumn.Delete

'''''''''''''   ORGANIZAR COLUNA LOTE E VALOR   '''''''''''''''
    Cells.Find("Prazo Médio ").EntireColumn.Value = Cells.Find("Ton cal").EntireColumn.Value
    Cells.Find("Ton cal").Offset(0, 1).EntireColumn.Value = Cells.Find("Valor").EntireColumn.Value
    Range("A1").Select
    Cells.Find("Valor").EntireColumn.Delete

''''''''''''''''  AUTOFIT   ''''''''''''''
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit

End Sub