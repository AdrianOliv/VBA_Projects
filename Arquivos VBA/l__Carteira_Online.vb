Sub l__Carteira_Online()

    '' Barra de Progresso
    Dim total As Long
    Dim largura As Long
    Dim percent As Double
    
    '' Bloqueia atualização de tela para agilizar processo
    Application.ScreenUpdating = False

    total = 7
    Progresso.Show
    largura = Progresso.Evolucao.Width
    
    For i = 1 To total
        Select Case i
            Case 1
                exclui_colunas
            Case 2
                formata_celulas
            Case 3
                reverte_colunas
            Case 4
                organiza_ordens
            Case 5
                organiza_reservas
            Case 6
                filtro_az
            Case 7
                ajustar
        End Select
        
        DoEvents
        percent = i / total
        Progresso.Evolucao.Width = percent * largura
        Progresso.Valor = Round(percent * 100, 1) & "%"
    Next
    
    '' Libera atualização de tela para agilizar processo
    Application.ScreenUpdating = True
    Unload Progresso
    msg = MsgBox("Concluído", vbInformation)
    
End Sub



Function ajustar()
    ' ajustes finais
    
    ActiveCell.Offset(1, 0).EntireRow.Delete
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
End Function


Function exclui_colunas()
    ' exclui as colunas desnecessárias
    
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
End Function


Function filtro_az()
    '' filtro AZ das ordens

    Range("A1000").End(xlUp).Select
    
    While ActiveCell.Row <> 1:
        ActiveCell.End(xlUp).Offset(0, 2).Select
        ActiveCell.CurrentRegion.Sort Key1:=Cells.Find("Linha"), Order1:=xlAscending, Header:=xlNo
        ActiveCell.Offset(0, -2).Select
        ActiveCell.CurrentRegion.Sort Key1:=Cells.Find("Nº Pedido"), Order1:=xlAscending, Header:=xlNo
        ActiveCell.End(xlUp).Select
    Wend
End Function


Function formata_celulas()
    ' formata a visualização das celulas
    
    Range("A1").CurrentRegion.Select
    Selection.Replace _
    What:="", Replacement:="-", _
    SearchOrder:=xlByColumns, MatchCase:=True
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


Function organiza_ordens()
    '' Organiza pelos tipos de ordem

    Range("I1").CurrentRegion.Sort Key1:=Cells.Find("Tipo de Pedido"), Order1:=xlDescending, Header:=xlYes
    Range("I2").Select
    While IsEmpty(ActiveCell.Value) <> True
        If ActiveCell.Offset(-1, 0).Value <> ActiveCell.Value Then
            ActiveCell.EntireRow.Insert
            ActiveCell.Offset(1, 0).Select
        End If
        ActiveCell.Offset(1, 0).Select
    Wend
End Function


Function organiza_reservas()
    '' organiza as ordens reservadas

    Cells.Find("01-MI - Venda - MI").Offset(0, 1).Select
    ActiveCell.CurrentRegion.Sort Key1:=Cells.Find("Status Reserva"), Order1:=xlDescending, Header:=xlNo
    ActiveCell.Offset(1, 0).Select
    Cells.Find("RESERVADO", ActiveCell).EntireRow.Insert
End Function


Function reverte_colunas()
    ' troca as colunas lote e valor de lugar
    
    Cells.Find("Prazo Médio ").EntireColumn.Value = Cells.Find("Ton cal").EntireColumn.Value
    Cells.Find("Ton cal").Offset(0, 1).EntireColumn.Value = Cells.Find("Valor").EntireColumn.Value
    Range("A1").Select
    Cells.Find("Valor").EntireColumn.Delete
End Function


Function sleep(x)
    '' cria um tempo de espera de x segundos
    '' 26.08.23 - Função removida do programa principal
    
    newHour = Hour(Now())
    newMinute = Minute(Now())
    newSecond = Second(Now()) + x
    waitTime = TimeSerial(newHour, newMinute, newSecond)
    Application.Wait waitTime
End Function