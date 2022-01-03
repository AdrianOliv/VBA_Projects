Sub l__Subs_Trib()
    
    retira_colunas
    coluna_valor
    l_fim = Range("A1").End(xlDown).Row

    valor_produto (l_fim)
    icms (l_fim)
    formatar
    soma
    cfop_ast (l_fim)
    filtro
    autofitt
    aba_apuracao (l_fim)
    aba_st
    mva
    valores
    formata_st
    num_notas
    autofitt
    finalizacao

End Sub



Function retira_colunas()
    '' Retirar colunas desnecessárias

    Cells.Find("Depósito").EntireColumn.Delete
    Cells.Find("Data Saída").EntireColumn.Delete
    Cells.Find("Condição Pagamento").EntireColumn.Delete
    Cells.Find("Prazo Médio").EntireColumn.Delete
    Cells.Find("Canal Venda").EntireColumn.Delete
    Cells.Find("CNPJ").EntireColumn.Delete
    Cells.Find("Código Cliente").EntireColumn.Delete
    Cells.Find("Nome Cliente").EntireColumn.Delete
    Cells.Find("Cidade").EntireColumn.Delete
    Cells.Find("UF").EntireColumn.Delete
    Cells.Find("Tipo Contribuinte").EntireColumn.Delete
    Cells.Find("Grupo Cliente").EntireColumn.Delete
    Cells.Find("Mercado").EntireColumn.Delete
    Cells.Find("Formato").EntireColumn.Delete
    Cells.Find("Acabamento").EntireColumn.Delete
    Cells.Find("Peso Líquido").EntireColumn.Delete
    Cells.Find("Valor Base ICMS Normal").EntireColumn.Delete
    Cells.Find("%ICMS Normal").EntireColumn.Delete
    Cells.Find("Valor ICMS Normal").EntireColumn.Delete
    Cells.Find("Valor Base ICMS ST").EntireColumn.Delete
    Cells.Find("%ICMS ST").EntireColumn.Delete
    Cells.Find("Valor ICMS ST").EntireColumn.Delete
    Cells.Find("Valor Base IPI").EntireColumn.Delete
    Cells.Find("%IPI").EntireColumn.Delete
    Cells.Find("Valor IPI").EntireColumn.Delete
    Cells.Find("Valor Base PIS").EntireColumn.Delete
    Cells.Find("%PIS").EntireColumn.Delete
    Cells.Find("Valor PIS").EntireColumn.Delete
    Cells.Find("Valor Base COFINS").EntireColumn.Delete
    Cells.Find("%COFINS").EntireColumn.Delete
    Cells.Find("Valor COFINS").EntireColumn.Delete
    Cells.Find("Preço Unitário").EntireColumn.Delete
    Cells.Find("Valor Líquido").EntireColumn.Delete
End Function


Function coluna_valor()
    '' Ajusta a coluna valor para calculos

    Cells.Find("Preço Unitário Líquido").Value = "Preço Unit."
    Cells.Find("Valor Produto").Offset(0, -1).EntireColumn.Value = Cells.Find("Preço Unit.").EntireColumn.Value
    Cells.Find("Valor Produto").Offset(0, 1).EntireColumn.Value = ""
    Cells.Find("Valor Produto").Offset(0, 1).Value = "ICMS"
    Cells.Find("Valor Produto").Value = "Valor NF"
    Cells.Find("Valor NF").EntireColumn.Insert
    Cells.Find("Valor NF").Offset(0, -1).Value = "Valor Prod."
End Function


Function valor_produto(l_fim)
    '' Calcula valor do produto

    Cells.Find("Valor Prod.").Select
    x = 1
    While ActiveCell.Row <> l_fim
        Cells.Find("Valor Prod.").Offset(x, 0).Select
        ActiveCell.Value = (ActiveCell.Offset(0, -1) * ActiveCell.Offset(0, -2))
        x = x + 1
    Wend
End Function


Function icms(l_fim)
    '' Calcular o valor do ICMS

    Cells.Find("ICMS").Select
    x = 1
    While x <> l_fim
        ActiveCell.Offset(x, 0).Value = 1 - ((Cells.Find("Valor NF").Offset(x, 0).Value) / (Cells.Find("Valor Prod.").Offset(x, 0).Value))
        ActiveCell.Offset(x, 0).NumberFormat = "#%"
        If ActiveCell.Offset(x, 0).Value > (0.06) And ActiveCell.Offset(x, 0).Value < (0.075) Then
            ActiveCell.Offset(x, 0).Value = 0.07
        Else
            ActiveCell.Offset(x, 0).Value = 0.04
            Cells.Find("NCM").Offset(x, 0).Value = Cells.Find("NCM").Offset(x, 0).Value & "(4%)"
        End If
        x = x + 1
    Wend
End Function


Function formatar()
    '' Formata as células

    ActiveSheet.Range("a1").CurrentRegion.Select
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

    '' BORDAS
    For k = 5 To 12
        If k > 6 Then
            With Selection.Borders.Item(k)
            .LineStyle = xlContinuos
            .ColorIndex = 0
            .Weight = xlThin
            End With
        End If
    Next k

    '' RENOMEAR ABA ATIVA
    ActiveSheet.Name = "Imposto_Nota"
End Function


Function soma()
    '' Soma dos valores de Produto e NF

    '' VALOR PRODUTO
    Cells.Find("Valor Prod.").Select
    ActiveCell.End(xlDown).Offset(1, 0).Interior.Color = 9621584
    ActiveCell.End(xlDown).Offset(1, 0).Value = WorksheetFunction.Sum(Selection.Resize(ActiveCell.End(xlDown).Row, Selection.Columns.Count).Value)

    '' VALOR NOTA FISCAL
    Cells.Find("Valor NF").Select
    ActiveCell.End(xlDown).Offset(1, 0).Interior.Color = 9621584
    ActiveCell.End(xlDown).Offset(1, 0).Value = WorksheetFunction.Sum(Selection.Resize(ActiveCell.End(xlDown).Row, Selection.Columns.Count).Value)
End Function


Function cfop_ast(l_fim)
    '' Marca as notas fiscal de assistência técnica

    Cells.Find("CFOP").Select
    x = 1
    While x <> l_fim + 1
        If ActiveCell.Offset(x, 0).Value = "6949" Then
            ActiveCell.Offset(x, 0).EntireRow.Font.Color = vbRed
        End If
            x = x + 1
    Wend
End Function


Function filtro()
    '' Filtro AZ

    Range("A1").CurrentRegion.Sort Key1:=Cells.Find("Linha Nota"), Order1:=xlAscending, Header:=xlYes
    Range("A1").CurrentRegion.Sort Key1:=Cells.Find("Nota Fiscal"), Order1:=xlAscending, Header:=xlYes
End Function


Function autofitt()
    '' Autofit

    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    Range("A1").Select
End Function


Function aba_apuracao(l_fim)
    '' Formatar nova planilha de apuração

    '' CRIA NOVA ABA
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Apuracao"

    '' COPIAR INFORMAÇÕES
    Sheets(1).Range("A1").CurrentRegion.Copy
    Range("A1").PasteSpecial
    
    '' EXCLUIR CFOP 6949
    Cells.Find("CFOP").Select
    x = 1
    While x <> l_fim
        If ActiveCell.Offset(x, 0).Value = "6949" Then
            ActiveCell.Offset(x, 0).EntireRow.Delete
        Else
            x = x + 1
        End If
    Wend
    
    '' EXCLUIR Amostras
    Cells.Find("Tipo Ordem").Select
    x = 1
    While x <> l_fim
        If ActiveCell.Offset(x, 0).Value = "50-Ambiente p/Exposição PBS" Then
            ActiveCell.Offset(x, 0).EntireRow.Delete
        Else
            x = x + 1
        End If
    Wend
    
    autofitt
End Function


Function aba_st()
    '' Criação da planilha de ST

    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "ST"

    '' COPIAR NCM
    Sheets(2).Cells.Find("NCM").EntireColumn.Copy
    Range("A1").PasteSpecial
   
    '' FILTRO AZ
    Range("A1").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
   
    '' APAGAR NCMS IGUAIS
    linha_fim = Range("A1").End(xlDown).Row
    ActiveSheet.Range("$A$1:$A$" & linha_fim).RemoveDuplicates Columns:=1, Header:= _
    xlNo

    '' INSERIR OUTRAS COLUNAS
    Range("B1").Value = "MVA"
    Range("C1").Value = "Valor Prod."
    Range("D1").Value = "Valor NF"
    Range("E1").Value = "ICMS"
    Range("F1").Value = "MVA Aplicada"
    Range("G1").Value = "ICMS (18%)"
    Range("H1").Value = "ICMS/ST"
End Function


Function mva()
    '' Adicionando MVA - arquivo TXT

    '' Criacao das variáveis
    Dim arquivo As String, texto As String, linhaTexto As String
    Dim l_ncm As String
    user = Environ("USERNAME")

    '' Especifica uma variável e abre ela
    arquivo = "C:\Users\" & user & "\Documents\Não Apagar - VBA\MVA.txt"
    Open arquivo For Input As #1

    '' Coloca todo o texto em uma variavel e fecha arquivo
    Do Until EOF(1)
        Line Input #1, linhaTexto
        texto = texto & linhaTexto
    Loop
    Close #1
    
    l_fim = Range("a1").End(xlDown).Row
    x = 1
    While x <> l_fim
        ncm = ActiveCell.Offset(x, 0).Value
        l_ncm = InStr(texto, ncm)
        ActiveCell.Offset(x, 1).Value = Mid(texto, l_ncm + 13, 6)
        ActiveCell.Offset(x, 1).NumberFormat = "#.##%"
        x = x + 1
    Wend
End Function


Function valores()
    '' Adicionar os valores de Produto, Nota Fiscal
    '' ICMS, MVA Aplicada, ICMS 18% e ST

    l_fim = Range("A1").End(xlDown).Row

    '' VALOR PRODUTO E NF
    x = 1
    While x <> l_fim
        ncm = ActiveCell.Offset(x, 0).Value
        ActiveCell.Offset(x, 2).Value = WorksheetFunction.SumIf(Sheets(2).Cells.Find("NCM").EntireColumn, ncm, Sheets(2).Cells.Find("Valor Prod.").EntireColumn)
        ActiveCell.Offset(x, 3).Value = WorksheetFunction.SumIf(Sheets(2).Cells.Find("NCM").EntireColumn, ncm, Sheets(2).Cells.Find("Valor NF").EntireColumn)
        x = x + 1
    Wend
    
    '' ICMS
    Cells.Find("ICMS").Select
    x = 1
    While x <> l_fim
        ActiveCell.Offset(x, 0).Value = ActiveCell.Offset(x, -2).Value - ActiveCell.Offset(x, -1).Value
        If (ActiveCell.Offset(x, 0).Value / ActiveCell.Offset(x, -2)) < (0.06) Then
            ActiveCell.Offset(x, 0).Value = ActiveCell.Offset(x, -2) * 0.04
        End If
        x = x + 1
    Wend

    '' MVA Aplicada
    Cells.Find("MVA Aplicada").Select
    x = 1
    While x <> l_fim
        If Cells.Find("MVA").Offset(x, 0).Value = 0 Or IsNumeric(Cells.Find("MVA").Offset(x, 0).Value) = False Then
            ActiveCell.Offset(x, 0).Value = "-"
        Cells.Find("MVA").Offset(x, 0).Value = "-"
        Else
            ActiveCell.Offset(x, 0).Value = Cells.Find("Valor NF").Offset(x, 0).Value * (1 + Cells.Find("MVA").Offset(x, 0).Value)
        End If
        x = x + 1
    Wend

    '' ICMS 18%
    Cells.Find("ICMS (18%)").Select
    x = 1
    While x <> l_fim
        If ActiveCell.Offset(x, -1).Value = "-" Then
            ActiveCell.Offset(x, 0).Value = "-"
        Else
            ActiveCell.Offset(x, 0).Value = ActiveCell.Offset(x, -1).Value * 0.18
        End If
        x = x + 1
    Wend

    '' ICMS ST
    Cells.Find("ICMS/ST").Select
    x = 1
    While x <> l_fim
        If Cells.Find("ICMS (18%)").Offset(x, 0).Value = "-" Then
            ActiveCell.Offset(x, 0).Value = "-"
        Else
            ActiveCell.Offset(x, 0).Value = Cells.Find("ICMS (18%)").Offset(x, 0).Value - Cells.Find("ICMS").Offset(x, 0).Value
        End If
        x = x + 1
    Wend
End Function


Function formata_st()
    '' Formatar a planilha de ST

    '' BORDAS
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

    '' SOMA VALOR PROD
    Cells.Find("Valor Prod.").Select
    ActiveCell.End(xlDown).Offset(1, 0).Interior.Color = 9621584
    ActiveCell.End(xlDown).Offset(1, 0).Value = WorksheetFunction.Sum(Selection.Resize(ActiveCell.End(xlDown).Row, Selection.Columns.Count).Value)
    
    '' SOMA VALOR NF
    Cells.Find("Valor NF").Select
    ActiveCell.End(xlDown).Offset(1, 0).Interior.Color = 9621584
    ActiveCell.End(xlDown).Offset(1, 0).Value = WorksheetFunction.Sum(Selection.Resize(ActiveCell.End(xlDown).Row, Selection.Columns.Count).Value)
    
    '' VALOR ST
    Cells.Find("ICMS/ST").Select
    ActiveCell.End(xlDown).Offset(1, 0).Interior.Color = 9621584
    ActiveCell.End(xlDown).Offset(1, 0).Value = WorksheetFunction.Sum(Selection.Resize(ActiveCell.End(xlDown).Row, Selection.Columns.Count).Value)

    '' CONFIGURA CELULAS PARA MOEDA
    Cells.Find("Valor Prod.").Select
    Selection.Resize(ActiveCell.End(xlDown).Row, ActiveCell.EntireColumn.Count).Style = "Currency"
    Cells.Find("Valor NF").Select
    Selection.Resize(ActiveCell.End(xlDown).Row, ActiveCell.EntireColumn.Count).Style = "Currency"
    Cells.Find("ICMS").Select
    Selection.Resize(ActiveCell.End(xlDown).Row, ActiveCell.EntireColumn.Count).Style = "Currency"
    Cells.Find("MVA Aplicada").Select
    Selection.Resize(ActiveCell.End(xlDown).Row, ActiveCell.EntireColumn.Count).Style = "Currency"
    Cells.Find("ICMS (18%)").Select
    Selection.Resize(ActiveCell.End(xlDown).Row, ActiveCell.EntireColumn.Count).Style = "Currency"
    Cells.Find("ICMS/ST").Select
    Selection.Resize(ActiveCell.End(xlDown).Row, ActiveCell.EntireColumn.Count).Style = "Currency"
End Function


Function num_notas()
    '' Confere o número de notas fiscais

    Sheets(1).Activate
    Cells.Find("Nota Fiscal").Select
    Selection.Resize(ActiveCell.End(xlDown).Row, Selection.Columns.Count).Copy
    Sheets(3).Activate
    Range("K1").PasteSpecial
 
    l_fim = Range("K1").End(xlDown).Row
    Range("K1").Select
    x = 1
    cont = 0
    While x <> l_fim
        If ActiveCell.Offset(x, 0).Value <> ActiveCell.Offset(x - 1, 0).Value Then
            cont = cont + 1
        End If
        x = x + 1
    Wend
    Range("K1").EntireColumn.Delete
    Range("K1").Value = cont & " Notas Fiscais"
End Function


Function finalizacao()
    '' Finalizar trabalho

    '' CABECALHO DE CADA ABA
    For Each aba In Worksheets
        aba.Activate
        Range("A1").Select
        Selection.Resize(Selection.Rows.Count, Selection.End(xlToRight).Column).Interior.ColorIndex = 43
        ActiveWindow.DisplayGridlines = False
    Next aba
    Sheets(1).Activate

    '' SALVAR SEM VBA
    My_Date = Cells.Find("Data Emissão").Offset(1, 0).Value
    user = Environ("USERNAME")
    ActiveWorkbook.SaveAs Filename:="C:\Users\" & user & "\Downloads\" & Month(My_Date) & " - ST " & Format(My_Date, "mmm") & "_" & Format(My_Date, "yy"), FileFormat:=xlWorkbookDefault
    Fim = MsgBox("Programa Finalizado!!", vbExclamation)
End Function