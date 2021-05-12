Sub l__VBA_SubstituicaoTributaria()

''''''''''''''RETIRAR COLUNAS DESNCESSESSÁRIAS'''''''''
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

'''''''''''''   MOVER COLUNA VALOR '''''''''''''''
    Cells.Find("Preço Unitário Líquido").Offset(0, 1).EntireColumn.Value = Cells.Find("Peso Bruto").EntireColumn.Value
    Cells.Find("Preço Unitário Líquido").Value = "Preço Unit."
    Cells.Find("Valor Produto").Offset(0, -1).EntireColumn.Value = Cells.Find("Preço Unit.").EntireColumn.Value
    Cells.Find("Peso Bruto").Offset(0, -1).EntireColumn.Value = ""
    Cells.Find("Peso Bruto").Offset(0, -1).Value = "ICMS"
    Cells.Find("Peso Bruto").Offset(0, 1).Value = "Frete"
    Cells.Find("Peso Bruto").Offset(0, 2).Value = "Royalties"
    Cells.Find("Valor Produto").Value = "Valor NF"
    Cells.Find("Valor NF").EntireColumn.Insert
    Cells.Find("Valor NF").Offset(0, -1).Value = "Valor Prod."

''''''''''''''''' DEFINE ULTIMA LINHA DA TABELA ''''''''''''''
    l_fim = Range("A1").End(xlDown).Row

'''''''''''''''' COLOCAR CALCULO VALOR PRODUTO '''''''''''''''''
    Cells.Find("Valor Prod.").Select
    x = 1
    While ActiveCell.Row <> l_fim
        Cells.Find("Valor Prod.").Offset(x, 0).Select
        ActiveCell.Value = (ActiveCell.Offset(0, -1) * ActiveCell.Offset(0, -2))
        x = x + 1
    Wend

'''''''''''''''''''''' COLOCAR CALCULO VALOR ICMS''''''''''''
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
    
''''''''''''''''COLOCAR CALCULO FRETE ''''''''''''''
    Cells.Find("Frete").Select
    x = 1
    While ActiveCell.Row <> l_fim
        Cells.Find("Frete").Offset(x, 0).Select
        If Cells.Find("Tipo Ordem").Offset(x, 0).Value = "27 - Reposição" Then   '' alterado do original
            ActiveCell.Value = (Cells.Find("Peso Bruto").Offset(x, 0).Value / 1000) * 133
        Else
            ActiveCell.Value = (Cells.Find("Peso Bruto").Offset(x, 0).Value / 1000) * 733
        End If
        x = x + 1
    Wend

''''''''''''''''COLOCAR CALCULO ROYALTIES ''''''''''''''
    Cells.Find("Royalties").Select
    x = 1
    While ActiveCell.Row <> l_fim
        Cells.Find("Royalties").Offset(x, 0).Select
        If (Cells.Find("Tipo Ordem").Offset(x, 0).Value = "01-MI - Venda - MI") Or (Cells.Find("Tipo Ordem").Offset(x, 0).Value = "151 - Venda Programada") Then
            ActiveCell.Value = Cells.Find("Valor Prod.").Offset(x, 0).Value * 0.33
        Else
            ActiveCell.Value = "-"
        End If
        x = x + 1
    Wend

''''''''''''''''''' FORMATA CELULAS '''''''''''''''''
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
    Selection.Font.Underline = xlUnderlineStyleSingle
    Selection.Font.Underline = xlUnderlineStyleNone
    
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

'''''''''''''''' SOMA VALOR PROD '''''''''''''
    Cells.Find("Valor Prod.").Select
    ActiveCell.End(xlDown).Offset(1, 0).Interior.Color = 9621584
    ActiveCell.End(xlDown).Offset(1, 0).Value = WorksheetFunction.Sum(Selection.Resize(ActiveCell.End(xlDown).Row, Selection.Columns.Count).Value)

'''''''''''''''' SOMA Valor NF '''''''''''''
    Cells.Find("Valor NF").Select
    ActiveCell.End(xlDown).Offset(1, 0).Interior.Color = 9621584
    ActiveCell.End(xlDown).Offset(1, 0).Value = WorksheetFunction.Sum(Selection.Resize(ActiveCell.End(xlDown).Row, Selection.Columns.Count).Value)

'''''''''''''''' SOMA Frete '''''''''''''
    Cells.Find("Frete").Select
    ActiveCell.End(xlDown).Offset(1, 0).Interior.Color = 9621584
    ActiveCell.End(xlDown).Offset(1, 0).Value = WorksheetFunction.Sum(Selection.Resize(ActiveCell.End(xlDown).Row, Selection.Columns.Count).Value)

'''''''''''''''' SOMA Royalties '''''''''''''
    Cells.Find("Royalties").Select
    ActiveCell.End(xlDown).Offset(1, 0).Interior.Color = 9621584
    ActiveCell.End(xlDown).Offset(1, 0).Value = WorksheetFunction.Sum(Selection.Resize(ActiveCell.End(xlDown).Row, Selection.Columns.Count).Value)

''''''''''' MARCAR CFOP 6949'''''''''''''''
    Cells.Find("CFOP").Select
    x = 1
    While x <> l_fim + 1
        If ActiveCell.Offset(x, 0).Value = "6949" Then
            ActiveCell.Offset(x, 0).EntireRow.Font.Color = vbRed
        End If
            x = x + 1
    Wend

''''''''''''''ORGANIZAR NOTA FISCAL DE A - Z '''''''''''''''''''
Range("A1").CurrentRegion.Sort Key1:=Cells.Find("Linha Nota"), Order1:=xlAscending, Header:=xlYes
Range("A1").CurrentRegion.Sort Key1:=Cells.Find("Nota Fiscal"), Order1:=xlAscending, Header:=xlYes

''''''''''''''AUTOFIT''''''''''''
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    Range("A1").Select
    
'''''''''''''''''NOMEAR ABA ATIVA''''''''''''
    ActiveSheet.Name = "Imposto_Nota"

''''''''''''''CRIA NOVA ABA'''''''''''''''
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Apuracao"

''''''''''''''''' COPIAR PARA APURAÇÃO''''''''''''
    Sheets(1).Select
    Range("A1").CurrentRegion.Copy
    Sheets(2).Select
    Range("A1").PasteSpecial

''''''''''''''''' EXCLUIR CFOP 6949''''''''''''
Cells.Find("CFOP").Select
    x = 1
    While x <> l_fim
        If ActiveCell.Offset(x, 0).Value = "6949" Then
        ActiveCell.Offset(x, 0).EntireRow.Delete
    Else
        x = x + 1
    End If
        
    Wend

''''''''''''''ORGANIZAR DE A - Z '''''''''''''''''''
Range("A1").CurrentRegion.Sort Key1:=Cells.Find("ICMS"), Order1:=xlAscending, Header:=xlYes
Range("A1").CurrentRegion.Sort Key1:=Cells.Find("NCM"), Order1:=xlAscending, Header:=xlYes

''''''''''''''AUTOFIT SHEET 2''''''''''''
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    Range("A1").Select
    
''''''''''''''CRIA NOVA ABA'''''''''''''''
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "ST"

''''''''''''''' COPIAR NCM PARA PLANILHA ST'''''''''''''''''
   Sheets(2).Cells.Find("NCM").EntireColumn.Copy
   Sheets(3).Select
   Range("A1").PasteSpecial
   
''''''''''''''''' FILTRO AZ NCM'''''''''''''''
Range("A1").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
   
''''''''''''''''''' APAGAR NCMS IGUAIS ''''''''''''''''''''''''''
    l_fim = Range("A1").End(xlDown).Row
    Range("A1").Select
    x = 1
    While x <> l_fim
        If ActiveCell.Offset(x, 0).Value = ActiveCell.Offset(x - 1, 0).Value Then
            ActiveCell.Offset(x, 0).EntireRow.Delete
            l_fim = l_fim - 1
        Else
            x = x + 1
        End If
    Wend

'''''''''''''''' INSERIR OUTRAS COLUNAS''''''''''''''''
    Range("B1").Value = "MVA"
    Range("C1").Value = "Valor Prod."
    Range("D1").Value = "Valor NF"
    Range("E1").Value = "ICMS"
    Range("F1").Value = "MVA Aplicada"
    Range("G1").Value = "ICMS (18%)"
    Range("H1").Value = "ICMS/ST"

''''''''''''''''''' ADICIONANDO MVA PELO ARQUIVO MVA.TXT'''''''''''''''''
    '''''''''' criacao das variáveis '''''''''''''''''''''''''
    Dim arquivo As String, texto As String, linhaTexto As String
    Dim l_ncm As String
    user = Environ("USERNAME")
    '''''''''''''' ele especifica uma variável e abre ela''''''''''
    arquivo = "C:\Users\" & user & "\Documents\MVA.txt"
    Open arquivo For Input As #1

    ''' aqui ele coloca todo o texto em uma variavel e fecha arquivo'''
    Do Until EOF(1)
        Line Input #1, linhaTexto
        texto = texto & linhaTexto
    Loop
    Close #1
    
    x = 1
    While x <> l_fim
        ncm = ActiveCell.Offset(x, 0).Value
        l_ncm = InStr(texto, ncm)
        ActiveCell.Offset(x, 1).Value = Mid(texto, l_ncm + 13, 6)
        ActiveCell.Offset(x, 1).NumberFormat = "#.##%"
        x = x + 1
    Wend

'''''''''''''''' SOMA SE '''''''''''''''''''''
    x = 1
    While x <> l_fim
        ncm = ActiveCell.Offset(x, 0).Value
        ActiveCell.Offset(x, 2).Value = WorksheetFunction.SumIf(Sheets(2).Cells.Find("NCM").EntireColumn, ncm, Sheets(2).Cells.Find("Valor Prod.").EntireColumn)
        ActiveCell.Offset(x, 3).Value = WorksheetFunction.SumIf(Sheets(2).Cells.Find("NCM").EntireColumn, ncm, Sheets(2).Cells.Find("Valor NF").EntireColumn)
        x = x + 1
    Wend
    
''''''''''''''''''''' ICMS ''''''''''''''''''''
    Cells.Find("ICMS").Select
    x = 1
    While x <> l_fim
        ActiveCell.Offset(x, 0).Value = ActiveCell.Offset(x, -2).Value - ActiveCell.Offset(x, -1).Value
        If (ActiveCell.Offset(x, 0).Value / ActiveCell.Offset(x, -2)) < (0.06) Then
            ActiveCell.Offset(x, 0).Value = ActiveCell.Offset(x, -2) * 0.04
        End If
        x = x + 1
    Wend

''''''''''''''''''''' MVA Aplicada ''''''''''''''''''''
    Cells.Find("MVA Aplicada").Select
    x = 1
    While x <> l_fim
        If Cells.Find("MVA").Offset(x, 0).Value = 0 Then
            ActiveCell.Offset(x, 0).Value = "-"
        Cells.Find("MVA").Offset(x, 0).Value = "-"
        Else
            ActiveCell.Offset(x, 0).Value = Cells.Find("Valor NF").Offset(x, 0).Value * (1 + Cells.Find("MVA").Offset(x, 0).Value)
        End If
        x = x + 1
    Wend

''''''''''''''''''''' ICMS 18% ''''''''''''''''''''
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

''''''''''''''''ICMS ST ''''''''''''''''''''
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

'''''''''''''''' SOMA VALOR PROD '''''''''''''
    Cells.Find("Valor Prod.").Select
    ActiveCell.End(xlDown).Offset(1, 0).Interior.Color = 9621584
    ActiveCell.End(xlDown).Offset(1, 0).Value = WorksheetFunction.Sum(Selection.Resize(ActiveCell.End(xlDown).Row, Selection.Columns.Count).Value)
    
    ''''''''''''' SOMA VALOR NF ''''''''''''''''
    Cells.Find("Valor NF").Select
    ActiveCell.End(xlDown).Offset(1, 0).Interior.Color = 9621584
    ActiveCell.End(xlDown).Offset(1, 0).Value = WorksheetFunction.Sum(Selection.Resize(ActiveCell.End(xlDown).Row, Selection.Columns.Count).Value)
    
    '''''''''''' VALOR ST''''''''''''
    Cells.Find("ICMS/ST").Select
    ActiveCell.End(xlDown).Offset(1, 0).Interior.Color = 9621584
    ActiveCell.End(xlDown).Offset(1, 0).Value = WorksheetFunction.Sum(Selection.Resize(ActiveCell.End(xlDown).Row, Selection.Columns.Count).Value)

'''''''''''''' CONFIGURA CELULAS PARA MOEDA''''''''''''''''
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
    
'''''''''''''''''''' CONFERE O NUM DE NOTAS ''''''''''''''
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
   Range("K1").Value = cont & " NFs"

''''''''''''''AUTOFIT SHEET 3''''''''''''
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    Range("A1").Select

'''''''''''''' CABECALHO DE CADA ABA '''''''''''''''
    For Each aba In Worksheets
        aba.Activate
        Range("A1").Select
        Selection.Resize(Selection.Rows.Count, Selection.End(xlToRight).Column).Interior.ColorIndex = 44
    Next aba
    Sheets(1).Activate

'''''''''''SALVAR SEM VBA''''''''''''''
    My_Date = Cells.Find("Data Emissão").Offset(1, 0).Value
    ActiveWorkbook.SaveAs Filename:="C:\Users\" & user & "\Downloads\" & Month(My_Date) & " - ST " & Format(My_Date, "mmm") & "_" & Format(My_Date, "yy"), FileFormat:=xlWorkbookDefault
    Fim = MsgBox("Projeto Concluído!!", vbExclamation)
End Sub