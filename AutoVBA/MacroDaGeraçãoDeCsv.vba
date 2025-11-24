Option Explicit

' ===== Função para verificar se há célula vazia em um intervalo =====
Function RangeVazio(ws As Worksheet, rng As Range) As Boolean
    Dim cel As Range
    For Each cel In rng
        If IsEmpty(cel.Value) Then
            RangeVazio = True
            Exit Function
        End If
    Next cel
    RangeVazio = False
End Function

' ===== Adiciona linhas da aba ADD_Linhas na aba CSV =====
Sub ADD_Linha()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim wsCSV As Worksheet
    Dim wsAdd As Worksheet
    Set wsCSV = wb.Sheets("CSV")
    Set wsAdd = wb.Sheets("ADD_Linhas")
    
    ' Cabeçalho fixo
    wsCSV.Range("A1:K1").Value = Array("FaixaInicio", "FaixaFinal", "Intervalo", "Percurso", "TempTerm", "Frota", "Linha", "Dia", "Sentido", "Oso", "LinhaOso")
    
    ' Verifica campos obrigatórios H2:J2
    If RangeVazio(wsAdd, wsAdd.Range("H2:J2")) Then
        MsgBox "Faltam informações das faixas.", vbExclamation
        Exit Sub
    End If
    
    ' Ler dados
    Dim lastAddRow As Long
    lastAddRow = wsAdd.Cells(wsAdd.Rows.Count, "A").End(xlUp).Row
    
    If lastAddRow < 2 Then
        MsgBox "Nenhuma faixa preenchida.", vbExclamation
        Exit Sub
    End If
    
    Dim dados As Variant
    dados = wsAdd.Range("A2:K" & lastAddRow).Value
    
    ' Dados fixos da primeira linha
    Dim linhaValor As Variant, diaValor As Variant, sentidoBruto As Variant
    linhaValor = dados(1, 8)   ' Coluna H
    diaValor = dados(1, 9)     ' Coluna I
    sentidoBruto = dados(1, 10) ' Coluna J
    
    Dim sentidoValor As String
    If VarType(sentidoBruto) = vbString Then
        If InStr(sentidoBruto, "0") > 0 Then
            sentidoValor = "0"
        ElseIf InStr(sentidoBruto, "1") > 0 Then
            sentidoValor = "1"
        ElseIf InStr(sentidoBruto, "2") > 0 Then
            sentidoValor = "2"
        Else
            sentidoValor = Trim(sentidoBruto)
        End If
    Else
        sentidoValor = sentidoBruto
    End If
    
    ' Formatar hora HHMM
    Dim i As Long
    Dim novasLinhas() As Variant
    ReDim novasLinhas(1 To lastAddRow - 1, 1 To 11)
    
    Dim count As Long: count = 0
    
    For i = 1 To UBound(dados, 1)
        ' Ignora linhas vazias
        If IsEmpty(dados(i, 1)) Or IsEmpty(dados(i, 2)) Then GoTo NextRow
        
        count = count + 1
        novasLinhas(count, 1) = FormatHora(dados(i, 1))
        novasLinhas(count, 2) = FormatHora(dados(i, 2))
        novasLinhas(count, 3) = dados(i, 3)
        novasLinhas(count, 4) = dados(i, 4)
        novasLinhas(count, 5) = dados(i, 5)
        novasLinhas(count, 6) = dados(i, 6)
        novasLinhas(count, 7) = linhaValor
        novasLinhas(count, 8) = diaValor
        novasLinhas(count, 9) = sentidoValor
        novasLinhas(count, 10) = "" ' Oso
        novasLinhas(count, 11) = "" ' LinhaOso
NextRow:
    Next i
    
    If count = 0 Then
        MsgBox "Nenhuma linha válida encontrada.", vbExclamation
        Exit Sub
    End If
    
    ' Adiciona após cabeçalho
    Dim ultimaLinhaCSV As Long
    ultimaLinhaCSV = wsCSV.Cells(wsCSV.Rows.Count, "A").End(xlUp).Row
    wsCSV.Range("A" & ultimaLinhaCSV + 1).Resize(count, 11).Value = novasLinhas
    
    ' Limpa ADD_Linhas (mantendo cabeçalho)
    wsAdd.Range("A2:K" & lastAddRow).ClearContents
    
    MsgBox "Linhas adicionadas à aba CSV com sucesso!", vbInformation
End Sub

' ===== Função para formatar hora HHMM =====
Function FormatHora(valor As Variant) As String
    If IsDate(valor) Then
        FormatHora = Format(valor, "HHMM")
    ElseIf IsNumeric(valor) Then
        FormatHora = Right("0000" & valor, 4)
    Else
        ' Caso seja texto tipo 8:00 ou 08:00
        Dim partes() As String
        partes = Split(valor, ":")
        If UBound(partes) >= 1 Then
            FormatHora = Right("0" & partes(0), 2) & partes(1)
        Else
            FormatHora = ""
        End If
    End If
End Function

' ===== Adiciona OSO e LinhaOso =====
Sub ADD_OSO()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim wsCSV As Worksheet
    Dim wsAdd As Worksheet
    Set wsCSV = wb.Sheets("CSV")
    Set wsAdd = wb.Sheets("ADD_OSOs")
    
    ' Cabeçalho fixo
    wsCSV.Range("A1:K1").Value = Array("FaixaInicio", "FaixaFinal", "Intervalo", "Percurso", "TempTerm", "Frota", "Linha", "Dia", "Sentido", "Oso", "LinhaOso")
    
    Dim lastAddRow As Long
    lastAddRow = wsAdd.Cells(wsAdd.Rows.Count, "A").End(xlUp).Row
    If lastAddRow < 2 Then
        MsgBox "Nenhum dado em ADD_OSOs", vbExclamation
        Exit Sub
    End If
    
    Dim dados As Variant
    dados = wsAdd.Range("A2:B" & lastAddRow).Value
    
    Dim novos() As Variant
    ReDim novos(1 To UBound(dados, 1), 1 To 2)
    
    Dim i As Long, count As Long
    count = 0
    
    For i = 1 To UBound(dados, 1)
        Dim oso As String, linha As String
        oso = dados(i, 1)
        linha = dados(i, 2)
        
        ' ignora linha completamente vazia
        If oso = "" And linha = "" Then GoTo NextOSO
        
        ' verifica preenchimento conjunto
        If (oso = "" Xor linha = "") Then
            MsgBox "Linha " & (i + 1) & ": preencha OSO e Linha juntos.", vbExclamation
            GoTo NextOSO
        End If
        
        ' valida OSO (6 dígitos numéricos)
        If Not oso Like "######" Then
            MsgBox "Linha " & (i + 1) & ": OSO inválido (" & oso & ")", vbExclamation
            GoTo NextOSO
        End If
        
        ' valida Linha (4 ou 6 caracteres)
        If Not (Len(linha) = 4 Or Len(linha) = 6) Then
            MsgBox "Linha " & (i + 1) & ": Linha inválida (" & linha & ")", vbExclamation
            GoTo NextOSO
        End If
        
        count = count + 1
        novos(count, 1) = oso
        novos(count, 2) = linha
        
NextOSO:
    Next i
    
    If count = 0 Then
        MsgBox "Nenhum dado válido para adicionar.", vbExclamation
        Exit Sub
    End If
    
    ' Adiciona abaixo das OSOs existentes na aba CSV
    Dim ultimaLinhaCSV As Long
    ultimaLinhaCSV = wsCSV.Cells(wsCSV.Rows.Count, "J").End(xlUp).Row
    Dim inicioNova As Long
    inicioNova = IIf(ultimaLinhaCSV < 2, 2, ultimaLinhaCSV + 1)
    
    wsCSV.Range("J" & inicioNova).Resize(count, 2).Value = novos
    
    ' Limpa ADD_OSOs
    wsAdd.Range("A2:B" & lastAddRow).ClearContents
    
    MsgBox count & " OSO(s) adicionada(s) à aba CSV com sucesso!", vbInformation
End Sub

' ===== Limpar colunas 1 a 9 =====
Sub Limpar_Linha()
    Dim wsCSV As Worksheet
    Set wsCSV = ThisWorkbook.Sheets("CSV")
    
    Dim ultimaLinhaCSV As Long
    ultimaLinhaCSV = wsCSV.Cells(wsCSV.Rows.Count, "A").End(xlUp).Row
    If ultimaLinhaCSV < 2 Then
        MsgBox "Não há dados para limpar.", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Deseja realmente limpar as FAIXAS e demais dados?", vbYesNo + vbQuestion) = vbYes Then
        wsCSV.Range("A2:I" & ultimaLinhaCSV).ClearContents
        MsgBox "Colunas de Linhas e Faixas limpas com sucesso!", vbInformation
    End If
End Sub

' ===== Limpar colunas 10 e 11 =====
Sub Limpar_OSO()
    Dim wsCSV As Worksheet
    Set wsCSV = ThisWorkbook.Sheets("CSV")
    
    Dim ultimaLinhaCSV As Long
    ultimaLinhaCSV = wsCSV.Cells(wsCSV.Rows.Count, "J").End(xlUp).Row
    If ultimaLinhaCSV < 2 Then
        MsgBox "Não há OSOs para limpar.", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Deseja realmente limpar as OSOs?", vbYesNo + vbQuestion) = vbYes Then
        wsCSV.Range("J2:K" & ultimaLinhaCSV).ClearContents
        MsgBox "Colunas OSO e LinhaOso limpas com sucesso!", vbInformation
    End If
End Sub
