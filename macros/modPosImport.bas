Attribute VB_Name = "modPosImport"
' =====================================================================
' modPosImport
' Pos-processamento apos o import: executa o import por sheet, pinta a
' descricao, calcula o peso, confere o total da aba (C58) contra o
' total da origem, ajusta a fonte e remove linhas vazias.
' A conferencia so roda quando a matriz de pedidos e passada e o total
' nao esta desativado (V1 <> 0); no uso manual por botao ela e ignorada.
' A caixa de dialogo de divergencia aparece uma unica vez ao final.
' =====================================================================
Option Explicit

Private Const POS_PRIMEIRA_LINHA As Long = 12
Private Const POS_ULTIMA_LINHA As Long = 53
Private Const POS_COLUNA_TEXTO As String = "D"
Private Const POS_COLUNA_QUANTIDADE As String = "C"
Private Const POS_CELULA_TOTAL_PEDIDO As String = "C58"
Private Const POS_EXECUTAR_IMPORT As Boolean = True
Private Const POS_CALCULAR_PESO As Boolean = True
Private Const POS_SHEET_MODELO As String = "template"

' True replica o comportamento original (Rows.Delete). Atencao: isso
' sobe as linhas 54+ e desloca a fonte W11:W56 numa reexecucao.
Private Const POS_EXCLUIR_LINHA_INTEIRA As Boolean = True

Public Sub ExecutarImportEPintar(Optional ByVal matrizPedidos As Variant, Optional ByVal conferirTotais As Boolean = True)
    Dim nomesSelecionados As Variant
    nomesSelecionados = NomesSheetsSelecionadas()
    If IsEmpty(nomesSelecionados) Then Exit Sub

    ProcessarNomesSheets nomesSelecionados, True, matrizPedidos, conferirTotais
End Sub

Public Sub ExecutarImportEPintarTodasExcetoTemplate(Optional ByVal matrizPedidos As Variant, Optional ByVal conferirTotais As Boolean = True)
    Dim nomesVisiveis As Variant
    nomesVisiveis = NomesSheetsVisiveis()
    If IsEmpty(nomesVisiveis) Then Exit Sub

    ProcessarNomesSheets nomesVisiveis, False, matrizPedidos, conferirTotais
End Sub

Private Sub ProcessarNomesSheets(ByVal nomesSheets As Variant, ByVal usarSelecaoMultipla As Boolean, Optional ByVal matrizPedidos As Variant, Optional ByVal conferirTotais As Boolean = True)
    Dim nomeSheetOrigem As String
    nomeSheetOrigem = NomeSheetAtiva()

    Dim registrosConferencia As Collection
    Set registrosConferencia = New Collection

    EntrarModoRapido
    On Error GoTo Erro

    Dim nomeSheet As Variant
    For Each nomeSheet In nomesSheets
        ProcessarSheetPorNome CStr(nomeSheet), matrizPedidos, registrosConferencia, conferirTotais
    Next nomeSheet

    SairModoRapido
    RestaurarSelecao nomesSheets, nomeSheetOrigem, usarSelecaoMultipla

    If ConferenciaTemDivergencia(registrosConferencia) Then
        MsgBox MontarRelatorioConferencia(registrosConferencia), vbExclamation, "Conferencia de totais"
    End If
    Exit Sub

Erro:
    AbortarModoRapido
    MsgBox "Falha no pos-import: " & Err.Description, vbCritical
End Sub

Private Sub ProcessarSheetPorNome(ByVal nomeSheet As String, ByVal matrizPedidos As Variant, registrosConferencia As Collection, ByVal conferirTotais As Boolean)
    Dim sheetPlanilha As Worksheet
    Set sheetPlanilha = ObterWorksheet(nomeSheet)
    If sheetPlanilha Is Nothing Then Exit Sub
    If StrComp(sheetPlanilha.Name, POS_SHEET_MODELO, vbTextCompare) = 0 Then Exit Sub

    ProcessarSheetImportPinturaPeso sheetPlanilha, matrizPedidos, registrosConferencia, conferirTotais
End Sub

Private Function ProcessarSheetImportPinturaPeso(sheetPlanilha As Worksheet, ByVal matrizPedidos As Variant, registrosConferencia As Collection, ByVal conferirTotais As Boolean) As Boolean
    On Error GoTo Erro

    If POS_EXECUTAR_IMPORT Then
        If Not SubstituirNaSheet(sheetPlanilha) Then Exit Function
    End If

    PintarSheet sheetPlanilha
    If POS_CALCULAR_PESO Then AtualizarPeso sheetPlanilha
    RegistrarConferenciaTotal sheetPlanilha, matrizPedidos, registrosConferencia, conferirTotais
    AjustarFonteCelulaPedido sheetPlanilha
    RemoverLinhasVaziasSomenteComItens sheetPlanilha

    ProcessarSheetImportPinturaPeso = True
    Exit Function

Erro:
    Debug.Print "Sheet " & sheetPlanilha.Name & " ignorada: " & Err.Description
End Function

' Le C58 antes da exclusao de linhas: a exclusao desloca a celula e
' mudaria o endereco de leitura.
Private Sub RegistrarConferenciaTotal(sheetPlanilha As Worksheet, ByVal matrizPedidos As Variant, registrosConferencia As Collection, ByVal conferirTotais As Boolean)
    If Not conferirTotais Then Exit Sub
    If IsEmpty(matrizPedidos) Then Exit Sub

    Dim totalOrigem As Variant
    totalOrigem = TotalOrigemDaSheet(matrizPedidos, sheetPlanilha.Name)
    If IsEmpty(totalOrigem) Then Exit Sub

    RecalcularSheet sheetPlanilha

    Dim totalPlanilha As Double
    totalPlanilha = ConverterNumero(sheetPlanilha.Range(POS_CELULA_TOTAL_PEDIDO).Value2)

    registrosConferencia.Add NovoRegistroConferencia(sheetPlanilha.Name, CDbl(totalOrigem), totalPlanilha)
End Sub

Private Function TotalOrigemDaSheet(ByVal matrizPedidos As Variant, ByVal nomeSheet As String) As Variant
    Dim indice As Long
    For indice = LBound(matrizPedidos, 2) To UBound(matrizPedidos, 2)
        If StrComp(MontarNomeSheet(CStr(matrizPedidos(IDX_PEDIDO, indice))), nomeSheet, vbTextCompare) = 0 Then
            TotalOrigemDaSheet = matrizPedidos(IDX_TOTAL, indice)
            Exit Function
        End If
    Next indice
End Function

Private Sub AtualizarPeso(sheetPlanilha As Worksheet)
    RecalcularSheet sheetPlanilha
    CalcularPesoPortasNaSheet sheetPlanilha
End Sub

Private Sub PintarSheet(sheetAtual As Worksheet)
    Dim intervaloTexto As Range
    Set intervaloTexto = sheetAtual.Range( _
        POS_COLUNA_TEXTO & POS_PRIMEIRA_LINHA & ":" & _
        POS_COLUNA_TEXTO & POS_ULTIMA_LINHA)

    ColorirPalavrasNoIntervalo intervaloTexto
End Sub

Private Sub RemoverLinhasVaziasSomenteComItens(sheetAtual As Worksheet)
    If Not TemQuantidadeNoIntervalo(sheetAtual) Then Exit Sub

    Dim linhasVazias As Range
    Set linhasVazias = LinhasSemQuantidade(sheetAtual)
    If linhasVazias Is Nothing Then Exit Sub

    ExcluirLinhas linhasVazias
End Sub

Private Sub ExcluirLinhas(linhasVazias As Range)
    If POS_EXCLUIR_LINHA_INTEIRA Then
        linhasVazias.EntireRow.Delete
        Exit Sub
    End If

    linhasVazias.Delete Shift:=xlUp
End Sub

' Union em 1 chamada de Delete: antes eram ate 42 exclusoes encadeadas,
' cada uma recalculando e deslocando a planilha inteira.
Private Function LinhasSemQuantidade(sheetAtual As Worksheet) As Range
    Dim quantidades As Variant
    quantidades = LerQuantidades(sheetAtual)

    Dim linhasVazias As Range
    Dim indice As Long

    For indice = LBound(quantidades, 1) To UBound(quantidades, 1)
        If Len(ConverterTexto(quantidades(indice, 1))) = 0 Then
            Set linhasVazias = UnirFaixaItens(linhasVazias, sheetAtual, POS_PRIMEIRA_LINHA + indice - 1)
        End If
    Next indice

    Set LinhasSemQuantidade = linhasVazias
End Function

Private Function UnirFaixaItens(faixaAtual As Range, sheetAtual As Worksheet, ByVal linha As Long) As Range
    Dim faixaLinha As Range
    Set faixaLinha = sheetAtual.Range("B" & linha & ":M" & linha)

    If faixaAtual Is Nothing Then
        Set UnirFaixaItens = faixaLinha
        Exit Function
    End If

    Set UnirFaixaItens = Application.Union(faixaAtual, faixaLinha)
End Function

Private Function TemQuantidadeNoIntervalo(sheetAtual As Worksheet) As Boolean
    Dim quantidades As Variant
    quantidades = LerQuantidades(sheetAtual)

    Dim indice As Long
    For indice = LBound(quantidades, 1) To UBound(quantidades, 1)
        If Len(ConverterTexto(quantidades(indice, 1))) > 0 Then
            TemQuantidadeNoIntervalo = True
            Exit Function
        End If
    Next indice
End Function

Private Function LerQuantidades(sheetAtual As Worksheet) As Variant
    LerQuantidades = sheetAtual.Range( _
        POS_COLUNA_QUANTIDADE & POS_PRIMEIRA_LINHA & ":" & _
        POS_COLUNA_QUANTIDADE & POS_ULTIMA_LINHA).Value2
End Function

Private Function NomesSheetsSelecionadas() As Variant
    Dim total As Long
    total = ActiveWindow.SelectedSheets.Count
    If total = 0 Then Exit Function

    Dim nomes() As String
    ReDim nomes(0 To total - 1)

    Dim indice As Long
    For indice = 1 To total
        nomes(indice - 1) = ActiveWindow.SelectedSheets(indice).Name
    Next indice

    NomesSheetsSelecionadas = nomes
End Function

Private Function NomesSheetsVisiveis() As Variant
    Dim nomes() As String
    ReDim nomes(0 To ThisWorkbook.Worksheets.Count - 1)

    Dim contador As Long
    Dim sheetPlanilha As Worksheet

    For Each sheetPlanilha In ThisWorkbook.Worksheets
        contador = contador + AdicionarSeVisivel(nomes, contador, sheetPlanilha)
    Next sheetPlanilha

    If contador = 0 Then Exit Function
    ReDim Preserve nomes(0 To contador - 1)
    NomesSheetsVisiveis = nomes
End Function

Private Function AdicionarSeVisivel(ByRef nomes() As String, ByVal contador As Long, sheetPlanilha As Worksheet) As Long
    If sheetPlanilha.Visible <> xlSheetVisible Then Exit Function
    If StrComp(sheetPlanilha.Name, POS_SHEET_MODELO, vbTextCompare) = 0 Then Exit Function

    nomes(contador) = sheetPlanilha.Name
    AdicionarSeVisivel = 1
End Function

Private Function ObterWorksheet(ByVal nomeSheet As String) As Worksheet
    Dim sheetEncontrada As Object

    On Error Resume Next
    Set sheetEncontrada = ThisWorkbook.Sheets(nomeSheet)
    On Error GoTo 0

    If sheetEncontrada Is Nothing Then Exit Function
    If Not TypeOf sheetEncontrada Is Worksheet Then Exit Function

    Set ObterWorksheet = sheetEncontrada
End Function

Private Function NomeSheetAtiva() As String
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Function
    NomeSheetAtiva = ActiveSheet.Name
End Function

Private Sub RestaurarSelecao(ByVal nomesSheets As Variant, ByVal nomeSheetOrigem As String, ByVal usarSelecaoMultipla As Boolean)
    On Error Resume Next

    If Not usarSelecaoMultipla Then
        ThisWorkbook.Worksheets(nomeSheetOrigem).Select
        Exit Sub
    End If

    SelecionarNomes nomesSheets
End Sub

Private Sub SelecionarNomes(ByVal nomesSheets As Variant)
    If UBound(nomesSheets) < 0 Then Exit Sub

    If UBound(nomesSheets) = 0 Then
        ThisWorkbook.Sheets(CStr(nomesSheets(0))).Select
        Exit Sub
    End If

    ThisWorkbook.Sheets(nomesSheets).Select
End Sub

