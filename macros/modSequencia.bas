Attribute VB_Name = "modSequencia"
' =====================================================================
' modSequencia
' Aplica a sequencia de carregamento em P6 e a descricao em P9,
' incluindo cor e negrito.
' O "On Error Resume Next" por operacao foi trocado por verificacao
' previa de protecao: erro de escrita agora aparece em vez de sumir.
' =====================================================================
Option Explicit

Private Const SEQ_CELULA_SEQUENCIA As String = "P6"
Private Const SEQ_CELULA_DESCRICAO As String = "P9"
Private Const SEQ_MINIMO As Long = 1
Private Const SEQ_MAXIMO As Long = 10

Public Sub AplicarSequencia(sheetDestino As Worksheet, ByVal sequencia As Long)
    If sheetDestino Is Nothing Then Exit Sub

    If sheetDestino.ProtectContents Then
        Debug.Print "Sequencia nao aplicada: " & sheetDestino.Name & " esta protegida."
        Exit Sub
    End If

    Dim celulaSequencia As Range
    Dim celulaDescricao As Range
    Set celulaSequencia = ObterCelulaOperativa(sheetDestino.Range(SEQ_CELULA_SEQUENCIA))
    Set celulaDescricao = ObterCelulaOperativa(sheetDestino.Range(SEQ_CELULA_DESCRICAO))

    LimparSequencia celulaSequencia, celulaDescricao
    If Not SequenciaValida(sequencia) Then Exit Sub

    EscreverSequencia celulaSequencia, celulaDescricao, sequencia
End Sub

Private Function SequenciaValida(ByVal sequencia As Long) As Boolean
    SequenciaValida = (sequencia >= SEQ_MINIMO And sequencia <= SEQ_MAXIMO)
End Function

Private Sub LimparSequencia(celulaSequencia As Range, celulaDescricao As Range)
    RemoverCor celulaSequencia
    RemoverCor celulaDescricao
    DefinirNegrito celulaSequencia, False
    LimparCelula celulaSequencia
    LimparCelula celulaDescricao
End Sub

Private Sub EscreverSequencia(celulaSequencia As Range, celulaDescricao As Range, ByVal sequencia As Long)
    Dim corSequencia As Long
    corSequencia = SequenciaParaCor(sequencia)

    DefinirValor celulaSequencia, sequencia
    DefinirNegrito celulaSequencia, True
    DefinirCor celulaSequencia, corSequencia

    DefinirValor celulaDescricao, SequenciaParaDescricao(sequencia)
    DefinirCor celulaDescricao, corSequencia
End Sub

' Em celula mesclada, so a ancora aceita escrita.
Private Function ObterCelulaOperativa(celula As Range) As Range
    If celula Is Nothing Then Exit Function

    If Not celula.MergeCells Then
        Set ObterCelulaOperativa = celula
        Exit Function
    End If

    If celula.Address = celula.MergeArea.Cells(1, 1).Address Then Set ObterCelulaOperativa = celula
End Function

Private Sub RemoverCor(celula As Range)
    If celula Is Nothing Then Exit Sub
    celula.MergeArea.Interior.ColorIndex = xlNone
End Sub

Private Sub DefinirCor(celula As Range, ByVal cor As Long)
    If celula Is Nothing Then Exit Sub
    celula.MergeArea.Interior.Color = cor
End Sub

Private Sub DefinirNegrito(celula As Range, ByVal ativar As Boolean)
    If celula Is Nothing Then Exit Sub
    celula.Font.Bold = ativar
End Sub

Private Sub LimparCelula(celula As Range)
    If celula Is Nothing Then Exit Sub
    celula.MergeArea.ClearContents
End Sub

Private Sub DefinirValor(celula As Range, ByVal valor As Variant)
    If celula Is Nothing Then Exit Sub
    celula.Value2 = valor
End Sub

Public Function SequenciaParaCor(ByVal sequencia As Long) As Long
    Select Case sequencia
        Case 1: SequenciaParaCor = RGB(255, 255, 0)
        Case 2: SequenciaParaCor = RGB(0, 63, 242)
        Case 3: SequenciaParaCor = RGB(246, 94, 0)
        Case 4: SequenciaParaCor = RGB(117, 219, 255)
        Case 5: SequenciaParaCor = RGB(255, 0, 0)
        Case 6: SequenciaParaCor = RGB(0, 54, 24)
        Case 7: SequenciaParaCor = RGB(180, 0, 171)
        Case 8: SequenciaParaCor = RGB(44, 255, 33)
        Case 9: SequenciaParaCor = RGB(131, 60, 12)
        Case 10: SequenciaParaCor = RGB(36, 1, 99)
    End Select
End Function

Public Function SequenciaParaDescricao(ByVal sequencia As Long) As String
    Select Case sequencia
        Case 1: SequenciaParaDescricao = "AMARELO"
        Case 2: SequenciaParaDescricao = "AZUL ESCURO"
        Case 3: SequenciaParaDescricao = "LARANJA"
        Case 4: SequenciaParaDescricao = "AZUL CLARO"
        Case 5: SequenciaParaDescricao = "VERMELHO"
        Case 6: SequenciaParaDescricao = "VERDE ESCURO"
        Case 7: SequenciaParaDescricao = "ROSA CHOQUE"
        Case 8: SequenciaParaDescricao = "VERDE CLARO"
        Case 9: SequenciaParaDescricao = "MARROM"
        Case 10: SequenciaParaDescricao = "ROXO"
    End Select
End Function
