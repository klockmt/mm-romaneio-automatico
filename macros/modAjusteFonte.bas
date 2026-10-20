Attribute VB_Name = "modAjusteFonte"
' =====================================================================
' modAjusteFonte
' Reduz o tamanho da fonte de uma celula ate o texto caber inteiramente
' dentro dela (largura e altura), sem deixar conteudo cortado.
' Usado no pos-import, antes de remover as linhas em branco, sobre a
' celula E61 de cada pedido.
'
' Medicao: nao ha API nativa no VBA para medir a largura renderizada de
' um texto. A tecnica usada e criar uma Shape.Textbox temporaria fora
' da area visivel, com a mesma fonte, deixar o TextFrame2 se
' autoajustar (AutoSize/WordWrap) e comparar o resultado com o espaco
' disponivel na celula. E a abordagem padrao para esse problema em VBA
' e nao depende de referencias externas (ActiveX/Forms).
' Alternativa descartada: Range.ShrinkToFit nativo do Excel e mais
' simples, mas e um efeito puramente visual (nao altera Font.Size);
' calculos, impressao e exportacao que leem o tamanho da fonte
' ignoram esse ajuste, o que nao atende "nao ficar texto fora" de
' forma confiavel fora da tela do Excel.
' =====================================================================
Option Explicit

Private Const AJUSTE_SHEET_MODELO As String = "template"
Private Const AJUSTE_CELULA_PEDIDO As String = "E61"
Private Const AJUSTE_TAMANHO_MINIMO As Long = 6
Private Const AJUSTE_MARGEM_PONTOS As Double = 2

Public Sub AjustarFonteCelulaPedido(sheetDestino As Worksheet)
    If sheetDestino Is Nothing Then Exit Sub
    AjustarFonteParaCaber sheetDestino.Range(AJUSTE_CELULA_PEDIDO)
End Sub

Public Sub AjustarFonteParaCaber(celula As Range, Optional ByVal tamanhoMinimo As Long = AJUSTE_TAMANHO_MINIMO)
    Dim celulaAncora As Range
    Set celulaAncora = ResolverCelulaAncora(celula)
    If celulaAncora Is Nothing Then Exit Sub

    Dim texto As String
    texto = ConverterTextoBruto(celulaAncora.Value2)
    If Len(texto) = 0 Then Exit Sub

    On Error GoTo Erro
    ReduzirAteCaber celulaAncora, texto, tamanhoMinimo
    Exit Sub

Erro:
    Debug.Print "Ajuste de fonte falhou em " & celulaAncora.Address & ": " & Err.Description
End Sub

Private Sub ReduzirAteCaber(celulaAncora As Range, ByVal texto As String, ByVal tamanhoMinimo As Long)
    celulaAncora.Font.Size = TamanhoFonteBase(celulaAncora)

    Dim tamanho As Long
    For tamanho = CLng(celulaAncora.Font.Size) To tamanhoMinimo Step -1
        celulaAncora.Font.Size = tamanho
        If TextoCabeNaCelula(celulaAncora, texto) Then Exit Sub
    Next tamanho
End Sub

' Sempre parte do tamanho do template, nunca do tamanho atual da
' celula: evita que reexecucoes sucessivas encolham a fonte sem
' nunca voltar a crescer quando o texto fica mais curto.
Private Function TamanhoFonteBase(celulaAncora As Range) As Long
    Dim sheetModelo As Worksheet
    Set sheetModelo = SheetModeloSeExistir()

    If sheetModelo Is Nothing Then
        TamanhoFonteBase = CLng(celulaAncora.Font.Size)
        Exit Function
    End If

    TamanhoFonteBase = CLng(sheetModelo.Range(celulaAncora.Address).Font.Size)
End Function

Private Function SheetModeloSeExistir() As Worksheet
    Dim sheetModelo As Object

    On Error Resume Next
    Set sheetModelo = ThisWorkbook.Worksheets(AJUSTE_SHEET_MODELO)
    On Error GoTo 0

    If sheetModelo Is Nothing Then Exit Function
    If Not TypeOf sheetModelo Is Worksheet Then Exit Function

    Set SheetModeloSeExistir = sheetModelo
End Function

Private Function ResolverCelulaAncora(celula As Range) As Range
    If celula Is Nothing Then Exit Function

    If Not celula.MergeCells Then
        Set ResolverCelulaAncora = celula
        Exit Function
    End If

    If celula.Address = celula.MergeArea.Cells(1, 1).Address Then Set ResolverCelulaAncora = celula.MergeArea.Cells(1, 1)
End Function

Private Function TextoCabeNaCelula(celulaAncora As Range, ByVal texto As String) As Boolean
    Dim areaCelula As Range
    Set areaCelula = celulaAncora.MergeArea

    Dim larguraDisponivel As Double
    Dim alturaDisponivel As Double
    larguraDisponivel = areaCelula.Width - AJUSTE_MARGEM_PONTOS
    alturaDisponivel = areaCelula.Height - AJUSTE_MARGEM_PONTOS

    Dim caixaMedicao As Shape
    Set caixaMedicao = CriarCaixaMedicao(celulaAncora, texto, larguraDisponivel)

    TextoCabeNaCelula = (caixaMedicao.Width <= larguraDisponivel) And (caixaMedicao.Height <= alturaDisponivel)
    caixaMedicao.Delete
End Function

Private Function CriarCaixaMedicao(celulaAncora As Range, ByVal texto As String, ByVal larguraColuna As Double) As Shape
    Dim caixa As Shape
    Set caixa = celulaAncora.Worksheet.Shapes.AddTextbox(msoTextOrientationHorizontal, -1000, -1000, 10, 10)

    If celulaAncora.WrapText Then caixa.Width = larguraColuna

    ConfigurarTextoMedicao caixa, celulaAncora, texto
    Set CriarCaixaMedicao = caixa
End Function

Private Sub ConfigurarTextoMedicao(caixa As Shape, celulaAncora As Range, ByVal texto As String)
    With caixa.TextFrame2
        .TextRange.Text = texto
        .TextRange.Font.Name = celulaAncora.Font.Name
        .TextRange.Font.Bold = ParaTriState(CBool(celulaAncora.Font.Bold))
        .TextRange.Font.Size = celulaAncora.Font.Size
        .WordWrap = ParaTriState(celulaAncora.WrapText)
        .AutoSize = msoAutoSizeShapeToFitText
    End With
End Sub

Private Function ParaTriState(ByVal ativo As Boolean) As MsoTriState
    ParaTriState = msoFalse
    If ativo Then ParaTriState = msoTrue
End Function
