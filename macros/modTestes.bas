Attribute VB_Name = "modTestes"
' =====================================================================
' modTestes
' Testes unitarios das funcoes puras (conversao, traducao, peso, mapa
' de cores, nomes de sheet). Rodar com F5 em ExecutarTestes.
' Nao ha mocks de Worksheet porque o nucleo foi extraido para funcoes
' que recebem String/array; as fixtures sao matrizes nomeadas.
' =====================================================================
Option Explicit

Private totalTestes As Long
Private totalFalhas As Long
Private relatorioFalhas As String

Public Sub ExecutarTestes()
    ReiniciarContadores

    TestarConversaoNumero
    TestarConversaoSequencia
    TestarArredondamentoChave
    TestarSubstituicoes
    TestarSubstituicaoNaoIdempotente
    TestarTipoPorta
    TestarTotalPesoPortas
    TestarMapaCores
    TestarNomeSheet
    TestarLetrasIniciais

    MsgBox MontarResumo(), IIf(totalFalhas = 0, vbInformation, vbCritical)
End Sub

' Teste manual/integracao: AjustarFonteParaCaber depende de Shapes do
' Excel (medicao de texto), entao nao ha como isolar em fake/stub sem
' abrir a aplicacao. Rodar com a celula alvo selecionada.
Public Sub TestarAjustarFonteManual()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Selecione a celula a testar antes de rodar.", vbExclamation
        Exit Sub
    End If

    Dim tamanhoAntes As Double
    tamanhoAntes = Selection.Font.Size

    AjustarFonteParaCaber Selection

    MsgBox "Fonte antes: " & tamanhoAntes & vbCrLf & _
           "Fonte depois: " & Selection.Font.Size, vbInformation
End Sub

Private Sub TestarConversaoNumero()
    AfirmarDouble "numero decimal pt-BR", ConverterNumero("1.234,56"), 1234.56
    AfirmarDouble "numero decimal simples", ConverterNumero("35,9"), 35.9
    AfirmarDouble "numero ja numerico", ConverterNumero(12), 12
    AfirmarDouble "texto invalido vira zero", ConverterNumero("abc"), 0
    AfirmarDouble "vazio vira zero", ConverterNumero(Empty), 0
    AfirmarDouble "booleano vira zero", ConverterNumero(True), 0
End Sub

Private Sub TestarConversaoSequencia()
    AfirmarLong "sequencia inteira", ConverterSequencia("3"), 3
    AfirmarLong "sequencia zero explicita", ConverterSequencia("0"), 0
    AfirmarLong "sequencia com decimal e rejeitada", ConverterSequencia(2.5), 0
    AfirmarLong "sequencia com texto e rejeitada", ConverterSequencia("Xa"), 0
End Sub

Private Sub TestarArredondamentoChave()
    AfirmarLong "chave meio pra cima", ArredondarChave(819.5), 820
    AfirmarLong "chave exata", ArredondarChave(1220), 1220
    AfirmarLong "chave negativa e zerada", ArredondarChave(-5), 0
End Sub

Private Sub TestarSubstituicoes()
    AfirmarTexto "sigla de porta", AplicarSubstituicoes("PTA LISA"), "Porta LISA"
    AfirmarTexto "material DPL", AplicarSubstituicoes("PTA DPL 5C"), "Porta DURAPLAC 5C"
    AfirmarTexto "largura 210", AplicarSubstituicoes("X 210X82X35X"), "X 2,10 x 0,82 x 0,35 x"
    AfirmarTexto "vazio continua vazio", AplicarSubstituicoes(""), ""
End Sub

' Documenta por que BD_PRESERVAR_ORIGEM existe: reaplicar a tabela
' sobre um texto ja traduzido corrompe o resultado (" 40" -> " FENIX 40").
Private Sub TestarSubstituicaoNaoIdempotente()
    Dim primeiraPassada As String
    primeiraPassada = AplicarSubstituicoes("PTA CAM 40")

    AfirmarVerdadeiro "traducao nao e idempotente", _
        AplicarSubstituicoes(primeiraPassada) <> primeiraPassada
End Sub

Private Sub TestarTipoPorta()
    AfirmarTexto "tipo DBL", IdentificarTipoPorta("Porta DBL 820"), "DBL"
    AfirmarTexto "tipo COL", IdentificarTipoPorta("porta colmeia"), "COL"
    AfirmarTexto "tipo padrao", IdentificarTipoPorta("Porta Lisa"), "LISA"
    AfirmarDouble "peso DBL 820", ObterPesoPorta("DBL", 820), 35.9
    AfirmarDouble "tamanho fora da tabela", ObterPesoPorta("LISA", 777), 0
End Sub

Private Sub TestarTotalPesoPortas()
    Dim quantidades As Variant, nomes As Variant, tamanhos As Variant
    quantidades = MatrizColuna(Array(2, 1, 0, "3"))
    nomes = MatrizColuna(Array("Porta DBL", "Porta COL", "Porta Lisa", "Porta Lisa"))
    tamanhos = MatrizColuna(Array(820, 700, 900, 600))

    ' 2*35,9 + 1*13 + 0 + 3*18,5 = 140,3
    AfirmarDouble "total do peso", TotalPesoPortas(quantidades, nomes, tamanhos), 140.3
End Sub

Private Sub TestarMapaCores()
    Dim cores() As Long
    cores = CalcularMapaCores("Porta DBL")

    AfirmarLong "DBL pintado de vermelho", cores(7), RGB(255, 0, 0)
    AfirmarLong "texto comum sem cor", cores(1), -1

    cores = CalcularMapaCores("HDF 4-2")
    AfirmarLong "HDF em verde claro", cores(1), RGB(144, 238, 144)
    AfirmarLong "sequencia 4 em vermelho", cores(5), RGB(255, 0, 0)
End Sub

Private Sub TestarNomeSheet()
    AfirmarTexto "nome simples", MontarNomeSheet("12345"), "Pedido_12345"
    AfirmarTexto "caracter invalido trocado", MontarNomeSheet("123/45"), "Pedido_123_45"
    AfirmarLong "nome limitado a 31", Len(MontarNomeSheet(String$(60, "9"))), 31
End Sub

Private Sub TestarLetrasIniciais()
    AfirmarTexto "letra de coluna", ExtrairLetrasIniciais("L1"), "L"
    AfirmarTexto "coluna dupla", ExtrairLetrasIniciais("AI"), "AI"
    AfirmarTexto "sem letras", ExtrairLetrasIniciais("12"), ""
End Sub

' Converte Array() em matriz (1 To n, 1 To 1), o formato devolvido por
' Range.Value2, para que os testes usem a mesma interface da planilha.
Private Function MatrizColuna(ByVal valores As Variant) As Variant
    Dim matriz() As Variant
    ReDim matriz(1 To UBound(valores) - LBound(valores) + 1, 1 To 1)

    Dim indice As Long
    For indice = LBound(valores) To UBound(valores)
        matriz(indice - LBound(valores) + 1, 1) = valores(indice)
    Next indice

    MatrizColuna = matriz
End Function

Private Sub ReiniciarContadores()
    totalTestes = 0
    totalFalhas = 0
    relatorioFalhas = ""
End Sub

Private Function MontarResumo() As String
    MontarResumo = "Testes: " & totalTestes & " | Falhas: " & totalFalhas
    If Len(relatorioFalhas) > 0 Then MontarResumo = MontarResumo & vbCrLf & vbCrLf & relatorioFalhas
End Function

Private Sub AfirmarTexto(ByVal descricao As String, ByVal obtido As String, ByVal esperado As String)
    RegistrarResultado descricao, (obtido = esperado), obtido, esperado
End Sub

Private Sub AfirmarLong(ByVal descricao As String, ByVal obtido As Long, ByVal esperado As Long)
    RegistrarResultado descricao, (obtido = esperado), CStr(obtido), CStr(esperado)
End Sub

Private Sub AfirmarDouble(ByVal descricao As String, ByVal obtido As Double, ByVal esperado As Double)
    RegistrarResultado descricao, (Abs(obtido - esperado) < 0.000001), CStr(obtido), CStr(esperado)
End Sub

Private Sub AfirmarVerdadeiro(ByVal descricao As String, ByVal condicao As Boolean)
    RegistrarResultado descricao, condicao, CStr(condicao), "True"
End Sub

Private Sub RegistrarResultado(ByVal descricao As String, ByVal aprovado As Boolean, ByVal obtido As String, ByVal esperado As String)
    totalTestes = totalTestes + 1
    If aprovado Then Exit Sub

    totalFalhas = totalFalhas + 1
    relatorioFalhas = relatorioFalhas & "FALHA: " & descricao & " | obtido=" & obtido & " esperado=" & esperado & vbCrLf
    Debug.Print "FALHA: " & descricao & " | obtido=" & obtido & " esperado=" & esperado
End Sub
