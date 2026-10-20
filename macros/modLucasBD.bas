Attribute VB_Name = "modLucasBD"
' =====================================================================
' modLucasBD
' Importacao/transformacao dos dados do banco.
' Espera a fonte, limpa a area de itens, traduz as siglas em memoria e
' escreve V -> B, Y -> C, W(traduzido) -> D.
' As ~50 chamadas Selection.Replace sobre a coluna W inteira foram
' trocadas por Replace() em array: 1 leitura + 1 escrita por coluna.
' =====================================================================
Option Explicit

Private Const BD_ORIGEM_CODIGO As String = "V11:V56"
Private Const BD_ORIGEM_QUANTIDADE As String = "Y11:Y56"
Private Const BD_ORIGEM_DESCRICAO As String = "W11:W56"
Private Const BD_DESTINO_CODIGO As String = "B12"
Private Const BD_DESTINO_QUANTIDADE As String = "C12"
Private Const BD_DESTINO_DESCRICAO As String = "D12"
Private Const BD_AREA_LIMPEZA As String = "B12:M53"
Private Const BD_CELULA_FOCO As String = "P3"

' True = a coluna W nao e reescrita. Evita dupla traducao ao reexecutar
' (" 40" -> " FENIX 40" -> " FENIX FENIX 40") e preserva formulas da fonte.
Private Const BD_PRESERVAR_ORIGEM As Boolean = True

Private tabelaSubstituicoesCache As Variant

Public Sub Substituir()
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub

    If SubstituirNaSheet(ActiveSheet) Then PosicionarFoco ActiveSheet
End Sub

Public Function SubstituirNaSheet(sheetAlvo As Worksheet) As Boolean
    If sheetAlvo Is Nothing Then Exit Function
    If Not EsperarFonteBanco(sheetAlvo) Then Exit Function

    EntrarModoRapido
    On Error GoTo Erro

    LimparAreaItens sheetAlvo
    PreencherDestino sheetAlvo

    SairModoRapido
    RecalcularSheet sheetAlvo
    CalcularPesoPortasNaSheet sheetAlvo

    SubstituirNaSheet = True
    Exit Function

Erro:
    AbortarModoRapido
    MsgBox "Falha em Substituir (" & sheetAlvo.Name & "): " & Err.Description, vbCritical
End Function

Public Sub Apagar()
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    LimparAreaItens ActiveSheet
End Sub

Public Sub Copiar()
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub

    PreencherDestino ActiveSheet
    PosicionarFoco ActiveSheet
End Sub

Private Sub LimparAreaItens(sheetAlvo As Worksheet)
    sheetAlvo.Range(BD_AREA_LIMPEZA).ClearContents
End Sub

Private Sub PreencherDestino(sheetAlvo As Worksheet)
    CopiarValores sheetAlvo.Range(BD_ORIGEM_CODIGO), sheetAlvo.Range(BD_DESTINO_CODIGO)
    CopiarValores sheetAlvo.Range(BD_ORIGEM_QUANTIDADE), sheetAlvo.Range(BD_DESTINO_QUANTIDADE)
    EscreverDescricoes sheetAlvo
End Sub

' Select so e valido na sheet ativa; fica restrito as entradas manuais.
Private Sub PosicionarFoco(sheetAlvo As Worksheet)
    If Not ActiveSheet Is sheetAlvo Then Exit Sub
    sheetAlvo.Range(BD_CELULA_FOCO).Select
End Sub

Private Sub CopiarValores(origem As Range, destino As Range)
    destino.Resize(origem.Rows.Count, origem.Columns.Count).Value2 = origem.Value2
End Sub

Private Sub EscreverDescricoes(sheetAlvo As Worksheet)
    Dim descricoes As Variant
    descricoes = sheetAlvo.Range(BD_ORIGEM_DESCRICAO).Value2

    Dim linha As Long
    For linha = LBound(descricoes, 1) To UBound(descricoes, 1)
        descricoes(linha, 1) = AplicarSubstituicoes(ConverterTextoBruto(descricoes(linha, 1)))
    Next linha

    If Not BD_PRESERVAR_ORIGEM Then sheetAlvo.Range(BD_ORIGEM_DESCRICAO).Value2 = descricoes
    sheetAlvo.Range(BD_DESTINO_DESCRICAO).Resize(UBound(descricoes, 1), 1).Value2 = descricoes
End Sub

' Publica para teste unitario: nucleo da traducao, sem planilha.
Public Function AplicarSubstituicoes(ByVal textoOriginal As String) As String
    If Len(textoOriginal) = 0 Then Exit Function

    Dim tabela As Variant
    tabela = TabelaSubstituicoes()

    Dim resultado As String
    resultado = textoOriginal

    Dim indice As Long
    For indice = LBound(tabela, 2) To UBound(tabela, 2)
        resultado = Replace(resultado, tabela(0, indice), tabela(1, indice), 1, -1, vbTextCompare)
    Next indice

    AplicarSubstituicoes = Trim$(resultado)
End Function

Private Function TabelaSubstituicoes() As Variant
    If IsEmpty(tabelaSubstituicoesCache) Then tabelaSubstituicoesCache = MontarTabelaSubstituicoes()
    TabelaSubstituicoes = tabelaSubstituicoesCache
End Function

' A ordem dos blocos e significativa e replica a ordem original.
Private Function MontarTabelaSubstituicoes() As Variant
    Dim tabela() As String
    ReDim tabela(0 To 1, 0 To 199)

    Dim contador As Long
    AdicionarPares tabela, contador, ParesAcabamento()
    AdicionarPares tabela, contador, ParesLarguras()
    AdicionarPares tabela, contador, ParesEspessuras()
    AdicionarPares tabela, contador, ParesEspessurasDecimais()

    ReDim Preserve tabela(0 To 1, 0 To contador - 1)
    MontarTabelaSubstituicoes = tabela
End Function

Private Sub AdicionarPares(ByRef tabela() As String, ByRef contador As Long, ByVal pares As Variant)
    Dim indice As Long
    For indice = LBound(pares) To UBound(pares) Step 2
        tabela(0, contador) = CStr(pares(indice))
        tabela(1, contador) = CStr(pares(indice + 1))
        contador = contador + 1
    Next indice
End Sub

Private Function ParesAcabamento() As Variant
    ParesAcabamento = Array( _
        "PTA ", "Porta ", " NENC ", " NÃO Encabeçada ", _
        "ACAB ", "Porta ", " CB ", " ", _
        " ENC ", " Encabeçada ", " SS ", " ", _
        " EB ", " ", " DUR ", " DURATREE ", _
        " DPL ", " DURAPLAC ", " COL ", " COLMÉIA ", _
        " ANG ", " Em Angelim ", " CUR ", " Em Curupixá ", _
        " TAU ", " Em Tauari ", " JEQ ", " Em Jequitibá ", _
        " IPE ", " Em Ipê Champagne ", " 40", " FENIX 40", _
        " DBL ", " Doorblanks Brasileiro ", _
        "Porta Não Encab. Doorblanks Brasileiro ", "Doorblanks Brasileiro ")
End Function

Private Function ParesLarguras() As Variant
    ParesLarguras = Array( _
        " 200X", " 2,00 x", " 207X", " 2,07 x", _
        " 208X", " 2,08 x", " 210X", " 2,10 x", _
        " 220X", " 2,20 x", " 240X", " 2,40 x")
End Function

Private Function ParesEspessuras() As Variant
    ParesEspessuras = Array( _
        "X31X", "x 0,31 x", "X35X", "x 0,35 x", _
        "X36X", "x 0,36 x", "X40X", "x 0,40 x", _
        "X41X", "x 0,41 x", "X46X", "x 0,46 x", _
        "X60X", "x 0,60 x", "X62X", "x 0,62 x", _
        "X70X", "x 0,70 x", "X72X", "x 0,72 x", _
        "X80X", "x 0,80 x", "X82X", "x 0,82 x", _
        "X90X", "x 0,90 x", "X92X", "x 0,92 x", _
        "X100X", "x 1,00 x", "X102X", "x 1,02 x", _
        "X110X", "x 1,10 x", "X112X", "x 1,12 x", _
        "X120X", "x 1,20 x", "X122X", "x 1,22 x", _
        "X130X", "x 1,30 x", "X132X", "x 1,32 x")
End Function

Private Function ParesEspessurasDecimais() As Variant
    ParesEspessurasDecimais = Array( _
        "X3,00", "x 0,030", "X3,50", "x 0,035", _
        "X4,00", "x 0,040", "X4,50", "x 0,045")
End Function
