Attribute VB_Name = "modColorirItens"
' =====================================================================
' modColorirItens
' Pinta termos especificos nas descricoes dos itens.
' Estrategia: monta em memoria um mapa cor-por-caractere e aplica
' apenas as faixas contiguas resultantes. Characters(...).Font e a
' chamada COM mais cara do projeto; assim ela cai de ~1 por ocorrencia
' para 1 por bloco de cor.
' =====================================================================
Option Explicit

Private Const COR_NEUTRA As Long = -1
Private Const LIMITE_CARACTERES As Long = 255

Private termosVermelhosCache As Variant
Private avaliadorSequencia As Object

Public Sub ColorirPalavrasSelecao()
    If TypeName(Selection) <> "Range" Then Exit Sub
    ColorirPalavrasNoIntervalo Selection
End Sub

Public Sub ColorirPalavrasNoIntervalo(intervalo As Range)
    If intervalo Is Nothing Then Exit Sub

    EntrarModoRapido
    On Error GoTo Erro

    Dim celula As Range
    For Each celula In intervalo.Cells
        ColorirCelula celula
    Next celula

    SairModoRapido
    Exit Sub

Erro:
    AbortarModoRapido
    MsgBox "Falha ao colorir descricoes: " & Err.Description, vbCritical
End Sub

Private Sub ColorirCelula(ByRef celula As Range)
    Dim textoCelula As String
    textoCelula = TextoColorivel(celula)
    If Len(textoCelula) = 0 Then Exit Sub

    Dim coresPorCaracter() As Long
    coresPorCaracter = CalcularMapaCores(textoCelula)

    AplicarMapaCores celula, coresPorCaracter
End Sub

' Characters nao formata resultado de formula e fica instavel acima de
' 255 caracteres; nesses casos o texto e truncado ou a celula ignorada.
Private Function TextoColorivel(ByRef celula As Range) As String
    If celula Is Nothing Then Exit Function
    If celula.HasFormula Then Exit Function
    If IsError(celula.Value2) Then Exit Function

    Dim textoCelula As String
    textoCelula = ConverterTextoBruto(celula.Value2)
    If Len(textoCelula) = 0 Then Exit Function

    TextoColorivel = Left$(textoCelula, LIMITE_CARACTERES)
End Function

' Publica para permitir teste unitario sem planilha.
Public Function CalcularMapaCores(ByVal textoCelula As String) As Long()
    Dim coresPorCaracter() As Long
    ReDim coresPorCaracter(1 To MaiorQueZero(Len(textoCelula)))
    PreencherNeutro coresPorCaracter

    If Len(textoCelula) = 0 Then
        CalcularMapaCores = coresPorCaracter
        Exit Function
    End If

    MarcarTermosVermelhos textoCelula, coresPorCaracter
    MarcarTermosEspecificos textoCelula, coresPorCaracter
    MarcarSequencia4 textoCelula, coresPorCaracter

    CalcularMapaCores = coresPorCaracter
End Function

Private Function MaiorQueZero(ByVal numero As Long) As Long
    MaiorQueZero = numero
    If numero < 1 Then MaiorQueZero = 1
End Function

Private Sub PreencherNeutro(ByRef coresPorCaracter() As Long)
    Dim indice As Long
    For indice = LBound(coresPorCaracter) To UBound(coresPorCaracter)
        coresPorCaracter(indice) = COR_NEUTRA
    Next indice
End Sub

Private Sub MarcarTermosVermelhos(ByVal textoCelula As String, ByRef coresPorCaracter() As Long)
    Dim termo As Variant
    For Each termo In TermosVermelhos()
        MarcarTermo textoCelula, coresPorCaracter, CStr(termo), RGB(255, 0, 0)
    Next termo
End Sub

Private Sub MarcarTermosEspecificos(ByVal textoCelula As String, ByRef coresPorCaracter() As Long)
    MarcarGrupo textoCelula, coresPorCaracter, Array("FRESNO", "FRE"), RGB(0, 128, 0)
    MarcarGrupo textoCelula, coresPorCaracter, Array("DURATREE", "DRT"), RGB(0, 176, 240)
    MarcarGrupo textoCelula, coresPorCaracter, Array("IMBUIA", "IMB"), RGB(0, 32, 96)
    MarcarGrupo textoCelula, coresPorCaracter, Array("MOGNO", "MOG"), RGB(112, 48, 160)
    MarcarGrupo textoCelula, coresPorCaracter, Array("HDF"), RGB(144, 238, 144)
End Sub

Private Sub MarcarGrupo(ByVal textoCelula As String, ByRef coresPorCaracter() As Long, ByVal termos As Variant, ByVal cor As Long)
    Dim termo As Variant
    For Each termo In termos
        MarcarTermo textoCelula, coresPorCaracter, CStr(termo), cor
    Next termo
End Sub

Private Sub MarcarTermo(ByVal textoCelula As String, ByRef coresPorCaracter() As Long, ByVal termo As String, ByVal cor As Long)
    Dim comprimento As Long
    comprimento = Len(termo)
    If comprimento = 0 Then Exit Sub

    Dim posicao As Long
    posicao = InStr(1, textoCelula, termo, vbTextCompare)

    Do While posicao > 0
        PreencherFaixa coresPorCaracter, posicao, comprimento, cor
        posicao = InStr(posicao + comprimento, textoCelula, termo, vbTextCompare)
    Loop
End Sub

Private Sub MarcarSequencia4(ByVal textoCelula As String, ByRef coresPorCaracter() As Long)
    Dim avaliador As Object
    Set avaliador = RegexSequencia4()
    If Not avaliador.Test(textoCelula) Then Exit Sub

    Dim ocorrencia As Object
    For Each ocorrencia In avaliador.Execute(textoCelula)
        PreencherFaixa coresPorCaracter, ocorrencia.FirstIndex + 1, ocorrencia.Length, RGB(255, 0, 0)
    Next ocorrencia
End Sub

Private Sub PreencherFaixa(ByRef coresPorCaracter() As Long, ByVal inicio As Long, ByVal comprimento As Long, ByVal cor As Long)
    Dim fim As Long
    fim = inicio + comprimento - 1
    If fim > UBound(coresPorCaracter) Then fim = UBound(coresPorCaracter)

    Dim indice As Long
    For indice = inicio To fim
        coresPorCaracter(indice) = cor
    Next indice
End Sub

Private Sub AplicarMapaCores(ByRef celula As Range, ByRef coresPorCaracter() As Long)
    celula.Font.ColorIndex = xlAutomatic

    Dim inicioFaixa As Long
    inicioFaixa = LBound(coresPorCaracter)

    Dim indice As Long
    For indice = inicioFaixa + 1 To UBound(coresPorCaracter)
        If coresPorCaracter(indice) <> coresPorCaracter(inicioFaixa) Then
            PintarFaixa celula, inicioFaixa, indice - inicioFaixa, coresPorCaracter(inicioFaixa)
            inicioFaixa = indice
        End If
    Next indice

    PintarFaixa celula, inicioFaixa, UBound(coresPorCaracter) - inicioFaixa + 1, coresPorCaracter(inicioFaixa)
End Sub

Private Sub PintarFaixa(ByRef celula As Range, ByVal inicio As Long, ByVal comprimento As Long, ByVal cor As Long)
    If cor = COR_NEUTRA Then Exit Sub
    If comprimento <= 0 Then Exit Sub

    celula.Characters(inicio, comprimento).Font.Color = cor
End Sub

Private Function TermosVermelhos() As Variant
    If IsEmpty(termosVermelhosCache) Then
        termosVermelhosCache = Array( _
            "5C", "7C", "2400", "2440", "2200", "CAM", "/", "+", _
            "DURAPLAC", "DPL", "COLMÉIA", "FENIX", "REG4", _
            "ML", "US", "CL", "DBL")
    End If

    TermosVermelhos = termosVermelhosCache
End Function

' RegExp instanciado uma unica vez: CreateObject por celula custava
' mais que toda a varredura de termos.
Private Function RegexSequencia4() As Object
    If avaliadorSequencia Is Nothing Then
        Set avaliadorSequencia = CreateObject("VBScript.RegExp")
        avaliadorSequencia.Pattern = "\b4\d*(-\d+)?\b"
        avaliadorSequencia.Global = True
        avaliadorSequencia.IgnoreCase = True
    End If

    Set RegexSequencia4 = avaliadorSequencia
End Function
