Attribute VB_Name = "modConversao"
' =====================================================================
' modConversao
' Conversoes seguras de valores vindos de celulas: numero, texto e
' sequencia. Nenhuma funcao aqui toca em Range: entrada e Variant,
' o que torna o modulo testavel sem planilha.
' =====================================================================
Option Explicit

Public Function ConverterNumero(ByVal valor As Variant) As Double
    If Not ValorUtilizavel(valor) Then Exit Function
    If VarType(valor) = vbBoolean Then Exit Function

    If VarType(valor) <> vbString Then
        If IsNumeric(valor) Then ConverterNumero = CDbl(valor)
        Exit Function
    End If

    ConverterNumero = NumeroDeTexto(CStr(valor))
End Function

Public Function ConverterTexto(ByVal valor As Variant) As String
    If Not ValorUtilizavel(valor) Then Exit Function
    ConverterTexto = Trim$(CStr(valor))
End Function

Public Function ConverterTextoBruto(ByVal valor As Variant) As String
    If Not ValorUtilizavel(valor) Then Exit Function
    ConverterTextoBruto = CStr(valor)
End Function

Public Function ConverterSequencia(ByVal valor As Variant) As Long
    Dim texto As String
    texto = ConverterTexto(valor)
    If Len(texto) = 0 Then Exit Function

    Dim numeroDecimal As Double
    numeroDecimal = ConverterNumero(valor)

    If numeroDecimal = 0 And texto <> "0" Then Exit Function
    If numeroDecimal <> Int(numeroDecimal) Then Exit Function
    If Abs(numeroDecimal) > 2147483647# Then Exit Function

    ConverterSequencia = CLng(numeroDecimal)
End Function

' Arredondamento meio-pra-cima: Round() do VBA usa banker's rounding
' e erra a chave de tamanho da tabela de pesos (ex.: 820,5).
Public Function ArredondarChave(ByVal numero As Double) As Long
    If numero <= 0 Then Exit Function
    If numero > 2147483000# Then Exit Function

    ArredondarChave = CLng(Int(numero + 0.5))
End Function

Private Function ValorUtilizavel(ByVal valor As Variant) As Boolean
    If IsObject(valor) Then Exit Function
    If IsArray(valor) Then Exit Function
    If IsError(valor) Then Exit Function
    If IsNull(valor) Then Exit Function
    If IsEmpty(valor) Then Exit Function

    ValorUtilizavel = True
End Function

' Val() sempre usa ponto como decimal, entao independe do locale.
Private Function NumeroDeTexto(ByVal texto As String) As Double
    Dim limpo As String
    limpo = NormalizarSeparadores(texto)
    If Len(limpo) = 0 Then Exit Function
    If Not IsNumeric(limpo) Then Exit Function

    NumeroDeTexto = Val(limpo)
End Function

Private Function NormalizarSeparadores(ByVal texto As String) As String
    Dim limpo As String
    limpo = Replace(Trim$(texto), " ", "")
    limpo = Replace(limpo, Chr$(160), "")

    ' Se ha virgula, o ponto so pode ser separador de milhar.
    If InStr(1, limpo, ",") > 0 Then limpo = Replace(limpo, ".", "")

    NormalizarSeparadores = Replace(limpo, ",", ".")
End Function
