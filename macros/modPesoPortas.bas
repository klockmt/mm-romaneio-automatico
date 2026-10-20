Attribute VB_Name = "modPesoPortas"
' =====================================================================
' modPesoPortas
' Calculo do peso das portas.
' Le quantidade em C, descricao em D e tamanho em AI (3 leituras de
' intervalo, antes eram ~126 leituras de celula) e escreve em U2.
' O nucleo (TotalPesoPortas) recebe arrays, nao Range: testavel.
' =====================================================================
Option Explicit

Private Const PESO_PRIMEIRA_LINHA As Long = 12
Private Const PESO_ULTIMA_LINHA As Long = 53
Private Const PESO_COLUNA_QUANTIDADE As String = "C"
Private Const PESO_COLUNA_NOME As String = "D"
Private Const PESO_COLUNA_TAMANHO As String = "AI"
Private Const PESO_CELULA_TOTAL As String = "U2"

Public Sub CalcularPesoPortas()
    Dim sheetAtiva As Worksheet
    If Not TentarObterSheetAtiva(sheetAtiva) Then Exit Sub

    CalcularPesoPortasNaSheet sheetAtiva
End Sub

Public Sub CalcularPesoPortasNaSheet(sheetPortas As Worksheet)
    If sheetPortas Is Nothing Then Exit Sub

    On Error GoTo Erro
    sheetPortas.Range(PESO_CELULA_TOTAL).Value2 = TotalPesoDaSheet(sheetPortas)
    Exit Sub

Erro:
    Debug.Print "Peso nao calculado em " & sheetPortas.Name & ": " & Err.Description
End Sub

Public Sub AgendarCalculoPesoPortas()
    On Error GoTo Erro
    Application.OnTime EarliestTime:=Now + TimeValue("00:00:01"), _
                       Procedure:=NomeProcedimentoCompleto("CalcularPesoPortas")
    Exit Sub

Erro:
    Debug.Print "Falha ao agendar o calculo de peso: " & Err.Description
End Sub

Public Sub TestarCalculoPesoPortas()
    Dim sheetAtiva As Worksheet
    If Not TentarObterSheetAtiva(sheetAtiva) Then Exit Sub

    RecalcularSheet sheetAtiva
    CalcularPesoPortasNaSheet sheetAtiva

    MsgBox "Total calculado em " & PESO_CELULA_TOTAL & ": " & _
           sheetAtiva.Range(PESO_CELULA_TOTAL).Value2, vbOKOnly
End Sub

Public Sub DiagnosticarPesoPortas()
    Dim sheetAtiva As Worksheet
    If Not TentarObterSheetAtiva(sheetAtiva) Then Exit Sub

    RecalcularSheet sheetAtiva

    Dim quantidades As Variant, nomes As Variant, tamanhos As Variant
    LerItens sheetAtiva, quantidades, nomes, tamanhos

    Dim total As Double
    total = TotalPesoPortas(quantidades, nomes, tamanhos)
    sheetAtiva.Range(PESO_CELULA_TOTAL).Value2 = total

    MsgBox MontarDiagnostico(quantidades, nomes, tamanhos, total), vbOKOnly
End Sub

Private Function MontarDiagnostico(ByVal quantidades As Variant, ByVal nomes As Variant, _
                                   ByVal tamanhos As Variant, ByVal total As Double) As String
    Dim linhasComQuantidade As Long
    Dim pesosDesconhecidos As Long
    ContarInconsistencias quantidades, nomes, tamanhos, linhasComQuantidade, pesosDesconhecidos

    MontarDiagnostico = "Linhas com quantidade: " & linhasComQuantidade & vbCrLf & _
                        "Linhas sem peso na tabela: " & pesosDesconhecidos & vbCrLf & _
                        "Total calculado: " & total
End Function

Private Sub ContarInconsistencias(ByVal quantidades As Variant, ByVal nomes As Variant, ByVal tamanhos As Variant, _
                                  ByRef linhasComQuantidade As Long, ByRef pesosDesconhecidos As Long)
    Dim indice As Long
    For indice = LBound(quantidades, 1) To UBound(quantidades, 1)
        If ConverterNumero(quantidades(indice, 1)) > 0 Then
            linhasComQuantidade = linhasComQuantidade + 1
            If PesoUnitarioLinha(nomes(indice, 1), tamanhos(indice, 1)) = 0 Then pesosDesconhecidos = pesosDesconhecidos + 1
        End If
    Next indice
End Sub

Private Function TotalPesoDaSheet(sheetPortas As Worksheet) As Double
    Dim quantidades As Variant, nomes As Variant, tamanhos As Variant
    LerItens sheetPortas, quantidades, nomes, tamanhos

    TotalPesoDaSheet = TotalPesoPortas(quantidades, nomes, tamanhos)
End Function

Private Sub LerItens(sheetPortas As Worksheet, ByRef quantidades As Variant, ByRef nomes As Variant, ByRef tamanhos As Variant)
    quantidades = LerColunaItens(sheetPortas, PESO_COLUNA_QUANTIDADE)
    nomes = LerColunaItens(sheetPortas, PESO_COLUNA_NOME)
    tamanhos = LerColunaItens(sheetPortas, PESO_COLUNA_TAMANHO)
End Sub

Private Function LerColunaItens(sheetPortas As Worksheet, ByVal coluna As String) As Variant
    LerColunaItens = sheetPortas.Range(coluna & PESO_PRIMEIRA_LINHA & ":" & coluna & PESO_ULTIMA_LINHA).Value2
End Function

' Publica para teste unitario: aceita matrizes (1 To n, 1 To 1).
Public Function TotalPesoPortas(ByVal quantidades As Variant, ByVal nomes As Variant, ByVal tamanhos As Variant) As Double
    Dim total As Double
    Dim indice As Long

    For indice = LBound(quantidades, 1) To UBound(quantidades, 1)
        total = total + PesoLinha(quantidades(indice, 1), nomes(indice, 1), tamanhos(indice, 1))
    Next indice

    TotalPesoPortas = total
End Function

Private Function PesoLinha(ByVal quantidade As Variant, ByVal nomePorta As Variant, ByVal tamanho As Variant) As Double
    Dim unidades As Double
    unidades = ConverterNumero(quantidade)
    If unidades <= 0 Then Exit Function

    PesoLinha = unidades * PesoUnitarioLinha(nomePorta, tamanho)
End Function

Private Function PesoUnitarioLinha(ByVal nomePorta As Variant, ByVal tamanho As Variant) As Double
    PesoUnitarioLinha = ObterPesoPorta(IdentificarTipoPorta(ConverterTexto(nomePorta)), ConverterNumero(tamanho))
End Function

Public Function IdentificarTipoPorta(ByVal nomePorta As String) As String
    Dim nomeNormalizado As String
    nomeNormalizado = UCase$(nomePorta)

    If InStr(1, nomeNormalizado, "DBL") > 0 Then
        IdentificarTipoPorta = "DBL"
        Exit Function
    End If

    If InStr(1, nomeNormalizado, "COL") > 0 Then
        IdentificarTipoPorta = "COL"
        Exit Function
    End If

    IdentificarTipoPorta = "LISA"
End Function

Public Function ObterPesoPorta(ByVal tipoPorta As String, ByVal tamanho As Double) As Double
    Dim tamanhoChave As Long
    tamanhoChave = ArredondarChave(tamanho)
    If tamanhoChave = 0 Then Exit Function

    Select Case tipoPorta
        Case "DBL": ObterPesoPorta = PesoDoorblanks(tamanhoChave)
        Case "COL": ObterPesoPorta = PesoColmeia(tamanhoChave)
        Case Else: ObterPesoPorta = PesoLisa(tamanhoChave)
    End Select
End Function

Private Function TentarObterSheetAtiva(ByRef sheetAtiva As Worksheet) As Boolean
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Function

    Set sheetAtiva = ActiveSheet
    TentarObterSheetAtiva = True
End Function

Private Function NomeProcedimentoCompleto(ByVal nomeProcedimento As String) As String
    NomeProcedimentoCompleto = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & nomeProcedimento
End Function

Private Function PesoDoorblanks(ByVal tamanhoCentimetros As Long) As Double
    Select Case tamanhoCentimetros
        Case 600: PesoDoorblanks = 23.4
        Case 700: PesoDoorblanks = 25
        Case 800: PesoDoorblanks = 33.5
        Case 900: PesoDoorblanks = 35.7
        Case 1000: PesoDoorblanks = 38.5
        Case 1100: PesoDoorblanks = 43.6
        Case 1200: PesoDoorblanks = 46.1
        Case 620: PesoDoorblanks = 24
        Case 720: PesoDoorblanks = 25.4
        Case 820: PesoDoorblanks = 35.9
        Case 920: PesoDoorblanks = 37.5
        Case 1020: PesoDoorblanks = 40
        Case 1120: PesoDoorblanks = 42
        Case 1220: PesoDoorblanks = 45
        Case Else: PesoDoorblanks = 0
    End Select
End Function

Private Function PesoColmeia(ByVal tamanhoCentimetros As Long) As Double
    Select Case tamanhoCentimetros
        Case 600: PesoColmeia = 11.1
        Case 700: PesoColmeia = 13
        Case 800: PesoColmeia = 14.2
        Case 900: PesoColmeia = 15.7
        Case 1000: PesoColmeia = 17.8
        Case 1100: PesoColmeia = 19
        Case 1200: PesoColmeia = 21
        Case 620: PesoColmeia = 11.4
        Case 720: PesoColmeia = 13.3
        Case 820: PesoColmeia = 14.5
        Case 920: PesoColmeia = 16.1
        Case 1020: PesoColmeia = 18.1
        Case 1120: PesoColmeia = 19.5
        Case 1220: PesoColmeia = 22
        Case Else: PesoColmeia = 0
    End Select
End Function

Private Function PesoLisa(ByVal tamanhoCentimetros As Long) As Double
    Select Case tamanhoCentimetros
        Case 600: PesoLisa = 18.5
        Case 700: PesoLisa = 20.5
        Case 800: PesoLisa = 23.7
        Case 900: PesoLisa = 26.5
        Case 1000: PesoLisa = 29.2
        Case 1100: PesoLisa = 33
        Case 1200: PesoLisa = 35.8
        Case 620: PesoLisa = 18.7
        Case 720: PesoLisa = 21.1
        Case 820: PesoLisa = 23.6
        Case 920: PesoLisa = 26.8
        Case 1020: PesoLisa = 28.9
        Case 1120: PesoLisa = 33.1
        Case 1220: PesoLisa = 35.9
        Case Else: PesoLisa = 0
    End Select
End Function
