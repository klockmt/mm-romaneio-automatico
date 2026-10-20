Attribute VB_Name = "modEsperaBanco"
' =====================================================================
' modEsperaBanco
' Aguarda a fonte do banco ficar pronta antes do Substituir processar
' os dados. Tambem tenta atualizar QueryTables/ListObjects.
' Timeout total mantido em 20s; a granularidade caiu de 1s para 250ms.
' Application.Run foi removido (o refresh e feito via objeto).
' =====================================================================
Option Explicit

Private Const ESPERA_TIMEOUT_SEGUNDOS As Double = 20
Private Const ESPERA_INTERVALO_SEGUNDOS As Double = 0.25
Private Const ESPERA_INTERVALO_FONTE_1 As String = "Y11:Y56"
Private Const ESPERA_INTERVALO_FONTE_2 As String = "W11:W56"

Public Function EsperarFonteBanco(sheetAtual As Worksheet) As Boolean
    If sheetAtual Is Nothing Then Exit Function

    If FonteBancoProntaInterna(sheetAtual) Then
        EsperarFonteBanco = True
        Exit Function
    End If

    If Not ConsultasBancoAtivas(sheetAtual) Then AtualizarBanco sheetAtual

    EsperarFonteBanco = AguardarFonte(sheetAtual)
End Function

Public Function FonteBancoPronta(sheetAtual As Worksheet) As Boolean
    If sheetAtual Is Nothing Then Exit Function
    FonteBancoPronta = FonteBancoProntaInterna(sheetAtual)
End Function

Private Function AguardarFonte(sheetAtual As Worksheet) As Boolean
    Dim limite As Double
    limite = Timer + ESPERA_TIMEOUT_SEGUNDOS

    Do While Timer <= limite
        If FonteBancoProntaInterna(sheetAtual) Then
            AguardarFonte = True
            Exit Function
        End If

        If Not AguardarSegundos(ESPERA_INTERVALO_SEGUNDOS) Then Exit Function
    Loop
End Function

' Devolve False na virada de meia-noite (Timer reinicia em 0),
' evitando loop infinito.
Private Function AguardarSegundos(ByVal segundos As Double) As Boolean
    Dim inicio As Double
    inicio = Timer

    Do While Timer - inicio < segundos
        DoEvents
        If Timer < inicio Then Exit Function
    Loop

    AguardarSegundos = True
End Function

Private Function FonteBancoProntaInterna(sheetAtual As Worksheet) As Boolean
    If ConsultasBancoAtivas(sheetAtual) Then Exit Function
    FonteBancoProntaInterna = FonteBancoDisponivel(sheetAtual)
End Function

Private Function FonteBancoDisponivel(sheetAtual As Worksheet) As Boolean
    If IntervaloTemTexto(sheetAtual.Range(ESPERA_INTERVALO_FONTE_1)) Then
        FonteBancoDisponivel = True
        Exit Function
    End If

    FonteBancoDisponivel = IntervaloTemTexto(sheetAtual.Range(ESPERA_INTERVALO_FONTE_2))
End Function

' Le o intervalo inteiro de uma vez: 46 leituras COM viram 1.
Private Function IntervaloTemTexto(intervalo As Range) As Boolean
    Dim valores As Variant
    valores = intervalo.Value2

    If Not IsArray(valores) Then
        IntervaloTemTexto = (Len(ConverterTexto(valores)) > 0)
        Exit Function
    End If

    IntervaloTemTexto = MatrizTemTexto(valores)
End Function

Private Function MatrizTemTexto(ByVal valores As Variant) As Boolean
    Dim linha As Long
    For linha = LBound(valores, 1) To UBound(valores, 1)
        If LinhaTemTexto(valores, linha) Then
            MatrizTemTexto = True
            Exit Function
        End If
    Next linha
End Function

Private Function LinhaTemTexto(ByVal valores As Variant, ByVal linha As Long) As Boolean
    Dim coluna As Long
    For coluna = LBound(valores, 2) To UBound(valores, 2)
        If Len(ConverterTexto(valores(linha, coluna))) > 0 Then
            LinhaTemTexto = True
            Exit Function
        End If
    Next coluna
End Function

Private Function ConsultasBancoAtivas(sheetAtual As Worksheet) As Boolean
    If QueryTablesAtualizando(sheetAtual) Then
        ConsultasBancoAtivas = True
        Exit Function
    End If

    ConsultasBancoAtivas = ListObjectsAtualizando(sheetAtual)
End Function

Private Function QueryTablesAtualizando(sheetAtual As Worksheet) As Boolean
    Dim consulta As QueryTable
    For Each consulta In sheetAtual.QueryTables
        If consulta.Refreshing Then
            QueryTablesAtualizando = True
            Exit Function
        End If
    Next consulta
End Function

Private Function ListObjectsAtualizando(sheetAtual As Worksheet) As Boolean
    Dim tabela As ListObject
    For Each tabela In sheetAtual.ListObjects
        If TabelaAtualizando(tabela) Then
            ListObjectsAtualizando = True
            Exit Function
        End If
    Next tabela
End Function

Private Function TabelaAtualizando(tabela As ListObject) As Boolean
    If tabela.QueryTable Is Nothing Then Exit Function
    TabelaAtualizando = tabela.QueryTable.Refreshing
End Function

Private Sub AtualizarBanco(sheetAtual As Worksheet)
    Dim consulta As QueryTable
    For Each consulta In sheetAtual.QueryTables
        AtualizarQueryTable consulta
    Next consulta

    Dim tabela As ListObject
    For Each tabela In sheetAtual.ListObjects
        AtualizarListObject tabela
    Next tabela

    Application.CalculateUntilAsyncQueriesDone
    sheetAtual.Calculate
End Sub

' Erro isolado por consulta: uma fonte quebrada nao aborta as demais.
Private Sub AtualizarQueryTable(consulta As QueryTable)
    On Error Resume Next
    consulta.Refresh BackgroundQuery:=False

    If Err.Number <> 0 Then
        Debug.Print "Falha no refresh da QueryTable: " & Err.Description
        Err.Clear
    End If

    On Error GoTo 0
End Sub

Private Sub AtualizarListObject(tabela As ListObject)
    If tabela.QueryTable Is Nothing Then Exit Sub

    On Error Resume Next
    tabela.QueryTable.Refresh BackgroundQuery:=False

    If Err.Number <> 0 Then
        Debug.Print "Falha no refresh da tabela " & tabela.Name & ": " & Err.Description
        Err.Clear
    End If

    On Error GoTo 0
End Sub
