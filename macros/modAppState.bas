Attribute VB_Name = "modAppState"
' =====================================================================
' modAppState
' Centraliza o modo de execucao rapida (ScreenUpdating, Calculation,
' EnableEvents, DisplayAlerts) com contagem de aninhamento.
' Existe porque macros chamam umas as outras: sem contador, a macro
' interna restaurava o ambiente no meio da externa.
' =====================================================================
Option Explicit

Private Type ConfiguracaoExcel
    atualizacaoTela As Boolean
    eventosAtivos As Boolean
    alertasAtivos As Boolean
    modoCalculo As XlCalculation
End Type

Private configuracaoOriginal As ConfiguracaoExcel
Private profundidadeModoRapido As Long

Public Sub EntrarModoRapido()
    profundidadeModoRapido = profundidadeModoRapido + 1
    If profundidadeModoRapido > 1 Then Exit Sub

    SalvarConfiguracaoAtual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
End Sub

Public Sub SairModoRapido()
    If profundidadeModoRapido = 0 Then Exit Sub

    profundidadeModoRapido = profundidadeModoRapido - 1
    If profundidadeModoRapido > 0 Then Exit Sub

    RestaurarConfiguracao
End Sub

' Usado somente por handlers de erro: derruba qualquer aninhamento pendente.
Public Sub AbortarModoRapido()
    profundidadeModoRapido = 0
    RestaurarConfiguracao
End Sub

Public Function ModoRapidoAtivo() As Boolean
    ModoRapidoAtivo = (profundidadeModoRapido > 0)
End Function

' Calculo esta em manual durante o modo rapido; formulas de apoio
' (ex.: coluna AI do peso) precisam de recalculo explicito.
Public Sub RecalcularSheet(sheetAlvo As Worksheet)
    If sheetAlvo Is Nothing Then Exit Sub

    sheetAlvo.Calculate
    Application.CalculateUntilAsyncQueriesDone
End Sub

Private Sub SalvarConfiguracaoAtual()
    configuracaoOriginal.atualizacaoTela = Application.ScreenUpdating
    configuracaoOriginal.eventosAtivos = Application.EnableEvents
    configuracaoOriginal.alertasAtivos = Application.DisplayAlerts
    configuracaoOriginal.modoCalculo = Application.Calculation
End Sub

Private Sub RestaurarConfiguracao()
    Application.Calculation = ModoCalculoRestaurado()
    Application.DisplayAlerts = configuracaoOriginal.alertasAtivos
    Application.EnableEvents = configuracaoOriginal.eventosAtivos
    Application.ScreenUpdating = configuracaoOriginal.atualizacaoTela
End Sub

' Type nao inicializado devolve 0, que nao e um XlCalculation valido.
Private Function ModoCalculoRestaurado() As XlCalculation
    If configuracaoOriginal.modoCalculo = 0 Then
        ModoCalculoRestaurado = xlCalculationAutomatic
        Exit Function
    End If

    ModoCalculoRestaurado = configuracaoOriginal.modoCalculo
End Function
