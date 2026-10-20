Attribute VB_Name = "modConferenciaTotais"
' =====================================================================
' modConferenciaTotais
' Compara o total de cada pedido na origem com o total calculado na
' aba do pedido e monta o relatorio de divergencias.
' Nucleo puro: recebe valores, nao toca em Range.
' =====================================================================
Option Explicit

Private Const CONF_TOLERANCIA_TOTAL As Double = 0.01
Private Const IDX_REG_PEDIDO As Long = 0
Private Const IDX_REG_ORIGEM As Long = 1
Private Const IDX_REG_PLANILHA As Long = 2
Private Const IDX_REG_BATEU As Long = 3

Public Function TotaisBatem(ByVal totalOrigem As Double, ByVal totalPlanilha As Double) As Boolean
    TotaisBatem = (Abs(totalOrigem - totalPlanilha) <= CONF_TOLERANCIA_TOTAL)
End Function

Public Function NovoRegistroConferencia(ByVal nomePedido As String, ByVal totalOrigem As Double, ByVal totalPlanilha As Double) As Variant
    NovoRegistroConferencia = Array(nomePedido, totalOrigem, totalPlanilha, TotaisBatem(totalOrigem, totalPlanilha))
End Function

Public Function ConferenciaTemDivergencia(ByVal registros As Collection) As Boolean
    Dim registro As Variant
    For Each registro In registros
        If Not CBool(registro(IDX_REG_BATEU)) Then
            ConferenciaTemDivergencia = True
            Exit Function
        End If
    Next registro
End Function

Public Function MontarRelatorioConferencia(ByVal registros As Collection) As String
    Dim divergentes As String
    Dim conferem As String
    Dim registro As Variant

    For Each registro In registros
        If CBool(registro(IDX_REG_BATEU)) Then
            conferem = conferem & LinhaRelatorio(registro, False)
        Else
            divergentes = divergentes & LinhaRelatorio(registro, True)
        End If
    Next registro

    If Len(divergentes) = 0 Then divergentes = "(nenhum)" & vbCrLf
    If Len(conferem) = 0 Then conferem = "(nenhum)" & vbCrLf

    MontarRelatorioConferencia = "NAO BATERAM:" & vbCrLf & divergentes & vbCrLf & "BATERAM:" & vbCrLf & conferem
End Function

Private Function LinhaRelatorio(ByVal registro As Variant, ByVal mostrarOrigem As Boolean) As String
    LinhaRelatorio = CStr(registro(IDX_REG_PEDIDO)) & " | total na planilha: " & Format$(CDbl(registro(IDX_REG_PLANILHA)), "#,##0.00")
    If mostrarOrigem Then LinhaRelatorio = LinhaRelatorio & " | total na origem: " & Format$(CDbl(registro(IDX_REG_ORIGEM)), "#,##0.00")
    LinhaRelatorio = LinhaRelatorio & vbCrLf
End Function
