Attribute VB_Name = "modRomaneio"
' =====================================================================
' modRomaneio
' Le o arquivo de origem, extrai carga/pedido/total/sequencia e cria ou
' atualiza as abas de pedido no arquivo de destino.
' Coluna da sequencia configuravel em TEMPLATE!T1 (T1 = 0 desativa).
' Coluna do total configuravel em TEMPLATE!V1 (V1 = 0 desativa).
' Leitura da origem em bloco unico; workbook de origem fechado ao final
' se foi aberto pelo macro; T1/V1 lidos uma unica vez por execucao.
' =====================================================================
Option Explicit

Private Const ROM_LINHA_INICIO As Long = 8
Private Const ROM_COL_CARGA As Long = 2
Private Const ROM_COL_PEDIDO As Long = 3
Private Const ROM_COL_TOTAL As Long = 11
Private Const ROM_COL_SEQUENCIA_PADRAO As Long = 12
Private Const ROM_SHEET_MODELO As String = "template"
Private Const ROM_CELULA_CARGA As String = "P5"
Private Const ROM_CELULA_PEDIDO As String = "P3"
Private Const ROM_CONFIG_COLUNA_SEQUENCIA As String = "T1"
Private Const ROM_CONFIG_COLUNA_TOTAL As String = "V1"

Public Const IDX_CARGA As Long = 0
Public Const IDX_PEDIDO As Long = 1
Public Const IDX_TOTAL As Long = 2
Public Const IDX_SEQUENCIA As Long = 3

Public Sub ProcessarPedidos()
    Dim caminhoOrigem As String
    caminhoOrigem = SelecionarCaminhoOrigem()
    If Len(caminhoOrigem) = 0 Then Exit Sub

    ExecutarRomaneio caminhoOrigem
End Sub

Private Sub ExecutarRomaneio(ByVal caminhoOrigem As String)
    Dim abertoAgora As Boolean
    Dim workbookOrigem As Workbook
    Set workbookOrigem = AbrirArquivoOrigem(caminhoOrigem, abertoAgora)
    If workbookOrigem Is Nothing Then Exit Sub

    On Error GoTo Finalizar
    ProcessarWorkbookOrigem workbookOrigem

Finalizar:
    If Err.Number <> 0 Then MsgBox "Falha no romaneio: " & Err.Description, vbCritical
    If abertoAgora Then workbookOrigem.Close SaveChanges:=False
End Sub

Private Sub ProcessarWorkbookOrigem(workbookOrigem As Workbook)
    Dim sheetOrigem As Worksheet
    Set sheetOrigem = IdentificarSheetOrigem(workbookOrigem)

    If sheetOrigem Is Nothing Then
        MsgBox "Nao foi possivel identificar a sheet de origem.", vbCritical
        Exit Sub
    End If

    ExtrairEProcessar sheetOrigem
End Sub

Private Sub ExtrairEProcessar(sheetOrigem As Worksheet)
    Dim colunaSequencia As Long
    colunaSequencia = ObterColunaSequencia()

    Dim colunaTotal As Long
    colunaTotal = ObterColunaTotal()

    Dim matrizPedidos As Variant
    matrizPedidos = ExtrairPedidos(sheetOrigem, colunaSequencia, colunaTotal)

    If IsEmpty(matrizPedidos) Then
        MsgBox "Nenhum pedido extraido.", vbCritical
        Exit Sub
    End If

    If Not ValidarPedidos(matrizPedidos, colunaTotal) Then Exit Sub
    If GerarSheetsPedido(matrizPedidos, colunaSequencia) < 0 Then Exit Sub

    ExecutarImportEPintarTodasExcetoTemplate matrizPedidos, colunaTotal > 0
End Sub

' Rejeita a carga quando nenhum pedido tem total: sinal de coluna
' errada na origem. Com V1 = 0 a validacao e pulada de proposito.
Private Function ValidarPedidos(ByVal matrizPedidos As Variant, ByVal colunaTotal As Long) As Boolean
    If colunaTotal = 0 Then
        ValidarPedidos = True
        Exit Function
    End If

    Dim totalOrigem As Double
    Dim indice As Long

    For indice = LBound(matrizPedidos, 2) To UBound(matrizPedidos, 2)
        totalOrigem = totalOrigem + CDbl(matrizPedidos(IDX_TOTAL, indice))
    Next indice

    Debug.Print "Pedidos: " & (UBound(matrizPedidos, 2) + 1) & " | Total origem: " & totalOrigem
    If totalOrigem > 0 Then ValidarPedidos = True

    If Not ValidarPedidos Then MsgBox "Total da origem zerado. Confira a coluna de total.", vbCritical
End Function

Private Function SelecionarCaminhoOrigem() As String
    Dim caminhoSelecionado As Variant
    caminhoSelecionado = Application.GetOpenFilename( _
        FileFilter:="Arquivos Excel (*.xlsx;*.xls;*.xlsb), *.xlsx;*.xls;*.xlsb", _
        Title:="Selecione o arquivo de origem")

    If VarType(caminhoSelecionado) = vbBoolean Then Exit Function
    SelecionarCaminhoOrigem = CStr(caminhoSelecionado)
End Function

Private Function AbrirArquivoOrigem(ByVal caminhoOrigem As String, ByRef abertoAgora As Boolean) As Workbook
    Dim nomeArquivo As String
    nomeArquivo = Dir(caminhoOrigem)

    Dim workbookExistente As Workbook
    On Error Resume Next
    Set workbookExistente = Workbooks(nomeArquivo)
    On Error GoTo 0

    If Not workbookExistente Is Nothing Then
        Set AbrirArquivoOrigem = workbookExistente
        Exit Function
    End If

    Set AbrirArquivoOrigem = AbrirSomenteLeitura(caminhoOrigem, abertoAgora)
End Function

Private Function AbrirSomenteLeitura(ByVal caminhoOrigem As String, ByRef abertoAgora As Boolean) As Workbook
    On Error GoTo Erro
    Set AbrirSomenteLeitura = Workbooks.Open(caminhoOrigem, ReadOnly:=True, UpdateLinks:=0)
    abertoAgora = True
    Exit Function

Erro:
    MsgBox "Nao foi possivel abrir o arquivo de origem: " & Err.Description, vbCritical
End Function

Private Function IdentificarSheetOrigem(workbookOrigem As Workbook) As Worksheet
    Dim sheetAtual As Worksheet

    For Each sheetAtual In workbookOrigem.Worksheets
        If SheetTemRegistros(sheetAtual) Then
            Set IdentificarSheetOrigem = sheetAtual
            Exit Function
        End If
    Next sheetAtual

    If workbookOrigem.Worksheets.Count > 0 Then Set IdentificarSheetOrigem = workbookOrigem.Worksheets(1)
End Function

Private Function SheetTemRegistros(sheetOrigem As Worksheet) As Boolean
    Dim cargaInicial As String
    Dim pedidoInicial As String

    cargaInicial = ConverterTexto(sheetOrigem.Cells(ROM_LINHA_INICIO, ROM_COL_CARGA).Value2)
    pedidoInicial = ConverterTexto(sheetOrigem.Cells(ROM_LINHA_INICIO, ROM_COL_PEDIDO).Value2)

    SheetTemRegistros = (Len(cargaInicial) > 0 Or Len(pedidoInicial) > 0)
End Function

Private Function ExtrairPedidos(sheetOrigem As Worksheet, ByVal colunaSequencia As Long, ByVal colunaTotal As Long) As Variant
    Dim ultimaLinha As Long
    ultimaLinha = UltimaLinhaOrigem(sheetOrigem, colunaSequencia, colunaTotal)
    If ultimaLinha < ROM_LINHA_INICIO Then Exit Function

    Dim bloco As Variant
    bloco = sheetOrigem.Range( _
        sheetOrigem.Cells(ROM_LINHA_INICIO, 1), _
        sheetOrigem.Cells(ultimaLinha, ColunaMaximaLeitura(colunaSequencia, colunaTotal))).Value2

    ExtrairPedidos = MontarMatrizPedidos(bloco, colunaSequencia, colunaTotal)
End Function

' Colunas desativadas valem 0 e nao podem entrar no maximo; o piso e a
' coluna do pedido, pois carga e pedido sao sempre obrigatorios.
Private Function ColunaMaximaLeitura(ByVal colunaSequencia As Long, ByVal colunaTotal As Long) As Long
    ColunaMaximaLeitura = ROM_COL_PEDIDO
    If colunaTotal > ColunaMaximaLeitura Then ColunaMaximaLeitura = colunaTotal
    If colunaSequencia > ColunaMaximaLeitura Then ColunaMaximaLeitura = colunaSequencia
End Function

Private Function UltimaLinhaOrigem(sheetOrigem As Worksheet, ByVal colunaSequencia As Long, ByVal colunaTotal As Long) As Long
    Dim linhaCarga As Long, linhaPedido As Long, linhaTotal As Long, linhaSequencia As Long

    linhaCarga = UltimaLinhaColuna(sheetOrigem, ROM_COL_CARGA)
    linhaPedido = UltimaLinhaColuna(sheetOrigem, ROM_COL_PEDIDO)
    linhaTotal = 1
    If colunaTotal > 0 Then linhaTotal = UltimaLinhaColuna(sheetOrigem, colunaTotal)
    linhaSequencia = 1
    If colunaSequencia > 0 Then linhaSequencia = UltimaLinhaColuna(sheetOrigem, colunaSequencia)

    UltimaLinhaOrigem = Application.WorksheetFunction.Max(linhaCarga, linhaPedido, linhaTotal, linhaSequencia)
End Function

Private Function UltimaLinhaColuna(sheetOrigem As Worksheet, ByVal coluna As Long) As Long
    UltimaLinhaColuna = sheetOrigem.Cells(sheetOrigem.Rows.Count, coluna).End(xlUp).Row
End Function

Private Function MontarMatrizPedidos(ByVal bloco As Variant, ByVal colunaSequencia As Long, ByVal colunaTotal As Long) As Variant
    Dim matrizPedidos() As Variant
    ReDim matrizPedidos(IDX_CARGA To IDX_SEQUENCIA, 0 To UBound(bloco, 1) - 1)

    Dim contador As Long
    Dim linha As Long

    For linha = LBound(bloco, 1) To UBound(bloco, 1)
        contador = contador + AdicionarPedido(bloco, linha, colunaSequencia, colunaTotal, matrizPedidos, contador)
    Next linha

    If contador = 0 Then Exit Function
    ReDim Preserve matrizPedidos(IDX_CARGA To IDX_SEQUENCIA, 0 To contador - 1)
    MontarMatrizPedidos = matrizPedidos
End Function

Private Function AdicionarPedido(ByVal bloco As Variant, ByVal linha As Long, ByVal colunaSequencia As Long, _
                                 ByVal colunaTotal As Long, ByRef matrizPedidos() As Variant, ByVal posicao As Long) As Long
    Dim numeroPedido As String
    numeroPedido = ConverterTexto(bloco(linha, ROM_COL_PEDIDO))
    If Len(numeroPedido) = 0 Then Exit Function

    matrizPedidos(IDX_CARGA, posicao) = ConverterTexto(bloco(linha, ROM_COL_CARGA))
    matrizPedidos(IDX_PEDIDO, posicao) = numeroPedido

    If colunaTotal > 0 Then
        matrizPedidos(IDX_TOTAL, posicao) = ConverterNumero(bloco(linha, colunaTotal))
    Else
        matrizPedidos(IDX_TOTAL, posicao) = 0
    End If

    matrizPedidos(IDX_SEQUENCIA, posicao) = SequenciaDaLinha(bloco, linha, colunaSequencia)

    AdicionarPedido = 1
End Function

Private Function SequenciaDaLinha(ByVal bloco As Variant, ByVal linha As Long, ByVal colunaSequencia As Long) As Long
    If colunaSequencia = 0 Then Exit Function
    If colunaSequencia > UBound(bloco, 2) Then Exit Function

    SequenciaDaLinha = ConverterSequencia(bloco(linha, colunaSequencia))
End Function

Private Function GerarSheetsPedido(ByVal matrizPedidos As Variant, ByVal colunaSequencia As Long) As Long
    If Not SheetExiste(ROM_SHEET_MODELO) Then
        MsgBox "Sheet modelo nao encontrada.", vbCritical
        GerarSheetsPedido = -1
        Exit Function
    End If

    Dim sheetModelo As Worksheet
    Set sheetModelo = ThisWorkbook.Worksheets(ROM_SHEET_MODELO)

    EntrarModoRapido
    On Error GoTo Erro

    GerarSheetsPedido = CriarSheetsDaMatriz(sheetModelo, matrizPedidos, colunaSequencia)
    SairModoRapido
    Exit Function

Erro:
    AbortarModoRapido
    MsgBox "Falha ao gerar as abas de pedido: " & Err.Description, vbCritical
    GerarSheetsPedido = -1
End Function

Private Function CriarSheetsDaMatriz(sheetModelo As Worksheet, ByVal matrizPedidos As Variant, ByVal colunaSequencia As Long) As Long
    Dim indice As Long
    For indice = LBound(matrizPedidos, 2) To UBound(matrizPedidos, 2)
        CriarSheetsDaMatriz = CriarSheetsDaMatriz + CriarSheetPedido( _
            sheetModelo, _
            CStr(matrizPedidos(IDX_CARGA, indice)), _
            CStr(matrizPedidos(IDX_PEDIDO, indice)), _
            CLng(matrizPedidos(IDX_SEQUENCIA, indice)), _
            colunaSequencia > 0)
    Next indice
End Function

Private Function CriarSheetPedido(sheetModelo As Worksheet, ByVal numeroCarga As String, ByVal numeroPedido As String, _
                                  ByVal sequencia As Long, ByVal usarSequencia As Boolean) As Long
    Dim nomeSheet As String
    nomeSheet = MontarNomeSheet(numeroPedido)

    Dim criadaAgora As Boolean
    Dim sheetDestino As Worksheet
    Set sheetDestino = ObterOuCopiarSheet(sheetModelo, nomeSheet, criadaAgora)
    If sheetDestino Is Nothing Then Exit Function

    sheetDestino.Range(ROM_CELULA_CARGA).Value2 = numeroCarga
    sheetDestino.Range(ROM_CELULA_PEDIDO).Value2 = numeroPedido
    If usarSequencia Then AplicarSequencia sheetDestino, sequencia

    If criadaAgora Then CriarSheetPedido = 1
End Function

Private Function ObterOuCopiarSheet(sheetModelo As Worksheet, ByVal nomeSheet As String, ByRef criadaAgora As Boolean) As Worksheet
    If SheetExiste(nomeSheet) Then
        Set ObterOuCopiarSheet = ThisWorkbook.Worksheets(nomeSheet)
        ObterOuCopiarSheet.Visible = xlSheetVisible
        Exit Function
    End If

    sheetModelo.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set ObterOuCopiarSheet = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    ObterOuCopiarSheet.Name = nomeSheet
    ObterOuCopiarSheet.Visible = xlSheetVisible
    criadaAgora = True
End Function

Public Function MontarNomeSheet(ByVal numeroPedido As String) As String
    Dim nomeBase As String
    nomeBase = "Pedido_" & Trim$(numeroPedido)

    MontarNomeSheet = Left$(RemoverCaracteresInvalidos(nomeBase), 31)
End Function

Public Function RemoverCaracteresInvalidos(ByVal texto As String) As String
    Dim caracteresInvalidos As Variant
    caracteresInvalidos = Array("\", "/", "?", "*", "[", "]", ":")

    Dim resultado As String
    resultado = texto

    Dim caracterInvalido As Variant
    For Each caracterInvalido In caracteresInvalidos
        resultado = Replace(resultado, CStr(caracterInvalido), "_")
    Next caracterInvalido

    RemoverCaracteresInvalidos = resultado
End Function

Private Function SheetExiste(ByVal nomeSheet As String) As Boolean
    Dim sheetProcurada As Object

    On Error Resume Next
    Set sheetProcurada = ThisWorkbook.Worksheets(nomeSheet)
    On Error GoTo 0

    SheetExiste = Not sheetProcurada Is Nothing
End Function

Private Function ObterColunaSequencia() As Long
    ObterColunaSequencia = LerColunaConfigurada(ROM_CONFIG_COLUNA_SEQUENCIA, ROM_COL_SEQUENCIA_PADRAO, True)
End Function

Private Function ObterColunaTotal() As Long
    ObterColunaTotal = LerColunaConfigurada(ROM_CONFIG_COLUNA_TOTAL, ROM_COL_TOTAL, True)
End Function

Private Function LerColunaConfigurada(ByVal enderecoConfig As String, ByVal colunaPadrao As Long, ByVal zeroDesativa As Boolean) As Long
    If Not SheetExiste(ROM_SHEET_MODELO) Then
        LerColunaConfigurada = colunaPadrao
        Exit Function
    End If

    Dim configuracao As String
    configuracao = ConverterTexto(ThisWorkbook.Worksheets(ROM_SHEET_MODELO).Range(enderecoConfig).Value2)

    If Len(configuracao) = 0 Then
        LerColunaConfigurada = colunaPadrao
    ElseIf IsNumeric(configuracao) Then
        LerColunaConfigurada = ColunaNumerica(configuracao, colunaPadrao, zeroDesativa)
    Else
        LerColunaConfigurada = ColunaPorLetra(configuracao, colunaPadrao)
    End If
End Function

Private Function ColunaNumerica(ByVal configuracao As String, ByVal colunaPadrao As Long, ByVal zeroDesativa As Boolean) As Long
    Dim coluna As Long
    coluna = CLng(ConverterNumero(configuracao))

    If coluna = 0 Then
        If zeroDesativa Then ColunaNumerica = 0 Else ColunaNumerica = colunaPadrao
        Exit Function
    End If

    If coluna < 1 Or coluna > ThisWorkbook.Worksheets(ROM_SHEET_MODELO).Columns.Count Then
        ColunaNumerica = colunaPadrao
        Exit Function
    End If

    ColunaNumerica = coluna
End Function

Private Function ColunaPorLetra(ByVal configuracao As String, ByVal colunaPadrao As Long) As Long
    Dim letras As String
    letras = ExtrairLetrasIniciais(configuracao)

    If Len(letras) = 0 Then
        ColunaPorLetra = colunaPadrao
        Exit Function
    End If

    On Error GoTo Erro
    ColunaPorLetra = ThisWorkbook.Worksheets(ROM_SHEET_MODELO).Range(letras & "1").Column
    Exit Function

Erro:
    ColunaPorLetra = colunaPadrao
End Function

Public Function ExtrairLetrasIniciais(ByVal texto As String) As String
    Dim letras As String
    Dim caracter As String
    Dim indice As Long

    For indice = 1 To Len(texto)
        caracter = LetraValida(Mid$(texto, indice, 1))
        If Len(caracter) = 0 And Len(letras) > 0 Then Exit For
        letras = letras & caracter
    Next indice

    ExtrairLetrasIniciais = letras
End Function

Private Function LetraValida(ByVal caracter As String) As String
    If caracter >= "A" And caracter <= "Z" Then LetraValida = caracter
    If caracter >= "a" And caracter <= "z" Then LetraValida = caracter
End Function

