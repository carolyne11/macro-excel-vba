Attribute VB_Name = "Módulo1"
Sub Importar_Dados_Com_Mes()
macro -Excel - VBA
    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim arquivo As Variant
    Dim ultimaLinha As Long
    Dim linhaDestino As Long

    Dim colFornecedor As Variant
    Dim colDescricao As Variant
    Dim colConta As Variant
    Dim colValor As Variant
    Dim colCusto As Variant
    Dim colOrdem As Variant

    Dim mesInformado As String
    Dim dataMes As Date

    ' ?? Perguntar o mês
    mesInformado = InputBox("Informe o mês no formato MM/AAAA (ex: 01/2026):", "Mês de Referência")

    If mesInformado = "" Then Exit Sub

    dataMes = DateSerial(Right(mesInformado, 4), Left(mesInformado, 2), 1)

    ' ?? Selecionar arquivo
    arquivo = Application.GetOpenFilename("Arquivos Excel (*.xlsx), *.xlsx")

    If arquivo = False Then Exit Sub

    Set wbOrigem = Workbooks.Open(arquivo)
    Set wsOrigem = wbOrigem.Sheets(1)

    ' ?? Aba destino
    Set wsDestino = ThisWorkbook.Sheets(1)

    ' ?? Encontrar colunas (IGUAL ao código que funciona)
    colFornecedor = Application.Match("Fornecedor", wsOrigem.Rows(1), 0)
    colDescricao = Application.Match("Descrição Conta Contábil", wsOrigem.Rows(1), 0)
    colConta = Application.Match("Conta Contábil", wsOrigem.Rows(1), 0)
    colValor = Application.Match("Valor BRL", wsOrigem.Rows(1), 0)
    colCusto = Application.Match("Centro de Custo", wsOrigem.Rows(1), 0)
    colOrdem = Application.Match("Ordem Interna", wsOrigem.Rows(1), 0)
    

    ' ?? Última linha
    ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, colFornecedor).End(xlUp).Row

    ' ?? Cabeçalhos (agora com Mês)
    wsDestino.Range("A1:F2").Value = Array("Mês", "Fornecedor", "Descrição", "Conta", "Valor", "Centro De Custos", "Ordem Interna")

    ' ?? Começar SEMPRE na linha 2 (como no código antigo)
   linhaDestino = wsDestino.Cells(wsDestino.Rows.Count, 1).End(xlUp).Row + 1
If linhaDestino < 2 Then linhaDestino = 2


    ' ?? Copiar dados
    Dim i As Long
   For i = 2 To ultimaLinha

    ' Só copia se houver fornecedor (linha com dados)
    If Trim(wsOrigem.Cells(i, colFornecedor).Value) <> "" Then

        wsDestino.Cells(linhaDestino, 1).Value = dataMes
        wsDestino.Cells(linhaDestino, 2).Value = wsOrigem.Cells(i, colFornecedor).Value
        wsDestino.Cells(linhaDestino, 3).Value = wsOrigem.Cells(i, colDescricao).Value
        wsDestino.Cells(linhaDestino, 4).Value = wsOrigem.Cells(i, colConta).Value
        wsDestino.Cells(linhaDestino, 5).Value = wsOrigem.Cells(i, colValor).Value
        wsDestino.Cells(linhaDestino, 6).Value = wsOrigem.Cells(i, colCusto).Value
        wsDestino.Cells(linhaDestino, 7).Value = wsOrigem.Cells(i, colOrdem).Value

        linhaDestino = linhaDestino + 1

    End If

Next i

    
wsDestino.Range(wsDestino.Cells(2, 1), wsDestino.Cells(linhaDestino - 1, 1)).NumberFormat = "mm/yyyy"


    wbOrigem.Close False

    MsgBox "Importação concluída com sucesso!", vbInformation

End Sub





