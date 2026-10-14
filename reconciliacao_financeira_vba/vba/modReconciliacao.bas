Attribute VB_Name = "modReconciliacao"
Option Explicit

' Reconciliação financeira automatizada — dados fictícios.
' Abas esperadas:
'   Extrato: ID_Transacao | Data_Pagamento | Fornecedor | Nota_Fiscal | Valor | Moeda | Status_Banco
'   Notas_Fiscais: Nota_Fiscal | Data_Emissao | Fornecedor | Valor_Nota | Moeda | Referencia_Servico | Status_Nota
' A macro cria/atualiza a aba "Reconciliação" e não altera as abas de origem.

Public Sub ReconciliarPagamentos()
    Dim wsE As Worksheet, wsN As Worksheet, wsR As Worksheet
    Dim lastE As Long, lastN As Long, i As Long, j As Long, outRow As Long
    Dim invoiceE As String, invoiceN As String, supplierE As String, supplierN As String
    Dim currencyE As String, currencyN As String
    Dim amountE As Currency, amountN As Currency
    Dim bankStatus As String, invoiceStatus As String
    Dim found As Boolean, amountMatch As Boolean, supplierMatch As Boolean
    Dim exactInvoice As Boolean, dateGap As Long
    Dim matchedRows As Object, matchedBank As Object
    Dim key As String, statusResult As String, detail As String
    Dim tolerance As Currency

    On Error GoTo TrataErro
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    tolerance = 0.01 ' Tolerância monetária: 1 centavo.

    Set wsE = ThisWorkbook.Worksheets("Extrato")
    Set wsN = ThisWorkbook.Worksheets("Notas_Fiscais")
    Set wsR = GetOrCreateSheet("Reconciliação")
    Set matchedRows = CreateObject("Scripting.Dictionary")
    Set matchedBank = CreateObject("Scripting.Dictionary")

    wsR.Cells.Clear
    wsR.Range("A1:K1").Value = Array("ID_Transacao", "Nota_Extrato", "Fornecedor_Extrato", _
        "Valor_Extrato", "Nota_Livro", "Fornecedor_Livro", "Valor_Nota", _
        "Diferença_Valor", "Diferença_Dias", "Resultado", "Detalhes")
    wsR.Range("A1:K1").Font.Bold = True
    wsR.Range("A1:K1").Interior.Color = RGB(31, 78, 121)
    wsR.Range("A1:K1").Font.Color = RGB(255, 255, 255)

    lastE = wsE.Cells(wsE.Rows.Count, "A").End(xlUp).Row
    lastN = wsN.Cells(wsN.Rows.Count, "A").End(xlUp).Row
    outRow = 2

    ' Percorre cada linha do extrato e procura a melhor correspondência no livro.
    For i = 2 To lastE
        invoiceE = NormalizeText(wsE.Cells(i, "D").Value)
        supplierE = NormalizeText(wsE.Cells(i, "C").Value)
        currencyE = UCase$(Trim$(CStr(wsE.Cells(i, "F").Value)))
        amountE = CCur(Val(Replace(CStr(wsE.Cells(i, "E").Value), ",", ".")))
        bankStatus = UCase$(Trim$(CStr(wsE.Cells(i, "G").Value)))
        found = False
        exactInvoice = False
        detail = ""

        ' Estornos devem ser reportados separadamente, sem reconciliar como pagamento efetivo.
        If bankStatus = "ESTORNADO" Or bankStatus = "REVERSED" Then
            statusResult = "REVISAR: ESTORNO"
            detail = "Transação estornada no extrato; validar lançamento contábil."
            WriteResult wsR, outRow, wsE, i, 0, statusResult, detail, "", "", "", "", ""
            outRow = outRow + 1
            GoTo NextBankRow
        End If

        ' Prioridade 1: nota fiscal + fornecedor + moeda.
        ' Evita reutilizar a mesma linha de nota para duas transações.
        For j = 2 To lastN
            key = CStr(j)
            If Not matchedRows.Exists(key) Then
                invoiceN = NormalizeText(wsN.Cells(j, "A").Value)
                supplierN = NormalizeText(wsN.Cells(j, "C").Value)
                currencyN = UCase$(Trim$(CStr(wsN.Cells(j, "E").Value)))
                amountN = CCur(Val(Replace(CStr(wsN.Cells(j, "D").Value), ",", ".")))
                invoiceStatus = UCase$(Trim$(CStr(wsN.Cells(j, "G").Value)))

                If invoiceE <> "" And invoiceE = invoiceN _
                   And supplierE = supplierN And currencyE = currencyN Then
                    found = True
                    exactInvoice = True
                    Exit For
                End If
            End If
        Next j

        If found Then
            Dim supplierN2 As String, currencyN2 As String, invoiceN2 As String
            Dim dateN As Variant, dateE As Variant
            supplierN2 = NormalizeText(wsN.Cells(j, "C").Value)
            currencyN2 = UCase$(Trim$(CStr(wsN.Cells(j, "E").Value)))
            invoiceN2 = CStr(wsN.Cells(j, "A").Value)
            amountN = CCur(Val(Replace(CStr(wsN.Cells(j, "D").Value), ",", ".")))
            invoiceStatus = UCase$(Trim$(CStr(wsN.Cells(j, "G").Value)))
            dateE = wsE.Cells(i, "B").Value
            dateN = wsN.Cells(j, "B").Value
            dateGap = 0
            If IsDate(dateE) And IsDate(dateN) Then dateGap = Abs(DateDiff("d", CDate(dateE), CDate(dateN)))

            If invoiceStatus = "ESTORNADA" Or invoiceStatus = "REVERSED" Then
                statusResult = "REVISAR: NOTA ESTORNADA"
                detail = "A nota correspondente está marcada como estornada."
            ElseIf invoiceStatus = "PENDENTE" Or invoiceStatus = "PENDING" Then
                statusResult = "REVISAR: NOTA PENDENTE"
                detail = "A nota correspondente ainda consta como pendente."
            ElseIf Abs(amountE - amountN) <= tolerance Then
                statusResult = "CONCILIADO"
                detail = "Nota, fornecedor, moeda e valor conferem."
            Else
                statusResult = "DIVERGÊNCIA DE VALOR"
                detail = "Nota, fornecedor e moeda conferem, mas os valores divergem."
            End If

            matchedRows(CStr(j)) = True
            matchedBank(CStr(i)) = True
            WriteResult wsR, outRow, wsE, i, j, statusResult, detail, _
                wsN.Cells(j, "A").Value, wsN.Cells(j, "C").Value, amountN, _
                amountE - amountN, dateGap
        Else
            ' Prioridade 2: busca alternativa por fornecedor + moeda + valor.
            ' Serve para encontrar possíveis divergências no número da nota.
            Dim candidate As Long, candidates As Long
            candidate = 0: candidates = 0
            For j = 2 To lastN
                key = CStr(j)
                If Not matchedRows.Exists(key) Then
                    supplierN = NormalizeText(wsN.Cells(j, "C").Value)
                    currencyN = UCase$(Trim$(CStr(wsN.Cells(j, "E").Value)))
                    amountN = CCur(Val(Replace(CStr(wsN.Cells(j, "D").Value), ",", ".")))
                    invoiceStatus = UCase$(Trim$(CStr(wsN.Cells(j, "G").Value)))
                    If supplierE = supplierN And currencyE = currencyN _
                       And Abs(amountE - amountN) <= tolerance _
                       And invoiceStatus <> "ESTORNADA" And invoiceStatus <> "REVERSED" Then
                        candidate = j
                        candidates = candidates + 1
                    End If
                End If
            Next j

            If candidates = 1 Then
                j = candidate
                statusResult = "REVISAR: NOTA DIFERENTE"
                detail = "Fornecedor, moeda e valor conferem, mas o número da nota não."
                matchedRows(CStr(j)) = True
                matchedBank(CStr(i)) = True
                WriteResult wsR, outRow, wsE, i, j, statusResult, detail, _
                    wsN.Cells(j, "A").Value, wsN.Cells(j, "C").Value, _
                    CCur(Val(Replace(CStr(wsN.Cells(j, "D").Value), ",", "."))), _
                    amountE - CCur(Val(Replace(CStr(wsN.Cells(j, "D").Value), ",", "."))), ""
            ElseIf candidates > 1 Then
                statusResult = "REVISAR: MÚLTIPLAS CANDIDATAS"
                detail = "Mais de uma nota combina com fornecedor, moeda e valor; revisão manual necessária."
                WriteResult wsR, outRow, wsE, i, 0, statusResult, detail, "", "", "", "", ""
            Else
                statusResult = "SEM NOTA CORRESPONDENTE"
                detail = "Nenhuma nota disponível corresponde à nota, fornecedor e moeda do extrato."
                WriteResult wsR, outRow, wsE, i, 0, statusResult, detail, "", "", "", "", ""
            End If
        End If

        outRow = outRow + 1
NextBankRow:
    Next i

    ' Notas sem transação correspondente no extrato.
    For j = 2 To lastN
        If Not matchedRows.Exists(CStr(j)) Then
            statusResult = "SEM PAGAMENTO NO EXTRATO"
            detail = "Nota não foi associada a nenhuma transação do extrato."
            wsR.Cells(outRow, 5).Value = wsN.Cells(j, "A").Value
            wsR.Cells(outRow, 6).Value = wsN.Cells(j, "C").Value
            wsR.Cells(outRow, 7).Value = wsN.Cells(j, "D").Value
            wsR.Cells(outRow, 10).Value = statusResult
            wsR.Cells(outRow, 11).Value = detail
            outRow = outRow + 1
        End If
    Next j

    ' Formatação do relatório
    wsR.Columns("A:K").AutoFit
    wsR.Range("D:D").NumberFormat = """R$"" #,##0.00"
    wsR.Range("G:H").NumberFormat = """R$"" #,##0.00"
    wsR.Range("A1:K" & outRow - 1).AutoFilter
    wsR.Activate
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Reconciliação concluída. Registros analisados: " & (lastE - 1) & _
        ". Veja a aba 'Reconciliação'.", vbInformation
    Exit Sub

TrataErro:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Erro na reconciliação: " & Err.Description, vbExclamation
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Function NormalizeText(ByVal value As Variant) As String
    Dim s As String
    s = LCase$(Trim$(CStr(value)))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeText = s
End Function

Private Sub WriteResult(ByVal wsR As Worksheet, ByVal outRow As Long, _
    ByVal wsE As Worksheet, ByVal bankRow As Long, ByVal invoiceRow As Long, _
    ByVal resultText As String, ByVal detail As String, _
    ByVal invoiceNo As Variant, ByVal invoiceSupplier As Variant, _
    ByVal invoiceAmount As Variant, ByVal amountDiff As Variant, _
    ByVal dateGap As Variant)
    wsR.Cells(outRow, 1).Value = wsE.Cells(bankRow, 1).Value
    wsR.Cells(outRow, 2).Value = wsE.Cells(bankRow, 4).Value
    wsR.Cells(outRow, 3).Value = wsE.Cells(bankRow, 3).Value
    wsR.Cells(outRow, 4).Value = wsE.Cells(bankRow, 5).Value
    If invoiceRow > 0 Then
        wsR.Cells(outRow, 5).Value = invoiceNo
        wsR.Cells(outRow, 6).Value = invoiceSupplier
        wsR.Cells(outRow, 7).Value = invoiceAmount
        wsR.Cells(outRow, 8).Value = amountDiff
        wsR.Cells(outRow, 9).Value = dateGap
    End If
    wsR.Cells(outRow, 10).Value = resultText
    wsR.Cells(outRow, 11).Value = detail
End Sub
