Option Explicit

' ===== CONFIGURAÇÕES =====
Private Const DEST_SHEET As String = "Matriz" ' Nome da aba de destino (planilha do setor)
Private Const DEST_HEADER_ROW As Long = 1
Private Const DEST_START_ROW As Long = 2

Private Const TMP_SHEET_NAME As String = "__TMP_PIPEFY"
Private Const NEW_ROW_FILL_COLOR As Long = 15773696 ' RGB(224, 236, 255) — azul claro; use RGB() se preferir

' ===== ENTRADA PRINCIPAL =====
Public Sub AtualizarComprasComPipefy()
    On Error GoTo fail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim csvPath As String
    csvPath = SelecionarArquivoCSV("Selecione o CSV exportado do Pipefy")
    If Len(csvPath) = 0 Then GoTo cleanup

    Dim wsTmp As Worksheet
    Set wsTmp = ImportarCSVParaAbaTemporaria(csvPath, TMP_SHEET_NAME)

    Dim srcH As Object: Set srcH = MapearCabecalhos(wsTmp, 1)

    Dim wsDst As Worksheet
    On Error Resume Next
    Set wsDst = ThisWorkbook.Worksheets(DEST_SHEET)
    On Error GoTo fail
    If wsDst Is Nothing Then
        MsgBox "A aba de destino '" & DEST_SHEET & "' não foi encontrada.", vbCritical
        GoTo cleanup
    End If

    Dim dstH As Object: Set dstH = MapearCabecalhos(wsDst, DEST_HEADER_ROW)

    ' Descobrir coluna da chave no destino (preferência: N° SOLICITAÇÃO)
    Dim dstKeyCol As Long
    dstKeyCol = DescobrirColunaPorAliases(wsDst, dstH, Array( _
        "N° SOLICITAÇÃO", "Nº SOLICITAÇÃO", "N° DA SOLICITAÇÃO", "Nº DA SOLICITAÇÃO", _
        "N° DA", "Nº DA", "N° SOLICITACAO", "Nº SOLICITACAO" _
    ))
    If dstKeyCol <= 0 Then
        MsgBox "Não encontrei a coluna da chave (ex.: 'N° SOLICITAÇÃO'). Renomeie o cabeçalho e tente novamente.", vbCritical
        GoTo cleanup
    End If

    ' Mapa de linhas do destino por chave (valor da coluna 'N° SOLICITAÇÃO')
    Dim mapDst As Object: Set mapDst = MapearLinhasPorChave(wsDst, dstKeyCol, DEST_START_ROW)

    ' Índices dos campos destino (somente os mapeados)
    Dim cSTATUS As Variant: cSTATUS = ColunaDestino(dstH, "STATUS")

    ' chave gravada no destino (N° SOLICITAÇÃO)
    Dim cNSOL As Variant
    cNSOL = ColunaDestinoPreferindo(dstH, Array("N° SOLICITAÇÃO", "Nº SOLICITAÇÃO", "N° DA SOLICITAÇÃO", "Nº DA SOLICITAÇÃO", "N° DA", "Nº DA"))

    Dim cSOLICITANTE As Variant: cSOLICITANTE = ColunaDestino(dstH, "SOLICITANTE")
    Dim cDT_SOL As Variant: cDT_SOL = ColunaDestino(dstH, "DATA DA SOLICITAÇÃO")
    Dim cOC As Variant: cOC = ColunaDestinoPreferindo(dstH, Array("ORDEM DE COMPRA", "ORDEM  DE COMPRA", "ORDEM\nDE COMPRA"))
    Dim cPREV As Variant: cPREV = ColunaDestinoPreferindo(dstH, Array("PREVISÃO DE ENTREGA", "PREVISÃO  DE ENTREGA", "PREVISÃO\nDE ENTREGA"))
    Dim cDT_REQ As Variant: cDT_REQ = ColunaDestinoPreferindo(dstH, Array("DATA DA REQUISIÇÃO", "DATA  DA REQUISIÇÃO", "DATA\nDA REQUISIÇÃO"))
    Dim cNF As Variant: cNF = ColunaDestinoPreferindo(dstH, Array("N° NF", "Nº NF"))

    ' Novos campos solicitados
    Dim cUD As Variant: cUD = ColunaDestinoPreferindo(dstH, Array("UNIDADE DEPARTAMENTAL", "UNIDADE  DEPARTAMENTAL"))
    Dim cCR As Variant: cCR = ColunaDestinoPreferindo(dstH, Array("CENTRO DE RESPONSABILIDADE", "CENTRO  DE RESPONSABILIDADE"))

    ' Última coluna com cabeçalho (para pintar a linha toda)
    Dim lastHeaderCol As Long: lastHeaderCol = wsDst.Cells(DEST_HEADER_ROW, wsDst.Columns.Count).End(xlToLeft).Column

    Dim lastSrcRow As Long: lastSrcRow = wsTmp.Cells(wsTmp.Rows.Count, 1).End(xlUp).Row
    Dim r As Long, atualizadas As Long, criadas As Long

    For r = 2 To lastSrcRow
        If LinhaVazia(wsTmp, r) Then GoTo proxima

        ' Ler campos do Pipefy
        Dim fase As String: fase = GetCell(wsTmp, r, srcH, "Fase atual")

        ' CHAVE: Código (preferido) -> fallback Nº da Solicitação
        Dim codigo As String: codigo = NormalizarChave(GetCell(wsTmp, r, srcH, "Código"))
        Dim nroSolicPF As String
        If Len(codigo) = 0 Then
            nroSolicPF = NormalizarChave(GetCell(wsTmp, r, srcH, "Nº da Solicitação"))
            If Len(nroSolicPF) = 0 Then nroSolicPF = NormalizarChave(GetCell(wsTmp, r, srcH, "N° da Solicitação"))
        End If
        Dim chave As String: chave = IIf(Len(codigo) > 0, codigo, nroSolicPF)
        If Len(chave) = 0 Then GoTo proxima ' sem chave, ignora

        Dim criador As String: criador = GetCell(wsTmp, r, srcH, "Criador")

        Dim dtSol As Variant: dtSol = ParseDateFlex(GetCell(wsTmp, r, srcH, "Data da solicitação"))
        Dim oc As String: oc = GetCell(wsTmp, r, srcH, "Ordem de compra")
        Dim prev As Variant: prev = ParseDateFlex(GetCell(wsTmp, r, srcH, "Previsão de entrega"))

        ' DATA DA REQUISIÇÃO no Pipefy — várias possibilidades
        Dim dtReq As Variant
        dtReq = ParseDateFlex(GetCell(wsTmp, r, srcH, "DATA DA REQUISIÇÃO"))
        If IsEmptyOrBlank(dtReq) Then dtReq = ParseDateFlex(GetCell(wsTmp, r, srcH, "Data")) ' alguns relatórios usam "Data"
        If IsEmptyOrBlank(dtReq) Then dtReq = ParseDateFlex(GetCell(wsTmp, r, srcH, "Data entrega"))

        Dim nf As String: nf = GetCell(wsTmp, r, srcH, "Número NF")

        ' Unidade Departamental (UD) e CR
        Dim ud As String
        ud = GetCell(wsTmp, r, srcH, "Unidade Departamental (UD)")
        If Len(Trim$(ud)) = 0 Then ud = GetCell(wsTmp, r, srcH, "Unidade Departamental") ' variação
        If Len(Trim$(ud)) = 0 Then ud = GetCell(wsTmp, r, srcH, "Área Demandante")      ' fallback

        Dim cr As String
        cr = GetCell(wsTmp, r, srcH, "Centro de Responsabilidade (CR)")
        If Len(Trim$(cr)) = 0 Then cr = GetCell(wsTmp, r, srcH, "Centro de Responsabilidade")

        Dim dstRow As Long
        If mapDst.Exists(chave) Then
            ' Atualiza linha existente (SOBRESCREVE campos mapeados, exceto SOLICITANTE se for pipebot)
            dstRow = CLng(mapDst(chave))
            SafeSet wsDst, cSTATUS, dstRow, fase

            ' grava a chave na coluna "N° SOLICITAÇÃO"
            SafeSet wsDst, cNSOL, dstRow, chave

            ' SOLICITANTE: não sobrescrever se vier como pipebot
            If Not IsPipebot(criador) Then
                SafeSet wsDst, cSOLICITANTE, dstRow, criador
            End If

            SafeSet wsDst, cDT_SOL, dstRow, dtSol
            SafeSet wsDst, cOC, dstRow, oc
            SafeSet wsDst, cPREV, dstRow, prev
            SafeSet wsDst, cDT_REQ, dstRow, dtReq
            SafeSet wsDst, cNF, dstRow, nf

            ' UD e CR
            SafeSet wsDst, cUD, dstRow, ud
            SafeSet wsDst, cCR, dstRow, cr

            atualizadas = atualizadas + 1
        Else
            ' Cria nova linha e pinta azul
            dstRow = wsDst.Cells(wsDst.Rows.Count, 1).End(xlUp).Row + 1
            If dstRow < DEST_START_ROW Then dstRow = DEST_START_ROW

            SafeSet wsDst, cSTATUS, dstRow, fase
            SafeSet wsDst, cNSOL, dstRow, chave

            ' Em linha nova, pode escrever SOLICITANTE mesmo se pipebot (regra de não sobrescrever vale só para existentes)
            SafeSet wsDst, cSOLICITANTE, dstRow, criador

            SafeSet wsDst, cDT_SOL, dstRow, dtSol
            SafeSet wsDst, cOC, dstRow, oc
            SafeSet wsDst, cPREV, dstRow, prev
            SafeSet wsDst, cDT_REQ, dstRow, dtReq
            SafeSet wsDst, cNF, dstRow, nf

            ' UD e CR
            SafeSet wsDst, cUD, dstRow, ud
            SafeSet wsDst, cCR, dstRow, cr

            ' Pintar a linha toda de azul
            On Error Resume Next
            wsDst.Range(wsDst.Cells(dstRow, 1), wsDst.Cells(dstRow, lastHeaderCol)).Interior.Color = NEW_ROW_FILL_COLOR
            On Error GoTo 0

            ' Indexa para evitar duplicar
            mapDst.Add chave, dstRow
            criadas = criadas + 1
        End If

proxima:
    Next r

    ' Limpa aba temporária
    Application.DisplayAlerts = False
    wsTmp.Delete
    Application.DisplayAlerts = True

    MsgBox "Concluído." & vbCrLf & _
           "Atualizadas: " & atualizadas & vbCrLf & _
           "Criadas: " & criadas, vbInformation
    GoTo cleanup

fail:
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical

cleanup:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub

' ===== APOIO =====

Private Function SelecionarArquivoCSV(prompt As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = prompt
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "CSV", "*.csv"
        If .Show = -1 Then SelecionarArquivoCSV = .SelectedItems(1)
    End With
End Function

Private Function ImportarCSVParaAbaTemporaria(csvPath As String, tmpName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(tmpName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = tmpName
    Else
        ws.Cells.Clear
    End If

    Dim qt As QueryTable
    Set qt = ws.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=ws.Range("A1"))
    With qt
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 65001
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileSemicolonDelimiter = True
        .TextFileDecimalSeparator = ","
        .TextFileThousandsSeparator = "."
        .TextFileTrailingMinusNumbers = True
        .AdjustColumnWidth = True
        .Refresh BackgroundQuery:=False
    End With

    Set ImportarCSVParaAbaTemporaria = ws
End Function

Private Function MapearCabecalhos(ws As Worksheet, headerRow As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' TextCompare

    Dim lastCol As Long: lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, h As String, norm As String
    For c = 1 To lastCol
        h = LimparQuebras(CStr(ws.Cells(headerRow, c).Value))
        If Len(h) > 0 Then
            norm = NormalizarHeader(h)
            If Not d.Exists(h) Then d.Add h, c
            If Not d.Exists(norm) Then d.Add norm, c
        End If
    Next c
    Set MapearCabecalhos = d
End Function

Private Function ColunaDestino(dstHeaders As Object, destHeader As String) As Variant
    If dstHeaders Is Nothing Then ColunaDestino = CVErr(xlErrNA): Exit Function
    If dstHeaders.Exists(destHeader) Then
        ColunaDestino = dstHeaders(destHeader): Exit Function
    End If
    Dim norm As String: norm = NormalizarHeader(LimparQuebras(destHeader))
    If dstHeaders.Exists(norm) Then
        ColunaDestino = dstHeaders(norm)
    Else
        ColunaDestino = CVErr(xlErrNA)
    End If
End Function

Private Function ColunaDestinoPreferindo(dstHeaders As Object, aliases As Variant) As Variant
    Dim i As Long, c As Variant
    For i = LBound(aliases) To UBound(aliases)
        c = ColunaDestino(dstHeaders, CStr(aliases(i)))
        If Not IsError(c) Then ColunaDestinoPreferindo = c: Exit Function
    Next i
    ColunaDestinoPreferindo = CVErr(xlErrNA)
End Function

Private Function DescobrirColunaPorAliases(ws As Worksheet, dstHeaders As Object, aliases As Variant) As Long
    Dim i As Long, c As Variant
    For i = LBound(aliases) To UBound(aliases)
        c = ColunaDestino(dstHeaders, CStr(aliases(i)))
        If Not IsError(c) Then DescobrirColunaPorAliases = CLng(c): Exit Function
    Next i
    DescobrirColunaPorAliases = -1
End Function

Private Function MapearLinhasPorChave(ws As Worksheet, keyCol As Long, startRow As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    Dim r As Long, k As String
    For r = startRow To lastRow
        k = NormalizarChave(CStr(ws.Cells(r, keyCol).Value))
        If Len(k) > 0 Then If Not d.Exists(k) Then d.Add k, r
    Next r
    Set MapearLinhasPorChave = d
End Function

Private Function GetCell(ws As Worksheet, r As Long, headers As Object, headerName As String) As String
    Dim v As Variant
    If headers.Exists(headerName) Then
        v = ws.Cells(r, headers(headerName)).Value
    Else
        Dim norm As String: norm = NormalizarHeader(LimparQuebras(headerName))
        If headers.Exists(norm) Then
            v = ws.Cells(r, headers(norm)).Value
        Else
            v = ""
        End If
    End If
    GetCell = CStr(v)
End Function

Private Function LinhaVazia(ws As Worksheet, r As Long) As Boolean
    LinhaVazia = (Application.WorksheetFunction.CountA(ws.Rows(r)) = 0)
End Function

Private Function NormalizarChave(ByVal s As String) As String
    s = Replace$(s, Chr$(160), " ")
    NormalizarChave = Trim$(s)
End Function

Private Sub SafeSet(ws As Worksheet, c As Variant, r As Long, v As Variant)
    On Error Resume Next
    If Not IsError(c) And c > 0 Then ws.Cells(r, c).Value = v
    On Error GoTo 0
End Sub

Private Function IsEmptyOrBlank(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then IsEmptyOrBlank = True: Exit Function
    If VarType(v) = vbString Then
        IsEmptyOrBlank = (Trim$(CStr(v)) = "")
    Else
        IsEmptyOrBlank = False
    End If
End Function

Private Function IsPipebot(ByVal s As String) As Boolean
    Dim t As String: t = LCase$(Trim$(s))
    IsPipebot = (InStr(t, "pipebot") > 0)
End Function

' ---- Normalização de cabeçalhos e datas ----

Private Function LimparQuebras(ByVal s As String) As String
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    LimparQuebras = Trim$(s)
End Function

Private Function NormalizarHeader(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = " " & s & " "
    s = Replace$(s, " DE ", " ")
    s = Replace$(s, " DA ", " ")
    s = Replace$(s, " DO ", " ")
    s = Replace$(s, " DAS ", " ")
    s = Replace$(s, " DOS ", " ")
    s = Trim$(s)
    s = RemoverAcentos(s)
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then out = out & ch
    Next i
    NormalizarHeader = out
End Function

Private Function RemoverAcentos(ByVal s As String) As String
    Dim ac As String, no As String, i As Long
    ac = "ÁÀÂÃÄáàâãäÉÈÊËéèêëÍÌÎÏíìîïÓÒÔÕÖóòôõöÚÙÛÜúùûüÇçÑñ"
    no = "AAAAAaaaaaEEEEeeeeIIIIiiiiOOOOOoooooUUUUuuuuCcNn"
    For i = 1 To Len(ac)
        s = Replace$(s, Mid$(ac, i, 1), Mid$(no, i, 1))
    Next i
    s = Replace$(s, "º", "")
    s = Replace$(s, "°", "")
    RemoverAcentos = s
End Function

Private Function ParseDateFlex(ByVal s As String) As Variant
    s = Trim$(s)
    If Len(s) = 0 Then ParseDateFlex = "": Exit Function
    On Error GoTo falha
    ' Tenta YYYY-MM-DD
    If InStr(s, "-") > 0 And Len(s) >= 10 Then
        Dim y As Long: y = Val(Left$(s, 4))
        If y >= 1900 And y <= 9999 Then
            ParseDateFlex = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
            Exit Function
        End If
    End If
    ' Tenta DD/MM/YYYY
    If InStr(s, "/") > 0 Then
        Dim p() As String: p = Split(s, "/")
        If UBound(p) = 2 Then
            ParseDateFlex = DateSerial(CInt(p(2)), CInt(p(1)), CInt(p(0)))
            Exit Function
        End If
    End If
falha:
    ParseDateFlex = s ' deixa como texto se não conseguir converter
End Function
