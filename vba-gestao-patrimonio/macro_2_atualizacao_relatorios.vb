Sub SincronizarInventarioPorHistorico()
    Dim wsInv As Worksheet, wsHist As Worksheet
    Dim ultLinhaHist As Long, i As Long
    Dim linhaInv As Variant
    Dim numPatrimonio As String
    Dim totalAtualizados As Long, totalNovos As Long
    
    ' 1. Definindo as abas internas
    Set wsInv = ThisWorkbook.Sheets("INVENTARIO_ATUAL")
    Set wsHist = ThisWorkbook.Sheets("HISTORICO_TRANSFERENCIAS")
    
    ' Encontra a última linha baseada na Coluna C (Patrimônio) do Histórico
    ultLinhaHist = wsHist.Cells(wsHist.Rows.Count, 3).End(xlUp).Row 
    
    If ultLinhaHist < 2 Then
        MsgBox "Nenhum dado encontrado na tabela de Histórico para sincronizar.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    totalAtualizados = 0
    totalNovos = 0
    
    ' 2. Loop por todos os registros já existentes no Histórico
    For i = 2 To ultLinhaHist
        numPatrimonio = CStr(wsHist.Cells(i, 3).Value) ' Coluna 3: Patrimônio
        
        If Trim(numPatrimonio) <> "" Then
            
            ' ========================================================
            ' FASE 2: ATUALIZAR INVENTÁRIO_ATUAL <- HISTORICO_TRANSFERENCIAS
            ' ========================================================
            ' Procura o patrimônio na Coluna 1 do Inventário
            linhaInv = Application.Match(numPatrimonio, wsInv.Columns(1), 0)
            
            If IsError(linhaInv) Then
                ' ---> SE BEM NOVO (Não encontrado no Inventário) <---
                linhaInv = wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row + 1
                wsInv.Cells(linhaInv, 1).Value = numPatrimonio                 ' N° PATRIMÔNIO
                wsInv.Cells(linhaInv, 2).Value = wsHist.Cells(i, 4).Value      ' ITEM (Nome do item - grava apenas na criação)
                totalNovos = totalNovos + 1
                
                ' Pinta de verde as colunas de A até M (1 a 13) para novos itens
                wsInv.Range(wsInv.Cells(CLng(linhaInv), 1), wsInv.Cells(CLng(linhaInv), 13)).Interior.Color = RGB(226, 239, 218)
            Else
                ' ---> SE BEM EXISTENTE (Encontrado no Inventário) <---
                ' O campo ITEM (Col 2) NÃO é substituído para preservar o cadastro original.
                ' As colunas de Categorias, Marcas e Status fixos de lista são preservadas.
                totalAtualizados = totalAtualizados + 1
                
                ' Pinta de amarelo as colunas de A até M (1 a 13) para itens atualizados
                wsInv.Range(wsInv.Cells(CLng(linhaInv), 1), wsInv.Cells(CLng(linhaInv), 13)).Interior.Color = RGB(255, 242, 204)
            End If
            
            ' ---> ATUALIZAÇÃO DOS DADOS ATUAIS (Desde que o histórico possua a informação) <---
            If wsHist.Cells(i, 5).Value <> "" Then wsInv.Cells(CLng(linhaInv), 5).Value = wsHist.Cells(i, 5).Value   ' NÚMERO DE SÉRIE (Col 5)
            If wsHist.Cells(i, 10).Value <> "" Then wsInv.Cells(CLng(linhaInv), 6).Value = wsHist.Cells(i, 10).Value ' UNIDADE ATUAL <- Unidade de destino (Col 10)
            If wsHist.Cells(i, 11).Value <> "" Then wsInv.Cells(CLng(linhaInv), 7).Value = wsHist.Cells(i, 11).Value ' LOCALIZAÇÃO ATUAL <- Para Local (Col 11)
            If wsHist.Cells(i, 7).Value <> "" Then wsInv.Cells(CLng(linhaInv), 8).Value = wsHist.Cells(i, 7).Value   ' RESPONSÁVEL ATUAL <- Para Responsável (Col 7)
            
            ' REGISTROS DE DATAS E CHAMADOS
            wsInv.Cells(CLng(linhaInv), 10).Value = Date ' DATA ÚLTIMA ATUALIZAÇÃO (Carimbo de automação)
            
            If wsHist.Cells(i, 12).Value <> "" Then wsInv.Cells(CLng(linhaInv), 11).Value = wsHist.Cells(i, 12).Value ' ÚLTIMO CHAMADO <- Chamado (Col 12)
            If wsHist.Cells(i, 14).Value <> "" Then wsInv.Cells(CLng(linhaInv), 12).Value = wsHist.Cells(i, 14).Value ' ÚLTIMA TRANSFERÊNCIA <- Data conclusão (Col 14)
            If wsHist.Cells(i, 18).Value <> "" Then wsInv.Cells(CLng(linhaInv), 13).Value = wsHist.Cells(i, 18).Value ' OBSERVAÇÃO PADRONIZADA <- Status (Col 18)
            
        End If
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "Sincronização concluída com sucesso diretamente do Histórico!" & vbCrLf & _
           "Itens Atualizados: " & totalAtualizados & vbCrLf & _
           "Novos Itens Cadastrados: " & totalNovos, vbInformation, "Sincronização Concluída"
End Sub
