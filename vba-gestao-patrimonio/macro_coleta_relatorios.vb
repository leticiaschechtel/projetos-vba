Sub ImportarRelatorioPipefy()
    Dim wsInv As Worksheet, wsHist As Worksheet
    Dim wbPipefy As Workbook, wsPipefy As Worksheet
    Dim arquivoCSV As Variant
    Dim ultLinhaCSV As Long, i As Long
    Dim ultLinhaHist As Long
    Dim linhaInv As Variant
    Dim numPatrimonio As String
    
    ' 1. Definindo as abas do seu arquivo principal
    Set wsInv = ThisWorkbook.Sheets("INVENTARIO_ATUAL")
    Set wsHist = ThisWorkbook.Sheets("HISTORICO_TRANSFERENCIAS")
    
    ' 2. Abrir janela para selecionar o arquivo CSV do Pipefy
    arquivoCSV = Application.GetOpenFilename("Arquivos CSV (*.csv), *.csv", , "Selecione o Relatório do Pipefy")
    
    If arquivoCSV = False Then
        MsgBox "Importação cancelada.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 3. Abrir o arquivo CSV selecionado
    Set wbPipefy = Workbooks.Open(Filename:=arquivoCSV, Local:=True)
    Set wsPipefy = wbPipefy.Sheets(1)
    
    ultLinhaCSV = wsPipefy.Cells(wsPipefy.Rows.Count, 1).End(xlUp).Row
    
    ' 4. Loop pelos registros do Pipefy (Começa na linha 2 ignorando cabeçalho)
    For i = 2 To ultLinhaCSV
        
        ' ALINHAMENTO DE CHAVE PRIMÁRIA
        ' O N° Patrimônio está na Coluna 4 do Relatório Pipefy
        numPatrimonio = CStr(wsPipefy.Cells(i, 4).Value) 
        
        If Trim(numPatrimonio) <> "" Then
            
            ' ========================================================
            ' FASE 1: GRAVAR NA TABELA DE HISTÓRICO (HISTORICO_TRANSFERENCIAS <- RELATORIO_PIPEFY)
            ' ========================================================
            ultLinhaHist = wsHist.Cells(wsHist.Rows.Count, 1).End(xlUp).Row + 1
            
            wsHist.Cells(ultLinhaHist, 1).Value = Date ' Data atual da execução
            wsHist.Cells(ultLinhaHist, 2).Value = wsPipefy.Cells(i, 3).Value  ' TIPO DE MOVIMENTAÇÃO (Col 3)
            wsHist.Cells(ultLinhaHist, 3).Value = numPatrimonio               ' Patrimônio (Col 4)
            wsHist.Cells(ultLinhaHist, 4).Value = wsPipefy.Cells(i, 5).Value  ' Nome do item (Col 5)
            wsHist.Cells(ultLinhaHist, 5).Value = wsPipefy.Cells(i, 6).Value  ' NÚMERO DE SÉRIE (Col 6)
            wsHist.Cells(ultLinhaHist, 6).Value = wsPipefy.Cells(i, 7).Value  ' De Responsável (Col 7)
            wsHist.Cells(ultLinhaHist, 7).Value = wsPipefy.Cells(i, 8).Value  ' Para Responsável (Col 8)
            wsHist.Cells(ultLinhaHist, 8).Value = wsPipefy.Cells(i, 10).Value  ' Unidade atual (Col 10)
            wsHist.Cells(ultLinhaHist, 9).Value = wsPipefy.Cells(i, 9).Value ' De Local (Col 9)
            wsHist.Cells(ultLinhaHist, 10).Value = wsPipefy.Cells(i, 12).Value' Unidade de destino (Col 12)
            wsHist.Cells(ultLinhaHist, 11).Value = wsPipefy.Cells(i, 11).Value' Para Local (Col 11)
            wsHist.Cells(ultLinhaHist, 12).Value = wsPipefy.Cells(i, 13).Value' Chamado (Col 13)
            wsHist.Cells(ultLinhaHist, 13).Value = wsPipefy.Cells(i, 19).Value' Motivos de espera (Col 19)
            wsHist.Cells(ultLinhaHist, 14).Value = wsPipefy.Cells(i, 17).Value' Data conclusão (Col 17)
            wsHist.Cells(ultLinhaHist, 15).Value = wsPipefy.Cells(i, 18).Value' Status (Observação) (Col 18)
            wsHist.Cells(ultLinhaHist, 16).Value = wsPipefy.Cells(i, 20).Value' Responsável Transporte (Col 20)
            wsHist.Cells(ultLinhaHist, 17).Value = wsPipefy.Cells(i, 14).Value' Placa (Col 14)
            wsHist.Cells(ultLinhaHist, 18).Value = wsPipefy.Cells(i, 16).Value' Status (Col 18)
            wsHist.Cells(ultLinhaHist, 19).Value = wsPipefy.Cells(i, 15).Value' TIPO DE DOCUMENTO (Col 19)

            
            ' ========================================================
            ' FASE 2: ATUALIZAR INVENTÁRIO_ATUAL (INVENTARIO_ATUAL <- HISTORICO)
            ' ========================================================
            ' Procura o patrimônio na Coluna 1 do Inventário
            linhaInv = Application.Match(numPatrimonio, wsInv.Columns(1), 0)
            
            If IsError(linhaInv) Then
                ' ---> SE BEM NOVO (Não encontrou) <---
                linhaInv = wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row + 1
                wsInv.Cells(linhaInv, 1).Value = numPatrimonio                     ' N° PATRIMÔNIO
                wsInv.Cells(linhaInv, 2).Value = wsPipefy.Cells(i, 5).Value       ' ITEM (Grava apenas se for novo)
                
                wsInv.Range(wsInv.Cells(CLng(linhaInv), 1), wsInv.Cells(CLng(linhaInv), 13)).Interior.Color = RGB(226, 239, 218) ' Verde
            Else
                ' ---> SE BEM EXISTENTE (Encontrou) <---
                ' O campo ITEM (Col 2) NÃO é substituído. As categorias de lista são ignoradas pelo VBA.
                wsInv.Range(wsInv.Cells(CLng(linhaInv), 1), wsInv.Cells(CLng(linhaInv), 13)).Interior.Color = RGB(255, 242, 204) ' Amarelo
            End If
            
            ' ---> ATUALIZAÇÕES GERAIS PARA AMBOS (Desde que o Pipefy não esteja vazio) <---
            If wsPipefy.Cells(i, 5).Value <> "" Then wsInv.Cells(CLng(linhaInv), 5).Value = wsPipefy.Cells(i, 5).Value   ' NÚMERO DE SÉRIE
            If wsPipefy.Cells(i, 11).Value <> "" Then wsInv.Cells(CLng(linhaInv), 6).Value = wsPipefy.Cells(i, 11).Value ' UNIDADE ATUAL (Unidade de destino)
            If wsPipefy.Cells(i, 10).Value <> "" Then wsInv.Cells(CLng(linhaInv), 7).Value = wsPipefy.Cells(i, 10).Value ' LOCALIZAÇÃO ATUAL (Para Local)
            If wsPipefy.Cells(i, 7).Value <> "" Then wsInv.Cells(CLng(linhaInv), 8).Value = wsPipefy.Cells(i, 7).Value   ' RESPONSÁVEL ATUAL (Para Responsável)
            
            ' REGISTROS DE DATAS E STATUS
            ' Carimba a data do momento em que o código foi executado
            wsInv.Cells(CLng(linhaInv), 10).Value = Date ' DATA ÚLTIMA ATUALIZAÇÃO (Automático)
            
            If wsPipefy.Cells(i, 12).Value <> "" Then wsInv.Cells(CLng(linhaInv), 11).Value = wsPipefy.Cells(i, 12).Value ' ÚLTIMO CHAMADO
            If wsPipefy.Cells(i, 19).Value <> "" Then wsInv.Cells(CLng(linhaInv), 12).Value = wsPipefy.Cells(i, 19).Value ' ÚLTIMA TRANSFERÊNCIA
            If wsPipefy.Cells(i, 18).Value <> "" Then wsInv.Cells(CLng(linhaInv), 13).Value = wsPipefy.Cells(i, 18).Value ' OBSERVAÇÃO PADRONIZADA (Status)
            
        End If
    Next i 
    
    ' 5. Fechar o CSV do Pipefy sem salvar alterações
    wbPipefy.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    MsgBox "Dados importados e validados! Os itens já cadastrados foram preservados com sucesso.", vbInformation, "Automação Pipefy Concluída"
End Sub
