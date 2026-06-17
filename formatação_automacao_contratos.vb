# Formatação de contratos 
' Converte alta quantidade de arquivos em Word para contratos formatados em PDF com finalidade de Assinatura Digital. 
' Atualiza de maneira automática os nomes dos arquivos.
---

Sub SalvarUmPorUm()
    Dim docBase As Document
    Dim i As Integer
    Dim total As Integer
    Dim nomeArquivo As String
    Dim caminhoSalvar As String
    
    ' Definir modelo da pagina principal
    Set docBase = ActiveDocument
    
    ' Pasta p_ salvar os arquivos
    caminhoSalvar = " nome_pasta  "
    
    With docBase.MailMerge
        ' Total de pessoas da listagem
        .DataSource.ActiveRecord = wdLastRecord
        total = .DataSource.ActiveRecord
        .DataSource.ActiveRecord = wdFirstRecord
        
        ' Loop da listagem
        For i = 1 To total
            .DataSource.ActiveRecord = i
            nomeArquivo = .DataSource.DataFields("COLABORADOR").Value
            
            ' Otimizacao caracteres windows
            nomeArquivo = Replace(nomeArquivo, "/", "-")
            nomeArquivo = Replace(nomeArquivo, "\", "-")
            nomeArquivo = Replace(nomeArquivo, ":", "")
            nomeArquivo = Replace(nomeArquivo, "?", "")
            nomeArquivo = Replace(nomeArquivo, "*", "")
            nomeArquivo = Replace(nomeArquivo, """", "")
            nomeArquivo = Replace(nomeArquivo, "<", "")
            nomeArquivo = Replace(nomeArquivo, ">", "")
            nomeArquivo = Replace(nomeArquivo, "|", "")
            
            ' Evita erro de mesclagem do word
            .Destination = wdSendToNewDocument
            .SuppressBlankLines = True
            .DataSource.FirstRecord = i
            .DataSource.LastRecord = i
            
            ' Executa a criação do novo documento temporário
            .Execute Pause:=False
            
            ' O novo documento gerado vira o arquivo ativo. P_ salvar o arquivo em PDF 
            ActiveDocument.ExportAsFixedFormat OutputFileName:= _
                caminhoSalvar & nomeArquivo & ".pdf", _
                ExportFormat:=wdExportFormatPDF, _
                OpenAfterExport:=False, _
                OptimizeFor:=wdExportOptimizeForPrint, _
                Range:=wdExportAllDocument, _
                Item:=wdExportDocumentContent
            
            ' Fecha a página isolada sem salvar (o PDF já foi criado)
            ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
            
        Next i
    End With
    
    MsgBox "Sucesso!", vbInformation, "Concluído"
End Sub
