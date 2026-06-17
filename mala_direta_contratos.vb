
Option Explicit

Sub GerarRequisicoesPDF()
    Dim docMaster As Document
    Dim docFinal As Document
    Dim nRegistros As Integer
    Dim i As Integer
    Dim nomeArquivo As String
    Dim caminhoPasta As String
    Dim campoNome As String

    ' Define o documento que contém os campos de mesclagem
    Set docMaster = ActiveDocument
    
    ' Define o nome da coluna no Excel que será o nome do arquivo PDF
    ' Deve ser igual ao cabeçalho da planilha
    campoNome = "OBJETO"

    ' Tenta capturar a pasta onde o Word está salvo, se a pasta for externa tende a dar erro. 
    On Error Resume Next
    caminhoPasta = docMaster.Path & "\"
    If caminhoPasta = "\" Then
        MsgBox "Por favor, salve seu documento Word em uma pasta no computador antes de continuar.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' 1. Configura o documento como Mala Direta (Tipo Carta)
    docMaster.MailMerge.MainDocumentType = wdFormLetters

    ' 2. Verifica se há uma planilha vinculada
    If docMaster.MailMerge.DataSource.Name = "" Then
        MsgBox "Nenhuma planilha Excel vinculada. Por favor, conecte os dados na aba 'Correspondências' primeiro.", vbExclamation
        Exit Sub
    End If

    ' 3. Identifica a quantidade de registros na planilha
    docMaster.MailMerge.DataSource.ActiveRecord = wdLastRecord
    nRegistros = docMaster.MailMerge.DataSource.ActiveRecord
    docMaster.MailMerge.DataSource.ActiveRecord = wdFirstRecord

    ' --- INÍCIO DO PROCESSO POR LINHA ---
    For i = 1 To nRegistros
        docMaster.MailMerge.DataSource.ActiveRecord = i
        
        ' Pega o nome do arquivo da coluna selecionada
        nomeArquivo = docMaster.MailMerge.DataSource.DataFields(campoNome).Value
        nomeArquivo = LimparNomeArquivo(nomeArquivo)
        
        ' Se o nome estiver vazio, define um padrão para não dar erro
        If nomeArquivo = "" Then nomeArquivo = "Item_" & i

        ' 4. Executa a mesclagem APENAS da linha atual para um NOVO documento
        ' Isso garante que o PDF contenha o texto real e não os campos « »
        With docMaster.MailMerge
            .Destination = wdSendToNewDocument
            .SuppressBlankLines = True
            .DataSource.FirstRecord = i
            .DataSource.LastRecord = i
            .Execute Pause:=False
        End With

        ' O comando acima abre um novo documento (chamado Cartas1, Cartas2, etc.)
        ' Agora vamos salvar esse documento novo que está ativo
        Set docFinal = ActiveDocument
        
        ' 5. Exporta o documento mesclado como PDF
        docFinal.ExportAsFixedFormat _
            OutputFileName:=caminhoPasta & nomeArquivo & ".pdf", _
            ExportFormat:=wdExportFormatPDF, _
            OpenAfterExport:=False, _
            OptimizeFor:=wdExportOptimizeForPrint, _
            Range:=wdExportAllDocument

        ' 6. Fecha o documento mesclado sem salvar (para ficar apenas o PDF)
        docFinal.Close SaveChanges:=False
        
        docMaster.Activate
    Next i
    ' --- FIM ---

    MsgBox "Sucesso! " & nRegistros & " Termos de Referência foram gerados em PDF na pasta: " & vbCrLf & caminhoPasta, vbInformation, "Automação Concluída"
End Sub

' Função auxiliar para remover caracteres proibidos pelo Windows em nomes de arquivos
Function LimparNomeArquivo(ByVal texto As String) As String
    Dim i As Integer
    Dim caracteresProibidos As Variant
    caracteresProibidos = Array("/", "\", ":", "*", "?", "<", ">", "|", Chr(34))
    
    For i = LBound(caracteresProibidos) To UBound(caracteresProibidos)
        texto = Replace(texto, caracteresProibidos(i), "-")
    Next i
    
    ' Remove espaços extras no início ou fim
    LimparNomeArquivo = Trim(texto)
End Function
