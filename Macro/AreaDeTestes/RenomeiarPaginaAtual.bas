Sub RenomearPaginaAtual()
    Dim doc As Document
    Dim pagina As Page
    Dim NomeArquivo As String
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Renomeia a página atual para "DESIGN"
    pagina.Name = "DESIGN"
    
    ' Verifica se o nome do documento segue o padrão "Sem Título-" seguido de um número
    If InStr(1, doc.Name, "Sem título-") = 1 Then
        ' O nome do documento segue o padrão, então definimos o nome do documento como "SISBolt_"
        NomeArquivo = "SISBolt_"
    Else
        ' O nome do documento não segue o padrão, então mantemos o nome atual do documento
        NomeArquivo = Replace(doc.Name, ".cdr", "")
    End If
    
    ' Define o nome do documento
    doc.Name = NomeArquivo
    
    ' Exclui a camada "Camada 1" se existir
    ExcluirCamada1
End Sub

Sub RenomearDocumentoAtual(NomeArquivo As String)
    Dim doc As Document
    Dim pagina As Page
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument

    If NomeArquivo == "SISBolt_" Then
        ' Verifica se o nome do documento segue o padrão "Sem Título-" seguido de um número
        If InStr(1, doc.Name, "Sem título-") = 1 Then
            ' O nome do documento segue o padrão, então definimos o nome do documento como "SISBolt_"
            NomeArquivo = "SISBolt_"
        Else
            ' O nome do documento não segue o padrão, então mantemos o nome atual do documento
            NomeArquivo = Replace(doc.Name, ".cdr", "")
        End If
    Else
        ' O nome do documento não segue o padrão, então mantemos o nome atual do documento
        NomeArquivo = Replace(doc.Name, ".cdr", "")    
    End If
End Sub