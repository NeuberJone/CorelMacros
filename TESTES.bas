Attribute VB_Name = "TESTES"
Sub adicionaID()
    Dim doc As Document
    Dim pagina As page
    Dim grupo As shape
    Dim subShape As shape
    Dim caminhoArquivo As String
    Dim nomePasta As String, nomeArquivo As String, nomeLista As String
    Dim textoFinal As String
    Dim i As Integer
    Dim baseNome As String

    ' Referência ao documento ativo
    Set doc = ActiveDocument

    ' Verifica se o documento possui pelo menos 2 páginas
    If doc.Pages.Count < 2 Then
        MsgBox "O documento precisa ter pelo menos 2 páginas.", vbExclamation
        Exit Sub
    End If

    ' Define a página 2
    Set pagina = doc.Pages(2)

    ' Pega o caminho do arquivo atual
    caminhoArquivo = doc.FullFileName
    If caminhoArquivo = "" Then
        MsgBox "Salve o arquivo antes de executar a macro.", vbExclamation
        Exit Sub
    End If

    ' Extrai o nome da pasta (pedido)
    nomePasta = Mid(caminhoArquivo, InStrRev(caminhoArquivo, "\", InStrRev(caminhoArquivo, "\") - 1) + 1, _
                    InStrRev(caminhoArquivo, "\") - InStrRev(caminhoArquivo, "\", InStrRev(caminhoArquivo, "\") - 1) - 1)
    
    ' Extrai o nome do arquivo
    nomeArquivo = Mid(caminhoArquivo, InStrRev(caminhoArquivo, "\") + 1)

    ' Remove a extensão .cdr
    baseNome = Mid(nomeArquivo, 1, Len(nomeArquivo) - 4)

    ' Pega só o que vem depois de "SISBolt_"
    If Left(baseNome, 8) = "SISBolt_" Then
        nomeLista = Mid(baseNome, 9)
    Else
        nomeLista = baseNome
    End If

    ' Monta o texto final
    textoFinal = nomePasta & " - " & nomeLista

    ' Percorre todos os objetos da página 2
    For i = pagina.shapes.Count To 1 Step -1
        Set grupo = pagina.shapes(i)

        If grupo.Type = cdrGroupShape Then
            For Each subShape In grupo.shapes
                If subShape.Type = cdrTextShape And UCase(subShape.Name) = "ID" Then
                    subShape.Text.Story = textoFinal
                End If
            Next subShape
        End If
    Next i

    MsgBox "Identificação adicionada às peças", vbInformation
End Sub


