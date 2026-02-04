Sub RenomeiaTipoDePeca()
    Dim doc As Document
    Dim pagina As page
    Dim camada As layer
    Dim nomeOriginal As String
    Dim partes() As String
    Dim novoNome As String
    Dim tipoAtual As String
    Dim novoTipo As String
    Dim tiposValidos As String
    Dim tiposArray() As String
    Dim i As Integer
    Dim opcoesPrimeira As String
    Dim opcoesSegunda As String
    Dim alterarTodos As Boolean
    
    ' Lista de tipos válidos
    tiposValidos = "CAM1,CAM2,REG1,COL1,CAL1,SHO1"
    tiposArray = Split(tiposValidos, ",")
    
    ' Opções para escolha do tipo atual (inclui "Tudo")
    opcoesPrimeira = "1 - CAM1 (Camisa Manga Curta)" & vbNewLine & _
                     "2 - CAM2 (Camisa Manga Longa)" & vbNewLine & _
                     "3 - REG1 (Regata)" & vbNewLine & _
                     "4 - COL1 (Colete)" & vbNewLine & _
                     "5 - CAL1 (Calça)" & vbNewLine & _
                     "6 - SHO1 (Short)" & vbNewLine & _
                     "7 - Tudo (Todas as peças)"
    
    ' Opções para escolha do novo tipo (sem "Tudo")
    opcoesSegunda = "1 - CAM1 (Camisa Manga Curta)" & vbNewLine & _
                    "2 - CAM2 (Camisa Manga Longa)" & vbNewLine & _
                    "3 - REG1 (Regata)" & vbNewLine & _
                    "4 - COL1 (Colete)" & vbNewLine & _
                    "5 - CAL1 (Calça)" & vbNewLine & _
                    "6 - SHO1 (Short)"
    
    ' Seleção do tipo a modificar (permite "Tudo")
    tipoAtual = EscolherTipo("Selecione o tipo que deseja modificar:", tiposValidos, opcoesPrimeira, True)
    If tipoAtual = "" Then Exit Sub ' Sai se o usuário cancelar
    
    ' Verifica se o usuário escolheu "Tudo"
    alterarTodos = (tipoAtual = "Tudo")
    
    ' Seleção do novo tipo (NÃO permite "Tudo")
    novoTipo = EscolherTipo("Selecione o novo tipo para substituir " & tipoAtual & ":", tiposValidos, opcoesSegunda, False)
    If novoTipo = "" Then Exit Sub ' Sai se o usuário cancelar

    ' Referência ao documento ativo
    Set doc = ActiveDocument

    ' Percorre todas as páginas do documento
    For i = 1 To doc.Pages.Count
        Set pagina = doc.Pages(i)

        ' Percorre todas as camadas da página atual
        For Each camada In pagina.Layers
            nomeOriginal = camada.Name ' Obtém o nome original da camada
            
            ' Divide o nome da camada em partes separadas pelo ponto
            partes = Split(nomeOriginal, ".")
            
            ' Verifica se devemos alterar todas ou apenas um tipo específico
            If UBound(partes) >= 1 Then
                If alterarTodos Then
                    ' Altera qualquer tipo válido
                    If ExisteNaLista(partes(0), tiposArray) Then
                        partes(0) = novoTipo
                        camada.Name = Join(partes, ".") ' Renomeia a camada
                    End If
                Else
                    ' Altera apenas o tipo selecionado
                    If partes(0) = tipoAtual Then
                        partes(0) = novoTipo
                        camada.Name = Join(partes, ".") ' Renomeia a camada
                    End If
                End If
            End If
        Next camada
    Next i

    ' Limpeza de variáveis
    Set doc = Nothing
    Set pagina = Nothing
End Sub

' Função para exibir a lista de opções e capturar a escolha
Function EscolherTipo(mensagem As String, tiposValidos As String, opcoes As String, permiteTudo As Boolean) As String
    Dim escolha As String
    Dim tipoArray() As String
    Dim index As Integer
    
    ' Exibe opções
    escolha = InputBox(mensagem & vbNewLine & vbNewLine & opcoes, "Escolha um tipo")

    ' Define o limite superior (6 se não permitir "Tudo", 7 se permitir)
    Dim limiteSuperior As Integer
    If permiteTudo Then
        limiteSuperior = 7 ' Permite "Tudo"
    Else
        limiteSuperior = 6 ' Somente os tipos válidos
    End If
    
    ' Verifica se a escolha é válida
    If Not IsNumeric(escolha) Or Val(escolha) < 1 Or Val(escolha) > limiteSuperior Then
        MsgBox "Opção inválida ou cancelada!", vbExclamation
        EscolherTipo = ""
        Exit Function
    End If

    ' Converte número para tipo correspondente
    tipoArray = Split(tiposValidos, ",")
    
    ' Se for opção 7 e for permitido, retorna "Tudo"
    If Val(escolha) = 7 And permiteTudo Then
        EscolherTipo = "Tudo"
    Else
        index = Val(escolha) - 1
        EscolherTipo = tipoArray(index)
    End If
End Function

' Função para verificar se um valor está na lista de tipos válidos
Function ExisteNaLista(valor As String, lista() As String) As Boolean
    Dim i As Integer
    For i = LBound(lista) To UBound(lista)
        If lista(i) = valor Then
            ExisteNaLista = True
            Exit Function
        End If
    Next i
    ExisteNaLista = False
End Function