Attribute VB_Name = "Correcoes"
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


Sub RemoveCodigosNoNu()
    Dim doc As Document
    Dim pagina As page
    Dim s As shape
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Função para renomear objetos dentro de grupos e subgrupos
    Call RenameObjects(pagina.shapes)

End Sub

Sub RenameObjects(ByVal shapes As shapes)
    Dim s As shape
    Dim innerShapes As shapes

    ' Percorrer todos os objetos
    For Each s In shapes
        ' Se o objeto for um grupo, chamar a função para o grupo
        If s.Type = cdrGroupShape Then
            ' Se o objeto for um grupo, chamar a função para percorrer os objetos do grupo
            Set innerShapes = s.shapes
            Call RenameObjects(innerShapes)  ' Chama a função para renomear os objetos dentro do grupo
        Else
            ' Verificar se o nome do objeto é "NO" e renomeá-lo para "TextoNome"
            If s.Name = "NO" Then
                s.Name = "TextoNome"
            End If
            
            ' Verificar se o nome do objeto é "NU" e renomeá-lo para "TextoNumero"
            If s.Name = "NU" Then
                s.Name = "TextoNumero"
            End If
        End If
    Next s
End Sub



Sub AjustarLarguraObjetos()
    Dim sr As shapeRange
    Set sr = ActiveSelectionRange
    
    ' Verifica se há uma seleção ativa
    If sr.Count > 0 Then
        Dim sh As shape
        For Each sh In sr
            ' Verifica se o objeto possui o nome "NO" e largura superior a 850.394pt
            If sh.Name = "NO" And sh.SizeWidth > 850.394 Then
                ' Define a largura para 850.394pt
                sh.SizeWidth = 850.394
            End If
            ' Verifica se o objeto é um grupo ou PowerClip
            If sh.Type = cdrGroupShape Or sh.Type = cdrPowerClipShape Then
                ' Chama a função para verificar e ajustar a largura dos objetos dentro do grupo ou PowerClip
                AjustarLarguraObjetosRecursivo sh
            End If
        Next sh
    Else
        MsgBox "Nenhum objeto selecionado.", vbExclamation
    End If
End Sub

Sub AjustarLarguraObjetosRecursivo(grp As shape)
    Dim sh As shape
    For Each sh In grp.shapes
        ' Verifica se o objeto possui o nome "NO" e largura superior a 850.394pt
        If sh.Name = "NO" And sh.SizeWidth > 850.394 Then
            ' Define a largura para 850.394pt
            sh.SizeWidth = 850.394
        End If
        ' Verifica se o objeto é um grupo ou PowerClip
        If sh.Type = cdrGroupShape Or sh.Type = cdrPowerClipShape Then
            ' Chama recursivamente a função para verificar e ajustar a largura dos objetos dentro do grupo ou PowerClip
            AjustarLarguraObjetosRecursivo sh
        End If
    Next sh
End Sub




Sub RenomearObjetosNoMold()

    'Sub para renomear objetos nos molds
    Dim doc As Document
    Dim pagina As page
    Dim camada As layer
    Dim grupo As shape
    Dim subItem As shape
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Percorre todas as camadas na página
    For Each camada In pagina.Layers
        ' Percorre todos os grupos na camada atual
        For Each grupo In camada.shapes.FindShapes(, cdrGroupShape)
            ' Verifica se o grupo contém subitens
            If grupo.shapes.Count > 0 Then
                ' Percorre todos os subitens no grupo
                For Each subItem In grupo.shapes
                    ' Verifica o nome do subitem e a camada em que está
                    If Not NomeEstaNaLista(subItem.Name) Then
                        ' Renomeia o subitem com o nome da camada após o segundo ponto
                        subItem.Name = ObterNomeCamadaDepoisDoSegundoPonto(camada.Name)
                    End If
                Next subItem
            End If
        Next grupo
    Next camada
End Sub

Sub RenomearCamadas()
    Dim doc As Document
    Dim pagina As page
    Dim camada As layer
    Dim partes As Variant
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Percorre todas as camadas na página
    For Each camada In pagina.Layers
        ' Divide o nome da camada em partes usando o ponto como delimitador
        partes = Split(camada.Name, ".")
        
        ' Verifica se a camada possui pelo menos três partes (indicativos)
        If UBound(partes) >= 2 Then
            ' Renomeia a camada substituindo o último indicativo pelo segundo
            camada.Name = partes(0) & "." & partes(1) & "." & partes(2) & "." & partes(1)
        End If
    Next camada
End Sub

Sub RenomearObjetosNoGrid()
    Dim doc As Document
    Dim pagina As page
    Dim camada As layer
    Dim objeto As shape
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Percorre todas as camadas na página
    For Each camada In pagina.Layers
        ' Percorre todos os objetos na camada atual
        For Each objeto In camada.shapes
            ' Obtém o nome da camada após o primeiro ponto
            Dim nomeCamada As String
            nomeCamada = ObterNomeCamadaDepoisDoPrimeiroPonto(camada.Name)
            
            ' Renomeia o objeto com o nome da camada após o primeiro ponto
            objeto.Name = nomeCamada
        Next objeto
    Next camada
End Sub

Sub RedimensionarEMoverObjetos()
    Dim doc As Document
    Dim pagina As page
    Dim objeto As shape
    Dim larguraObjetos As Double
    Dim alturaObjetos As Double
    Dim larguraPagina As Double
    Dim alturaPagina As Double
    Dim novoX As Double
    Dim novoY As Double
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Obter largura e altura da página
    larguraPagina = pagina.SizeWidth
    alturaPagina = pagina.SizeHeight
    
    ' Seleciona todas as formas na página
    pagina.shapes.All.CreateSelection
    
    ' Inicializa as variáveis de dimensão
    larguraObjetos = pagina.shapes.All.SizeWidth
    alturaObjetos = pagina.shapes.All.SizeHeight
    
    ' Redimensionar a página
    pagina.SizeWidth = larguraObjetos + 30
    pagina.SizeHeight = alturaObjetos + 30
End Sub

Sub MoverObjetosParaCamada()
    Dim doc As Document
    Dim pagina As page
    Dim objeto As shape
    Dim camada As layer
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Percorre todos os objetos na página
    For Each objeto In pagina.shapes
        ' Verifica se há uma camada com o mesmo nome do objeto
        On Error Resume Next
        Set camada = pagina.Layers(objeto.Name)
        On Error GoTo 0
        
        ' Move o objeto para a camada correspondente
        If Not camada Is Nothing Then
            objeto.MoveToLayer camada
        End If
    Next objeto
End Sub

Sub AgrupaECentralizaObjetos()
    ActivePage.shapes.All.CreateSelection
    Dim s1 As shape
    Set s1 = ActiveSelection.Group
    s1.AlignAndDistribute 3, 3, 2, 0, False, 2
End Sub

Sub DesagrupaObjetos()
    Dim OrigSelection As shapeRange
    Set OrigSelection = ActiveSelectionRange
    Dim grp1 As shapeRange
    Set grp1 = OrigSelection.UngroupEx
End Sub

Sub CentralizaObjetosNaPagina()
    Dim OrigSelection As shapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection.AlignAndDistribute 3, 3, 2, 0, False, 2
End Sub

Sub RenomearDocumentoAtual(nomeArquivo As String)
    Dim doc As Document
    Dim pagina As page
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument

    If nomeArquivo = "SISBolt_" Then
        ' Verifica se o nome do documento segue o padrão "Sem Título-" seguido de um número
        If InStr(1, doc.Name, "Sem título-") = 1 Then
            ' O nome do documento segue o padrão, então definimos o nome do documento como "SISBolt_"
            doc.Name = "SISBolt_"
        Else
            ' O nome do documento não segue o padrão, então mantemos o nome atual do documento
            doc.Name = Replace(doc.Name, ".cdr", "")
        End If
    Else
        ' O nome do documento não segue o padrão, então definimos o nome do documento como o nome fornecido
        doc.Name = nomeArquivo
    End If
End Sub

Sub RenomeiaPaginaAtual(NomePagina As String)
    Dim doc As Document
    Dim pagina As page
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Renomeia a página atual com o nome fornecido
    pagina.Name = NomePagina
End Sub

Function ObterNomeCamadaDepoisDoPrimeiroPonto(nomeCamada As String) As String
    Dim partes() As String
    partes = Split(nomeCamada, ".")
    If UBound(partes) >= 1 Then
        ObterNomeCamadaDepoisDoPrimeiroPonto = partes(1)
    Else
        ObterNomeCamadaDepoisDoPrimeiroPonto = ""
    End If
End Function

Function ObterNomeCamadaDepoisDoSegundoPonto(nomeCamada As String) As String
    Dim partes() As String
    partes = Split(nomeCamada, ".")
    If UBound(partes) >= 2 Then
        ObterNomeCamadaDepoisDoSegundoPonto = partes(3)
    Else
        ObterNomeCamadaDepoisDoSegundoPonto = ""
    End If
End Function

Function NomeEstaNaLista(nome As String) As Boolean
    Dim nomesPermitidos As Variant
    Dim nomeMaiusculo As String
    
    ' Converte o nome para maiúsculas para fazer a comparação sem diferenciar maiúsculas de minúsculas
    nomeMaiusculo = UCase(nome)
    
    ' Lista de nomes permitidos
    nomesPermitidos = Array("PP", "P", "M", "G", "GG", "XGG", "XXGG", "PLPP", "BLP", "BLM", "BLG", "BLGG", _
                            "BLXGG", "BLXXGG", "2A", "4A", "6A", "8A", "12A", "X", "TAG", "MARCAÇÃO")
    
    ' Verifica se o nome está na lista de nomes permitidos
    NomeEstaNaLista = False
    For i = LBound(nomesPermitidos) To UBound(nomesPermitidos)
        If nomeMaiusculo = nomesPermitidos(i) Then
            NomeEstaNaLista = True
            Exit Function
        End If
    Next i
End Function


Sub MoverParaProximaCamada()
    Dim sh As shape
    Dim currentPage As page
    Dim nextLayerName As String
    Dim layer As layer
    Dim newLayer As layer
    
    ' Verifica se há um objeto selecionado
    If ActiveSelection.shapes.Count <> 1 Then
        MsgBox "Selecione exatamente um objeto."
        Exit Sub
    End If
    
    ' Obtém a forma selecionada
    Set sh = ActiveSelection.shapes(1)
    
    ' Obtém a página atual
    Set currentPage = ActivePage
    
    ' Verifica se há outras camadas além de #Layer 1 e Linhas-guia
    If ExisteOutraCamada(currentPage) Then
        ' Define o nome da próxima camada
        nextLayerName = ProximaCamada(currentPage)
    Else
        ' Cria uma nova camada
        Set newLayer = currentPage.CreateLayer("Nova Camada")
        nextLayerName = newLayer.Name
    End If
    
    ' Move a forma para a próxima camada
    sh.MoveToLayer nextLayerName

    
    ' Exclui todas as camadas com o nome #Layer 1
    For Each layer In currentPage.Layers
        If layer.Name = "#Layer 1" Then
            layer.Delete
        End If
    Next layer
End Sub

Function ProximaCamada(page As page) As String
    Dim layer As layer
    Dim layerNames As String
    Dim nextLayerName As String
    
    ' Concatena os nomes de todas as camadas
    For Each layer In page.Layers
        If layer.Name <> "#Layer 1" And layer.Name <> "Linhas-guia" Then
            layerNames = layerNames & layer.Name & ","
        End If
    Next layer
    
    ' Obtém o nome da próxima camada
    nextLayerName = Split(layerNames, ",")(0)
    
    ' Retorna o nome da próxima camada
    ProximaCamada = nextLayerName
End Function

Function ExisteOutraCamada(page As page) As Boolean
    Dim layer As layer
    
    ' Verifica se há camadas além de #Layer 1 e Linhas-guia
    For Each layer In page.Layers
        If layer.Name <> "#Layer 1" And layer.Name <> "Linhas-guia" Then
            ExisteOutraCamada = True
            Exit Function
        End If
    Next layer
    
    ExisteOutraCamada = False
End Function

