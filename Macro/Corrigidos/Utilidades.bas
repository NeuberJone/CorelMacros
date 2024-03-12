Sub ExcluirCamadasVazias()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As Layer
    Dim camadasExcluidas As Integer
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Inicializa o contador de camadas excluídas
    camadasExcluidas = 0
    
    ' Percorre todas as camadas na página
    For Each camada In pagina.Layers
        ' Verifica se a camada não é especial e se está vazia
        If Not camada.IsSpecialLayer And camada.shapes.Count = 0 Then
            ' Exclui a camada
            camada.Delete
            ' Incrementa o contador de camadas excluídas
            camadasExcluidas = camadasExcluidas + 1
        End If
    Next camada
End Sub

Sub ImportarClipBoard()
    ' Grabado el 05/03/2024
    Dim impopt As StructImportOptions
    Set impopt = CreateStructImportOptions
    With impopt
        .MaintainLayers = True
        With .ColorConversionOptions
            .SourceColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%"
            .TargetColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%"
        End With
    End With
    Dim impflt As ImportFilter
    Set impflt = ActivePage.ActiveLayer.ImportEx("C:\Users\Design 4\Downloads\clipboard.ai", 1283, impopt)
    impflt.Finish
    Dim s1 As shape
    Set s1 = ActiveShape
End Sub

Sub ObterPosicaoXY()
    Dim selecao As ShapeRange
    Dim objeto As shape
    
    ' Verificar se há uma seleção
    If Not ActiveSelectionRange Is Nothing Then
        ' Referência à seleção
        Set selecao = ActiveSelectionRange
        ' Verificar se há apenas um objeto selecionado
        If selecao.Count = 1 Then
            ' Obter o objeto selecionado
            Set objeto = selecao(1)
            ' Obter as coordenadas X e Y do objeto
            Dim posX As Double
            Dim posY As Double
            posX = objeto.PositionX
            posY = objeto.PositionY
            ' Exibir as coordenadas X e Y em uma caixa de mensagem
            MsgBox "Posição X: " & posX & vbCrLf & "Posição Y: " & posY
        Else
            MsgBox "Selecione apenas um objeto.", vbExclamation
        End If
    Else
        MsgBox "Nenhum objeto selecionado.", vbExclamation
    End If
End Sub




Function VerificarGerenciadorObjetos() As Boolean
    ' Exibe uma mensagem para o usuário
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("O Gerenciador de Objetos está aberto?", vbYesNo, "Verificar Gerenciador de Objetos")
    
    ' Retorna true se o usuário selecionou "Sim" e false se selecionou "Não"
    VerificarGerenciadorObjetos = (resposta = vbYes)
End Function

Function ObterNomeDoLayout() As String
    Dim userInput As String
    
    ' Abre a caixa de diálogo para entrada de texto
    userInput = InputBox("Qual o nome do Layout?")
    
    ' Retorna o nome inserido pelo usuário
    ObterNomeDoLayout = userInput
End Function


























Sub VerificarCamadasDuplicadas()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As Layer
    Dim nomesCamadas As Object
    Dim nomeCamada As String
    Dim duplicadas As Boolean
    
    ' Cria um objeto dicionário para armazenar os nomes das camadas
    Set nomesCamadas = CreateObject("Scripting.Dictionary")
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Percorre todas as camadas na página
    For Each camada In pagina.Layers
        ' Verifica se o nome da camada já está no dicionário
        If nomesCamadas.Exists(camada.Name) Then
            ' Se o nome da camada já existe, exibe uma mensagem informando que há um nome de camada duplicado
            MsgBox "A camada " & camada.Name & " está duplicada."
            duplicadas = True
        Else
            ' Se o nome da camada não existe, adiciona ao dicionário
            nomesCamadas.Add camada.Name, True
        End If
    Next camada
    
    ' Verifica se foram encontradas camadas duplicadas
    If Not duplicadas Then
        MsgBox "Tirando o que tá errado, o resto tá certo."
    End If
    
    ' Limpa o dicionário
    nomesCamadas.RemoveAll
End Sub

Sub MoverCamadaParaTopo()
    Dim pagina As Page
    'Dim camada As Layer
    Dim primeiraCamada As Layer
    
    ' Referência à página ativa
    Set pagina = ActiveDocument.ActivePage
    
    ' Obter a primeira camada
    Set primeiraCamada = pagina.Layers(2)
    
    ' Mover a camada selecionada para o topo
    ActiveLayer.MoveAbove primeiraCamada
End Sub

Sub MoverCamadaParaFundo()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As Layer
    Dim ultimaCamada As String
    Dim quantidadeDeCamadas As Integer
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Chama a função ContarCamadas para obter o número de camadas
    quantidadeDeCamadas = ContarCamadas(pagina)
    
    ' Armazena o nome da última camada
    ultimaCamada = pagina.Layers(quantidadeDeCamadas).Name
    
    ActiveLayer.MoveBelow ActivePage.Layers(ultimaCamada)
End Sub

Function ContarCamadas(pagina As Page) As Integer
    ' Retorna o número de camadas na página
    ContarCamadas = pagina.Layers.Count
End Function















