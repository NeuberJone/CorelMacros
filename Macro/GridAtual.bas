Sub RenomearObjetosComNomeDaCamada()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As Layer
    Dim objeto As Shape
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Percorrer todas as camadas na página
    For Each camada In pagina.Layers
        ' Percorrer todos os objetos na camada
        For Each objeto In camada.shapes
            ' Renomear o objeto com o nome da camada
            objeto.Name = camada.Name
        Next objeto
    Next camada
    
    'MsgBox "Objetos foram renomeados com sucesso!"
End Sub

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
    
    ' Exibe uma mensagem com o número de camadas excluídas
    'MsgBox camadasExcluidas & " camadas vazias foram excluídas."
End Sub

Sub RenomearPaginaAtual()
    Dim doc As Document
    Dim pagina As Page
    Dim nomeArquivo As String
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Renomeia a página atual para "DESIGN"
    pagina.Name = "DESIGN"
    
    ' Verifica se o nome do documento segue o padrão "Sem Título-" seguido de um número
    If InStr(1, doc.Name, "Sem título-") = 1 Then
        ' O nome do documento segue o padrão, então definimos o nome do documento como "SISBolt_"
        nomeArquivo = "SISBolt_"
    Else
        ' O nome do documento não segue o padrão, então mantemos o nome atual do documento
        nomeArquivo = Replace(doc.Name, ".cdr", "")
    End If
    
    ' Define o nome do documento
    doc.Name = nomeArquivo
    
    ' Exclui a camada "Camada 1" se existir
    ExcluirCamada1
End Sub

Sub RenomearObjetos()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As Layer
    Dim objeto As Shape
    
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
    
    'MsgBox "Objetos foram renomeados conforme as especificações."
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

Sub MoverObjetosParaCamada()
    Dim doc As Document
    Dim pagina As Page
    Dim objeto As Shape
    Dim camada As Layer
    
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
    
    'MsgBox "Objetos movidos para as camadas correspondentes com sucesso!"
End Sub

Sub AjustaPagina()
    Dim doc As Document
    Dim pagina As Page
    Dim objeto As Shape
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
    
    ' Limpa a seleção
    'pagina.ClearSelection
    
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
    
    ' Calcular nova posição para os objetos
    novoX = larguraPagina / 2
    novoY = alturaPagina / 2
    
    ActivePage.shapes.All.CreateSelection
    ActiveSelection.Move novoX, novoY
    
    Utilidades.RenomearPaginaAtual
    
    'MsgBox "Página ajustada com sucesso!"
End Sub

Sub PreparaGrid()
    ' Verifica se o gerenciador de objetos está aberto
    Dim gerenciadorAberto As Boolean
    gerenciadorAberto = VerificarGerenciadorObjetos()

    ' Se o gerenciador de objetos estiver aberto, continue com o código
    If gerenciadorAberto Then
        RenomearObjetosComNomeDaCamada
        ActivePage.shapes.All.CreateSelection
        AgrupaECentralizaObjetos
        AjustaPagina
        CentralizaObjetosNaPagina
        DesagrupaObjetos
        MoverObjetosParaCamada
        RenomearObjetos
        RenomearPaginaAtual
        ExcluirCamadasVazias
        MoverObjetosParaCamada
        RenomearObjetos
        MsgBox "Grid Preparado com sucesso! Eu acho..."
    Else
        MsgBox "Então deu ruim, abre a bagaça e roda denovo"
    End If
End Sub

Function VerificarGerenciadorObjetos() As Boolean
    ' Exibe uma mensagem para o usuário
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("O Gerenciador de Objetos está aberto?", vbYesNo, "Verificar Gerenciador de Objetos")
    
    ' Retorna true se o usuário selecionou "Sim" e false se selecionou "Não"
    VerificarGerenciadorObjetos = (resposta = vbYes)
End Function


Sub AgrupaECentralizaObjetos()
    ' Recorded 28/02/2024
    ActivePage.shapes.All.CreateSelection
    Dim s1 As Shape
    Set s1 = ActiveSelection.Group
    s1.AlignAndDistribute 3, 3, 2, 0, False, 2
End Sub

Sub DesagrupaObjetos()
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    Dim grp1 As ShapeRange
    Set grp1 = OrigSelection.UngroupEx
End Sub

Sub CentralizaObjetosNaPagina()
    ' Recorded 28/02/2024
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection.AlignAndDistribute 3, 3, 2, 0, False, 2
End Sub

Sub ObterPosicaoXY()
    Dim selecao As ShapeRange
    Dim objeto As Shape
    
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


