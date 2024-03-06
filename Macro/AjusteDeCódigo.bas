Sub DeslocarParaCima()
    Dim selecao As ShapeRange
    Dim objeto As Shape
    Dim altura As Double
    Dim deslocamento As Double
    Dim xAtual As Double
    Dim yAtual As Double

    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage

    ' Referência os objetos ativos
    Set objetos = doc.pagina.ActiveSelection
    
    ' Inicializa as variável de dimensão
    alturaObjetos = pagina.shapes.All.SizeHeight
    deslocamento = alturaObjetos + 2

    ObterPosicaoXY (xAtual, yAtual)

    ConverteCmEmMedidaCorel (deslocamento)

    NovoY = yAtual - deslocamento

    ActiveSelection.Move novoX, novoY

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
    
    ' Obter largura do objeto
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



















Sub DeslocarParaCima()
    Dim doc As Document
    Dim pagina As Page
    Dim selecao As ShapeRange
    Dim objeto As Shape
    Dim alturaObjetos As Double
    Dim deslocamento As Double
    Dim xAtual As Double
    Dim yAtual As Double
    Dim novoX As Double
    Dim novoY As Double
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage

    ' Referência os objetos ativos
    Set selecao = ActiveSelectionRange
    
    ' Inicializa as variáveis de dimensão
    alturaObjetos = pagina.shapes.All.SizeHeight
    deslocamento = alturaObjetos + 2

    ' Obter posição X e Y do objeto selecionado
    ObterPosicaoXY xAtual, yAtual

    ' Converter deslocamento de cm para a unidade do CorelDRAW
    deslocamento = ConverteCmEmMedidaCorel(deslocamento)

    ' Calcular nova posição Y
    novoY = yAtual - deslocamento

    ' Mover a seleção para a nova posição
    selecao.Move 0, novoY
End Sub

Sub ObterPosicaoXY(ByRef posX As Double, ByRef posY As Double)
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
            posX = objeto.PositionX
            posY = objeto.PositionY
        End If
    End If
End Sub

Function ConverteCmEmMedidaCorel(ByVal valorCm As Double) As Double
    ' Assumindo que 1 cm é aproximadamente 28.35 pontos (pt) no CorelDRAW
    ' Pode ser necessário ajustar esse valor dependendo da configuração do seu documento
    Const pontosPorCm As Double = 28.35
    ' Calcula o valor em pontos (pt)
    ConverteCmEmMedidaCorel = valorCm * pontosPorCm
End Function

