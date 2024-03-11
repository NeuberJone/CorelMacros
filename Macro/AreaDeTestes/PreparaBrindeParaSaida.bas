Sub PrepraraBrindeParaSaida()
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
    
    ' Seleciona todas as formas na página
    pagina.shapes.All.CreateSelection
    
    ' Inicializa as variáveis de dimensão
    larguraObjetos = pagina.shapes.All.SizeWidth
    alturaObjetos = pagina.shapes.All.SizeHeight
    
    ' Redimensionar a página
    pagina.SizeWidth = larguraObjetos
    pagina.SizeHeight = alturaObjetos
    
    ActivePage.shapes.All.CreateSelection
    ActiveSelection.AlignAndDistribute 3, 3, 2, 0, False, 2

End Sub
