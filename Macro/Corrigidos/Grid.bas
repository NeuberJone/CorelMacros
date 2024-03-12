Attribute VB_Name = "Grid"
Sub AjustaTamanhoDaPagina()
    Dim doc As Document
    Dim pagina As Page
    
    'Dim objeto As shape
    
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
    
    ' C