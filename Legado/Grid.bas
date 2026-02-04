Attribute VB_Name = "Grid"
Sub AjustaTamanhoDaPagina()
    Dim doc As Document
    Dim pagina As page
    
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
    
    ' Calcular nova posição para os objetos
    novoX = larguraPagina / 2
    novoY = alturaPagina / 2
    
    ActivePage.shapes.All.CreateSelection
    ActiveSelection.Move novoX, novoY
    
    Correcoes.RenomeiaPaginaAtual "DESIGN"
    
    Utilidades.OrganizarCamadas
End Sub

Sub PreparaGrid()
    ' Verifica se o gerenciador de objetos está aberto
    Dim gerenciadorAberto As Boolean
    
    'AbrirJanelaObjetos
    
    gerenciadorAberto = VerificarGerenciadorObjetos()

    ' Se o gerenciador de objetos estiver aberto, continue com o código
    If gerenciadorAberto Then
        RenomearObjetosComNomeDaCamada
        ActivePage.shapes.All.CreateSelection
        
        Correcoes.AgrupaECentralizaObjetos
        AjustaTamanhoDaPagina
        Correcoes.CentralizaObjetosNaPagina
        Correcoes.DesagrupaObjetos
        Correcoes.MoverObjetosParaCamada
        Correcoes.RenomearObjetosNoGrid
        Correcoes.RenomeiaPaginaAtual "DESIGN"
        Correcoes.RenomearDocumentoAtual "SISBolt_"
        Utilidades.ExcluirCamadasVazias
    Else
        MsgBox "Então deu ruim, abre a bagaça e roda denovo"
    End If
End Sub

Sub RenomearObjetosComNomeDaCamada()
    Dim doc As Document
    Dim pagina As page
    Dim camada As layer
    Dim objeto As shape
    
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
End Sub

Sub PreparaBandeiraParaSaida()
    Dim NomeDocumento As String
    NomeDocumento = "Bandeira"
    
    ActivePage.shapes.All.CreateSelection
        
    Correcoes.AgrupaECentralizaObjetos
    
    PreparaMaterialParaSaida
    
    Correcoes.RenomeiaPaginaAtual NomeDocumento
    NomeDocumento = NomeDocumento
    Correcoes.RenomearDocumentoAtual NomeDocumento
    SalvaArquivoMaterial NomeDocumento
End Sub

Sub PreparaFaixaDeCapitaoParaSaida()
    Dim NomeDocumento As String
    NomeDocumento = "Faixa De Capitão"
    
    ActivePage.shapes.All.CreateSelection
        
    Correcoes.AgrupaECentralizaObjetos
    
    PreparaMaterialParaSaida
    
    Correcoes.RenomeiaPaginaAtual NomeDocumento
    NomeDocumento = NomeDocumento
    Correcoes.RenomearDocumentoAtual NomeDocumento
    SalvaArquivoMaterial NomeDocumento
End Sub

Sub PreparaEscudoParaSaida()
    Dim NomeDocumento As String
    NomeDocumento = "Escudo"
    
    ActivePage.shapes.All.CreateSelection
        
    Correcoes.AgrupaECentralizaObjetos
    
    PreparaMaterialParaSaida
    
    Correcoes.RenomeiaPaginaAtual NomeDocumento
    NomeDocumento = NomeDocumento
    Correcoes.RenomearDocumentoAtual NomeDocumento
    SalvaArquivoMaterial NomeDocumento
End Sub

Sub SalvaArquivoMaterial(NomeDocumento As String)
    Dim OrigSelection As shapeRange
    Set OrigSelection = ActiveSelectionRange
    Dim SaveOptions As StructSaveAsOptions
    Dim nomeArquivo As String
    
    ' Obtém o nome do layout do usuário
    Dim nomeLayout As String
    nomeLayout = Utilidades.ObterNomeDoLayout()
    
    If nomeLayout = "" Then
        nomeArquivo = "Z:\Neuber\Saídas\" & NomeDocumento & ".cdr"
    Else
        nomeArquivo = "Z:\Neuber\Saídas\" & NomeDocumento & " - " & nomeLayout & ".cdr"
    End If
    
    ' Define as opções de salvamento
    Set SaveOptions = CreateStructSaveAsOptions
    With SaveOptions
        .EmbedVBAProject = False
        .Filter = cdrCDR
        .IncludeCMXData = False
        .Range = cdrAllPages
        .EmbedICCProfile = False
        .Version = 243
        .KeepAppearance = True
    End With
    
    ' Salva o documento no local escolhido pelo usuário
    ActiveDocument.SaveAs nomeArquivo, SaveOptions
    'ActiveDocument.Close
End Sub

Sub PreparaMaterialParaSaida()
    Dim doc As Document
    Dim pagina As page
    Dim larguraObjetos As Double
    Dim alturaObjetos As Double
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Referência à seleção ativa
    Set selecao = ActiveSelection
    
    ' Verifica se há objetos selecionados
    If selecao.shapes.Count > 0 Then
        ' Obter largura e altura dos objetos selecionados
        larguraObjetos = selecao.SizeWidth
        alturaObjetos = selecao.SizeHeight
        
        ' Redimensionar a página para o tamanho dos objetos selecionados
        pagina.SizeWidth = larguraObjetos
        pagina.SizeHeight = alturaObjetos
        Correcoes.CentralizaObjetosNaPagina
    Else
        MsgBox "Nenhum objeto selecionado."
    End If
End Sub





















