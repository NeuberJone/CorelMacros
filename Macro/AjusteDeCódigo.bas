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








Sub PreparaBandeiraParaSaida()
    PreparaBrindeParaSaida
    RenomeiaPaginaAtual "Bandeira"
    SalvaBrinde
End Sub

Sub PreparaFaixaCapitaoParaSaida()
    PreparaBrindeParaSaida
    RenomeiaPaginaAtual "Faixa de Capitão"
    SalvaBrinde
End Sub

Function ObterNomeDoLayout() As String
    Dim userInput As String
    
    ' Abre a caixa de diálogo para entrada de texto
    userInput = InputBox("Qual o nome do Layout?")
    
    ' Retorna o nome inserido pelo usuário
    ObterNomeDoLayout = userInput
End Function

Sub SalvaBrinde()
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    Dim SaveOptions As StructSaveAsOptions
    Dim NomeArquivo As String
    
    ' Obtém o nome do layout do usuário
    Dim nomeLayout As String
    nomeLayout = ObterNomeDoLayout()
    
    ' Obtém o nome do documento
    Dim nomeDocumento As String
    nomeDocumento = RenomeiaDocumentoAtual("Bandeira") ' Modifique conforme necessário
    
    ' Define o nome do arquivo
    NomeArquivo = "Z:\_Rodar\" & nomeLayout & "_" & nomeDocumento & ".cdr"
    
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
    ActiveDocument.SaveAs NomeArquivo, SaveOptions
    ActiveDocument.Close
End Sub









Sub PreparaBandeiraParaSaida()
    Dim NomeDocumento As String
    NomeDocuemnto = "Bandeira"
    PreparaBrindeParaSaida
    RenomeiaPaginaAtual NomeDocumento
    NomeDocumento = NomeDocumento
    RenomeiaDocumentoAtual NomeDocumento
    SalvaBrinde NomeDocumento
End Sub

Sub PreparaFaixaCapitaoParaSaida()
    Dim NomeDocumento As String
    NomeDocuemnto = "Faixa_de_Capitão"
    PreparaBrindeParaSaida
    RenomeiaPaginaAtual NomeDocumento
    NomeDocumento = NomeDocumento
    RenomeiaDocumentoAtual NomeDocumento
    SalvaBrinde NomeDocumento
End Sub

Sub RenomeiaDocumentoAtual(NomeArquivo As String)
    Dim doc As Document
    Dim pagina As Page
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument

    If NomeArquivo = "SISBolt_" Then
        ' Verifica se o nome do documento segue o padrão "Sem Título-" seguido de um número
        If InStr(1, doc.Name, "Sem título-") = 1 Then
            ' O nome do documento segue o padrão, então definimos o nome do documento como "SISBolt_"
            doc.Name = "SISBolt_"
        Else
            ' O nome do documento não segue o padrão, então mantemos o nome atual do documento
            doc.Name = Replace(doc.Name, ".cdr", "")
        End If
    Else
        doc.Name = NomeArquivo
        ' O nome do documento não segue o padrão, então definimos o nome do documento como o nome fornecido
        doc.Name = Replace(doc.Name, ".cdr", "")
    End If
End Sub

Private Sub SalvaBrinde(NomeDocumento As String)
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    Dim SaveOptions As StructSaveAsOptions
    Dim NomeArquivo As String
    
    ' Obtém o nome do layout do usuário
    Dim nomeLayout As String
    nomeLayout = ObterNomeDoLayout()
    
    ' Define o nome do arquivo
    NomeArquivo = "Z:\Neuber\Brindes\" & nomeLayout & "_" & NomeDocumento & ".cdr"
    
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
    ActiveDocument.SaveAs NomeArquivo, SaveOptions
    ActiveDocument.Close
End Sub







Sub ObterStringDoUsuario()
    Dim userInput As String
    
    ' Abre a caixa de diálogo para entrada de texto
    userInput = InputBox("Digite uma string:")
    
    ' Verifica se o usuário clicou em Cancelar ou deixou o campo em branco
    If userInput = "" Then
        MsgBox "Nenhuma string foi inserida."
    Else
        MsgBox "Você digitou: " & userInput
    End If
End Sub