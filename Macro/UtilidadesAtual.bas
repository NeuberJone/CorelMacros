Sub RenomearObjetos()
    Dim doc As Document
    Dim pagina As Page
    Dim s As Shape

    ' Referência ao documento ativo
    Set doc = ActiveDocument

    ' Referência à página ativa
    Set pagina = doc.ActivePage

    ' Percorrer todos os objetos na página
    For Each s In pagina.shapes
        ' Verificar se o nome do objeto é "NO" e renomeá-lo para "TextoNome"
        If s.Name = "NO" Then
            s.Name = "TextoNome"
        End If
        ' Verificar se o nome do objeto é "NU" e renomeá-lo para "TextoNumero"
        If s.Name = "NU" Then
            s.Name = "TextoNumero"
        End If
    Next s
End Sub

Sub RenomearPaginaAtualBackup()
    Dim doc As Document
    Dim pagina As Page
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Renomeia a página atual para "DESIGN"
    pagina.Name = "DESIGN"
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

Sub CorrigeGridRaglanGolaV()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As Layer
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Renomear camadas
    For Each camada In pagina.Layers
        camada.Name = Replace(camada.Name, "CostasRaglan", "COSTAS")
        camada.Name = Replace(camada.Name, "FrenteRaglanV", "FRENTEV")
        camada.Name = Replace(camada.Name, "renteRaglanV", "FRENTEV")
        camada.Name = Replace(camada.Name, "MangaRaglanDireita", "MD")
        camada.Name = Replace(camada.Name, "MangaRaglanEsquerda", "ME")
    Next camada
        
    ExcluirCamada1
    RenomearPaginaAtual
    
End Sub

Sub CorrigeGridRaglan()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As Layer
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Renomear camadas
    For Each camada In pagina.Layers
        camada.Name = Replace(camada.Name, "CostasRaglan", "COSTAS")
        camada.Name = Replace(camada.Name, "FrenteRaglan", "FRENTE")
        camada.Name = Replace(camada.Name, "renteRaglan", "FRENTE")
        camada.Name = Replace(camada.Name, "MangaRaglanDireita", "MD")
        camada.Name = Replace(camada.Name, "MangaRaglanEsquerda", "ME")
    Next camada
    
    ExcluirCamada1
    RenomearPaginaAtual
    
End Sub

Sub ExcluirCamada1()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As Layer
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Itera sobre as camadas na página ativa
    For Each camada In pagina.Layers
        ' Verifica se o nome da camada é "Camada 1"
        If camada.Name = "Camada 1" Then
            ' Exclui a camada
            camada.Delete
            Exit For ' Sai do loop assim que a camada for excluída
        End If
    Next camada
End Sub

Sub CorrigeLinhaOuro()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As Layer
    Dim objeto As Shape
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Renomear camadas
    For Each camada In pagina.Layers
        camada.Name = Replace(camada.Name, "PUNHOE", "PunhoEOuro")
        camada.Name = Replace(camada.Name, "PUNHOD", "PunhoDOuro")
        
        camada.Name = Replace(camada.Name, "MANGAE", "MangaEOuro")
        camada.Name = Replace(camada.Name, "MANGAD", "MangaDOuro")
        
        camada.Name = Replace(camada.Name, "S1", "StOuro1")
        camada.Name = Replace(camada.Name, "S2", "StOuro2")
        camada.Name = Replace(camada.Name, "S3", "StOuro3")
        camada.Name = Replace(camada.Name, "S4", "StOuro4")
        
        'camada.Name = Replace(camada.Name, "COSTA", "CostaOuro")
        
        ' Renomear objetos dentro da camada
        For Each objeto In camada.shapes
            objeto.Name = Replace(objeto.Name, "PUNHOE", "PunhoEOuro")
            objeto.Name = Replace(objeto.Name, "PUNHOD", "PunhoDOuro")
            
            objeto.Name = Replace(objeto.Name, "MANGAE", "MangaEOuro")
            objeto.Name = Replace(objeto.Name, "MANGAD", "MangaDOuro")
            
            objeto.Name = Replace(objeto.Name, "S1", "StOuro1")
            objeto.Name = Replace(objeto.Name, "S2", "StOuro2")
            objeto.Name = Replace(objeto.Name, "S3", "StOuro3")
            objeto.Name = Replace(objeto.Name, "S4", "StOuro4")
        Next objeto
    Next camada
    
    ' Exclui a camada "Camada 1" se existir
    ExcluirCamada1
    
    ' Renomeia a página atual
    RenomearPaginaAtual

    MsgBox "Grid Linha Ouro corrigido com sucesso!"

End Sub

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
    Dim s1 As Shape
    Set s1 = ActiveShape
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

Sub DeslocarParaCima()
    Dim doc As Document
    Dim pagina As Page
    Dim selecao As ShapeRange
    Dim alturaObjetos As Double
    Dim deslocamentoCm As Double
    Dim yAtual As Double
    Dim novoY As Double
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage

    ' Referência à seleção ativa
    Set selecao = ActiveSelectionRange
    
    ' Inicializa as variáveis de dimensão
    alturaObjetos = pagina.shapes.All.SizeHeight
    deslocamentoCm = -5 ' Defina o deslocamento em centímetros aqui
    
    ' Converta centímetros para milímetros
    Dim deslocamentoMm As Double
    deslocamentoMm = deslocamentoCm * 10
    
    ' Chame a função ConverteMmEmPontos com o valor em milímetros como parâmetro
    Dim deslocamentoPontos As Double
    deslocamentoPontos = ConverteMmEmPontos(deslocamentoMm)

    ' Obter posição Y do objeto selecionado
    ObterPosicaoY selecao, yAtual

    ' Calcular nova posição Y
    novoY = yAtual - deslocamentoPontos

    ' Mover a seleção para a nova posição
    selecao.Move 0, novoY
End Sub

Sub DeslocarParaBaixo()
    Dim doc As Document
    Dim pagina As Page
    Dim selecao As ShapeRange
    Dim alturaObjetos As Double
    Dim deslocamentoCm As Double
    Dim yAtual As Double
    Dim novoY As Double
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage

    ' Referência à seleção ativa
    Set selecao = ActiveSelectionRange
    
    ' Inicializa as variáveis de dimensão
    alturaObjetos = pagina.shapes.All.SizeHeight
    deslocamentoCm = -5 ' Defina o deslocamento em centímetros aqui
    
    ' Converta centímetros para milímetros
    Dim deslocamentoMm As Double
    deslocamentoMm = deslocamentoCm * 10
    
    ' Chame a função ConverteMmEmPontos com o valor em milímetros como parâmetro
    Dim deslocamentoPontos As Double
    deslocamentoPontos = ConverteMmEmPontos(deslocamentoMm)

    ' Obter posição Y do objeto selecionado
    ObterPosicaoY selecao, yAtual

    ' Calcular nova posição Y
    novoY = yAtual - deslocamentoPontos

    ' Mover a seleção para a nova posição
    selecao.Move 0, -novoY
End Sub

Sub ObterPosicaoY(ByVal selecao As ShapeRange, ByRef posY As Double)
    Dim objeto As Shape
    
    ' Verificar se há uma seleção
    If Not selecao Is Nothing Then
        ' Verificar se há apenas um objeto selecionado
        If selecao.Count = 1 Then
            ' Obter o objeto selecionado
            Set objeto = selecao(1)
            ' Obter a coordenada Y do objeto
            posY = objeto.PositionY
        End If
    End If
End Sub

Function ConverteMmEmPontos(ByVal valorMm As Double) As Double
    ' Assumindo que 1 mm é aproximadamente 2.83465 pontos (pt) no CorelDRAW
    ' Pode ser necessário ajustar esse valor dependendo da configuração do seu documento
    Const pontosPorMm As Double = 0.03937
    ' Calcula o valor em pontos (pt)
    ConverteMmEmPontos = valorMm * pontosPorMm
End Function

