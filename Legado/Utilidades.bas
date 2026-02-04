Attribute VB_Name = "Utilidades"
Sub LimpaObjetosGrid()
    Dim doc As Document
    Dim pagina As page
    Dim camada As layer
    Dim shape As shape
    Dim i As Integer

    ' Referência ao documento ativo
    Set doc = ActiveDocument

    ' Percorre todas as páginas do documento
    For i = 1 To doc.Pages.Count
        Set pagina = doc.Pages(i)

        ' Percorre todas as camadas da página atual
        For Each camada In pagina.Layers
            ' Percorre os objetos da camada de trás para frente (para evitar erro ao excluir)
            For Each shape In camada.shapes.All
                ' Verifica se o nome do objeto corresponde a "x" ou "tag"
                If shape.Name = "x" Or shape.Name = "tag" Then
                    shape.Delete
                End If
            Next shape
        Next camada
    Next i

    ' Limpeza de variáveis
    Set doc = Nothing
    Set pagina = Nothing
End Sub

Sub OrganizarCamadas()
    Dim doc As Document
    Dim pagina As page
    Dim camada As layer
    Dim listaCamadas As New Collection
    Dim prefixos As Variant
    Dim sufixos As Variant
    Dim i As Integer, j As Integer
    Dim nomeCamada As String
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Lista de prefixos na ordem desejada
    prefixos = Array("COL1.", "CAL1.", "CHO1.", "REG1.", "CAM2.", "CAM1.")
    
    ' Lista de sufixos na ordem desejada
    sufixos = Array(".C", ".FE", ".MA")
    
    ' Adiciona as camadas em suas respectivas coleções
    For i = LBound(prefixos) To UBound(prefixos)
        For j = LBound(sufixos) To UBound(sufixos)
            For Each camada In pagina.Layers
                nomeCamada = prefixos(i) & "*" & sufixos(j)
                If camada.Name Like nomeCamada Then
                    listaCamadas.Add camada
                End If
            Next camada
        Next j
    Next i
    
    ' Cria novas camadas na ordem desejada e move os objetos para essas camadas
    For i = 1 To listaCamadas.Count
        Set camada = listaCamadas(i)
        ' Mover a camada para o topo
        If i = 1 Then
            ' Se for a primeira camada, move para cima da camada 2
            camada.MoveAbove pagina.Layers(2)
        Else
            ' Caso contrário, move para cima da camada anterior da lista
            camada.MoveAbove listaCamadas(i - 1)
        End If
    Next i
    
    'MsgBox "Camadas organizadas com sucesso.", vbInformation
End Sub



Sub RenomearObjetosSelecionados()
    Dim sr As shapeRange
    Dim novoNome As String
    Dim obj As shape

    ' Verifica se há uma seleção ativa
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then
        MsgBox "Nenhum objeto selecionado.", vbExclamation
        Exit Sub
    End If

    ' Solicita ao usuário o novo nome para os objetos selecionados
    novoNome = InputBox("Digite o novo nome para os objetos selecionados:", "Renomear Objetos")

    ' Verifica se o usuário digitou um nome
    If Len(novoNome) = 0 Then
        MsgBox "Nenhum nome digitado. Operação cancelada.", vbExclamation
        Exit Sub
    End If

    ' Renomeia todos os objetos selecionados com o nome digitado
    For Each obj In sr
        obj.Name = novoNome
    Next obj

    MsgBox "Objetos renomeados com sucesso.", vbInformation
End Sub


Sub RenomearObjetosComNomeArquivoSimulador()
    Dim doc As Document
    Dim pagina As page
    Dim camada As layer
    Dim obj As shape
    Dim nomeArquivo As String
    Dim i As Integer
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Obtém o nome do arquivo sem a extensão
    nomeArquivo = Left(doc.fileName, InStrRev(doc.fileName, ".") - 1)
    
    ' Percorre todas as páginas no documento
    For Each pagina In doc.Pages
        ' Percorre todas as camadas na página
        For Each camada In pagina.Layers
            ' Percorre todos os objetos na camada
            For i = 1 To camada.shapes.Count
                Set obj = camada.shapes(i)
                obj.Name = nomeArquivo
            Next i
        Next camada
    Next pagina
    
    MsgBox "Objetos renomeados com sucesso.", vbInformation
End Sub


Sub ImportarGola()
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
    Set impflt = ActivePage.ActiveLayer.ImportEx("X:\Dados\SISBolt\GOLA.cdr", 1283, impopt)
    impflt.Finish
    Dim s1 As shape
    Set s1 = ActiveShape
    
End Sub

Sub ExcluirCamadasVazias()
    Dim doc As Document
    Dim pagina As page
    Dim camada As layer
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
    
    'Correcoes.MoverParaProximaCamada
    
End Sub

Sub ObterPosicaoXY()
    Dim selecao As shapeRange
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

Sub VerificarCamadasDuplicadas()
    Dim doc As Document
    Dim pagina As page
    Dim camada As layer
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

Sub MoverCamadaParaTopo()
    Dim pagina As page
    'Dim camada As Layer
    Dim primeiraCamada As layer
    
    ' Referência à página ativa
    Set pagina = ActiveDocument.ActivePage
    
    ' Obter a primeira camada
    Set primeiraCamada = pagina.Layers(2)
    
    ' Mover a camada selecionada para o topo
    ActiveLayer.MoveAbove primeiraCamada
End Sub

Sub NumerarObjetos()
    Dim doc As Document
    Dim pagina As page
    Dim shapes As shapes
    Dim numObjetos As Integer
    Dim i As Integer
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Referência aos objetos da página
    Set shapes = pagina.shapes
    
    ' Conta o número de objetos
    numObjetos = shapes.Count
    
    ' Loop reverso para numerar da última para a primeira
    For i = numObjetos To 1 Step -1
        ' Define o nome do objeto numerado
        shapes(i).Name = CStr(numObjetos - i + 1) & " " & shapes(i).Name
    Next i
End Sub

Sub ImportarArquivosECentralizar()
    ' Especifica o caminho da pasta onde os arquivos estão localizados
    Dim folderPath As String
    folderPath = "Z:\Neuber\Import"
    
    ' Verifica se a pasta existe
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "A pasta especificada não foi encontrada.", vbExclamation
        Exit Sub
    End If
    
    ' Variáveis para armazenar a posição central na página
    Dim centerX As Double
    Dim centerY As Double
    
    ' Define a posição central na página ativa
    centerX = ActivePage.SizeWidth / 2
    centerY = ActivePage.SizeHeight / 2
    
    ' Obtém uma lista de arquivos na pasta
    Dim fileName As String
    fileName = Dir(folderPath & "\*.cdr")
    
    ' Loop para importar cada arquivo na pasta
    Do While fileName <> ""
        ' Importa o arquivo
        Dim importedShape As shape
        Set importedShape = ActiveLayer.Import(folderPath & "\" & fileName)
        
        ' Centraliza o objeto importado na página
        importedShape.PositionX = centerX - importedShape.SizeWidth / 2
        importedShape.PositionY = centerY - importedShape.SizeHeight / 2
        
        ' Obtém o próximo arquivo na pasta
        fileName = Dir
    Loop
    
    MsgBox "Os arquivos foram importados e centralizados com sucesso.", vbInformation
End Sub

Sub CodigoTipoSanguineo()
    ' Macro para alterar o nome dos objetos selecionados com base em um critério
    Dim sr As shapeRange
    Set sr = ActiveSelectionRange
    
    If VerificarSeETexto() = True Then
        ' Para cada objeto na seleção
        Dim sh As shape
        For Each sh In sr
            If sh.Name = "TS" Then
                ' Se o nome do objeto for "TS", altere para "Texto Tipo Sanguíneo"
                sh.Name = "Texto Tipo Sanguíneo"
            Else
                ' Se não for "TS", altere para "TS"
                sh.Name = "TS"
            End If
        Next sh
    Else
        ' Se não for texto, exibe uma mensagem informando que o objeto é inválido
        MsgBox "O objeto selecionado não é um texto.", vbExclamation
    End If
End Sub

Sub CodigoNumero()
    Dim sr As shapeRange
    Set sr = ActiveSelectionRange
    
    If VerificarSeETexto() Then
        ' Para cada objeto na seleção
        Dim sh As shape
        For Each sh In sr
            If sh.Name = "NU" Then
                sh.Name = "Texto Número"
            Else
                sh.Name = "NU"
            End If
        Next sh
    Else
        ' Se não for texto, exibe uma mensagem informando que o objeto é inválido
        MsgBox "O objeto selecionado não é texto.", vbExclamation
    End If
End Sub

Sub CodigoNome()
    Dim sr As shapeRange
    Set sr = ActiveSelectionRange
    
    If VerificarSeETexto() Then
        ' Para cada objeto na seleção
        Dim sh As shape
        For Each sh In sr
            If sh.Name = "NO" Then
                sh.Name = "Texto Nome"
            Else
                sh.Name = "NO"
            End If
        Next sh
    Else
        ' Se não for texto, exibe uma mensagem informando que o objeto é inválido
        MsgBox "O objeto selecionado não é texto.", vbExclamation
    End If
End Sub

Sub CodigoApelido()
    Dim sr As shapeRange
    Set sr = ActiveSelectionRange
    
    If VerificarSeETexto() Then
        ' Para cada objeto na seleção
        Dim sh As shape
        For Each sh In sr
            If sh.Name = "AP" Then
                sh.Name = "Texto Apelido"
            Else
                sh.Name = "AP"
            End If
        Next sh
    Else
        ' Se não for texto, exibe uma mensagem informando que o objeto é inválido
        MsgBox "O objeto selecionado não é texto.", vbExclamation
    End If
End Sub

Function VerificarSeETexto() As Boolean
    ' Função para verificar se o objeto selecionado é um texto
    Dim sr As shapeRange
    Set sr = ActiveSelectionRange

    ' Verifica se há uma seleção ativa
    If sr.Count = 0 Then
        VerificarSeETexto = False
        Exit Function
    End If

    ' Verifica se o objeto é do tipo texto
    Dim sh As shape
    Dim isText As Boolean
    isText = False

    For Each sh In sr
        If sh.Type = cdrTextShape Then
            isText = True
            Exit For
        End If
    Next sh
    
    VerificarSeETexto = isText
End Function

Sub AdicionarMamilo()
    Dim sr As shapeRange
    Dim s1 As shape
    Dim xPos As Double
    Dim yPos As Double
    
    ' Verifica se há uma seleção ativa
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then
        MsgBox "Nenhum objeto selecionado.", vbExclamation
        Exit Sub
    End If
    
    ' Obter a posição X e Y do objeto selecionado
    xPos = sr.shapes(1).PositionX
    yPos = sr.shapes(1).PositionY

    ' Criar um novo objeto na posição centralizada
    Set s1 = ActiveLayer.CreateEllipse2(xPos, yPos, 1.589224, -1.589224)

    ' Aplicar as propriedades ao novo objeto
    s1.Fill.ApplyNoFill
    s1.Outline.SetPropertiesEx 0.007874, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#, Justification:=cdrOutlineJustificationMiddle

    ActiveDocument.ReferencePoint = cdrCenter
    s1.SetSize 0.07874, 0.07874
    s1.Fill.UniformColor.CMYKAssign 0, 0, 0, 0
    s1.Outline.SetNoOutline
    's1.Style.StringAssign "{""fill"":{""secondaryColor"":""CMYK,USER,0,0,0,0,100,00000000-0000-0000-0000-000000000000"",""primaryColor"":""CMYK255,USER,0,0,0,0,100,cccd19cb-4675-4a5e-8bda-d0bbbaab8af0"",""type"":""1"",""overprint"":""0""},""outline"":{""overprint"":""0"",""width"":""0"",""color"":""CMYK,USER,0,0,0,100,100,00000000-0000-0000-0000-000000000000""},""transparency"":{""mode"":""36""}}"

    ' Mover o objeto para a segunda camada
    s1.MoveToLayer ActivePage.Layers(2)
    
    s1.Name = "Marcação"

    ' Excluir o objeto original
    sr.shapes(1).Delete

End Sub



Sub ApagaGolaEPunhoDePesca()
    Dim doc As Document
    Dim pagina As page
    Dim camadaResult As layer
    Dim shape As shape
    Dim grupo As shape
    Dim subShape As shape
    Dim excluirGrupo As Boolean

    ' Referência ao documento ativo
    Set doc = ActiveDocument
    Set pagina = doc.ActivePage

    ' Referência à camada "RESULT"
    On Error Resume Next
    Set camadaResult = pagina.Layers("RESULT")
    On Error GoTo 0

    ' Verifica se a camada "RESULT" existe
    If camadaResult Is Nothing Then
        MsgBox "A camada 'RESULT' não foi encontrada.", vbExclamation
        Exit Sub
    End If

    ' Percorre todos os objetos da camada "RESULT"
    For Each shape In camadaResult.shapes
        ' Verifica se o objeto é um grupo
        If shape.Type = cdrGroupShape Then
            excluirGrupo = False
            
            ' Percorre todos os objetos dentro do grupo
            For Each subShape In shape.shapes
                ' Verifica se o nome de algum objeto dentro do grupo corresponde aos nomes especificados
                If subShape.Name = "PunhoPescaE" Or subShape.Name = "PunhoPescaD" Or subShape.Name = "GolaPesca" Then
                    excluirGrupo = True
                    Exit For
                End If
            Next subShape
            
            ' Exclui o grupo se um dos objetos corresponder aos nomes
            If excluirGrupo Then
                shape.Delete
            End If
        End If
    Next shape

End Sub


Sub Contorno()

    Dim OrigSelection As shapeRange
    Set OrigSelection = ActiveSelectionRange
  
    OrigSelection(1).shapes(2).Style.StringAssign "{""fill"":{""type"":""1"",""primaryColor"":""CMYK255,USER,0,0,0,0,100,cccd19cb-4675-4a5e-8bda-d0bbbaab8af0"",""overprint"":""0"",""winding"":""0"",""screenSpec"":""0,0,45000000,60,0""},""outline"":{""overprint"":""0"",""angle"":""0"",""screenSpec"":""0,0,45000000,60,0"",""matrix"":""1,0,0,0,1,0"",""leftArrowAttributes"":""0|0|0|0|0|0|0"",""behindFill"":""1"",""miterLimit"":""11.478339999999999321289578801952302455902099609375"",""width"":""28222"",""rightArrowAttributes"":""0|0|0|0|0|0|0"",""joinType"":""1"",""variableAttributes"":""0|0"",""projectMatrix"":""1,0,0,0,1,0"",""justification"":""0"",""endCaps"":""1"",""rightArrow"":""|0"",""scaleWithObject"":""1"",""dotLength"":""0"",""color"":""CMYK100,USER,255,255,255,255,100,00000000-0000-0000-0000-000000000000"",""dashDotSpec"":""0"",""dash" & "Adjust"":""0"",""overlapArrow"":""0"",""shareArrow"":""0"",""aspect"":""100"",""leftArrow"":""|0""},""transparency"":{}}"

End Sub

Sub DeixaSoCosta()
    Dim doc As Document
    Dim pagina As page
    Dim camadaResult As layer
    Dim shape As shape
    Dim grupo As shape
    Dim subShape As shape
    Dim excluirGrupo As Boolean

    ' Referência ao documento ativo
    Set doc = ActiveDocument
    Set pagina = doc.ActivePage

    ' Referência à camada "RESULT"
    On Error Resume Next
    Set camadaResult = pagina.Layers("RESULT")
    On Error GoTo 0

    ' Verifica se a camada "RESULT" existe
    If camadaResult Is Nothing Then
        MsgBox "A camada 'RESULT' não foi encontrada.", vbExclamation
        Exit Sub
    End If

    ' Percorre todos os objetos da camada "RESULT"
    For Each shape In camadaResult.shapes
        ' Verifica se o objeto é um grupo
        If shape.Type = cdrGroupShape Then
            excluirGrupo = False
            
            ' Percorre todos os objetos dentro do grupo
            For Each subShape In shape.shapes
                ' Verifica se o nome de algum objeto dentro do grupo corresponde aos nomes especificados
                If subShape.Name = "FRENTE" Or subShape.Name = "MANGAEP" Or subShape.Name = "MANGADP" Or subShape.Name = "PUNHOE" Or subShape.Name = "PUNHOD" Or subShape.Name = "SD" Or subShape.Name = "SE" Then
                    excluirGrupo = True
                    Exit For
                End If
            Next subShape
            
            ' Exclui o grupo se um dos objetos corresponder aos nomes
            If excluirGrupo Then
                shape.Delete
            End If
        End If
    Next shape

End Sub

Sub ContornoRetinho()
    Dim OrigSelection As shapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection(1).Style.StringAssign "{""fill"":{""type"":""1"",""overprint"":""0"",""primaryColor"":""CMYK255,USER,0,0,0,0,100,cccd19cb-4675-4a5e-8bda-d0bbbaab8af0"",""winding"":""0""},""outline"":{""type"":""0"",""overprint"":""0"",""screenSpec"":""0,0,45000000,60,0"",""angle"":""0"",""matrix"":""1,0,0,0,1,0"",""aspect"":""100"",""leftArrow"":""|0"",""width"":""40000"",""dashAdjust"":""0"",""joinType"":""1"",""rightArrowAttributes"":""0|0|0|0|0|0|0"",""dashDotSpec"":""0"",""shareArrow"":""0"",""rightArrow"":""|0"",""projectMatrix"":""1,0,0,0,1,0"",""miterLimit"":""5"",""scaleWithObject"":""1"",""color"":""CMYK,USER,0,0,0,100,100,00000000-0000-0000-0000-000000000000"",""overlapArrow"":""0"",""endCaps"":""0"",""behindFill"":""0"",""leftArrowAttributes"":""0|0|0|0|0|0|0"",""justification"":""2"",""dotLength"":""0""},""transparen" & "cy"":{}}"
End Sub

Sub AplicaContornoReto()
    Dim doc As Document
    Dim sr As shapeRange
    Dim shape As shape
    Dim subShape As shape

    ' Referência ao documento ativo
    Set doc = ActiveDocument

    ' Verifica se há objetos selecionados
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then
        MsgBox "Nenhum objeto selecionado.", vbExclamation
        Exit Sub
    End If

    ' Percorre todos os objetos selecionados
    For Each shape In sr.shapes
        ' Verifica se o objeto é um grupo
        If shape.Type = cdrGroupShape Then
            ' Percorre todos os objetos dentro do grupo
            For Each subShape In shape.shapes
                ' Verifica se o nome do objeto é "x"
                If subShape.Name = "x" Then
                    ' Seleciona o objeto e chama a sub ContornoRetinho
                    subShape.CreateSelection
                    ContornoRetinho
                End If
            Next subShape
        Else
            ' Verifica se o nome do objeto é "x"
            If shape.Name = "x" Then
                ' Seleciona o objeto e chama a sub ContornoRetinho
                shape.CreateSelection
                ContornoRetinho
            End If
        End If
    Next shape
End Sub

Sub CriaPaginaTemp()
    Dim p1 As page
    Dim lr1 As layer
    
    ' Insere uma nova página com as dimensões especificadas e nomeia como "temp"
    Set p1 = ActiveDocument.InsertPagesEx(1, False, ActivePage.index, 16.535433, 14.96063)
    p1.Name = "temp"
    p1.Activate
    
    ' Ativa e renomeia a camada "RESULT" na nova página "temp"
    On Error Resume Next
    Set lr1 = p1.Layers("RESULT")
    If lr1 Is Nothing Then
        Set lr1 = p1.Layers("Camada 1")
        lr1.Name = "temp"
        lr1.Activate
    Else
        lr1.Name = "temp"
        lr1.Activate
    End If
    On Error GoTo 0
End Sub


Sub ApagarPaginaTemp()
    Dim p As page
    
    ' Procura a página chamada "temp" no documento
    On Error Resume Next
    Set p = ActiveDocument.Pages("temp")
    On Error GoTo 0
    
    ' Verifica se a página "temp" foi encontrada
    If Not p Is Nothing Then
        p.Delete ' Apaga a página "temp"
    Else
        MsgBox "A página 'temp' não foi encontrada.", vbExclamation
    End If
End Sub

Sub CopiarCamadaResultParaTemp()
    Dim doc As Document
    Dim camadaResult As layer
    Dim camadaTemp As layer
    Dim camada As layer
    Dim encontradoResult As Boolean
    Dim encontradoTemp As Boolean
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Inicializa as variáveis de controle
    encontradoResult = False
    encontradoTemp = False
    
    ' Percorre todas as camadas do documento para encontrar "RESULT" e "temp"
    For Each camada In doc.Layers
        If camada.Name = "RESULT" Then
            Set camadaResult = camada
            encontradoResult = True
        ElseIf camada.Name = "temp" Then
            Set camadaTemp = camada
            encontradoTemp = True
        End If
    Next camada
    
    ' Verifica se a camada "RESULT" foi encontrada
    If Not encontradoResult Then
        MsgBox "A camada 'RESULT' não foi encontrada.", vbExclamation
        Exit Sub
    End If
    
    ' Verifica se a camada "temp" foi encontrada, caso contrário, cria a camada "temp"
    If Not encontradoTemp Then
        Set camadaTemp = doc.CreateLayer("temp")
    End If
    
    ' Copia todos os objetos da camada "RESULT" para a camada "temp"
    If camadaResult.shapes.Count > 0 Then
        ' Cria a seleção de todos os objetos na camada "RESULT"
        camadaResult.shapes.All.CreateSelection
        ActiveSelection.Copy ' Copia a seleção
        camadaTemp.Activate ' Ativa a camada de destino
        ActiveLayer.Paste ' Cola os objetos na camada "temp"
    Else
        MsgBox "Não há objetos na camada 'RESULT' para copiar.", vbExclamation
        Exit Sub
    End If
    
    MsgBox "Conteúdo da camada 'RESULT' copiado para a camada 'temp'.", vbInformation
End Sub
Sub AplicaContornoGrid()
    Dim doc As Document
    Dim pagina As page
    Dim camada As layer
    Dim shape As shape
    Dim i As Integer
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument

    ' Percorre todas as páginas do documento
    For i = 1 To doc.Pages.Count
        Set pagina = doc.Pages(i)

        ' Percorre todas as camadas da página atual
        For Each camada In pagina.Layers
            ' Percorre todos os objetos da camada
            For Each shape In camada.shapes
                ' Aplica o estilo ao objeto atual
                shape.Style.StringAssign "{""fill"":{""overprint"":""0"",""secondaryColor"":""CMYK,USER,0,0,0,0,100,00000000-0000-0000-0000-000000000000"",""type"":""1"",""primaryColor"":""CMYK255,USER,0,0,0,0,100,cccd19cb-4675-4a5e-8bda-d0bbbaab8af0""},""outline"":{""overprint"":""0"",""type"":""0"",""angle"":""0"",""screenSpec"":""0,0,45000000,60,0"",""overlapArrow"":""0"",""leftArrowAttributes"":""0|0|0|0|0|0|0"",""dotLength"":""0"",""width"":""10000"",""dashDotSpec"":""0"",""rightArrow"":""|0"",""endCaps"":""0"",""rightArrowAttributes"":""0|0|0|0|0|0|0"",""variableAttributes"":""0|0"",""scaleWithObject"":""0"",""projectMatrix"":""1,0,0,0,1,0"",""miterLimit"":""5"",""matrix"":""1,0,0,0,1,0"",""aspect"":""100"",""color"":""CMYK255,USER,0,0,0,255,100,00000000-0000-0000-0000-000000000000"",""behindFill"":""0"",""shareArrow"":""0"",""joinType"":""0"",""leftArrow"":""|0"",""dashAdjust"":""0"",""justification"":""1""},""transparency"":{}}"
            Next shape
        Next camada
    Next i

    ' Limpeza de variáveis
    Set doc = Nothing
    Set pagina = Nothing
End Sub
Sub RenomeiaCamadasNoGrid()
    Dim doc As Document
    Dim pagina As page
    Dim camada As layer
    Dim i As Integer
    Dim nomeOriginal As String
    Dim partes() As String
    Dim novoNome As String
    
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
            
            ' Verifica se há pelo menos 3 partes (o que significa 3 pontos no nome)
            If UBound(partes) >= 3 Then
                ' Mantém apenas as 3 primeiras partes e junta de volta com pontos
                novoNome = partes(0) & "." & partes(1) & "." & partes(2)
                camada.Name = novoNome ' Renomeia a camada
            End If
        Next camada
    Next i

    ' Limpeza de variáveis
    Set doc = Nothing
    Set pagina = Nothing
End Sub

Sub ConfiguraNovoGrid()
    Dim doc As Document
    Dim pagina As page
    Dim camada As layer
    Dim shape As shape
    Dim i As Integer

    ' Referência ao documento ativo
    Set doc = ActiveDocument

    ' Percorre todas as páginas do documento
    For i = 1 To doc.Pages.Count
        Set pagina = doc.Pages(i)

        ' Percorre todas as camadas da página atual
        For Each camada In pagina.Layers
            ' Percorre todos os objetos da camada
            For Each shape In camada.shapes
                ' Verifica se o objeto é um grupo e desagrupa
                If shape.Type = cdrGroupShape Then
                    shape.Ungroup
                End If
            Next shape
        Next camada
    Next i

    ' Chama as outras funções
    ExcluirCamadasVazias
    LimpaObjetosGrid
    AplicaContornoGrid
    RenomeiaCamadasNoGrid
    PreparaGrid
End Sub

Sub ConfiguraMoldNovo()

    ExcluirCamadasVazias
    Correcoes.RenomearCamadas
    Correcoes.RenomearObjetosNoMold
    
End Sub




