Sub RemoveCodigosNoNu()
    Dim doc As Document
    Dim pagina As Page
    Dim s As shape

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

Sub RenomearObjetosNoMold()

    'Sub para renomear objetos nos molds
    Dim doc As Document
    Dim pagina As Page
    Dim camada As Layer
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
    Dim pagina As Page
    Dim camada As Layer
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
    Dim pagina As Page
    Dim camada As Layer
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
    Dim pagina As Page
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
    Dim pagina As Page
    Dim objeto As shape
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
End Sub

Sub AgrupaECentralizaObjetos()
    ActivePage.shapes.All.CreateSelection
    Dim s1 As shape
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
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection.AlignAndDistribute 3, 3, 2, 0, False, 2
End Sub

Sub RenomearDocumentoAtual(NomeArquivo As String)
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
        ' O nome do documento não segue o padrão, então definimos o nome do documento como o nome fornecido
        doc.Name = NomeArquivo
    End If
End Sub

Sub RenomeiaPaginaAtual(NomePagina As String)
    Dim doc As Document
    Dim pagina As Page
    
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
                            "BLXGG", "BLXXGG", "2A", "4A", "6A", "8A", "12A", "X", "TAG")
    
    ' Verifica se o nome está na lista de nomes permitidos
    NomeEstaNaLista = False
    For i = LBound(nomesPermitidos) To UBound(nomesPermitidos)
        If nomeMaiusculo = nomesPermitidos(i) Then
            NomeEstaNaLista = True
            Exit Function
        End If
    Next i
End Function











