Sub RenomearObjetosAgrupados()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As layer
    Dim grupo As Shape
    Dim subItem As Shape
    
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
    
    MsgBox "Objetos agrupados foram renomeados conforme as especificações."
End Sub

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

Function ObterNomeCamadaDepoisDoSegundoPonto(nomeCamada As String) As String
    Dim partes() As String
    partes = Split(nomeCamada, ".")
    If UBound(partes) >= 2 Then
        ObterNomeCamadaDepoisDoSegundoPonto = partes(3)
    Else
        ObterNomeCamadaDepoisDoSegundoPonto = ""
    End If
End Function

Sub RenomearCamadas()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As layer
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
    
    MsgBox "Camadas foram renomeadas conforme as especificações."
End Sub

Sub ExcluirCamadasVazias()
    Dim doc As Document
    Dim pagina As Page
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
    
    ' Exibe uma mensagem com o número de camadas excluídas
    MsgBox camadasExcluidas & " camadas vazias foram excluídas."
End Sub

Sub OrganizarOrdemCamadas()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As layer
    Dim camada2 As layer
    Dim primeiraCamada As String
    Dim camadaAtual As layer
    
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Armazena o nome da primeira camada
    primeiraCamada = pagina.Layers(2).Name
    
    ' Percorre todas as camadas restantes
    For Each camada In pagina.Layers
        ' Verifica se a camada é do tipo C
        If camada.Name <> primeiraCamada Then
            If InStr(1, camada.Name, ".C.", vbTextCompare) > 0 Then
                ' Move a camada C acima da primeira camada
                'Se camada = ActiveLayert
                camada.MoveAbove ActivePage.Layers(primeiraCamada)
                'MsgBox camada.Name
            End If
        End If
    Next camada
    
    ' Percorre todas as camadas restantes
    For Each camada In pagina.Layers
        ' Verifica se a camada é do tipo FE
        If camada.Name <> primeiraCamada Then
            If InStr(1, camada.Name, ".FE.", vbTextCompare) > 0 Then
                ' Move a camada FE acima da primeira camada
                'Se camada = ActiveLayert
                camada.MoveAbove ActivePage.Layers(primeiraCamada)
                'MsgBox camada.Name
            End If
        End If
    Next camada
    
    ' Percorre todas as camadas restantes
    For Each camada In pagina.Layers
        ' Verifica se a camada é do tipo MA
        If camada.Name <> primeiraCamada Then
            If InStr(1, primeiraCamada, ".MA.", vbTextCompare) > 0 Then
                For Each camada2 In pagina.Layers
                    ' Verifica se a camada é do tipo PUNHO
                    If camada2.Name <> primeiraCamada Then
                        If InStr(1, camada.Name, ".PUNHOD.", vbTextCompare) > 0 Or InStr(1, camada.Name, ".PUNHOE.", vbTextCompare) > 0 Then
                            ' Move a camada primeiraCamada acima da camada2
                            ActivePage.Layers(primeiraCamada).MoveAbove camada2
                        End If
                    End If
                Next camada2
                ' Move a camada MA acima da primeira camada
                'Se camada = ActiveLayert
                'ActivePage.Layers(primeiraCamada).MoveAbove ActivePage.Layers(camada)
                'MsgBox camada.Name
            End If
        End If
    Next camada
    
    ' Percorre todas as camadas restantes
    For Each camada In pagina.Layers
        ' Verifica se a camada é do tipo FE
        If camada.Name <> primeiraCamada Then
            If InStr(1, primeiraCamada, ".FE.", vbTextCompare) > 0 Then
                For Each camada2 In pagina.Layers
                    ' Verifica se a camada é do tipo MA
                    If camada2.Name <> primeiraCamada Then
                        If InStr(1, camada2.Name, ".MA.", vbTextCompare) > 0 Then
                            ' Move a camada MA acima da primeira camada
                            Dim camadaMA As layer
                            Set camadaMA = ActivePage.Layers(primeiraCamada)
                            camadaMA.MoveAbove camada2
                        End If
                    End If
                Next camada2
                ' Move a camada MA acima da primeira camada
                'Se camada = ActiveLayert
                'ActivePage.Layers(primeiraCamada).MoveAbove ActivePage.Layers(camada)
                'MsgBox camada.Name
            End If
        End If
    Next camada
    
    ' Percorre todas as camadas restantes
    For Each camada In pagina.Layers
        ' Verifica se a camada é do tipo C
        If camada.Name <> primeiraCamada Then
            If InStr(1, primeiraCamada, ".C.", vbTextCompare) > 0 Then
                For Each camada2 In pagina.Layers
                    ' Verifica se a camada é do tipo FE
                    If camada2.Name <> primeiraCamada Then
                        If InStr(1, camada2.Name, ".FE.", vbTextCompare) > 0 Then
                            ' Move a camada MA acima da primeira camada
                            Dim camadaFE As layer
                            Set camadaFE = ActivePage.Layers(primeiraCamada)
                            camadaFE.MoveAbove camada2
                        End If
                    End If
                Next camada2
                ' Move a camada MA acima da primeira camada
                'Se camada = ActiveLayert
                'ActivePage.Layers(primeiraCamada).MoveAbove ActivePage.Layers(camada)
                'MsgBox camada.Name
            End If
        End If
    Next camada
    
    OrganizarPunhos
    
    MsgBox "A ordem das camadas foi organizada com sucesso."
    MsgBox primeiraCamada
End Sub

Sub OrganizarPunhos()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As layer
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
    
    ' Percorre todas as camadas restantes
    For Each camada In pagina.Layers
        ' Verifica se a camada é do tipo punho
        If camada.Name <> ultimaCamada Then
            If InStr(1, camada.Name, ".PUNHOD.", vbTextCompare) > 0 Or InStr(1, camada.Name, ".PUNHOE.", vbTextCompare) > 0 Then
                ' Move a camada de punho para o final
                camada.MoveBelow ActivePage.Layers(ultimaCamada)
            End If
        End If
    Next camada
    
    MsgBox "As camadas de punhos foram movidas para o final."
End Sub

Function ContarCamadas(pagina As Page) As Integer
    ' Retorna o número de camadas na página
    ContarCamadas = pagina.Layers.Count
End Function

Sub RenomearObjetos()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As layer
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
    
    MsgBox "Objetos foram renomeados conforme as especificações."
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

Sub CorrigeGridCamisaMangaCurta()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As layer
    Dim objeto As Shape
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Renomear camadas
    For Each camada In pagina.Layers
        camada.Name = Replace(camada.Name, ".FRENTEV.", ".MCFrenteV.")
        camada.Name = Replace(camada.Name, ".FRENTE.", ".MCFrente.")
        camada.Name = Replace(camada.Name, ".COSTAS.", ".MCCosta.")
        
        camada.Name = Replace(camada.Name, ".MANGAEP.", ".MangCEP.")
        camada.Name = Replace(camada.Name, ".MANGADP.", ".MangCDP.")
        camada.Name = Replace(camada.Name, ".MANGAE.", ".MangCE.")
        camada.Name = Replace(camada.Name, ".MANGAD.", ".MangCD.")
        
        camada.Name = Replace(camada.Name, ".PUNHOE.", ".PunhoE.")
        camada.Name = Replace(camada.Name, ".PUNHOD.", ".PunhoD.")
    Next camada
    
    RenomearObjetos

    MsgBox "Grid camisa manga curta corrigido com sucesso!"
End Sub

Sub CorrigeGridCamisaRaglanMC()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As layer
    Dim objeto As Shape
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Renomear camadas
    For Each camada In pagina.Layers
        camada.Name = Replace(camada.Name, ".FRENTEV.", ".FrenteRagV.")
        camada.Name = Replace(camada.Name, ".FRENTE.", ".FrenteRag.")
        camada.Name = Replace(camada.Name, ".COSTAS.", ".CostasRag.")
        
        camada.Name = Replace(camada.Name, ".MEP.", ".MangRagEP.")
        camada.Name = Replace(camada.Name, ".MDP.", ".MangRagDP.")
        camada.Name = Replace(camada.Name, ".ME.", ".MangRagE.")
        camada.Name = Replace(camada.Name, ".MD.", ".MangRagD.")
        
        'camada.Name = Replace(camada.Name, ".PUNHOE.", ".PunhoE.")
        'camada.Name = Replace(camada.Name, ".PUNHOD.", ".PunhoD.")
    Next camada
    
    RenomearObjetos

    MsgBox "Grid camisa manga curta corrigido com sucesso!"
End Sub
Sub CorrigeGridLinhaOuro()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As layer
    Dim objeto As Shape
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Renomear camadas
    For Each camada In pagina.Layers
        camada.Name = Replace(camada.Name, ".FRENTEV.", ".MCFrenteV.")
        camada.Name = Replace(camada.Name, ".FRENTE.", ".MCFrente.")
        camada.Name = Replace(camada.Name, ".COSTAS.", ".MCCosta.")
        
        camada.Name = Replace(camada.Name, ".MANGAEP.", ".MangCEP.")
        camada.Name = Replace(camada.Name, ".MANGADP.", ".MangCDP.")
        camada.Name = Replace(camada.Name, ".MANGAE.", ".MangCE.")
        camada.Name = Replace(camada.Name, ".MANGAD.", ".MangCD.")
        
        camada.Name = Replace(camada.Name, ".PUNHOE.", ".PunhoE.")
        camada.Name = Replace(camada.Name, ".PUNHOD.", ".PunhoD.")
    Next camada
    
    RenomearObjetos

    MsgBox "Grid camisa manga curta corrigido com sucesso!"
End Sub

Sub RedimensionarEMoverObjetos()
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
    
    MsgBox "Grid Preparado com sucesso!"
End Sub


