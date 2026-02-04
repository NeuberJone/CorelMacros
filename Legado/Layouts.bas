Attribute VB_Name = "Layouts"
Sub MoverEtiquetaParaTopo()
    Dim doc As Document
    Dim pagina As page
    Dim camadaInformacoes As layer
    Dim primeiraCamada As layer
    Dim etiqueta As shape
    Dim objeto As shape
    Dim encontrouEtiqueta As Boolean
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Referência à camada "Informações"
    On Error Resume Next
    Set camadaInformacoes = pagina.Layers("Informações")
    On Error GoTo 0
    
    ' Verifica se a camada "Informações" existe
    If camadaInformacoes Is Nothing Then
        MsgBox "A camada 'Informações' não foi encontrada.", vbExclamation
        Exit Sub
    End If
    
    ' Procurar o objeto "Etiqueta" na camada "Informações"
    encontrouEtiqueta = False
    For Each objeto In camadaInformacoes.shapes
        If objeto.Name = "Etiqueta" Then
            Set etiqueta = objeto
            encontrouEtiqueta = True
            Exit For
        End If
    Next objeto
    
    ' Verifica se o objeto "Etiqueta" foi encontrado
    If Not encontrouEtiqueta Then
        MsgBox "O objeto 'Etiqueta' não foi encontrado na camada 'Informações'.", vbExclamation
        Exit Sub
    End If
    
    ' Referência à primeira camada (assumindo que seja a camada 1)
    If pagina.Layers.Count < 1 Then
        MsgBox "Nenhuma camada encontrada na página.", vbExclamation
        Exit Sub
    End If
    Set primeiraCamada = pagina.Layers(2)
    
    ' Move o objeto "Etiqueta" para a primeira camada
    etiqueta.MoveToLayer primeiraCamada
    
    ' Move o objeto "Etiqueta" para o topo da primeira camada
    etiqueta.OrderToFront
    
    'MsgBox "O objeto 'Etiqueta' foi movido para o topo da primeira camada com sucesso.", vbInformation
End Sub

Sub MoverCamadas()
    Dim doc As Document
    Dim pagina As page
    Dim camadaFundo As layer
    Dim camadaInformacoes As layer
    Dim ultimaCamada As layer
    Dim encontrouFundo As Boolean
    Dim encontrouInformacoes As Boolean

    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    encontrouFundo = False
    encontrouInformacoes = False
    
    ' Procurar a camada "Fundo" e "Informações"
    For Each camada In pagina.Layers
        If camada.Name = "Fundo" Then
            Set camadaFundo = camada
            encontrouFundo = True
        ElseIf camada.Name = "Informações" Then
            Set camadaInformacoes = camada
            encontrouInformacoes = True
        End If
    Next camada
    
    ' Verificar se ambas as camadas foram encontradas
    If Not encontrouFundo Then
        'MsgBox "Camada 'Fundo' não encontrada.", vbExclamation
        Exit Sub
    End If
    
    If Not encontrouInformacoes Then
        'MsgBox "Camada 'Informações' não encontrada.", vbExclamation
        Exit Sub
    End If
    
    ' Mover a camada "Fundo" para a última posição
    Set ultimaCamada = pagina.Layers(pagina.Layers.Count)
    camadaFundo.MoveBelow ultimaCamada
    
    ' Mover a camada "Informações" logo acima da camada "Fundo"
    camadaInformacoes.MoveAbove camadaFundo
    
    'MsgBox "Camadas 'Fundo' e 'Informações' movidas com sucesso.", vbInformation
End Sub

Sub RenomearObjetosNaCamada1()
    Dim doc As Document
    Dim pagina As page
    Dim camada As layer
    Dim objeto As shape
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Referência à camada "Camada 1"
    Set camada = pagina.Layers("Camada 1")
    
    ' Percorre todos os objetos na camada "Camada 1"
    For Each objeto In camada.shapes
        ' Verifica os nomes dos objetos e os altera
        If objeto.Name = "NO" Then
            objeto.Name = "Texto Nome"
        ElseIf objeto.Name = "NU" Then
            objeto.Name = "Texto Numero"
        End If
    Next objeto
End Sub


Sub ImportarBaseStopBall()
    Dim impopt As StructImportOptions
    
    ExcluiFundoEInformacoes
    RenomearObjetosNaCamada1
    
    Set impopt = CreateStructImportOptions
    With impopt
        .MaintainLayers = True
        With .ColorConversionOptions
            .SourceColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%"
            .TargetColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%"
        End With
    End With
    Dim impflt As ImportFilter
    Set impflt = ActivePage.ActiveLayer.ImportEx("X:\Dados\Mockups Novos\Base - Stop Ball.cdr", 1283, impopt)
    impflt.Finish
    Dim s1 As shape
    Set s1 = ActiveShape
    
    MoverCamadas
    Utilidades.ExcluirCamadasVazias
    MoverEtiquetaParaTopo
    
End Sub

Sub AlteraEtiqueta(objetoParaDesocultar As String)
    Dim doc As Document
    Dim pagina As page
    Dim objeto As shape
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Percorre todos os objetos na página
    For Each objeto In pagina.shapes
        VerificarGrupo objeto, objetoParaDesocultar
    Next objeto
End Sub

Sub VerificarGrupo(objeto As shape, objetoParaDesocultar As String)
    Dim subObjeto As shape

    ' Verifica se o objeto é um grupo chamado "Etiqueta"
    If objeto.Type = cdrGroupShape And objeto.Name = "Etiqueta" Then
        ' Oculta os objetos especificados no grupo
        For Each subObjeto In objeto.shapes
            Select Case subObjeto.Name
                Case "Matheus", "Evandro", "Rogério", "Thiago Nery"
                    subObjeto.Visible = False
            End Select
        Next subObjeto

        ' Desoculta o objeto cujo nome é recebido como parâmetro
        For Each subObjeto In objeto.shapes
            If subObjeto.Name = objetoParaDesocultar Then
                subObjeto.Visible = True
            End If
        Next subObjeto
    ElseIf objeto.Type = cdrGroupShape Then
        ' Se o objeto é um grupo, chama a função recursivamente para verificar dentro do grupo
        For Each subObjeto In objeto.shapes
            VerificarGrupo subObjeto, objetoParaDesocultar
        Next subObjeto
    End If
End Sub


Sub OcultaVendedores()
    Dim doc As Document
    Dim pagina As page
    Dim camadaFundo As layer
    Dim s As shape

    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Verifica se a camada "Fundo" existe
    On Error Resume Next
    Set camadaFundo = pagina.Layers("Fundo")
    On Error GoTo 0
    
    If Not camadaFundo Is Nothing Then
        ' Percorre todos os objetos na camada "Fundo"
        For Each s In camadaFundo.shapes
            ' Verifica os nomes dos objetos e os oculta
            If s.Name = "Matheus" Or s.Name = "Thiago Nery" Or s.Name = "Rogério" Or s.Name = "Evandro" Then
                s.Visible = False
            End If
        Next s
    Else
        'MsgBox "A camada 'Fundo' não foi encontrada.", vbExclamation
    End If
End Sub

Sub OcultaFundos()
    Dim doc As Document
    Dim pagina As page
    Dim camadaFundo As layer
    Dim s As shape

    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Verifica se a camada "Fundo" existe
    On Error Resume Next
    Set camadaFundo = pagina.Layers("Fundo")
    On Error GoTo 0
    
    If Not camadaFundo Is Nothing Then
        ' Percorre todos os objetos na camada "Fundo"
        For Each s In camadaFundo.shapes
            ' Verifica os nomes dos objetos e os oculta
            If s.Name = "Fundo Linha Tradicional" Or s.Name = "Fundo Linha Bronze" Or s.Name = "Fundo Linha Prata" Or s.Name = "Fundo Linha Ouro" Or s.Name = "Fundo Linha Diamante" Then
                s.Visible = False
            End If
        Next s
    Else
        'MsgBox "A camada 'Fundo' não foi encontrada.", vbExclamation
    End If
End Sub

Sub BloqueiaVendedores()
    Dim doc As Document
    Dim pagina As page
    Dim camadaFundo As layer
    Dim s As shape

    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Verifica se a camada "Fundo" existe
    On Error Resume Next
    Set camadaFundo = pagina.Layers("Fundo")
    On Error GoTo 0
    
    If Not camadaFundo Is Nothing Then
        ' Percorre todos os objetos na camada "Fundo"
        For Each s In camadaFundo.shapes
            ' Verifica os nomes dos objetos e os oculta
            If s.Name = "Matheus" Or s.Name = "Thiago Nery" Or s.Name = "Rogério" Or s.Name = "Evandro" Then
                s.Locked = True
            End If
        Next s
    Else
        'MsgBox "A camada 'Fundo' não foi encontrada.", vbExclamation
    End If
End Sub

Sub DesbloqueiaVendedores()
    Dim doc As Document
    Dim pagina As page
    Dim camadaFundo As layer
    Dim s As shape

    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Verifica se a camada "Fundo" existe
    On Error Resume Next
    Set camadaFundo = pagina.Layers("Fundo")
    On Error GoTo 0
    
    If Not camadaFundo Is Nothing Then
        ' Percorre todos os objetos na camada "Fundo"
        For Each s In camadaFundo.shapes
            ' Verifica os nomes dos objetos e os oculta
            If s.Name = "Matheus" Or s.Name = "Thiago Nery" Or s.Name = "Rogério" Or s.Name = "Evandro" Then
                s.Locked = False
            End If
        Next s
    Else
        'MsgBox "A camada 'Fundo' não foi encontrada.", vbExclamation
    End If
End Sub

Sub BloqueiaFundos()
    Dim doc As Document
    Dim pagina As page
    Dim camadaFundo As layer
    Dim s As shape

    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Verifica se a camada "Fundo" existe
    On Error Resume Next
    Set camadaFundo = pagina.Layers("Fundo")
    On Error GoTo 0
    
    If Not camadaFundo Is Nothing Then
        ' Percorre todos os objetos na camada "Fundo"
        For Each s In camadaFundo.shapes
            ' Verifica os nomes dos objetos e os oculta
            If s.Name = "Fundo Linha Tradicional" Or s.Name = "Fundo Linha Bronze" Or s.Name = "Fundo Linha Prata" Or s.Name = "Fundo Linha Ouro" Or s.Name = "Fundo Linha Diamante" Then
                s.Locked = True
            End If
        Next s
    Else
        'MsgBox "A camada 'Fundo' não foi encontrada.", vbExclamation
    End If
End Sub

Sub DesbloqueiaFundos()
    Dim doc As Document
    Dim pagina As page
    Dim camadaFundo As layer
    Dim s As shape

    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Verifica se a camada "Fundo" existe
    On Error Resume Next
    Set camadaFundo = pagina.Layers("Fundo")
    On Error GoTo 0
    
    If Not camadaFundo Is Nothing Then
        ' Percorre todos os objetos na camada "Fundo"
        For Each s In camadaFundo.shapes
            ' Verifica os nomes dos objetos e os oculta
            If s.Name = "Fundo Linha Tradicional" Or s.Name = "Fundo Linha Bronze" Or s.Name = "Fundo Linha Prata" Or s.Name = "Fundo Linha Ouro" Or s.Name = "Fundo Linha Diamante" Then
                s.Locked = False
            End If
        Next s
    Else
        'MsgBox "A camada 'Fundo' não foi encontrada.", vbExclamation
    End If
End Sub

Sub MostraObjeto(nomeObjeto As String)
    Dim doc As Document
    Dim pagina As page
    Dim camadaFundo As layer
    Dim s As shape

    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Verifica se a camada "Fundo" existe
    On Error Resume Next
    Set camadaFundo = pagina.Layers("Fundo")
    On Error GoTo 0
    
    If Not camadaFundo Is Nothing Then
        ' Percorre todos os objetos na camada "Fundo"
        For Each s In camadaFundo.shapes
            ' Verifica os nomes dos objetos e os oculta
            If s.Name = nomeObjeto Then
                s.Visible = True
            End If
        Next s
        
    Else
        MsgBox "A camada 'Fundo' não foi encontrada.", vbExclamation
    End If
    
    If nomeObjeto = "Matheus" Or nomeObjeto = "Evandro" Or nomeObjeto = "Rogério" Or nomeObjeto = "Thiago Nery" Then
        AlteraEtiqueta nomeObjeto
    End If

End Sub

Sub MostraEvandro()
    DesbloqueiaVendedores
    OcultaVendedores
    MostraObjeto "Evandro"
    BloqueiaVendedores
End Sub

Sub MostraRogerio()
    DesbloqueiaVendedores
    OcultaVendedores
    MostraObjeto "Rogério"
    BloqueiaVendedores
End Sub

Sub MostraMatheus()
    DesbloqueiaVendedores
    OcultaVendedores
    MostraObjeto "Matheus"
    BloqueiaVendedores
End Sub

Sub MostraThiagoNery()
    DesbloqueiaVendedores
    OcultaVendedores
    MostraObjeto "Thiago Nery"
    BloqueiaVendedores
End Sub

Sub LinhaTradicional()
    DesbloqueiaFundos
    OcultaFundos
    MostraObjeto "Fundo Linha Tradicional"
    BloqueiaFundos
End Sub

Sub LinhaBronze()
    DesbloqueiaFundos
    OcultaFundos
    MostraObjeto "Fundo Linha Bronze"
    BloqueiaFundos
End Sub

Sub LinhaPrata()
    DesbloqueiaFundos
    OcultaFundos
    MostraObjeto "Fundo Linha Prata"
    BloqueiaFundos
End Sub

Sub LinhaOuro()
    DesbloqueiaFundos
    OcultaFundos
    MostraObjeto "Fundo Linha Ouro"
    BloqueiaFundos
End Sub

Sub LinhaDiamante()
    DesbloqueiaFundos
    OcultaFundos
    MostraObjeto "Fundo Linha Diamante"
    BloqueiaFundos
End Sub


Sub ExcluiFundoEInformacoes()
    Dim doc As Document
    Dim pagina As page
    Dim camadaFundo As layer
    Dim camadaInformacoes As layer
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Excluir a camada "Fundo" se existir
    On Error Resume Next
    Set camadaFundo = pagina.Layers("Fundo")
    If Not camadaFundo Is Nothing Then
        camadaFundo.Delete
    End If
    On Error GoTo 0
    
    ' Excluir a camada "Informações" se existir
    On Error Resume Next
    Set camadaInformacoes = pagina.Layers("Informações")
    If Not camadaInformacoes Is Nothing Then
        camadaInformacoes.Delete
    End If
    On Error GoTo 0
    
    ' Selecionar todos os objetos restantes na página ativa
    pagina.shapes.All.CreateSelection
End Sub

