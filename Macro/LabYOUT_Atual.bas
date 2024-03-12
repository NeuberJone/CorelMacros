Const Matheus As String = "Matheus"
Const TiagoTR10 As String = "TiagoTR10"
Const ThiagoNery As String = "ThiagoNery"
Const Rogerio As String = "Rogerio"
Const EvandroRocha As String = "EvandroRocha"

Dim LinhaTradicional As Boolean

Sub VerificaLinhaTradicional()

    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document
    

    Set d = ActiveDocument
    Set p = d.ActivePage
    
    ' Busca pela camada "Fundo"

    For Each l In p.Layers
        If l.Name = "Fundo" Then
            ' Verifica todos os objetos na camada "Fundo"
            For Each s In l.shapes
                If s.Name = "FundoTradicional" Then
                    If s.Visible = True Then
                        LinhaTradicional = True
                        Exit For
                    End If
                Else
                    LinhaTradicional = False
                End If
            Next s
        End If
    Next l
End Sub
Sub DesbloqueiaCamada(ByVal nomeCamada As String)
    Dim l As Layer
    
    ' Procura pela camada com o nome especificado
    For Each l In ActivePage.Layers
        If l.Name = nomeCamada Then
            ' Desbloqueia a camada
            l.Editable = True
            Exit Sub ' Sai do loop após encontrar a camada
        End If
    Next l
    
    ' Se a camada não for encontrada, exibe uma mensagem de aviso
    MsgBox "Camada não encontrada: " & nomeCamada
End Sub

Sub BloqueiaCamada(ByVal nomeCamada As String)
    Dim l As Layer
    
    ' Procura pela camada com o nome especificado
    For Each l In ActivePage.Layers
        If l.Name = nomeCamada Then
            ' Bloqueia a camada
            l.Editable = False
            Exit Sub ' Sai do loop após encontrar a camada
        End If
    Next l
    
    ' Se a camada não for encontrada, exibe uma mensagem de aviso
    MsgBox "Camada não encontrada: " & nomeCamada
End Sub


Function OrganizaVendedores() As String()
    ' Definindo um array com os nomes dos vendedores
    Dim Vendedores(1 To 5) As String
    Vendedores(1) = Matheus
    Vendedores(2) = TiagoTR10
    Vendedores(3) = ThiagoNery
    Vendedores(4) = Rogerio
    Vendedores(5) = EvandroRocha
    OrganizaVendedores = Vendedores
End Function

Sub OrganizaLinhas()
    ' Definindo um array com os nomes dos vendedores
    Dim linhas(1 To 5) As String
    linhas(1) = "Tradicional"
    linhas(2) = "Bronze"
    linhas(3) = "Prata"
    linhas(4) = "Ouro"
    linhas(5) = "Diamante"
End Sub

Sub AplicaContorno(ByVal s As shape)
    ' Aplica o contorno conforme especificado
    s.Outline.Width = 0.282 ' Define a espessura do contorno em cm
    s.Outline.SetProperties Color:=CreateCMYKColor(100, 100, 100, 100)
    ' Adicione outras propriedades de contorno aqui, se necessário
End Sub

Sub RemoveContorno(ByVal s As shape)
    ' Remove o contorno
    s.Outline.Clear
End Sub

Sub DefineContornoContato(ByVal tipoContorno As String)
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document
    Dim Vendedores() As String
    Dim i As Integer
    Dim nomeContato As String
    
    ' Organiza os vendedores e armazena em um array
    Vendedores = OrganizaVendedores()
    
    Set d = ActiveDocument
    Set p = d.ActivePage
    
    ' Desbloqueia a camada "Vendedores" para tornar os objetos visíveis
    DesbloqueiaCamada "Vendedores"
    
 ' Verifica se a camada "Vendedores" contém etiquetas para cada vendedor
    For Each l In p.Layers
        If l.Name = "Vendedores" Then
            ' Para cada vendedor, oculta a etiqueta correspondente
            For i = LBound(Vendedores) To UBound(Vendedores)
                nomeContato = "Contato" & Vendedores(i)
                For Each s In l.shapes
                    ' Verifica se o objeto é uma etiqueta correspondente ao nome do vendedor
                    If s.Name = nomeContato Then
                            If tipoContorno = "Tradicional" Then
                                ' Remove o contorno
                                RemoveContorno s
                            ElseIf tipoContorno = "Bronze" Or tipoContorno = "Prata" Or tipoContorno = "Ouro" Or tipoContorno = "Diamante" Then
                                ' Aplica o contorno
                                AplicaContorno s
                            End If
                    End If
                Next s
            Next i
            Exit For ' Sai do loop depois de encontrar a camada "Vendedores"
        End If
    Next l

    ' Bloqueia a camada "Vendedores" novamente para evitar edições acidentais
    BloqueiaCamada "Vendedores"
    
End Sub
Sub AplicaCorTradicional(ByVal s As shape)
    s.Fill.ApplyUniformFill CreateCMYKColor(100, 100, 100, 100)
End Sub

Sub AplicaCorLinhaBronze(ByVal s As shape)
    With ActivePage.Layers("Vendedores").shapes(1).Fill.ApplyFountainFill(CreateCMYKColor(0, 40, 60, 20), CreateCMYKColor(0, 40, 60, 20), cdrLinearFountainFill, 0#, 0, 0, 50, cdrCustomFountainFillBlend).Colors
        .Add CreateCMYKColor(7, 35, 65, 0), 30
        .Add CreateCMYKColor(4, 15, 39, 0), 52
        .Add CreateCMYKColor(11, 45, 67, 0), 76
    End With
End Sub

Sub AplicaCorLinhaPrata(ByVal s As shape)
    With ActivePage.Layers("Vendedores").shapes(1).Fill.ApplyFountainFill(CreateCMYKColor(0, 0, 0, 40), CreateCMYKColor(0, 0, 0, 40), cdrLinearFountainFill, 0#, 0, 0, 50, cdrCustomFountainFillBlend).Colors
        .Add CreateCMYKColor(0, 0, 0, 30), 30
        .Add CreateCMYKColor(0, 0, 0, 10), 52
        .Add CreateCMYKColor(0, 0, 0, 30), 76
    End With
End Sub

Sub AplicaCorLinhaOuro(ByVal s As shape)
    With s.Fill.ApplyFountainFill(CreateCMYKColor(36, 51, 100, 18), CreateCMYKColor(36, 51, 100, 18), cdrLinearFountainFill, 0#, 0, 0, 50, cdrCustomFountainFillBlend).Colors
        .Add CreateCMYKColor(4, 29, 78, 0), 30
        .Add CreateCMYKColor(0, 15, 48, 0), 52
        .Add CreateCMYKColor(5, 29, 79, 0), 76
    End With
End Sub

Sub AplicaCorLinhaDiamante(ByVal s As shape)
    With ActivePage.Layers("Vendedores").shapes(1).Fill.ApplyFountainFill(CreateCMYKColor(100, 20, 0, 0), CreateCMYKColor(100, 20, 0, 0), cdrLinearFountainFill, 0#, 0, 0, 50, cdrCustomFountainFillBlend).Colors
        .Add CreateCMYKColor(40, 0, 0, 0), 30
        .Add CreateCMYKColor(20, 0, 0, 0), 52
        .Add CreateCMYKColor(40, 0, 0, 0), 76
    End With
End Sub
Sub AplicaCorLinha(ByVal tipoLinha As String)
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document
    Dim Vendedores() As String
    Dim i As Integer
    
    ' Organiza os vendedores e armazena em um array
    Vendedores = OrganizaVendedores
    
    ' Referencia o documento, a página e a camada
    Set d = ActiveDocument
    Set p = d.ActivePage
    
    ' Desbloqueia as camadas "Fundo" e "Vendedores" para edição
    DesbloqueiaCamada "Fundo"
    DesbloqueiaCamada "Vendedores"
    
    ' Percorre as layers procurando pela camada "Vendedores"
    For Each l In p.Layers
        If l.Name = "Vendedores" Then
            ' Se encontrou a camada "Vendedores", percorre as shapes dentro dela
            For Each s In l.shapes
                ' Verifica se o nome da shape corresponde a um vendedor
                For i = LBound(Vendedores) To UBound(Vendedores)
                    If s.Name = "Contato" & Vendedores(i) Then
                        ' Aplica a cor correspondente ao tipo de linha
                        Select Case tipoLinha
                            Case "Tradicional"
                                s.Fill.ApplyUniformFill CreateCMYKColor(100, 100, 100, 100)
                            Case "Bronze"
                                AplicaCorLinhaBronze s
                            Case "Prata"
                                AplicaCorLinhaPrata s
                            Case "Ouro"
                                AplicaCorLinhaOuro s
                            Case "Diamante"
                                AplicaCorLinhaDiamante s
                            Case Else
                                MsgBox "Tipo de linha desconhecido: " & tipoLinha
                                Exit Sub
                        End Select
                    End If
                Next i
            Next s
            Exit For ' Sai do loop depois de encontrar a camada "Vendedores"
        End If
    Next l
    
    ' Percorre as layers procurando pela camada "Fundo"
    For Each l In p.Layers
        If l.Name = "Fundo" Then
            ' Se encontrou a camada "Fundo", percorre as shapes dentro dela
            For Each s In l.shapes
                ' Verifica se o nome da shape corresponde a "DetalheLinhaEspecial"
                If s.Name = "DetalheLinhaEspecial" Then
                    ' Aplica a cor correspondente ao tipo de linha
                    Select Case tipoLinha
                        Case "Tradicional"
                            AplicaCorTradicional s
                        Case "Bronze"
                            AplicaCorLinhaBronze s
                        Case "Prata"
                            AplicaCorLinhaPrata s
                        Case "Ouro"
                            AplicaCorLinhaOuro s
                        Case "Diamante"
                            AplicaCorLinhaDiamante s
                        Case Else
                            MsgBox "Tipo de linha desconhecido: " & tipoLinha
                            Exit Sub
                    End Select
                End If
            Next s
            Exit For ' Sai do loop depois de encontrar a camada "Fundo"
        End If
    Next l
    
    ' Bloqueia as camadas "Fundo" e "Vendedores" após a edição
    BloqueiaCamada "Fundo"
    BloqueiaCamada "Vendedores"
End Sub


Sub FundoLinhaTradicional()
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document
    Dim objetoEncontrado As Boolean
    
    Set d = ActiveDocument
    Set p = d.ActivePage
    
    ' Busca pela camada "Fundo"
    For Each l In p.Layers
        If l.Name = "Fundo" Then
            ' Desbloqueia a camada
            DesbloqueiaCamada "Fundo"
            
            ' Verifica todos os objetos na camada "Fundo"
            For Each s In l.shapes
                If s.Name = "FundoTradicional" Then
                    s.Visible = True
                    objetoEncontrado = True
                ElseIf s.Name = "FundoLinhaEspecial" Then
                    s.Visible = False
                    objetoEncontrado = True
                End If
            Next s
            
            BloqueiaCamada "Fundo"
            
            Exit For
        End If
    Next l
    
    ' Se a camada "Fundo" não for encontrada, exibe uma mensagem
    If Not objetoEncontrado Then
        MsgBox "Nenhum objeto com o nome 'FundoTradicional' ou 'FundoLinhaEspecial' foi encontrado na camada 'Fundo'."
    End If
End Sub

Sub FundoLinhaEspecial()
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document
    Dim objetoEncontrado As Boolean
    
    Set d = ActiveDocument
    Set p = d.ActivePage
    
    ' Busca pela camada "Fundo"
    For Each l In p.Layers
        If l.Name = "Fundo" Then
            ' Desbloqueia a camada
            DesbloqueiaCamada "Fundo"
            
            ' Verifica todos os objetos na camada "Fundo"
            For Each s In l.shapes
                If s.Name = "FundoTradicional" Then
                    s.Visible = False
                    objetoEncontrado = True
                ElseIf s.Name = "FundoLinhaEspecial" Then
                    s.Visible = True
                    objetoEncontrado = True
                End If
            Next s
            
            ' Bloqueia a camada novamente
            BloqueiaCamada "Fundo"
            
            Exit For
        End If
    Next l
    
    ' Se a camada "Fundo" não for encontrada, exibe uma mensagem
    If Not objetoEncontrado Then
        MsgBox "Nenhum objeto com o nome 'FundoTradicional' ou 'FundoLinhaEspecial' foi encontrado na camada 'Fundo'."
    End If
End Sub
    
Sub OcultaFundos()
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document

    Set d = ActiveDocument
    Set p = d.ActivePage

    ' Desbloqueia a camada
    DesbloqueiaCamada "Fundo"
    For Each l In p.Layers
        If l.Name = "Fundo" Then
            For Each s In l.shapes
                s.Visible = False
            Next s
            Exit For ' Sai do loop depois de encontrar a camada "Fundo"
        End If
    Next l
    
    ' Bloqueia a camada novamente
    BloqueiaCamada "Fundo"
End Sub
Sub OcultaSelos()
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document

    Set d = ActiveDocument
    Set p = d.ActivePage

    ' Desbloqueia a camada
    DesbloqueiaCamada "Selos"
    For Each l In p.Layers
        If l.Name = "Selos" Then
            For Each s In l.shapes
                s.Visible = False
            Next s
            Exit For ' Sai do loop depois de encontrar a camada "Selos"
        End If
    Next l
    
    ' Bloqueia a camada novamente
    BloqueiaCamada "Selos"
End Sub
Sub OcultaVendedores()
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document

    Set d = ActiveDocument
    Set p = d.ActivePage

    ' Desbloqueia a camada
    DesbloqueiaCamada "Vendedores"
    For Each l In p.Layers
        If l.Name = "Vendedores" Then
            For Each s In l.shapes
                s.Visible = False
            Next s
            Exit For ' Sai do loop depois de encontrar a camada "Vendedores"
        End If
    Next l
    
    ' Bloqueia a camada novamente
    BloqueiaCamada "Vendedores"
End Sub

Sub ExibeVendedor(ByVal nomeVendedor As String)
    Dim nomeContato As String
    Dim nomeEtiqueta As String
    
    ' Concatena "Contato" com o nome do vendedor
    nomeContato = "Contato" & nomeVendedor

    ' Concatena "Etiqueta" com o nome do vendedor
    nomeEtiqueta = "Etiqueta" & nomeVendedor
    
    ExibeContato nomeContato, nomeVendedor
    
    ' Verifica se a linha é tradicional antes de exibir a etiqueta
    If VerificaLinhaTradicional Then
        ExibeEtiqueta nomeEtiqueta
    End If
End Sub

Sub ExibeContato(ByVal nomeContato As String, ByVal nomeVendedor As String)
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document
    Dim vendedorExistente As Boolean

    Set d = ActiveDocument
    Set p = d.ActivePage
    
    ' Desbloqueia a camada "Vendedores" para tornar os objetos visíveis
    DesbloqueiaCamada "Vendedores"
    
    ' Verifica se o vendedor existe como um objeto na camada "Vendedores"
    For Each l In p.Layers
        If l.Name = "Vendedores" Then
            ' Se a camada "Vendedores" for encontrada, verifica os objetos nela para nomeContato
            For Each s In l.shapes
                ' Verifica se o objeto é um Contato correspondente ao nome do vendedor
                If s.Name = nomeContato Then
                    vendedorExistente = True
                    s.Visible = True ' Torna o objeto visível
                    Exit For ' Sai do loop depois de encontrar o objeto
                End If
            Next s
            
            Exit For ' Sai do loop depois de encontrar a camada "Vendedores"
        End If
    Next l

    ' Se o vendedor não existir, exibe uma mensagem de aviso
    If Not vendedorExistente Then
        MsgBox "Vendedor não encontrado: " & nomeVendedor
    End If
    
    ' Bloqueia a camada "Vendedores" novamente para evitar edições acidentais
    BloqueiaCamada "Vendedores"
End Sub

Sub ExibeEtiqueta(ByVal nomeEtiqueta As String)
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document

    Set d = ActiveDocument
    Set p = d.ActivePage
    
    ' Desbloqueia a camada "Vendedores" para tornar os objetos visíveis
    DesbloqueiaCamada "Vendedores"
    
    ' Verifica se a camada "Vendedores" contém a etiqueta correspondente
    For Each l In p.Layers
        If l.Name = "Vendedores" Then
            ' Se a camada "Vendedores" for encontrada, verifica os objetos nela para nomeEtiqueta
            For Each s In l.shapes
                ' Verifica se o objeto é uma etiqueta correspondente ao nome do vendedor
                If s.Name = nomeEtiqueta Then
                    s.Visible = True ' Torna o objeto visível
                    Exit For ' Sai do loop depois de encontrar o objeto
                End If
            Next s
            
            Exit For ' Sai do loop depois de encontrar a camada "Vendedores"
        End If
    Next l
    
    ' Bloqueia a camada "Vendedores" novamente para evitar edições acidentais
    BloqueiaCamada "Vendedores"
End Sub
Sub OcultaEtiquetas()
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document
    Dim Vendedores() As String
    Dim i As Integer
    Dim nomeEtiqueta As String

    ' Organiza os vendedores e armazena em um array
    Vendedores = OrganizaVendedores()
    
    Set d = ActiveDocument
    Set p = d.ActivePage
    
    ' Desbloqueia a camada "Vendedores" para tornar os objetos visíveis
    DesbloqueiaCamada "Vendedores"
    
    ' Verifica se a camada "Vendedores" contém etiquetas para cada vendedor
    For Each l In p.Layers
        If l.Name = "Vendedores" Then
            ' Para cada vendedor, oculta a etiqueta correspondente
            For i = LBound(Vendedores) To UBound(Vendedores)
                nomeEtiqueta = "Etiqueta" & Vendedores(i)
                For Each s In l.shapes
                    ' Verifica se o objeto é uma etiqueta correspondente ao nome do vendedor
                    If s.Name = nomeEtiqueta Then
                        s.Visible = False ' Oculta o objeto
                        Exit For ' Sai do loop depois de ocultar o objeto
                    End If
                Next s
            Next i
            Exit For ' Sai do loop depois de encontrar a camada "Vendedores"
        End If
    Next l
    
    ' Bloqueia a camada "Vendedores" novamente para evitar edições acidentais
    BloqueiaCamada "Vendedores"
End Sub


Sub DefineLinha(ByVal tipoLinha As String)
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document
    
    Dim seloLinha As String

    Set d = ActiveDocument
    Set p = d.ActivePage
    
    ' Concatena "SeloLinha" com o tipo da linha
    seloLinha = "SeloLinha" & tipoLinha
    DesbloqueiaCamada "Vendedores"
    
    If tipoLinha = "Tradicional" Then
        OcultaFundos
        FundoLinhaTradicional
        OcultaSelos
        ExibeEtiquetas
    Else
        ' Desbloqueia a camada "Selos" para tornar os objetos visíveis
        DesbloqueiaCamada "Selos"
        OcultaSelos
        
        For Each l In p.Layers
            If l.Name = "Selos" Then
                ' Se a camada "Selos" for encontrada, verifica os objetos nela para seloLinha
                For Each s In l.shapes
                    If s.Name = seloLinha Then
                        s.Visible = True ' Torna o objeto visível
                        ' Não utiliza o Exit For para continuar verificando os demais objetos
                    End If
                Next s
                Exit For ' Sai do loop depois de encontrar a camada "Selos"
            End If
        Next l
        OcultaEtiquetas
        OcultaFundos
        FundoLinhaEspecial
        DefineContornoContato tipoLinha
    End If
    
    AplicaCorLinha tipoLinha
    ' Bloqueia a camada "Selos" novamente para evitar edições acidentais
    BloqueiaCamada "Selos"
    BloqueiaCamada "Vendedores"
End Sub

Sub VendedorEvandro()
    OcultaVendedores
    ExibeVendedor ("EvandroRocha")
End Sub

Sub VendedorRogerio()
    OcultaVendedores
    ExibeVendedor ("Rogerio")
End Sub

Sub VendedorThiagoNery()
    OcultaVendedores
    ExibeVendedor ("ThiagoNery")
End Sub

Sub VendedorTiagoTR10()
    OcultaVendedores
    ExibeVendedor ("TiagoTR10")
End Sub

Sub VendedorMatheus()
    OcultaVendedores
    ExibeVendedor ("Matheus")
End Sub

Sub DefineLinhaTradicional()
    DefineLinha "Tradicional"
End Sub
Sub DefineLinhaBronze()
    DefineLinha "Bronze"
End Sub
Sub DefineLinhaPrata()
    DefineLinha "Prata"
End Sub
Sub DefineLinhaOuro()
    DefineLinha "Ouro"
End Sub
Sub DefineLinhaDiamante()
    DefineLinha "Diamante"
End Sub
