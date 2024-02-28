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
            For Each s In l.Shapes
                If s.Name = "FundoTradicional" Then
                    s.Visible = True
                    LinhaTradicional = True
                Else
                    LinhaTradicional = False
                End If
            Next s
            Exit For
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


Sub OrganizaVendedores()
    ' Definindo um array com os nomes dos vendedores
    Dim vendedores(1 To 5) As String
    vendedores(1) = Matheus
    vendedores(2) = TiagoTR10
    vendedores(3) = ThiagoNery
    vendedores(4) = Rogerio
    vendedores(5) = EvandroRocha
End Sub

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

Sub DefineContornoContato(tipoContorno As String)
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document
    Dim vendedor As Variant
    
    ' Chama a sub OrganizaVendedores para inicializar o array vendedores
    OrganizaVendedores
    
    Set d = ActiveDocument
    Set p = d.ActivePage
    
    For Each l In p.Layers
        If l.Name = "Vendedores" Then
            For Each vendedor In vendedores
                For Each s In l.Shapes
                    With s
                        ' Verifica se o objeto é um Contato
                        If .Type = cdrCurveShape And Left(.Name, 7) = "Contato" & vendedor Then
                            If tipoContorno = "Tradicional" Then
                                ' Remove o contorno
                                RemoveContorno s
                            ElseIf tipoContorno = "Bronze" Or tipoContorno = "Prata" Or tipoContorno = "Ouro" Or tipoContorno = "Diamante" Then
                                ' Aplica o contorno
                                AplicaContorno s
                            End If
                        End If
                    End With
                Next s
            Next vendedor
            Exit For ' Sai do loop depois de encontrar a camada "Vendedores"
        End If
    Next l
End Sub

Sub AplicaCorTradicional(ByVal s As shape)
    s.Fill.ApplyUniformFill CreateCMYKColor(100, 100, 100, 100)
End Sub

Sub AplicaCorLinhaBronze(ByVal s As shape)
    With ActivePage.Layers("Vendedores").Shapes(1).Fill.ApplyFountainFill(CreateCMYKColor(0, 40, 60, 20), CreateCMYKColor(0, 40, 60, 20), cdrLinearFountainFill, 0#, 0, 0, 50, cdrCustomFountainFillBlend).Colors
        .Add CreateCMYKColor(7, 35, 65, 0), 30
        .Add CreateCMYKColor(4, 15, 39, 0), 52
        .Add CreateCMYKColor(11, 45, 67, 0), 76
    End With
End Sub

Sub AplicaCorLinhaPrata(ByVal s As shape)
    With ActivePage.Layers("Vendedores").Shapes(1).Fill.ApplyFountainFill(CreateCMYKColor(0, 0, 0, 40), CreateCMYKColor(0, 0, 0, 40), cdrLinearFountainFill, 0#, 0, 0, 50, cdrCustomFountainFillBlend).Colors
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
    With ActivePage.Layers("Vendedores").Shapes(1).Fill.ApplyFountainFill(CreateCMYKColor(100, 20, 0, 0), CreateCMYKColor(100, 20, 0, 0), cdrLinearFountainFill, 0#, 0, 0, 50, cdrCustomFountainFillBlend).Colors
        .Add CreateCMYKColor(40, 0, 0, 0), 30
        .Add CreateCMYKColor(20, 0, 0, 0), 52
        .Add CreateCMYKColor(40, 0, 0, 0), 76
    End With
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
            For Each s In l.Shapes
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
            For Each s In l.Shapes
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
            For Each s In l.Shapes
                s.Visible = False
            Next s
            Exit For ' Sai do loop depois de encontrar a camada "Vendedores"
        End If
    Next l
    
    ' Bloqueia a camada novamente
    BloqueiaCamada "Vendedores"
End Sub


Sub ExibeVendedor(ByVal nomeVendedor As String)
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document
    Dim vendedorExistente As Boolean

    Set d = ActiveDocument
    Set p = d.ActivePage
    
    VerificaLinhaTradicional
    DesbloqueiaCamada "Vendedores"
    ' Verifica se o vendedor existe como uma camada
    For Each l In p.Layers
        If l.Name = "Vendedores" Then
            ' Se a camada "Vendedores" for encontrada, verifica os objetos nela
            For Each s In l.Shapes
                If Left(s.Name, Len("Contato")) = "Contato" And Right(s.Name, Len(nomeVendedor)) = nomeVendedor Then
                    vendedorExistente = True
                    s.Visible = True ' Torna o objeto visível
                    Exit For
                End If
                If LinhaTradicional = True Then
                    If Left(s.Name, Len("Etiqueta")) = "Etiqueta" And Right(s.Name, Len(nomeVendedor)) = nomeVendedor Then
                        vendedorExistente = True
                        s.Visible = True ' Torna o objeto visível
                        Exit For
                    End If
                End If
            Next s
            Exit For ' Sai do loop depois de encontrar a camada "Vendedores"
        End If
    Next l

    ' Se o vendedor não existir, exibe uma mensagem de aviso
    If Not vendedorExistente Then
        MsgBox "Vendedor não encontrado: " & nomeVendedor
    End If
    
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

