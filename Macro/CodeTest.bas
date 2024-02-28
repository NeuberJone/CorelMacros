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

Sub AplicaCorLinha(ByVal tipoLinha As String)
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document
    Dim vendedores() As String
    Dim i As Integer
    
    ' Organiza os vendedores e armazena em um array
    vendedores = OrganizaVendedores
    
    ' Referencia o documento, a página e a camada
    Set d = ActiveDocument
    Set p = d.ActivePage
    
    ' Desbloqueia a camada "Vendedores" para editar as shapes
    DesbloqueiaCamada "Vendedores"
    
    ' Percorre as layers procurando pela camada "Vendedores"
    For Each l In p.Layers
        If l.Name = "Vendedores" Then
            ' Se encontrou a camada "Vendedores", percorre as shapes dentro dela
            For Each s In l.Shapes
                ' Verifica se o nome da shape corresponde a um vendedor
                For i = LBound(vendedores) To UBound(vendedores)
                    If s.Name = "Contato" & vendedores(i) Then
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
                Next i
            Next s
            Exit For ' Sai do loop depois de encontrar a camada "Vendedores"
        End If
    Next l
    
    ' Bloqueia a camada "Vendedores" novamente para evitar edições acidentais
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
            For Each s In l.Shapes
                s.Visible = False
            Next s
            Exit For ' Sai do loop depois de encontrar a camada "Fundo"
        End If
    Next l
    
    ' Bloqueia a camada novamente
    BloqueiaCamada "Fundo"
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
    Dim nomeContato As String
    Dim nomeEtiqueta As String

    Set d = ActiveDocument
    Set p = d.ActivePage
    
    ' Concatena "Contato" com o nome do vendedor
    nomeContato = "Contato" & nomeVendedor

    ' Concatena "Etiqueta" com o nome do vendedor
    nomeEtiqueta = "Etiqueta" & nomeVendedor
    
    ' Chama a sub VerificaLinhaTradicional para atualizar a variável LinhaTradicional
    VerificaLinhaTradicional
    
    ' Desbloqueia a camada "Vendedores" para tornar os objetos visíveis
    DesbloqueiaCamada "Vendedores"
    
    ' Verifica se o vendedor existe como um objeto na camada "Vendedores"
    For Each l In p.Layers
        If l.Name = "Vendedores" Then
            ' Se a camada "Vendedores" for encontrada, verifica os objetos nela para nomeContato
            For Each s In l.Shapes
                ' Verifica se o objeto é um Contato correspondente ao nome do vendedor
                If s.Name = nomeContato Then
                    vendedorExistente = True
                    s.Visible = True ' Torna o objeto visível
                    Exit For ' Sai do loop depois de encontrar o objeto
                End If
            Next s
            
            ' Verifica os objetos na camada "Vendedores" para nomeEtiqueta
            For Each s In l.Shapes
                ' Verifica se o objeto é uma etiqueta correspondente ao nome do vendedor
                If LinhaTradicional = True And s.Name = nomeEtiqueta Then
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

Sub DefineLinha(ByVal tipoLinha As String)
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document
    
    Dim seloLinha As String

    Set d = ActiveDocument
    Set p = d.ActivePage
    
    ' Concatena "SeloLinha" com o tipo da linha
    nomeContato = "SeloLinha" & tipoLinha
    
    OcultaFundos
    VerificaLinhaTradicional
    
    If LinhaTradicional = False Then
        ' Desbloqueia a camada "Selos" para tornar os objetos visíveis
        DesbloqueiaCamada "Selos"
        
        For Each l In p.Layers
            If l.Name = "Selos" Then
                ' Se a camada "Selos" for encontrada, verifica os objetos nela para seloLinha
                For Each s In l.Shapes
                    
                    If s.Name = tipoLinha Then
                        s.Visible = True ' Torna o objeto visível
                        Exit For ' Sai do loop depois de encontrar o objeto
                    End If
                Next s
                
        For Each l In p.Layers
            If l.Name = "Selos" Then
                ' Se a camada "Selos" for encontrada, verifica os objetos nela para seloLinha
                For Each s In l.Shapes
                    
                    If s.Name = tipoLinha Then
                        s.Visible = True ' Torna o objeto visível
                        Exit For ' Sai do loop depois de encontrar o objeto
                    End If
                Next s
                
                Exit For ' Sai do loop depois de encontrar a camada "Selos"
            End If
        Next l
        
        FundoLinhaEspecial
        DefineContornoContato
        
    Else
        FundoLinhaTradicional
        RemoveContorno
    End If
    
    AplicaCorLinha tipoLinha
    ' Bloqueia a camada "Selos" novamente para evitar edições acidentais
    BloqueiaCamada "Selos"
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

