Sub Evandro()
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document

    Set d = ActiveDocument
    Set p = d.ActivePage

    ' Oculta as camadas
    For Each l In p.Layers
        If l.Name = "EvandroRocha" Then
            l.Visible = False
        End If
        
        If l.Name = "Rogerio" Then
            l.Visible = False
        End If
        
        If l.Name = "ThiagoNery" Then
            l.Visible = False
        End If
        
        If l.Name = "TiagoTR10" Then
            l.Visible = False
        End If
        
        If l.Name = "Matheus" Then
            l.Visible = False
        End If
    Next l

    ' Exibir a camada "Evandro"
    For Each l In p.Layers
        If l.Name = "EvandroRocha" Then
            l.Visible = True
        End If
    Next l
End Sub

Sub Rogerio()
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document

    Set d = ActiveDocument
    Set p = d.ActivePage

    ' Oculta as camadas
    For Each l In p.Layers
        If l.Name = "EvandroRocha" Then
            l.Visible = False
        End If
        
        If l.Name = "Rogerio" Then
            l.Visible = False
        End If
        
        If l.Name = "ThiagoNery" Then
            l.Visible = False
        End If
        
        If l.Name = "TiagoTR10" Then
            l.Visible = False
        End If
        
        If l.Name = "Matheus" Then
            l.Visible = False
        End If
    Next l

    ' Exibir a camada "Rogerio"
    For Each l In p.Layers
        If l.Name = "Rogerio" Then
            l.Visible = True
        End If
    Next l
End Sub

Sub ThiagoNery()
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document

    Set d = ActiveDocument
    Set p = d.ActivePage

    ' Oculta as camadas
    For Each l In p.Layers
        If l.Name = "EvandroRocha" Then
            l.Visible = False
        End If
        
        If l.Name = "Rogerio" Then
            l.Visible = False
        End If
        
        If l.Name = "ThiagoNery" Then
            l.Visible = False
        End If
        
        If l.Name = "TiagoTR10" Then
            l.Visible = False
        End If
        
        If l.Name = "Matheus" Then
            l.Visible = False
        End If
    Next l

    ' Exibir a camada "Thiago Nery"
    For Each l In p.Layers
        If l.Name = "ThiagoNery" Then
            l.Visible = True
        End If
    Next l
End Sub

Sub TiagoTR10()
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document

    Set d = ActiveDocument
    Set p = d.ActivePage

    ' Oculta as camadas
    For Each l In p.Layers
        If l.Name = "EvandroRocha" Then
            l.Visible = False
        End If
        
        If l.Name = "Rogerio" Then
            l.Visible = False
        End If
        
        If l.Name = "ThiagoNery" Then
            l.Visible = False
        End If
        
        If l.Name = "TiagoTR10" Then
            l.Visible = False
        End If
        
        If l.Name = "Matheus" Then
            l.Visible = False
        End If
    Next l

    ' Exibir a camada "Tiago TR10"
    For Each l In p.Layers
        If l.Name = "TiagoTR10" Then
            l.Visible = True
        End If
    Next l
End Sub

Sub Matheus()
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document

    Set d = ActiveDocument
    Set p = d.ActivePage

    ' Oculta as camadas
    For Each l In p.Layers
        If l.Name = "EvandroRocha" Then
            l.Visible = False
        End If
        
        If l.Name = "Rogerio" Then
            l.Visible = False
        End If
        
        If l.Name = "ThiagoNery" Then
            l.Visible = False
        End If
        
        If l.Name = "TiagoTR10" Then
            l.Visible = False
        End If
        
        If l.Name = "Matheus" Then
            l.Visible = False
        End If
    Next l

    ' Exibir a camada "Matheus"
    For Each l In p.Layers
        If l.Name = "Matheus" Then
            l.Visible = True
        End If
    Next l
End Sub

Sub OcultaVendedores()
    Dim s As shape
    Dim l As Layer
    Dim p As Page
    Dim d As Document

    Set d = ActiveDocument
    Set p = d.ActivePage

    ' Oculta as camadas
    For Each l In p.Layers
        If l.Name = "EvandroRocha" Then
            l.Visible = False
        End If
        
        If l.Name = "Rogerio" Then
            l.Visible = False
        End If
        
        If l.Name = "ThiagoNery" Then
            l.Visible = False
        End If
        
        If l.Name = "TiagoTR10" Then
            l.Visible = False
        End If
        
        If l.Name = "Matheus" Then
            l.Visible = False
        End If
    Next l
End Sub

Sub OcultarExibirObjetosMatheus()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As Layer
    Dim objeto As shape
    Dim encontrado As Boolean
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Ocultar todos os objetos na camada "Matheus"
    encontrado = False
    For Each camada In pagina.Layers
        If camada.Name = "Matheus" Then
            For Each objeto In camada.shapes
                objeto.Visible = False
            Next objeto
            encontrado = True
            Exit For
        End If
    Next camada
    
    ' Exibir os objetos "ContatoMatheusTradicional" e "EtiquetaMatheus"
    If encontrado Then
        ExibirObjetoNaCamada pagina, "Matheus", "ContatoMatheusTradicional"
        ExibirObjetoNaCamada pagina, "Matheus", "EtiquetaMatheus"
    Else
        MsgBox "Camada 'Matheus' não encontrada."
    End If
End Sub

Sub ExibirObjetoNaCamada(pagina As Page, nomeCamada As String, nomeObjeto As String)
    Dim camada As Layer
    Dim objeto As shape
    Dim encontrado As Boolean
    
    ' Procurar pela camada com o nome especificado
    encontrado = False
    For Each camada In pagina.Layers
        If camada.Name = nomeCamada Then
            camada.Visible = True
            encontrado = True
            Exit For
        End If
    Next camada
    
    If Not encontrado Then
        MsgBox "Camada """ & nomeCamada & """ não encontrada."
        Exit Sub
    End If
    
    ' Exibir o objeto na camada
    For Each objeto In pagina.shapes
        If objeto.Name = nomeObjeto Then
            objeto.Visible = True
            Exit Sub
        End If
    Next objeto
    
    MsgBox "Objeto """ & nomeObjeto & """ não encontrado na camada """ & nomeCamada & """."
End Sub

