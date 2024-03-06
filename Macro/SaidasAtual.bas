Sub VerificaDetalhesSaida()
    Dim doc As Document
    Dim pagina As Page
    Dim contadorNO As Integer
    Dim contadorNU As Integer
    Dim contadorAP As Integer
    Dim contadorTS As Integer
    Dim nomePagina As String
    Dim mensagem As String
    Dim s As Shape
    
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Verifica se a página atual já tem o nome "DESIGN"
    If pagina.Name = "DESIGN" Then
        MsgBox "A página atual já possui o nome 'DESIGN'."
    End If
    
    ' Verifica se já existe outra página com o nome "DESIGN"
    For Each pagina In doc.Pages
        If pagina.Name = "DESIGN" Then
            MsgBox "Já existe uma página com o nome 'DESIGN'."
        End If
    Next pagina
    
    ' Renomeia a página atual para "DESIGN"
    doc.ActivePage.Name = "DESIGN"
    
    ' Inicializa os contadores
    contadorNO = 0
    contadorNU = 0
    contadorAP = 0
    contadorTS = 0
    
    ' Chama a função para contar objetos em todas as formas recursivamente
    ContarObjetosRecursivamente pagina.shapes, contadorNO, contadorNU, contadorAP, contadorTS
    
    ' Constrói a mensagem com os resultados
    mensagem = "Resultados:" & vbCrLf & vbCrLf
    If contadorNO = 0 Then
        mensagem = mensagem & "Não possui nome" & vbCrLf
    Else
        mensagem = mensagem & "Quantidade de objetos com nome 'NO': " & contadorNO & vbCrLf
    End If
    
    If contadorNU = 0 Then
        mensagem = mensagem & "Não possui número" & vbCrLf
    Else
        mensagem = mensagem & "Quantidade de objetos com nome 'NU': " & contadorNU & vbCrLf
    End If
    
    If contadorAP = 0 Then
        mensagem = mensagem & "Não possui tipo sanguíneo" & vbCrLf
    Else
        mensagem = mensagem & "Quantidade de objetos com nome 'AP': " & contadorAP & vbCrLf
    End If
    
    If contadorTS = 0 Then
        mensagem = mensagem & "Não possui apelido" & vbCrLf
    Else
        mensagem = mensagem & "Quantidade de objetos com nome 'TS': " & contadorTS & vbCrLf
    End If
    
    ' Exibe a mensagem com os resultados
    MsgBox mensagem
End Sub

Sub ContarObjetosRecursivamente(ByRef shapes As shapes, ByRef contadorNO As Integer, ByRef contadorNU As Integer, ByRef contadorAP As Integer, ByRef contadorTS As Integer)
    Dim s As Shape
    Dim i As Integer
    
    For Each s In shapes
        Select Case s.Type
            Case cdrGroupShape ' Se é um grupo, chama a função recursivamente
                ContarObjetosRecursivamente s.GroupItems, contadorNO, contadorNU, contadorAP, contadorTS
            Case cdrPowerClipShape ' Se é um powerclip, chama a função recursivamente
                ContarObjetosRecursivamente s.PowerClip.shapes, contadorNO, contadorNU, contadorAP, contadorTS
            Case Else
                ' Verifica o nome do objeto e atualiza os contadores
                Select Case s.Name
                    Case "NO"
                        contadorNO = contadorNO + 1
                    Case "NU"
                        contadorNU = contadorNU + 1
                    Case "AP"
                        contadorAP = contadorAP + 1
                    Case "TS"
                        contadorTS = contadorTS + 1
                End Select
        End Select
    Next s
End Sub

