Attribute VB_Name = "Utilidades"
Sub ExcluirCamadasVazias()
    Dim doc As Document
    Dim pagina As Page
    Dim camada As Layer
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
    