Attribute VB_Name = "Correcoes"
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
    Set pagina = doc.Active