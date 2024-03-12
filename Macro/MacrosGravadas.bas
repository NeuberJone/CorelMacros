
Sub AjustaFrente()
    ' Recorded 11/02/2024
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    ActiveDocument.ReferencePoint = cdrCenter
    OrigSelection.SetSize 1.650339, 4.693335
    OrigSelection.SetPosition 0.915, 2.896661
End Sub
Sub AjustaCosta()
    ' Recorded 11/02/2024
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    ActiveDocument.ReferencePoint = cdrCenter
    OrigSelection.SetSize 1.658602, 4.693335
    OrigSelection.SetPosition 2.67, 2.896661
End Sub
Sub CorrigePosicao()
    ' Recorded 11/02/2024
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    ActiveDocument.ReferencePoint = cdrCenter
    OrigSelection.SetPosition 1.794563, 2.896661
End Sub
Sub AtualizarReferencias()
    Dim doc As Document
    Dim pg As Page
    Dim shape As shape
    Dim numPaginas As Integer
    
    ' Verifica se há um documento aberto
    If Documents.Count = 0 Then
        MsgBox "Não há nenhum documento aberto."
        Exit Sub
    End If
    
    ' Define o documento ativo
    Set doc = ActiveDocument
    
    ' Define o número total de páginas
    numPaginas = doc.Pages.Count
    
    ' Loop através de todas as páginas do documento
    For Each pg In doc.Pages
        ' Loop através de todos os objetos na página
        For Each shape In pg.shapes
            ' Verifica se o nome do objeto é "Referencia" e se é um texto artístico
            If shape.Type = cdrTextArtistic And shape.Name = "Referencia" Then
                ' Altera o texto para o número da página
                If pg.Index <= numPaginas Then
                    shape.Text.Story = CStr(pg.Index)
                Else
                    shape.Text.Story = "Erro: Página não encontrada"
                End If
            End If
        Next shape
    Next pg
    
    ' Atualiza a tela para refletir as mudanças
    doc.Refresh
End Sub

