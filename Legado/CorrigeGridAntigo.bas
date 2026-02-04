Attribute VB_Name = "CorrigeGridAntigo"
Sub CorrigeGridCamisaRaglan()
    Dim doc As Document
    Dim pagina As page
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
        
    Next camada
    
    Correcoes.RenomearObjetosNoGrid
    
End Sub

Sub CorrigeGridCamisaTradicional()
    Dim doc As Document
    Dim pagina As page
    Dim camada As layer
    Dim objeto As Shape
    
    ' Referência ao documento ativo
    Set doc = ActiveDocument
    
    ' Referência à página ativa
    Set pagina = doc.ActivePage
    
    ' Renomear camadas
    For Each camada In pagina.Layers
    
        camada.Name = Replace(camada.Name, ".FrenteRagV.", ".FRENTEV.")
        camada.Name = Replace(camada.Name, ".FrenteRag.", ".FRENTE.")
        camada.Name = Replace(camada.Name, ".CostasRag.", ".COSTAS.")
        
    Next camada
    
    Correcoes.RenomearObjetosNoGrid
    
End Sub
