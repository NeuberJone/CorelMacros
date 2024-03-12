export function RedimensionarEMoverObjetos() {
    let doc = app.activeDocument;
    let pagina = doc.activePage;
    
    let larguraObjetos, alturaObjetos, larguraPagina, alturaPagina, novoX, novoY;
    
    larguraPagina = pagina.sizeWidth;
    alturaPagina = pagina.sizeHeight;
    
    pagina.shapes.all.createSelection();
    larguraObjetos = pagina.shapes.all.sizeWidth;
    alturaObjetos = pagina.shapes.all.sizeHeight;
    
    pagina.sizeWidth = larguraObjetos + 30;
    pagina.sizeHeight = alturaObjetos + 30;
}
