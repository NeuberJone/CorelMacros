import { RenomeiaPaginaAtual } from './RenomearPaginaAtual.js'; // Importing RenomeiaPaginaAtual function if it's defined in a separate file

export function AjustaTamanhoDaPagina() {
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
    
    novoX = larguraPagina / 2;
    novoY = alturaPagina / 2;
    
    pagina.shapes.all.createSelection();
    activeSelection.move(novoX, novoY);
    
    RenomeiaPaginaAtual("DESIGN");
}
