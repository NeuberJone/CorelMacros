import { CentralizaObjetosNaPagina } from './CentralizaObjetosNaPagina.js';

export function PreparaMaterialParaSaida() {
    let doc = app.activeDocument;
    let pagina = doc.activePage;
    let larguraObjetos, alturaObjetos;
    
    let selecao = app.activeSelection;
    
    if (selecao.shapes.count > 0) {
        larguraObjetos = selecao.sizeWidth;
        alturaObjetos = selecao.sizeHeight;
        
        pagina.sizeWidth = larguraObjetos;
        pagina.sizeHeight = alturaObjetos;
        CentralizaObjetosNaPagina();
    } else {
        alert("Nenhum objeto selecionado.");
    }
}
