export function MoverObjetosParaCamada() {
    let doc = app.activeDocument;
    let pagina = doc.activePage;
    let objeto, camada;

    for (let i = 0; i < pagina.shapes.length; i++) {
        objeto = pagina.shapes[i];
        camada = pagina.layers[objeto.name];
        
        if (camada !== null && camada !== undefined) {
            objeto.moveToLayer(camada);
        }
    }
}
