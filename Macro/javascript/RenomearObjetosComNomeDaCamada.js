export function RenomearObjetosComNomeDaCamada() {
    let doc = app.activeDocument;
    let pagina = doc.activePage;
    
    for (let i = 0; i < pagina.layers.length; i++) {
        let camada = pagina.layers[i];
        for (let j = 0; j < camada.shapes.length; j++) {
            let objeto = camada.shapes[j];
            objeto.name = camada.name;
        }
    }
}
