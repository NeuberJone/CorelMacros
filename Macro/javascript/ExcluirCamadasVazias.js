export function ExcluirCamadasVazias() {
    let doc = app.activeDocument;
    let pagina = doc.activePage;
    let camadasExcluidas = 0;

    for (let i = pagina.layers.length - 1; i >= 0; i--) {
        let camada = pagina.layers[i];
        if (!camada.isSpecialLayer && camada.shapes.length === 0) {
            camada.remove();
            camadasExcluidas++;
        }
    }
}
