export function RenomearCamadas() {
    let doc = app.activeDocument;
    let pagina = doc.activePage;

    for (let i = 0; i < pagina.layers.length; i++) {
        let camada = pagina.layers[i];
        let partes = camada.name.split(".");

        if (partes.length >= 3) {
            camada.name = partes[0] + "." + partes[1] + "." + partes[2] + "." + partes[1];
        }
    }
}
