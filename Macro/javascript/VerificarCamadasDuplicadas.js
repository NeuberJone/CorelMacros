export function VerificarCamadasDuplicadas() {
    let doc = app.activeDocument;
    let pagina = doc.activePage;
    let nomesCamadas = {};
    let duplicadas = false;

    for (let i = 0; i < pagina.layers.length; i++) {
        let camada = pagina.layers[i];
        if (nomesCamadas[camada.name]) {
            alert("A camada " + camada.name + " está duplicada.");
            duplicadas = true;
        } else {
            nomesCamadas[camada.name] = true;
        }
    }

    if (!duplicadas) {
        alert("Tirando o que tá errado, o resto tá certo.");
    }
}
