export function RemoveCodigosNoNu() {
    let doc = app.activeDocument;
    let pagina = doc.activePage;

    for (let i = 0; i < pagina.shapes.length; i++) {
        let s = pagina.shapes[i];
        
        if (s.name === "NO") {
            s.name = "TextoNome";
        }
        if (s.name === "NU") {
            s.name = "TextoNumero";
        }
    }
}
