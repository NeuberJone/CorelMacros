export function RenomearDocumentoAtual(NomeArquivo) {
    let doc = app.activeDocument;

    if (NomeArquivo === "SISBolt_") {
        if (doc.name.startsWith("Sem título-")) {
            doc.name = "SISBolt_";
        } else {
            doc.name = doc.name.replace(".cdr", "");
        }
    } else {
        doc.name = NomeArquivo;
    }
}
