export function RenomearPaginaAtual(NomePagina) {
    let doc = app.activeDocument;
    let pagina = doc.activePage;

    pagina.name = NomePagina;
}
