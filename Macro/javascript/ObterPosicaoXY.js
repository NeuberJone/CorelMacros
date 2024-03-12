export function ObterPosicaoXY() {
    let selecao = app.activeSelectionRange;
    let objeto;

    if (selecao && selecao.length === 1) {
        objeto = selecao[0];
        let posX = objeto.positionX;
        let posY = objeto.positionY;
        alert("Posição X: " + posX + "\nPosição Y: " + posY);
    } else if (!selecao) {
        alert("Nenhum objeto selecionado.");
    } else {
        alert("Selecione apenas um objeto.");
    }
}
