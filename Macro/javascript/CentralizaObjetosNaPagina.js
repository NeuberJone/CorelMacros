export function CentralizaObjetosNaPagina() {
    let OrigSelection = app.activeSelectionRange;
    OrigSelection.alignAndDistribute(3, 3, 2, 0, false, 2);
}
