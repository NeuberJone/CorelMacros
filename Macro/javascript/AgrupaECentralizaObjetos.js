export function AgrupaECentralizaObjetos() {
    app.activePage.shapes.all.createSelection();
    let s1 = app.activeSelection.group();
    s1.alignAndDistribute(3, 3, 2, 0, false, 2);
}
