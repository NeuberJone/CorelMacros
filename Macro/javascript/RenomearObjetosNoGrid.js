import { ObterNomeCamadaDepoisDoPrimeiroPonto } from './ObterNomeCamadaDepoisDoPrimeiroPonto.js';


export function RenomearObjetosNoGrid() {
    let doc = app.activeDocument;
    let pagina = doc.activePage;

    for (let i = 0; i < pagina.layers.length; i++) {
        let camada = pagina.layers[i];
        let objetos = camada.shapes;

        for (let j = 0; j < objetos.length; j++) {
            let objeto = objetos[j];
            let nomeCamada = ObterNomeCamadaDepoisDoPrimeiroPonto(camada.name);
            objeto.name = nomeCamada;
        }
    }
}
