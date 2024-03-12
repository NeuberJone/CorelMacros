// Importar funções auxiliares
import { NomeEstaNaLista } from './NomeEstaNaLista.js'; // Importing RenomeiaPaginaAtual function if it's defined in a separate file
import { ObterNomeCamadaDepoisDoSegundoPonto } from './ObterNomeCamadaDepoisDoSegundoPonto.js'; // Importing RenomeiaPaginaAtual function if it's defined in a separate file

export function RenomearObjetosNoMold() {
    let doc = app.activeDocument;
    let pagina = doc.activePage;

    for (let i = 0; i < pagina.layers.length; i++) {
        let camada = pagina.layers[i];
        
        let grupos = camada.findShapes(null, cdrGroupShape);
        for (let j = 0; j < grupos.length; j++) {
            let grupo = grupos[j];
            if (grupo.shapes.count > 0) {
                let subItens = grupo.shapes;
                for (let k = 0; k < subItens.length; k++) {
                    let subItem = subItens[k];
                    if (!NomeEstaNaLista(subItem.name)) {
                        subItem.name = ObterNomeCamadaDepoisDoSegundoPonto(camada.name);
                    }
                }
            }
        }
    }
}
