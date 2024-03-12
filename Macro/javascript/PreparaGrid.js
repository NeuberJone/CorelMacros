import { VerificarGerenciadorObjetos } from './VerificarGerenciadorObjetos.js'; // Importing VerificarGerenciadorObjetos function if it's defined in a separate file
import { RenomearObjetosComNomeDaCamada } from './RenomearObjetosComNomeDaCamada.js';
import { AgrupaECentralizaObjetos } from './AgrupaECentralizaObjetos.js';
import { AjustaTamanhoDaPagina } from './AjustaTamanhoDaPagina.js';
import { CentralizaObjetosNaPagina } from './CentralizaObjetosNaPagina.js';
import { DesagrupaObjetos } from './DesagrupaObjetos.js';
import { MoverObjetosParaCamada } from './MoverObjetosParaCamada.js';
import { RenomearObjetosNoGrid } from './RenomearObjetosNoGrid.js';
import { RenomeiaPaginaAtual } from './Correcoes.js';
import { RenomearDocumentoAtual } from './Correcoes.js';
import { ExcluirCamadasVazias } from './Utilidades.js';

export function PreparaGrid() {
    let gerenciadorAberto = VerificarGerenciadorObjetos();

    if (gerenciadorAberto) {
        RenomearObjetosComNomeDaCamada();
        activePage.shapes.all.createSelection();
        AgrupaECentralizaObjetos();
        AjustaTamanhoDaPagina();
        CentralizaObjetosNaPagina();
        DesagrupaObjetos();
        MoverObjetosParaCamada();
        RenomearObjetosNoGrid();
        RenomeiaPaginaAtual("DESIGN");
        RenomearDocumentoAtual("SISBolt_");
        ExcluirCamadasVazias();
    } else {
        alert("Então deu ruim, abre a bagaça e roda denovo");
    }
}
