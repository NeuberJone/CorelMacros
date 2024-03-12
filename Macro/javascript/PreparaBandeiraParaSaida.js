import { PreparaMaterialParaSaida } from './PreparaMaterialParaSaida.js';
import { RenomeiaPaginaAtual } from './RenomeiaPaginaAtual.js';
import { RenomearDocumentoAtual } from '.RenomearDocumentoAtual.js';
import { SalvaArquivoMaterial } from './SalvaArquivoMaterial.js';

export function PreparaBandeiraParaSaida() {
    let NomeDocumento = "Bandeira";
    PreparaMaterialParaSaida();
    RenomeiaPaginaAtual(NomeDocumento);
    RenomearDocumentoAtual(NomeDocumento);
    SalvaArquivoMaterial(NomeDocumento);
}
