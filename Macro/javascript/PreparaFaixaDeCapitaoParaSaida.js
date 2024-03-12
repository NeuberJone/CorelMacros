import { PreparaMaterialParaSaida } from './PreparaMaterialParaSaida.js';
import { RenomeiaPaginaAtual } from './RenomeiaPaginaAtual.js';
import { RenomearDocumentoAtual } from '.RenomearDocumentoAtual.js';
import { SalvaArquivoMaterial } from './SalvaArquivoMaterial.js';

export function PreparaFaixaDeCapitaoParaSaida() {
    let NomeDocumento = "Faixa_De_Capitão";
    PreparaMaterialParaSaida();
    RenomeiaPaginaAtual(NomeDocumento);
    RenomearDocumentoAtual(NomeDocumento);
    SalvaArquivoMaterial(NomeDocumento);
}
