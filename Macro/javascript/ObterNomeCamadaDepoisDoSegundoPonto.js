// Função para obter o nome da camada após o segundo ponto
export function ObterNomeCamadaDepoisDoSegundoPonto(nomeCamada) {
    let partes = nomeCamada.split(".");
    if (partes.length > 2) {
        return partes[2];
    }
    return nomeCamada;
}
