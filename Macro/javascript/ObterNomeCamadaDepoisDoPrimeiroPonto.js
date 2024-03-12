export function ObterNomeCamadaDepoisDoPrimeiroPonto(nomeCamada) {
    let partes = nomeCamada.split(".");
    if (partes.length >= 2) {
        return partes[1];
    } else {
        return "";
    }
}
