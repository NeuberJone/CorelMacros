export function NomeEstaNaLista(nome) {
    let nomeMaiusculo = nome.toUpperCase();
    let nomesPermitidos = ["PP", "P", "M", "G", "GG", "XGG", "XXGG", "PLPP", "BLP", "BLM", "BLG", "BLGG",
                           "BLXGG", "BLXXGG", "2A", "4A", "6A", "8A", "12A", "X", "TAG"];

    for (let i = 0; i < nomesPermitidos.length; i++) {
        if (nomeMaiusculo === nomesPermitidos[i]) {
            return true;
        }
    }
    return false;
}
