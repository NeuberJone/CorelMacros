import { ObterNomeDoLayout } from './ObterNomeDoLayout.js';

export function SalvaArquivoMaterial(NomeDocumento) {
    let OrigSelection = app.activeSelectionRange;
    let SaveOptions = new StructSaveAsOptions();
    let NomeArquivo;
    
    let nomeLayout = ObterNomeDoLayout();
    
    NomeArquivo = "Z:\\Neuber\\Brindes\\" + nomeLayout + "_" + NomeDocumento + ".cdr";
    
    SaveOptions.embedVBAProject = false;
    SaveOptions.filter = cdrCDR;
    SaveOptions.includeCMXData = false;
    SaveOptions.range = cdrAllPages;
    SaveOptions.embedICCProfile = false;
    SaveOptions.version = 243;
    SaveOptions.keepAppearance = true;
    
    app.activeDocument.saveAs(NomeArquivo, SaveOptions);
    //app.activeDocument.close();
}
