import { CreateStructImportOptions } from './CreateStructImportOptions.js'; // Importing CreateStructImportOptions function if it's defined in a separate file

export function ImportarClipBoard() {
    // Grabado el 05/03/2024
    let impopt = CreateStructImportOptions();
    impopt.MaintainLayers = true;
    impopt.ColorConversionOptions = {
        SourceColorProfileList: "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%",
        TargetColorProfileList: "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%"
    };

    let impflt = activeDocument.activeLayer.importEx("C:\\Users\\Design 4\\Downloads\\clipboard.ai", 1283, impopt);
    impflt.finish();
    let s1 = activeShape;
}
