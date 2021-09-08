import fs from 'fs';
import AdmZip from 'adm-zip';
export function makeWorkingFile(targetExcel, workingDir) {
    if (fs.existsSync(workingDir))
        fs.rmdirSync(workingDir, { recursive: true }); //remove old files that have been parced at the same location.
    fs.mkdirSync(workingDir);
    const zip = new AdmZip(targetExcel);
    zip.extractAllTo(`${workingDir}/workingTemp`, true); //unzip excel template file to dump folder so that we can access xml files.
}
//# sourceMappingURL=makeWorkingFile.js.map