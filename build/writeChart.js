import AdmZip from 'adm-zip';
export function writeCharts(targetExcel, printPath) {
    const targetDir = targetExcel.tempDir;
    const zip = new AdmZip();
    zip.addLocalFolder(targetDir, '');
    zip.writeZip(printPath);
}
//# sourceMappingURL=writeChart.js.map