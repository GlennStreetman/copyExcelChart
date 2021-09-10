import AdmZip from 'adm-zip';
export function writeCharts(targetExcel, printPath) {
    return new Promise((resolve, reject) => {
        try {
            const targetDir = targetExcel.tempDir;
            const zip = new AdmZip();
            zip.addLocalFolder(targetDir, '');
            zip.writeZip(printPath);
            resolve(true);
        }
        catch (error) {
            console.log('Write chart file error. targetExcel: ', targetExcel, 'Error: ', error);
            reject(error);
        }
    });
}
//# sourceMappingURL=writeChart.js.map