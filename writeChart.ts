import AdmZip from 'adm-zip';

import { workbookChartDetails, worksheetObj } from 'readCharts'

export function writeCharts(targetExcel: workbookChartDetails, printPath: string) {
    const targetDir = targetExcel.tempDir

    const zip = new AdmZip();
    zip.addLocalFolder(targetDir, '')
    zip.writeZip(printPath);
}