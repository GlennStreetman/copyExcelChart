import AdmZip from 'adm-zip';

import { workbookChartDetails, worksheetObj } from 'readCharts'

export async function writeCharts(targetExcel: workbookChartDetails, printPath: string) {
    const targetDir = targetExcel.tempDir

    const zip = new AdmZip();
    zip.addLocalFolder(targetDir, '')
    zip.writeZip(printPath);

    return true
}