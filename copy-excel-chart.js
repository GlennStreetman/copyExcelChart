import {readCharts} from './build/readCharts.js'
import {copyChart} from './build/copyChart.js'
import {writeCharts} from './build/writeChart.js'

const copyExcelChart = {
    readCharts: readCharts,
    copyChart: copyChart,
    writeCharts: copyChart,
}

export default copyExcelChart