import {readCharts} from './build/readCharts.js'
import {copyChart} from './build/copyChart.js'
import {writeCharts} from './build/writeChart.js'
import util from 'util'
import fs from 'fs';

const testrun = async ()=>{
    fs.copyFileSync(`./testSource1.xlsx`, `./test.xlsx`)
    // fs.copyFileSync(`./testSource2.xlsx`, `./test.xlsx`)

    const source = await readCharts('./test.xlsx', './output', true)
    // console.log(util.inspect(source, false, null, true))
    console.log('source', source.summary())
    const output = await readCharts('./testOutput.xlsx', './output', false)
    // console.log('output', util.inspect(output, false, null, true))
    // console.log('output', output)
    copyChart(
        source,
        output,
        'sourceWorksheet',
        'chartEx1',
        'outputWorksheet',
        {
            [`'EBIT-US-TSLA'!$B$2:$B$22`]: `'outputWorksheet'!$B$2:$B$16`,
            [`'EBIT-US-TSLA'!$C$1`]: `'outputWorksheet'!$C$1`,
            [`'EBIT-US-TSLA'!$C$2:$C$22`]: `'outputWorksheet'!$C$2:$C$16`,
        },
    )
    writeCharts(output, './product.xlsx')
}

testrun()