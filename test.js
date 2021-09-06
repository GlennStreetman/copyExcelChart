import {readCharts} from './build/readCharts.js'
import {copyChart} from './build/copyChart.js'
import {writeCharts} from './build/writeChart.js'
import util from 'util'
import fs from 'fs';

const testrun = async ()=>{
    fs.copyFileSync(`./outputSource1.xlsx`, `./output.xlsx`)

    const source = await readCharts('./test.xlsx', './output', true)
    console.log(util.inspect(source, false, null, true))
    // console.log('source', source)
    const output = await readCharts('./output.xlsx', './output', false)
    console.log('output', util.inspect(output, false, null, true))
    // console.log('output', output)
    copyChart(
        source,
        output,
        'report-US-TSLA',
        'chartEx1',
        'test-EBIT',
        {
            [`'EBIT-US-TSLA'!$B$2:$B$22`]: `'test-EBIT'!$B$2:$B$22`,
            [`'EBIT-US-TSLA'!$C$1`]: `'test-EBIT'!$C$1`,
            [`'EBIT-US-TSLA'!$C$2:$C$22`]: `'test-EBIT'!$C$2:$C$22`,
        },
    )
    writeCharts(output, './product.xlsx')
}

testrun()