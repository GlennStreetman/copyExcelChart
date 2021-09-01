import {readCharts} from './build/readCharts.js'
import {copyChart} from './build/copyChart.js'
import util from 'util'
import fs from 'fs';

const testrun = async ()=>{
    fs.copyFileSync(`./outputSource2.xlsx`, `./output.xlsx`)

    const source = await readCharts('./test.xlsx', './output', true)
    // console.log(util.inspect(source, false, null, true))
    console.log('source', source)
    const output = await readCharts('./output.xlsx', './output', false)
    console.log('output', util.inspect(output, false, null, true))
    // console.log('output', output)
    // copyChart(
    //     source,
    //     output,
    //     'report-US-TSLA',
    //     'chart3',
    //     'test-EBIT',
    //     {},
    // )
}

testrun()