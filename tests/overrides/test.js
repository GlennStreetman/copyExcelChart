import { readCharts } from "./../../build/readCharts.js";
import { copyChart } from "./../../build/copyChart.js";
import { writeCharts } from "./../../build/writeChart.js";
import util from "util";
import fs from "fs";

export const overrides = async () => {
    //copy charts from one .xlsx worksheet to many .xlsx worksheets.
    console.log("starting test");
    if (!fs.existsSync("./tests/overrides/working")) fs.mkdirSync("./tests/overrides/working");

    const source = await readCharts("./tests/overrides/source.xlsx", "./tests/overrides/working/");
    // console.log(util.inspect(source, false, null, true))
    // console.log('source', source.summary())

    const output = await readCharts("./tests/overrides/target.xlsx", "./tests/overrides/working");
    // console.log(util.inspect(output, false, null, true))
    // console.log('output', output.summary())
    console.log(source.summary());
    const replaceCellRefs = source.summary().Sheet1["chart1"].reduce((acc, el) => {
        return { ...acc, [el]: el.replace("$B$5:$F$10", "$B$2:$F$10") };
    }, {});

    await copyChart(
        source,
        output,
        "Sheet1", //worksheet, in source file, that chart will be copied from
        "chart1", //chart that will be copied
        "Sheet1", //worksheet, in output file, that chart will be copied to
        replaceCellRefs //object containing key value pairs of cell references that will be replaced while chart is being copied.
    );

    await writeCharts(output, "./tests/overrides/product.xlsx");
    // fs.rmdirSync('./tests/copyFromOneToManySheets/working', { recursive: true })
};
