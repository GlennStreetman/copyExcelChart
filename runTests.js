import { copyFromOneToManySheets } from "./tests/copyFromOneToManySheets/test.js";
import { copyManyToOne } from "./tests/copyManyToOne/test.js";
import { copyNoOverrides } from "./tests/noOverridesOneToOne/test.js";
import { overrides } from "./tests/overrides/test.js";

copyFromOneToManySheets();
copyManyToOne();
copyNoOverrides();
overrides();
