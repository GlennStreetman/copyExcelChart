import fs from 'fs';
export function cleanup(path) {
    fs.rmdirSync(path, { recursive: true });
}
//# sourceMappingURL=cleanup.js.map