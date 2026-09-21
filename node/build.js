/** Convenience: run extract.js then translate.js for one run_dir. */
const { spawnSync } = require("child_process");
const path = require("path");

const runDir = process.argv[2];
if (!runDir) {
  console.error("Usage: node build.js <run_dir>");
  process.exit(1);
}

for (const script of ["extract.js", "translate.js"]) {
  const result = spawnSync("node", [path.join(__dirname, script), runDir], { stdio: "inherit" });
  if (result.status !== 0) process.exit(result.status || 1);
}
