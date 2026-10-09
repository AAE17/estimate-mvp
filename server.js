// SO+ entry point (Render/Docker run "node server.js").
// The app is built by core/server/app.js from the features listed in config/features.js.
const core = require("./core/server/app");

if (require.main === module) {
  core.start();
}

// Kept for compatibility with the old server.js export.
module.exports = {
  get bakeFormulaResults() { return require("./features/legacy/server").bakeFormulaResults; }
};
