// Feature manifest: the whole pre-restructure app (old server.js), mounted unchanged.
// It shrinks step by step as code moves into its own feature folders.
const legacy = require("./server");

module.exports = {
  name: "legacy",
  mount: function (app) { app.use(legacy.app); },
  onStart: legacy.onStart,
  exports: { bakeFormulaResults: legacy.bakeFormulaResults }
};
