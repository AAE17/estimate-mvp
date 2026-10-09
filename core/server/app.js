// Core kernel: builds the Express app from the feature list and starts it.
// Core never knows feature details; each feature folder exports { name, mount(app, core), onStart?, exports? }.
const express = require("express");
const path = require("path");
const env = require("../../config/env");
const featureNames = require("../../config/features");

const ROOT_DIR = path.join(__dirname, "..", "..");

function loadFeatures(names) {
  return (names || featureNames).map(function (name) {
    const f = require(path.join(ROOT_DIR, "features", name));
    if (!f || f.name !== name || typeof f.mount !== "function") {
      throw new Error("features/" + name + "/index.js must export { name: \"" + name + "\", mount(app) }");
    }
    return f;
  });
}

function createApp(names) {
  const app = express();
  const features = loadFeatures(names);
  const core = { rootDir: ROOT_DIR, env: env };
  features.forEach(function (f) { f.mount(app, core); });
  app.locals.features = features.map(function (f) { return f.name; });
  return { app: app, features: features };
}

function start(names) {
  const built = createApp(names);
  built.features.forEach(function (f) { if (typeof f.onStart === "function") f.onStart(); });
  built.app.listen(env.PORT, function () {
    console.log("ParaState MVP on " + env.PORT);
  });
  return built;
}

module.exports = { createApp: createApp, start: start, loadFeatures: loadFeatures };
