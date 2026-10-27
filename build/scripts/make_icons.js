// Render Material Design icons (react-icons/md) to PNGs for the deck.
// Usage: node make_icons.js OUT_DIR   (needs react, react-dom, react-icons, sharp on NODE_PATH)
const path = require("path");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const md = require("react-icons/md");

const out = process.argv[2];
const specs = [
  // [icon, color, name]
  ["MdTune", "FFFFFF", "tune_w"],
  ["MdShowChart", "FFFFFF", "chart_w"],
  ["MdCable", "FFFFFF", "cable_w"],
  ["MdBiotech", "FFFFFF", "biotech_w"],
  ["MdSmartToy", "FFFFFF", "robot_w"],
  ["MdHub", "FFFFFF", "hub_w"],
  ["MdAccountTree", "FFFFFF", "tree_w"],
  ["MdSpeed", "FFFFFF", "speed_w"],
  ["MdInsights", "FFFFFF", "insights_w"],
  ["MdPlayCircle", "FFFFFF", "play_w"],
  ["MdArticle", "FFFFFF", "article_w"],
  ["MdLanguage", "FFFFFF", "web_w"],
  ["MdRoute", "FFFFFF", "route_w"],
  ["MdLink", "FFFFFF", "link_w"],
  ["MdHelpOutline", "FFFFFF", "help_w"],
  ["MdDirectionsRun", "FFFFFF", "run_w"],
  ["MdFlightTakeoff", "FFFFFF", "flight_w"],
];
(async () => {
  for (const [icon, color, name] of specs) {
    const svg = ReactDOMServer.renderToStaticMarkup(
      React.createElement(md[icon], { color: "#" + color, size: 256 })
    );
    await sharp(Buffer.from(svg)).resize(256, 256).png().toFile(path.join(out, name + ".png"));
  }
  console.log("icons:", specs.length);
})();
