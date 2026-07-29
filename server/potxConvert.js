const { unzipSync, zipSync, strFromU8, strToU8 } = require("fflate");

const TEMPLATE_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml";
const PRESENTATION_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";

/**
 * PowerPoint.createPresentation(base64) rejects a genuine .potx's bytes
 * outright ("An internal error has occurred.") — confirmed directly
 * against a real file. Inspecting that file's raw [Content_Types].xml
 * showed why: a .potx declares its /ppt/presentation.xml part as
 * TEMPLATE_TYPE, not PRESENTATION_TYPE like an ordinary deck. That one
 * Override is the only thing distinguishing a template from a deck
 * (docProps/app.xml's <Template> field is unrelated — it names what a
 * document was created FROM, not whether it IS one). Patching just this
 * string and re-zipping is enough; confirmed working end-to-end against
 * the same real file.
 */
function convertPotxToPresentationBytes(potxBuffer) {
  const files = unzipSync(new Uint8Array(potxBuffer));
  const contentTypesXml = strFromU8(files["[Content_Types].xml"]);
  files["[Content_Types].xml"] = strToU8(contentTypesXml.replace(TEMPLATE_TYPE, PRESENTATION_TYPE));
  return zipSync(files);
}

module.exports = { convertPotxToPresentationBytes };
