/**
 * Production x2t worker (plain JS).
 *
 * Next static export copies utils/editor/x2t.worker.ts as
 * /_next/static/media/x2t.worker.*.ts (MIME linguist + nosniff),
 * so Chrome never starts that Worker. This file is real JS.
 */
/* eslint-disable */

var AVS_FILE_DOCUMENT_DOC = 0x0042;
var AVS_FILE_CROSSPLATFORM_PDFA = 0x0209;
var BASE_URL = self.location.origin + "/x2t-1/";

var x2t = null;
var initPromise = null;
var xmlPath = "/working/params.xml";

function initX2t() {
  if (x2t) return Promise.resolve();
  Object.assign(self, { __filename: BASE_URL });
  importScripts(BASE_URL + "x2t.js");
  x2t = self.Module;
  return new Promise(function (resolve, reject) {
    if (!x2t) {
      reject(new Error("x2t Module did not load"));
      return;
    }
    if (x2t.calledRun) {
      afterRuntime();
      resolve();
      return;
    }
    var previous = x2t.onRuntimeInitialized;
    var settled = false;
    function done() {
      if (settled) return;
      settled = true;
      afterRuntime();
      resolve();
    }
    x2t.onRuntimeInitialized = function () {
      if (typeof previous === "function") previous();
      done();
    };
    if (x2t.calledRun) {
      done();
      return;
    }
    var started = Date.now();
    var timer = setInterval(function () {
      if (x2t.calledRun) {
        clearInterval(timer);
        done();
      } else if (Date.now() - started > 30000) {
        clearInterval(timer);
        reject(new Error("x2t worker runtime init timed out"));
      }
    }, 50);
  });
}

function afterRuntime() {
  try {
    x2t.FS.mkdir("/working");
    x2t.FS.mkdir("/working/media");
    x2t.FS.mkdir("/working/fonts");
    x2t.FS.mkdir("/working/themes");
  } catch (err) {
    /* directories may already exist */
  }
}

function ensureInit() {
  if (!initPromise) initPromise = initX2t();
  return initPromise;
}

ensureInit().catch(function (err) {
  console.error("[x2t-worker] Auto-init failed:", err);
});

function cleanMedia() {
  try {
    var mediaFiles = x2t.FS.readdir("/working/media/");
    for (var i = 0; i < mediaFiles.length; i++) {
      var file = mediaFiles[i];
      if (file !== "." && file !== "..") {
        x2t.FS.unlink("/working/media/" + file);
      }
    }
  } catch (err) {
    console.error(err);
  }
}

function cleanupFiles(files) {
  for (var i = 0; i < files.length; i++) {
    try {
      x2t.FS.unlink(files[i]);
    } catch (err) {
      /* missing is fine */
    }
  }
  cleanMedia();
}

function readMedia() {
  var media = {};
  try {
    var files = x2t.FS.readdir("/working/media/");
    for (var i = 0; i < files.length; i++) {
      var file = files[i];
      if (file !== "." && file !== "..") {
        media[file] = x2t.FS.readFile("/working/media/" + file, {
          encoding: "binary",
        });
      }
    }
  } catch (e) {
    console.error(e);
  }
  return media;
}

function writeInputs(opts) {
  var params = {
    m_sFileFrom: opts.fileFrom,
    m_sThemeDir: "/working/themes",
    m_sFileTo: opts.fileTo,
    m_nFormatFrom: opts.formatFrom,
    m_nFormatTo: opts.formatTo,
    m_bIsPDFA: opts.formatTo === AVS_FILE_CROSSPLATFORM_PDFA,
    m_bIsNoBase64: false,
    m_sFontDir: "/working/fonts/",
  };
  var content = "";
  for (var key in params) {
    if (params[key]) content += "<" + key + ">" + params[key] + "</" + key + ">\n";
  }
  var xml =
    '<?xml version="1.0" encoding="utf-8"?>\n' +
    "<TaskQueueDataConvert\n" +
    '  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"\n' +
    '  xmlns:xsd="http://www.w3.org/2001/XMLSchema"\n' +
    ">\n" +
    content +
    "</TaskQueueDataConvert>";
  x2t.FS.writeFile(xmlPath, xml);
  if (opts.data) {
    x2t.FS.writeFile(opts.fileFrom, new Uint8Array(opts.data));
  }
  if (opts.media) {
    cleanMedia();
    for (var name in opts.media) {
      try {
        x2t.FS.writeFile("/working/" + name, opts.media[name]);
      } catch (err) {
        console.error(name, err);
      }
    }
  }
}

function convert(payload) {
  var fromPath = "/working/" + payload.fileFrom;
  var toPath = "/working/" + payload.fileTo;
  var files = [fromPath, toPath, xmlPath];

  writeInputs({
    fileFrom: fromPath,
    fileTo: toPath,
    formatFrom: payload.formatFrom,
    formatTo: payload.formatTo,
    data: payload.data,
    media: payload.media,
  });

  if (
    String(payload.fileFrom).endsWith(".doc") ||
    payload.formatFrom == AVS_FILE_DOCUMENT_DOC
  ) {
    var viaPath = fromPath + ".docx";
    writeInputs({ fileFrom: fromPath, fileTo: viaPath, data: null });
    x2t.ccall("main1", ["number"], ["string"], [xmlPath]);
    writeInputs({ fileFrom: viaPath, fileTo: toPath, data: null });
    files.push(viaPath);
  }

  try {
    var pathInfo = x2t.FS.analyzePath(toPath);
    if (pathInfo.exists) x2t.FS.unlink(toPath);
  } catch (err) {
    /* ignore */
  }

  try {
    x2t.ccall("main1", ["number"], ["string"], [xmlPath]);
  } catch (e) {
    console.error("ccall", e);
  }

  var output = null;
  try {
    output = x2t.FS.readFile(toPath);
  } catch (e) {
    console.error(e);
  }
  var outputMedia = readMedia();
  setTimeout(function () {
    cleanupFiles(files);
  });
  return { output: output, media: outputMedia };
}

self.onmessage = function (event) {
  var id = event.data && event.data.id;
  var type = event.data && event.data.type;
  var payload = event.data && event.data.payload;
  if (type !== "convert") {
    self.postMessage({ id: id, type: "error", error: "Unknown message type: " + type });
    return;
  }
  ensureInit()
    .then(function () {
      var result = convert(payload);
      var transferables = [];
      if (result.output) transferables.push(result.output.buffer);
      Object.keys(result.media).forEach(function (key) {
        transferables.push(result.media[key].buffer);
      });
      self.postMessage({ id: id, type: "convert:done", payload: result }, transferables);
    })
    .catch(function (error) {
      self.postMessage({
        id: id,
        type: "error",
        error: error instanceof Error ? error.message : String(error),
      });
    });
};

self.postMessage({ type: "ready" });
