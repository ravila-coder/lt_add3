// Analysis engine: Python (Pyodide) in a Web Worker, so the page never freezes.
// Runs app.py and drives it through bridge.py. Files are processed in this browser
// tab and are never uploaded anywhere.
const PYODIDE = "https://cdn.jsdelivr.net/pyodide/v0.27.7/full/";
importScripts(PYODIDE + "pyodide.js");

// Every file the engine needs (app.py, bridge.py, template, theme, the pinned wheels) is
// embedded in index.html by build.py and handed over in the first message, so the deploy
// is one file and nothing can go missing on the host.
const status = text => self.postMessage({ type: "status", text });
let gotAssets;
const assets = new Promise(r => (gotAssets = r));

const b64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));

async function boot() {
  status("Cargando Python...");
  const py = await loadPyodide({ indexURL: PYODIDE });
  status("Cargando librerías...");
  await py.loadPackage(["sqlite3", "pillow", "markupsafe", "jinja2"]);
  const files = await assets;
  py.FS.mkdirTree("/tmp/wheels");
  py.FS.mkdirTree("/app");
  for (const [name, data] of Object.entries(files)) {
    const dest = name.endsWith(".whl") ? "/tmp/wheels/" + name : "/app/" + name;
    py.FS.writeFile(dest, b64(data));
  }
  py.runPython(`
import os, site, sys, zipfile, importlib
sp = site.getsitepackages()[0]
for w in sorted(os.listdir('/tmp/wheels')):
    zipfile.ZipFile('/tmp/wheels/' + w).extractall(sp)
    os.remove('/tmp/wheels/' + w)
importlib.invalidate_caches()
sys.path.insert(0, '/app')
import flask, openpyxl
`);
  const call = py.pyimport("bridge").call;
  self.postMessage({ type: "ready" });
  return { py, call };
}

const engine = boot();
engine.catch(e => self.postMessage({ type: "fatal", text: String(e && e.message || e) }));

// Python is single threaded: run requests strictly one after another.
let queue = Promise.resolve();
self.onmessage = e => {
  const m = e.data;
  if (m.type === "assets") return gotAssets(m.files);
  queue = queue.then(() => handle(m));
};

async function handle({ id, method, path, body, file }) {
  try {
    const { py, call } = await engine;
    const data = file ? new Uint8Array(file.buf) : null;
    const res = call(method, path, body || null, file ? file.name : null, data);
    const [status, ctype, payload] = res.toJs();
    res.destroy();
    const buf = payload instanceof Uint8Array ? payload : new Uint8Array(payload);
    self.postMessage({ id, status, ctype, buf }, [buf.buffer]);
  } catch (err) {
    const text = String(err && err.message || err);
    // keep the last line of a Python traceback, it names the actual error
    const last = text.trim().split("\n").filter(Boolean).pop() || text;
    self.postMessage({ id, error: last, detail: text });
  }
}
