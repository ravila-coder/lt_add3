"""Bundle the app into one self-contained page: public/index.html.

    python source/build.py

Embeds worker.js and every engine file (app.py, bridge.py, the Excel template, the theme
and the pinned wheels) into index.html, so the deploy is a single file. Run it after any
change in source/ and commit the regenerated public/index.html.
"""
import base64
import json
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "public", "index.html")

# file in source/ -> name app.py expects on its own disk
FILES = {
    "py/app.py": "app.py",
    "py/bridge.py": "bridge.py",
    "py/template.xlsx": "Portfolio Analysis.xlsx",
    "py/Theme1.thmx": "Theme1.thmx",
}
FILES.update({f"wheels/{w}": w for w in sorted(os.listdir(os.path.join(HERE, "wheels")))
              if w.endswith(".whl")})


def main():
    page = open(os.path.join(HERE, "index.html"), encoding="utf-8").read()
    worker = open(os.path.join(HERE, "worker.js"), encoding="utf-8").read()
    files = {}
    for src, name in FILES.items():
        data = open(os.path.join(HERE, src), "rb").read()
        if name.endswith((".whl", ".xlsx", ".thmx")):
            zipfile.ZipFile(os.path.join(HERE, src)).testzip()   # fail the build, not the page
        files[name] = base64.b64encode(data).decode("ascii")
    assert "</script" not in worker.lower()
    for marker in ("/*WORKER*/", "/*FILES*/"):
        assert page.count(marker) == 1, marker
    page = page.replace("/*WORKER*/", worker).replace("/*FILES*/", json.dumps(files))
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        f.write(page)
    print(f"public/index.html  {os.path.getsize(OUT) / 1e6:.2f} MB  ({len(files)} embedded files)")


if __name__ == "__main__":
    main()
