const express = require("express");
const multer = require("multer");
const fs = require("node:fs/promises");
const path = require("node:path");
const os = require("node:os");
const { createReadStream } = require("node:fs");
const yauzl = require("yauzl");
const yazl = require("yazl");

const app = express();

const upload = multer({
  dest: os.tmpdir(),
  limits: {
    fileSize: 100 * 1024 * 1024,
  },
});

app.get("/health", (_req, res) => {
  res.status(200).send("ok");
});

app.post("/patch-pptx-transitions", upload.single("file"), async (req, res) => {
  let workdir = null;

  try {
    if (!req.file) {
      return res.status(400).json({ error: "Missing PPTX file in field 'file'" });
    }

    const slideDurationSeconds = Number(req.body.slideDuration ?? 5);
    const outputName = String(req.body.outputName ?? "patched-slides.pptx");

    if (!Number.isFinite(slideDurationSeconds) || slideDurationSeconds <= 0) {
      return res.status(400).json({ error: "slideDuration must be > 0" });
    }

    const advanceMs = Math.round(slideDurationSeconds * 1000);

    workdir = await fs.mkdtemp(path.join(os.tmpdir(), "pptx-patch-"));
    const inputPath = path.join(workdir, "input.pptx");
    const unzipDir = path.join(workdir, "unzipped");
    const outputPath = path.join(workdir, outputName);

    await fs.mkdir(unzipDir, { recursive: true });
    await fs.copyFile(req.file.path, inputPath);

    await unzipPptx(inputPath, unzipDir);
    await patchAllSlides(unzipDir, advanceMs);
    await zipDirectory(unzipDir, outputPath);

    const stat = await fs.stat(outputPath);
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${outputName}"`);
    res.setHeader("Content-Length", String(stat.size));

    const stream = createReadStream(outputPath);

    stream.on("error", (err) => {
      console.error("Stream error:", err);
      res.destroy(err);
    });

    res.on("close", async () => {
      if (req.file?.path) {
        await fs.rm(req.file.path, { force: true }).catch(() => {});
      }
      if (workdir) {
        await fs.rm(workdir, { recursive: true, force: true }).catch(() => {});
      }
    });

    return stream.pipe(res);
  } catch (err) {
    console.error(err);
    if (req.file?.path) {
      await fs.rm(req.file.path, { force: true }).catch(() => {});
    }
    if (workdir) {
      await fs.rm(workdir, { recursive: true, force: true }).catch(() => {});
    }
    return res.status(500).json({
      success: false,
      error: err?.message || "Failed to patch PPTX",
    });
  }
});

async function patchAllSlides(unzipDir, advanceMs) {
  const slidesDir = path.join(unzipDir, "ppt", "slides");

  let files = await fs.readdir(slidesDir);
  files = files
    .filter((name) => /^slide\d+\.xml$/i.test(name))
    .sort((a, b) => {
      const na = Number(a.match(/\d+/)?.[0] || 0);
      const nb = Number(b.match(/\d+/)?.[0] || 0);
      return na - nb;
    });

  if (!files.length) {
    throw new Error("No slide XML files found in PPTX");
  }

  for (const file of files) {
    const fullPath = path.join(slidesDir, file);
    let xml = await fs.readFile(fullPath, "utf8");
    xml = upsertFadeTransition(xml, advanceMs);
    await fs.writeFile(fullPath, xml, "utf8");
  }
}

function upsertFadeTransition(xml, advanceMs) {
  const transitionXml = `<p:transition advTm="${advanceMs}"><p:fade/></p:transition>`;

  // Remove existing transition block if present
  xml = xml.replace(/<p:transition\b[\s\S]*?<\/p:transition>/g, "");
  xml = xml.replace(/<p:transition\b[^>]*\/>/g, "");

  // Insert transition immediately after <p:cSld ...>...</p:cSld>
  if (/<p:cSld\b[\s\S]*?<\/p:cSld>/.test(xml)) {
    return xml.replace(
      /(<p:cSld\b[\s\S]*?<\/p:cSld>)/,
      `$1${transitionXml}`
    );
  }

  // Fallback: insert before closing </p:sld>
  if (/<\/p:sld>/.test(xml)) {
    return xml.replace(/<\/p:sld>/, `${transitionXml}</p:sld>`);
  }

  throw new Error("Unexpected slide XML structure");
}

function unzipPptx(zipPath, destDir) {
  return new Promise((resolve, reject) => {
    yauzl.open(zipPath, { lazyEntries: true }, (err, zipFile) => {
      if (err) return reject(err);

      zipFile.readEntry();

      zipFile.on("entry", async (entry) => {
        try {
          const outPath = path.join(destDir, entry.fileName);

          if (/\/$/.test(entry.fileName)) {
            await fs.mkdir(outPath, { recursive: true });
            zipFile.readEntry();
            return;
          }

          await fs.mkdir(path.dirname(outPath), { recursive: true });

          zipFile.openReadStream(entry, (streamErr, readStream) => {
            if (streamErr) return reject(streamErr);

            const chunks = [];
            readStream.on("data", (chunk) => chunks.push(chunk));
            readStream.on("end", async () => {
              try {
                await fs.writeFile(outPath, Buffer.concat(chunks));
                zipFile.readEntry();
              } catch (writeErr) {
                reject(writeErr);
              }
            });
            readStream.on("error", reject);
          });
        } catch (e) {
          reject(e);
        }
      });

      zipFile.on("end", resolve);
      zipFile.on("error", reject);
    });
  });
}

async function zipDirectory(sourceDir, outZipPath) {
  const zipfile = new yazl.ZipFile();
  const files = await walkFiles(sourceDir);

  for (const file of files) {
    const relPath = path.relative(sourceDir, file).split(path.sep).join("/");
    zipfile.addFile(file, relPath);
  }

  await new Promise((resolve, reject) => {
    zipfile.outputStream
      .pipe(require("node:fs").createWriteStream(outZipPath))
      .on("close", resolve)
      .on("error", reject);
    zipfile.end();
  });
}

async function walkFiles(dir) {
  const entries = await fs.readdir(dir, { withFileTypes: true });
  const files = [];

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      files.push(...(await walkFiles(fullPath)));
    } else if (entry.isFile()) {
      files.push(fullPath);
    }
  }

  return files;
}

const port = process.env.PORT || 10000;
app.listen(port, "0.0.0.0", () => {
  console.log(`PPTX patcher running on port ${port}`);
});