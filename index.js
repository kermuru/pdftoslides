const express = require("express");
const multer = require("multer");
const fs = require("node:fs/promises");
const path = require("node:path");
const os = require("node:os");
const { execFile } = require("node:child_process");
const { promisify } = require("node:util");

const execFileAsync = promisify(execFile);

const app = express();
const upload = multer({
  dest: os.tmpdir(),
  limits: {
    fileSize: 50 * 1024 * 1024,
  },
});

app.get("/health", (_req, res) => {
  res.status(200).send("ok");
});

app.post("/render-pdf-to-mp4", upload.single("file"), async (req, res) => {
  let workdir = null;

  try {
    const slideDuration = Number(req.body.slideDuration ?? 5);
    const fadeDuration = Number(req.body.fadeDuration ?? 1);
    const fps = Number(req.body.fps ?? 30);
    const width = Number(req.body.width ?? 1280);
    const height = Number(req.body.height ?? 720);
    const outputName = String(req.body.outputName ?? "slides-video.mp4");

    if (!req.file) {
      return res.status(400).json({ error: "Missing PDF file in field 'file'" });
    }

    if (!Number.isFinite(slideDuration) || slideDuration <= 0) {
      return res.status(400).json({ error: "slideDuration must be > 0" });
    }

    if (!Number.isFinite(fadeDuration) || fadeDuration <= 0 || fadeDuration >= slideDuration) {
      return res.status(400).json({
        error: "fadeDuration must be > 0 and less than slideDuration",
      });
    }

    workdir = await fs.mkdtemp(path.join(os.tmpdir(), "render-"));

    const pdfPath = path.join(workdir, "input.pdf");
    const outputPath = path.join(workdir, outputName);
    const framesDir = path.join(workdir, "frames");

    await fs.mkdir(framesDir, { recursive: true });
    await fs.copyFile(req.file.path, pdfPath);

    await execFileAsync("pdftoppm", [
      "-png",
      pdfPath,
      path.join(framesDir, "slide"),
    ]);

    const files = (await fs.readdir(framesDir))
      .filter((f) => f.endsWith(".png"))
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

    if (!files.length) {
      throw new Error("No PNG frames generated from PDF");
    }

    const imagePaths = files.map((f) => path.join(framesDir, f));
    const ffmpegArgs = buildFfmpegArgs({
      images: imagePaths,
      outputMp4: outputPath,
      slideDuration,
      fadeDuration,
      fps,
      width,
      height,
    });

    await execFileAsync("ffmpeg", ffmpegArgs, {
      maxBuffer: 50 * 1024 * 1024,
    });

    const mp4Buffer = await fs.readFile(outputPath);

    res.setHeader("Content-Type", "video/mp4");
    res.setHeader("Content-Disposition", `attachment; filename="${outputName}"`);
    res.setHeader("Content-Length", String(mp4Buffer.length));

    return res.status(200).send(mp4Buffer);
  } catch (err) {
    console.error(err);
    return res.status(500).json({
      success: false,
      error: err?.message || "Rendering failed",
    });
  } finally {
    if (req.file?.path) {
      await fs.rm(req.file.path, { force: true }).catch(() => {});
    }
    if (workdir) {
      await fs.rm(workdir, { recursive: true, force: true }).catch(() => {});
    }
  }
});

function buildFfmpegArgs({
  images,
  outputMp4,
  slideDuration,
  fadeDuration,
  fps,
  width,
  height,
}) {
  const args = ["-y"];
  const perInputDuration = slideDuration + fadeDuration;
  const filterParts = [];

  for (const img of images) {
    args.push("-loop", "1", "-t", String(perInputDuration), "-i", img);
  }

  for (let i = 0; i < images.length; i++) {
    filterParts.push(
      `[${i}:v]scale=${width}:${height}:force_original_aspect_ratio=decrease,` +
      `pad=${width}:${height}:(ow-iw)/2:(oh-ih)/2,setsar=1,format=yuv420p[v${i}]`
    );
  }

  if (images.length === 1) {
    filterParts.push(`[v0]copy[outv]`);
  } else {
    for (let i = 1; i < images.length; i++) {
      const left = i === 1 ? `[v0]` : `[x${i - 1}]`;
      const right = `[v${i}]`;
      const out = i === images.length - 1 ? `[outv]` : `[x${i}]`;
      const offset = slideDuration * i;

      filterParts.push(
        `${left}${right}xfade=transition=fade:duration=${fadeDuration}:offset=${offset}${out}`
      );
    }
  }

  args.push(
    "-filter_complex",
    filterParts.join(";"),
    "-map",
    "[outv]",
    "-r",
    String(fps),
    "-c:v",
    "libx264",
    "-pix_fmt",
    "yuv420p",
    outputMp4
  );

  return args;
}

const port = process.env.PORT || 10000;
app.listen(port, "0.0.0.0", () => {
  console.log(`Renderer API running on port ${port}`);
});