import sharp from "sharp";
import fs from "fs";
import path from "path";

const srcFile = path.resolve("public/editor_cube_transparent.png");
const outDir = path.resolve("public");

async function main() {
  console.log("Reading source image:", srcFile);
  const metadata = await sharp(srcFile).metadata();
  console.log(`Dimensions: ${metadata.width}x${metadata.height}`);

  const squareBuffer = await sharp(srcFile)
    .resize(512, 512, {
      fit: "contain",
      background: { r: 0, g: 0, b: 0, alpha: 0 },
    })
    .png()
    .toBuffer();

  await sharp(squareBuffer).resize(512, 512).png().toFile(path.join(outDir, "android-chrome-512x512.png"));
  await sharp(squareBuffer).resize(512, 512).png().toFile(path.join(outDir, "icon-512x512.png"));
  await sharp(squareBuffer).resize(512, 512).png().toFile(path.join(outDir, "icon.png"));
  await sharp(squareBuffer).resize(512, 512).png().toFile(path.join(outDir, "logo.png"));

  await sharp(squareBuffer).resize(192, 192).png().toFile(path.join(outDir, "android-chrome-192x192.png"));
  await sharp(squareBuffer).resize(192, 192).png().toFile(path.join(outDir, "icon-192x192.png"));

  await sharp(squareBuffer).resize(180, 180).png().toFile(path.join(outDir, "apple-touch-icon.png"));
  await sharp(squareBuffer).resize(180, 180).png().toFile(path.join(outDir, "apple-touch-icon-precomposed.png"));

  await sharp(squareBuffer).resize(32, 32).png().toFile(path.join(outDir, "favicon-32x32.png"));
  await sharp(squareBuffer).resize(16, 16).png().toFile(path.join(outDir, "favicon-16x16.png"));
  await sharp(squareBuffer).resize(48, 48).png().toFile(path.join(outDir, "favicon.png"));
  await sharp(squareBuffer).resize(32, 32).png().toFile(path.join(outDir, "favicon.ico"));

  console.log("Successfully generated all favicons from editor_cube_transparent.png");
}

main().catch(console.error);
