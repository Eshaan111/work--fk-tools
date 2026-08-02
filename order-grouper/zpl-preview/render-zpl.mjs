import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { init, zplToBase64Async } from "zpl-renderer-js/external";

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const [, , inputArgument, outputArgument] = process.argv;

if (!inputArgument || !outputArgument) {
  console.error("Usage: node render-zpl.mjs <input.zpl> <output.png>");
  process.exit(2);
}

const inputPath = path.resolve(inputArgument);
const outputPath = path.resolve(outputArgument);
const zpl = await fs.readFile(inputPath, "utf8");

const widthDotsMatch = zpl.match(/\^PW(\d+)/i);
const heightDotsMatch = zpl.match(/\^LL(\d+)/i);
const dotsPerMillimeter = 8;
const widthDots = Number(widthDotsMatch?.[1] ?? 812);
const heightDots = Number(heightDotsMatch?.[1] ?? 1218);
const widthMillimeters = widthDots / dotsPerMillimeter;
const heightMillimeters = heightDots / dotsPerMillimeter;

const wasmPath = path.join(
  scriptDirectory,
  "node_modules",
  "zpl-renderer-js",
  "dist",
  "zebrash.wasm",
);
const wasmBytes = await fs.readFile(wasmPath);
await init({ wasmBytes });

const base64 = await zplToBase64Async(
  zpl,
  widthMillimeters,
  heightMillimeters,
  dotsPerMillimeter,
  { grayscaleOutput: true },
);

await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.writeFile(outputPath, Buffer.from(base64, "base64"));
console.log(outputPath);
