import {
  access,
  mkdir,
  readFile,
  rename,
  rm,
  writeFile,
} from "node:fs/promises";
import path from "node:path";
import sharp from "sharp";

export interface LegacyCropDefinition {
  originalWidth: number;
  originalHeight: number;
  left: number;
  width: number;
  height: number;
  detectionScore: number;
}

export interface CleanedEvidenceResult extends LegacyCropDefinition {
  outputPath: string;
  reused: boolean;
  usable: boolean;
  reason?: string;
}

function pixelOffset(
  x: number,
  y: number,
  width: number,
  channels: number,
): number {
  return (y * width + x) * channels;
}

export async function detectLegacyConversationCrop(
  imagePath: string,
  configuredLeft?: number,
): Promise<LegacyCropDefinition> {
  const image = sharp(imagePath);
  const metadata = await image.metadata();
  if (metadata.format !== "png" || !metadata.width || !metadata.height) {
    throw new Error("Legacy evidence must be a readable PNG image");
  }
  if (configuredLeft !== undefined) {
    if (configuredLeft < 1 || configuredLeft >= metadata.width - 100) {
      throw new Error(
        `LEGACY_EVIDENCE_CROP_LEFT=${configuredLeft} is outside the ${metadata.width}px image`,
      );
    }
    return {
      originalWidth: metadata.width,
      originalHeight: metadata.height,
      left: configuredLeft,
      width: metadata.width - configuredLeft,
      height: metadata.height,
      detectionScore: 0,
    };
  }

  const { data, info } = await sharp(imagePath)
    .removeAlpha()
    .raw()
    .toBuffer({ resolveWithObject: true });
  const startX = Math.floor(info.width * 0.25);
  const endX = Math.floor(info.width * 0.5);
  const startY = Math.floor(info.height * 0.08);
  const endY = Math.floor(info.height * 0.93);
  const sampleCount = endY - startY;
  let best:
    | { x: number; dominantRatio: number; neighborDifference: number; score: number }
    | undefined;

  for (let x = startX; x < endX; x += 1) {
    const colorCounts = new Map<string, number>();
    let neighborDifference = 0;
    for (let y = startY; y < endY; y += 1) {
      const offset = pixelOffset(x, y, info.width, info.channels);
      const colorKey = `${data[offset] >> 3},${data[offset + 1] >> 3},${data[offset + 2] >> 3}`;
      colorCounts.set(colorKey, (colorCounts.get(colorKey) ?? 0) + 1);
      const leftOffset = offset - info.channels;
      const rightOffset = offset + info.channels;
      neighborDifference +=
        (Math.abs(data[leftOffset] - data[rightOffset]) +
          Math.abs(data[leftOffset + 1] - data[rightOffset + 1]) +
          Math.abs(data[leftOffset + 2] - data[rightOffset + 2])) /
        3;
    }
    const dominantRatio = Math.max(...colorCounts.values()) / sampleCount;
    const averageDifference = neighborDifference / sampleCount;
    const score = dominantRatio * 100 + averageDifference;
    if (!best || score > best.score) {
      best = {
        x,
        dominantRatio,
        neighborDifference: averageDifference,
        score,
      };
    }
  }

  if (
    !best ||
    best.dominantRatio < 0.85 ||
    best.neighborDifference < 15
  ) {
    throw new Error(
      "Could not safely detect the WhatsApp sidebar/conversation boundary",
    );
  }
  const left = best.x + 2;
  const width = info.width - left;
  if (width >= info.width || width < info.width * 0.5 || left < info.width * 0.2) {
    throw new Error("Detected legacy evidence crop is outside safe bounds");
  }
  return {
    originalWidth: info.width,
    originalHeight: info.height,
    left,
    width,
    height: info.height,
    detectionScore: Math.round(best.score * 100) / 100,
  };
}

export async function validateLegacyCropStrategy(
  representativePaths: string[],
  configuredLeft?: number,
): Promise<LegacyCropDefinition> {
  if (representativePaths.length < 3) {
    throw new Error("At least three representative evidence images are required");
  }
  const definitions = await Promise.all(
    representativePaths.map((imagePath) =>
      detectLegacyConversationCrop(imagePath, configuredLeft),
    ),
  );
  const first = definitions[0];
  if (
    definitions.some(
      (definition) =>
        definition.originalWidth !== first.originalWidth ||
        definition.originalHeight !== first.originalHeight ||
        Math.abs(definition.left - first.left) > 2,
    )
  ) {
    throw new Error("Representative evidence images do not share a safe crop boundary");
  }
  const sortedLefts = definitions.map(({ left }) => left).sort((a, b) => a - b);
  const left = sortedLefts[Math.floor(sortedLefts.length / 2)];
  return { ...first, left, width: first.originalWidth - left };
}

async function existingFile(filePath: string): Promise<boolean> {
  return access(filePath)
    .then(() => true)
    .catch(() => false);
}

export async function cropLegacyEvidence(
  inputPath: string,
  outputPath: string,
  expectedLeft: number,
  configuredLeft?: number,
): Promise<CleanedEvidenceResult> {
  const definition = await detectLegacyConversationCrop(
    inputPath,
    configuredLeft,
  );
  if (Math.abs(definition.left - expectedLeft) > 2) {
    return {
      ...definition,
      outputPath,
      reused: false,
      usable: false,
      reason: `Detected crop boundary ${definition.left}px differs from validated ${expectedLeft}px`,
    };
  }
  const effective = {
    ...definition,
    left: expectedLeft,
    width: definition.originalWidth - expectedLeft,
  };
  const cleanedBytes = await sharp(inputPath)
    .extract({
      left: effective.left,
      top: 0,
      width: effective.width,
      height: effective.height,
    })
    .png()
    .toBuffer();
  const reused =
    (await existingFile(outputPath)) &&
    (await readFile(outputPath)).equals(cleanedBytes);
  if (!reused) {
    await mkdir(path.dirname(outputPath), { recursive: true });
    const temporaryPath = `${outputPath}.${process.pid}.tmp`;
    try {
      await writeFile(temporaryPath, cleanedBytes);
      await rename(temporaryPath, outputPath);
    } catch (error) {
      await rm(temporaryPath, { force: true }).catch(() => undefined);
      throw error;
    }
  }

  const metadata = await sharp(outputPath).metadata();
  const statistics = await sharp(outputPath).stats();
  const visibleVariation = statistics.channels.some(
    (channel) => channel.stdev > 5,
  );
  const usable =
    metadata.format === "png" &&
    metadata.width === effective.width &&
    metadata.height === effective.height &&
    visibleVariation;
  return {
    ...effective,
    outputPath,
    reused,
    usable,
    reason: usable
      ? undefined
      : "Cleaned image dimensions or visible conversation content are invalid",
  };
}
