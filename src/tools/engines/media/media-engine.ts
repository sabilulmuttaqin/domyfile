import type { FFmpeg } from "@ffmpeg/ffmpeg";
import { asMediaProcessingError } from "./errors";
import {
  detectMediaSignature,
  formatFromProbe,
  getExpectedOutputCodec,
  getOutputExtension,
  getOutputMime,
  isOutputFormat,
  isSupportedInputProbe,
} from "./formats";
import {
  buildCompressCommand,
  buildConverterCommand,
  buildMergeCommand,
  buildMovToMp4Command,
  buildTrimVideoCommand,
  buildTrimCommand,
  buildVideoConverterCommand,
  buildVideoToMp3Command,
  getMediaOutputName,
  normalizeMediaOptions,
  resolveOutputFormat,
  validateTrimRange,
  validateVideoTrimRange,
} from "./options";
import {
  MediaProcessingError,
  type AudioMediaProbe,
  type AudioOutputFormat,
  type MediaEngine,
  type MediaErrorCode,
  type MediaOptions,
  type MediaProbe,
  type MediaProcessItem,
  type MediaProcessResult,
  type MediaProgressCallback,
  type MediaSignature,
  type MediaStreamProbe,
  type MediaToolSlug,
  type NormalizedMediaOptions,
  type VideoMediaProbe,
} from "./types";
import { mediaToolSlugs } from "./types";

const CORE_BASE_PATH = "/runtime/ffmpeg-core-lgpl-5.1.4";

type FfmpegModule = typeof import("@ffmpeg/ffmpeg");
type FfmpegUtilModule = typeof import("@ffmpeg/util");

type FfprobePayload = {
  streams?: Array<{
    codec_type?: string;
    codec_name?: string;
    codec_tag_string?: string;
    duration?: string | number;
    sample_rate?: string | number;
    channels?: string | number;
    bit_rate?: string | number;
    width?: string | number;
    height?: string | number;
  }>;
  packets?: Array<{
    flags?: string;
    pts_time?: string | number;
    dts_time?: string | number;
  }>;
  format?: {
    format_name?: string;
    duration?: string | number;
    bit_rate?: string | number;
  };
};

type Runtime = {
  ffmpeg: FFmpeg;
  fetchFile: FfmpegUtilModule["fetchFile"];
  activeProgress?: { durationSeconds?: number; callback?: MediaProgressCallback };
};

type ProbeMode = "any" | "audio" | "video" | "video-trim";

type JobInput = { path: string; file: File; data: Uint8Array };

let runtime: Runtime | undefined;
let runtimePromise: Promise<Runtime> | undefined;
let queueTail: Promise<void> = Promise.resolve();
let jobCounter = 0;

const videoTools = new Set<MediaToolSlug>(["trim-video", "video-to-mp3", "mov-to-mp4", "video-converter"]);
const isVideoTool = (tool: MediaToolSlug): tool is "trim-video" | "video-to-mp3" | "mov-to-mp4" | "video-converter" => videoTools.has(tool);

const runExclusive = async <T>(task: () => Promise<T>) => {
  let release!: () => void;
  const turn = new Promise<void>((resolve) => { release = resolve; });
  const previous = queueTail;
  queueTail = queueTail.then(() => turn);
  await previous;
  try {
    return await task();
  } finally {
    release();
  }
};

const throwIfAborted = (signal?: AbortSignal) => {
  if (signal?.aborted) throw new MediaProcessingError("CANCELLED", "Processing was cancelled.");
};

const getRuntimeUrls = () => {
  if (typeof globalThis.location === "undefined") throw new MediaProcessingError("ENGINE_LOAD_FAILED", "Media processing is only available in a browser tab.");
  const base = new URL(CORE_BASE_PATH, globalThis.location.origin).href;
  return {
    coreURL: `${base}/ffmpeg-core.js`,
    wasmURL: `${base}/ffmpeg-core.wasm`,
  };
};

const createRuntime = async (): Promise<Runtime> => {
  try {
    const [{ FFmpeg }, { fetchFile }] = await Promise.all([
      import("@ffmpeg/ffmpeg"),
      import("@ffmpeg/util"),
    ]) as [FfmpegModule, FfmpegUtilModule];
    const ffmpeg = new FFmpeg();
    const urls = getRuntimeUrls();
    const loadedRuntime: Runtime = { ffmpeg, fetchFile };
    ffmpeg.on("progress", ({ progress, time }) => {
      const active = loadedRuntime.activeProgress;
      if (!active?.callback) return;
      const duration = active.durationSeconds;
      const timeSeconds = Number.isFinite(time) && time > 0 ? time / 1_000_000 : undefined;
      const measured = duration && timeSeconds !== undefined ? timeSeconds / duration : progress;
      if (!Number.isFinite(measured)) return;
      active.callback(Math.min(1, Math.max(0, measured)));
    });
    await ffmpeg.load(urls);
    return loadedRuntime;
  } catch (error) {
    throw new MediaProcessingError("ENGINE_LOAD_FAILED", "The media engine could not be loaded in this browser.", error);
  }
};

const getRuntime = async (signal?: AbortSignal) => {
  throwIfAborted(signal);
  if (runtime?.ffmpeg.loaded) return runtime;
  if (!runtimePromise) {
    runtimePromise = createRuntime().then((loaded) => {
      runtime = loaded;
      return loaded;
    }).catch((error) => {
      runtimePromise = undefined;
      runtime = undefined;
      throw error;
    });
  }
  return runtimePromise;
};

const resetRuntime = () => {
  runtime?.ffmpeg.terminate();
  runtime = undefined;
  runtimePromise = undefined;
};

const toText = (data: Uint8Array | string) => typeof data === "string" ? data : new TextDecoder().decode(data);

const clearBytes = (bytes: Uint8Array) => {
  try {
    bytes.fill(0);
  } catch {
    // FFmpeg's worker may transfer the backing buffer; a detached buffer is already unreachable here.
  }
};

const parseNumber = (value: string | number | undefined) => {
  const parsed = typeof value === "number" ? value : Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
};

const getInputBytes = async (file: File, fetchFile: FfmpegUtilModule["fetchFile"]) => {
  try {
    const bytes = new Uint8Array(await fetchFile(file));
    return { bytes, signature: detectMediaSignature(bytes, file.name, file.type) };
  } catch (error) {
    throw new MediaProcessingError("CORRUPT_FILE", `“${file.name || "This file"}” could not be read.`, error);
  }
};

const validateSignatureForMode = (signature: MediaSignature, file: File, mode: ProbeMode) => {
  if (mode !== "video" && mode !== "video-trim") return;
  if (signature.container !== "mp4" && signature.container !== "mov") {
    throw new MediaProcessingError("UNSUPPORTED_CONTAINER", `“${file.name || "This file"}” is not a verified MP4 or MOV container.`);
  }
};

const getProbeJson = async (ffmpeg: FFmpeg, inputPath: string, probePath: string, signal?: AbortSignal) => {
  const exitCode = await ffmpeg.ffprobe([
    "-v", "error",
    "-show_entries", "stream=codec_type,codec_name,codec_tag_string,duration,sample_rate,channels,bit_rate,width,height:format=format_name,duration,bit_rate",
    "-of", "json",
    inputPath,
    "-o", probePath,
  ], -1, { signal });
  let data: Uint8Array | string;
  try {
    data = await ffmpeg.readFile(probePath, "utf8", { signal });
  } catch (error) {
    throw new MediaProcessingError("CORRUPT_FILE", "This media file could not be read by the browser media engine.", error);
  }
  try {
    const payload = JSON.parse(toText(data)) as FfprobePayload;
    if (exitCode !== 0 && !payload.streams?.length) throw new Error("FFprobe returned no media streams");
    return payload;
  } catch (error) {
    throw new MediaProcessingError("CORRUPT_FILE", "This media file did not contain readable stream metadata.", error);
  }
};

const getVideoKeyframeTimestamps = async (ffmpeg: FFmpeg, inputPath: string, probePath: string, signal?: AbortSignal) => {
  const exitCode = await ffmpeg.ffprobe([
    "-v", "error",
    "-select_streams", "v:0",
    "-show_packets",
    "-show_entries", "packet=pts_time,dts_time,flags",
    "-of", "json",
    inputPath,
    "-o", probePath,
  ], -1, { signal });
  let data: Uint8Array | string;
  try {
    data = await ffmpeg.readFile(probePath, "utf8", { signal });
  } catch (error) {
    throw new MediaProcessingError("CORRUPT_FILE", "The browser could not inspect this video's keyframe boundaries.", error);
  }
  try {
    const payload = JSON.parse(toText(data)) as FfprobePayload;
    const timestamps = (payload.packets ?? [])
      .filter((packet) => packet.flags?.includes("K"))
      .map((packet) => parseNumber(packet.pts_time ?? packet.dts_time))
      .filter((timestamp): timestamp is number => timestamp !== undefined);
    if (exitCode !== 0 || !timestamps.length) throw new Error("FFprobe returned no video keyframes");
    return timestamps;
  } catch (error) {
    throw new MediaProcessingError("CORRUPT_FILE", "The browser could not determine this video's keyframe boundaries.", error);
  }
};

const streamFromPayload = (stream: NonNullable<FfprobePayload["streams"]>[number]): MediaStreamProbe | null => {
  if (stream.codec_type !== "audio" && stream.codec_type !== "video") return null;
  if (!stream.codec_name) return null;
  return {
    type: stream.codec_type,
    codec: stream.codec_name,
    durationSeconds: parseNumber(stream.duration),
    sampleRate: parseNumber(stream.sample_rate),
    channels: parseNumber(stream.channels),
    bitRate: parseNumber(stream.bit_rate),
    width: parseNumber(stream.width),
    height: parseNumber(stream.height),
    codecTag: stream.codec_tag_string,
  };
};

const parseProbe = (payload: FfprobePayload, signature: MediaSignature, keyframeTimestamps?: number[]): MediaProbe => {
  const streams = payload.streams ?? [];
  const parsedStreams = streams.map(streamFromPayload);
  const audioStreams = parsedStreams.filter((stream): stream is MediaStreamProbe => stream?.type === "audio");
  const videoStreams = parsedStreams
    .filter((stream): stream is MediaStreamProbe => stream?.type === "video")
    .map((stream, index) => index === 0 && keyframeTimestamps ? { ...stream, keyframeTimestamps } : stream);
  const otherStreamCount = streams.length - audioStreams.length - videoStreams.length;
  const primaryAudio = audioStreams[0];
  const primaryVideo = videoStreams[0];
  const primary = primaryVideo ?? primaryAudio;
  const formatName = payload.format?.format_name ?? "";
  if (!primary?.codec) throw new MediaProcessingError("CORRUPT_FILE", "The browser could not find a readable media stream.");
  const durationSeconds = parseNumber(payload.format?.duration) ?? parseNumber(primary.durationSeconds);
  if (!durationSeconds || durationSeconds <= 0) throw new MediaProcessingError("CORRUPT_FILE", "The browser could not determine this media file's duration.");

  const base = {
    container: formatName,
    formatName,
    codec: primary.codec,
    durationSeconds,
    sampleRate: primaryAudio?.sampleRate,
    channels: primaryAudio?.channels,
    bitRate: primaryAudio?.bitRate ?? parseNumber(payload.format?.bit_rate),
    hasOnlyAudio: audioStreams.length > 0 && videoStreams.length === 0 && otherStreamCount === 0,
    hasAudio: audioStreams.length > 0,
    hasVideo: videoStreams.length > 0,
    audioStreams,
    videoStreams,
    otherStreamCount,
    signature,
  } satisfies Omit<MediaProbe, "kind" | "detectedFormat" | "detectedAudioFormat">;

  if (videoStreams.length > 0) {
    return {
      ...base,
      kind: "video",
      detectedFormat: signature.container === "mp4" || signature.container === "mov" ? signature.container : null,
      detectedAudioFormat: primaryAudio ? formatFromProbe({ container: formatName, formatName, codec: primaryAudio.codec }) ?? undefined : undefined,
    };
  }

  const detectedFormat = formatFromProbe({ container: formatName, formatName, codec: primary.codec });
  if (!detectedFormat) throw new MediaProcessingError("UNSUPPORTED_FORMAT", "This audio container or codec is not part of the verified DoMyFile audio format set.");
  return { ...base, kind: "audio", detectedFormat };
};

const getInputExtension = (file: File, signature: MediaSignature) => {
  if (signature.container !== "unknown") return signature.container;
  const extension = file.name.toLowerCase().split(".").pop() ?? "";
  return extension || "media";
};

const probeJob = async (file: File, signal?: AbortSignal, mode: ProbeMode = "any"): Promise<MediaProbe> => runExclusive(async () => {
  throwIfAborted(signal);
  const loaded = await getRuntime(signal);
  const jobId = ++jobCounter;
  const input = await getInputBytes(file, loaded.fetchFile);
  validateSignatureForMode(input.signature, file, mode);
  const inputPath = `/job-${jobId}-input.${getInputExtension(file, input.signature)}`;
  const probePath = `/job-${jobId}-probe.json`;
  const keyframePath = `/job-${jobId}-keyframes.json`;
  try {
    await loaded.ffmpeg.writeFile(inputPath, input.bytes, { signal });
    const payload = await getProbeJson(loaded.ffmpeg, inputPath, probePath, signal);
    const keyframeTimestamps = mode === "video-trim"
      ? await getVideoKeyframeTimestamps(loaded.ffmpeg, inputPath, keyframePath, signal)
      : undefined;
    return parseProbe(payload, input.signature, keyframeTimestamps);
  } catch (error) {
    const typedError = asMediaProcessingError(error, "CORRUPT_FILE");
    if (typedError.code === "PROCESSING_FAILED") throw new MediaProcessingError("CORRUPT_FILE", `“${file.name || "This file"}” could not be read.`, error);
    throw typedError;
  } finally {
    clearBytes(input.bytes);
    if (loaded.ffmpeg.loaded) {
      await loaded.ffmpeg.deleteFile(inputPath).catch(() => undefined);
      await loaded.ffmpeg.deleteFile(probePath).catch(() => undefined);
      await loaded.ffmpeg.deleteFile(keyframePath).catch(() => undefined);
    }
  }
});

const probeExistingPath = async (loaded: Runtime, outputPath: string, probePath: string, outputBytes: Uint8Array, outputName: string, outputMime: string, signal?: AbortSignal) => {
  const payload = await getProbeJson(loaded.ffmpeg, outputPath, probePath, signal);
  return parseProbe(payload, detectMediaSignature(outputBytes, outputName, outputMime));
};

const assertAudioProbe = (probe: MediaProbe): AudioMediaProbe => {
  if (!isSupportedInputProbe(probe)) {
    if (probe.kind === "video") throw new MediaProcessingError("UNSUPPORTED_FORMAT", "Choose an audio-only file. Video and other media streams are not supported by the audio tools.");
    throw new MediaProcessingError("UNSUPPORTED_FORMAT", "The selected file does not contain a supported audio-only stream.");
  }
  return probe;
};

const validateVideoProbe = (probe: MediaProbe, tool: "trim-video" | "video-to-mp3" | "mov-to-mp4" | "video-converter"): VideoMediaProbe => {
  if (probe.kind !== "video") throw new MediaProcessingError("UNSUPPORTED_STREAM_LAYOUT", "Choose a video file with the required video stream.");
  if (probe.signature.container !== "mp4" && probe.signature.container !== "mov") {
    throw new MediaProcessingError("UNSUPPORTED_CONTAINER", "Only MP4 and MOV container signatures are supported by this video tool.");
  }
  if (tool === "mov-to-mp4" && probe.signature.container !== "mov") {
    throw new MediaProcessingError("UNSUPPORTED_CONTAINER", "MOV to MP4 accepts QuickTime MOV files only.");
  }
  if (probe.otherStreamCount > 0 || probe.videoStreams.length !== 1 || probe.audioStreams.length > 1) {
    throw new MediaProcessingError("UNSUPPORTED_STREAM_LAYOUT", "This file has an unsupported stream layout. Use one video stream and at most one audio stream.");
  }
  const video = probe.videoStreams[0];
  if (!video || video.codec.toLowerCase() !== "h264") {
    throw new MediaProcessingError("INCOMPATIBLE_VIDEO_CODEC", `This tool supports H.264 video only; the selected file uses ${video?.codec ?? "an unknown codec"}.`);
  }
  const audio = probe.audioStreams[0];
  if (tool === "video-to-mp3" && !audio) {
    throw new MediaProcessingError("NO_AUDIO_STREAM", "This video does not contain an audio stream to extract.");
  }
  if (audio && audio.codec.toLowerCase() !== "aac") {
    throw new MediaProcessingError("INCOMPATIBLE_AUDIO_CODEC", `This tool supports AAC audio only; the selected file uses ${audio.codec}.`);
  }
  return probe;
};

const validateAudioOutput = (bytes: Uint8Array, outputName: string, outputFormat: AudioOutputFormat, probe: MediaProbe) => {
  const signature = detectMediaSignature(bytes, outputName, getOutputMime(outputFormat));
  const expectedContainer = outputFormat === "mp3" ? "mp3" : outputFormat === "wav" ? "wav" : outputFormat === "m4a" ? ["mp4", "mov"] : outputFormat;
  const validContainer = Array.isArray(expectedContainer) ? expectedContainer.includes(signature.container) : signature.container === expectedContainer;
  if (!validContainer) throw new MediaProcessingError("PROCESSING_FAILED", "The browser produced an audio file with an unexpected container.");
  if (probe.kind !== "audio" || !probe.hasOnlyAudio || probe.detectedFormat !== outputFormat) {
    throw new MediaProcessingError("PROCESSING_FAILED", "The browser produced an audio file in an unexpected format.");
  }
  if (probe.codec !== getExpectedOutputCodec(outputFormat)) {
    throw new MediaProcessingError("PROCESSING_FAILED", "The browser produced an audio file with an unexpected codec.");
  }
};

const validateVideoToMp3Output = (bytes: Uint8Array, outputName: string, probe: MediaProbe) => {
  const signature = detectMediaSignature(bytes, outputName, "audio/mpeg");
  if (signature.container !== "mp3" || probe.kind !== "audio" || !probe.hasOnlyAudio || probe.detectedFormat !== "mp3" || probe.codec !== "mp3") {
    throw new MediaProcessingError("PROCESSING_FAILED", "The browser produced an invalid MP3 audio output.");
  }
};

const validateMovToMp4Output = (bytes: Uint8Array, outputName: string, inputProbe: VideoMediaProbe, outputProbe: MediaProbe) => {
  const signature = detectMediaSignature(bytes, outputName, "video/mp4");
  const outputVideo = outputProbe.videoStreams[0];
  const inputVideo = inputProbe.videoStreams[0];
  const sameDimensions = inputVideo?.width === undefined || outputVideo?.width === undefined
    ? true
    : inputVideo.width === outputVideo.width && inputVideo.height === outputVideo.height;
  if (signature.container !== "mp4"
    || outputProbe.kind !== "video"
    || outputProbe.otherStreamCount > 0
    || outputProbe.videoStreams.length !== 1
    || outputProbe.audioStreams.length !== inputProbe.audioStreams.length
    || outputVideo?.codec !== "h264"
    || (inputProbe.audioStreams.length === 1 && outputProbe.audioStreams[0]?.codec !== "aac")
    || !sameDimensions) {
    throw new MediaProcessingError("PROCESSING_FAILED", "The browser could not validate the MP4 stream-copy output.");
  }
};

const validateVideoConverterOutput = (bytes: Uint8Array, outputName: string, outputFormat: "mp4" | "mov", inputProbe: VideoMediaProbe, outputProbe: MediaProbe) => {
  const signature = detectMediaSignature(bytes, outputName, outputFormat === "mov" ? "video/quicktime" : "video/mp4");
  const outputVideo = outputProbe.videoStreams[0];
  const inputVideo = inputProbe.videoStreams[0];
  const sameDimensions = inputVideo?.width === undefined || outputVideo?.width === undefined
    ? true
    : inputVideo.width === outputVideo.width && inputVideo.height === outputVideo.height;
  if (signature.container !== outputFormat
    || outputProbe.kind !== "video"
    || outputProbe.otherStreamCount > 0
    || outputProbe.videoStreams.length !== 1
    || outputProbe.audioStreams.length !== inputProbe.audioStreams.length
    || outputVideo?.codec !== "h264"
    || (inputProbe.audioStreams.length === 1 && outputProbe.audioStreams[0]?.codec !== "aac")
    || !sameDimensions) {
    throw new MediaProcessingError("PROCESSING_FAILED", "The browser could not validate the requested video remux output.");
  }
};

const validateTrimVideoOutput = (bytes: Uint8Array, outputName: string, inputProbe: VideoMediaProbe, outputProbe: MediaProbe) => {
  const signature = detectMediaSignature(bytes, outputName, "video/mp4");
  const outputVideo = outputProbe.videoStreams[0];
  const inputVideo = inputProbe.videoStreams[0];
  const sameDimensions = inputVideo?.width === undefined || outputVideo?.width === undefined
    ? true
    : inputVideo.width === outputVideo.width && inputVideo.height === outputVideo.height;
  if (signature.container !== "mp4"
    || outputProbe.kind !== "video"
    || outputProbe.otherStreamCount > 0
    || outputProbe.videoStreams.length !== 1
    || outputProbe.audioStreams.length !== inputProbe.audioStreams.length
    || outputVideo?.codec !== "h264"
    || (inputProbe.audioStreams.length === 1 && outputProbe.audioStreams[0]?.codec !== "aac")
    || !sameDimensions) {
    throw new MediaProcessingError("PROCESSING_FAILED", "The browser could not validate the keyframe-aligned MP4 trim output.");
  }
};

const percentReduction = (before: number, after: number) => Math.max(0, ((before - after) / before) * 100);

const durationTolerance = (duration: number) => Math.max(0.12, Math.min(0.75, duration * 0.05 + (duration <= 3 ? 0.15 : 0)));

const validateDuration = (actual: number, expected: number, message: string) => {
  if (Math.abs(actual - expected) > durationTolerance(expected)) throw new MediaProcessingError("PROCESSING_FAILED", message);
};

const getInputPath = (jobId: number, index: number, format: string) => `/job-${jobId}-input-${index}.${format}`;

const runJob = async <T>(inputs: JobInput[], outputPaths: string[], signal: AbortSignal | undefined, task: (loaded: Runtime, paths: string[]) => Promise<T>, durationSeconds?: number, onProgress?: MediaProgressCallback) => runExclusive(async () => {
  throwIfAborted(signal);
  const loaded = await getRuntime(signal);
  const paths = [...inputs.map((input) => input.path), ...outputPaths];
  loaded.activeProgress = { durationSeconds, callback: onProgress };
  try {
    for (const input of inputs) await loaded.ffmpeg.writeFile(input.path, input.data, { signal });
    return await task(loaded, paths);
  } catch (error) {
    const typedError = asMediaProcessingError(error);
    if (typedError.code === "CANCELLED") {
      resetRuntime();
      throw typedError;
    }
    throw typedError;
  } finally {
    loaded.activeProgress = undefined;
    inputs.forEach((input) => clearBytes(input.data));
    if (loaded.ffmpeg.loaded) {
      for (const path of paths) await loaded.ffmpeg.deleteFile(path).catch(() => undefined);
    }
  }
});

const processSingle = async (file: File, probe: MediaProbe, options: NormalizedMediaOptions, signal?: AbortSignal, onProgress?: MediaProgressCallback): Promise<MediaProcessItem> => {
  const jobId = ++jobCounter;
  const outputFormat = resolveOutputFormat(options.tool, probe, options);
  const outputName = getMediaOutputName(file.name, options.tool, outputFormat);
  const inputPath = getInputPath(jobId, 0, probe.signature.container === "unknown" ? "media" : probe.signature.container);
  const outputPath = `/job-${jobId}-output.${getOutputExtension(outputFormat)}`;
  const probePath = `/job-${jobId}-output-probe.json`;
  const videoProbe = isVideoTool(options.tool) ? validateVideoProbe(probe, options.tool) : undefined;
  const audioProbe = !videoProbe ? assertAudioProbe(probe) : undefined;
  const expectedDuration = options.tool === "trim-audio"
    ? validateTrimRange(options, probe.durationSeconds).duration
    : options.tool === "trim-video" && videoProbe
      ? validateVideoTrimRange(options, videoProbe).duration
      : probe.durationSeconds;
  const bytes = await runExclusive(async () => {
    const loaded = await getRuntime(signal);
    return new Uint8Array(await loaded.fetchFile(file));
  });

  return runJob([{ path: inputPath, file, data: bytes }], [outputPath, probePath], signal, async (loaded) => {
    let command;
    if (options.tool === "video-to-mp3" && videoProbe) {
      command = buildVideoToMp3Command(inputPath, outputPath, options);
    } else if (options.tool === "mov-to-mp4" && videoProbe) {
      command = buildMovToMp4Command(inputPath, outputPath);
    } else if (options.tool === "video-converter" && videoProbe) {
      if (outputFormat !== "mp4" && outputFormat !== "mov") throw new MediaProcessingError("INVALID_INPUT", "Choose MP4 or MOV for this verified video converter.");
      command = buildVideoConverterCommand(inputPath, outputPath, outputFormat);
    } else if (options.tool === "trim-video" && videoProbe) {
      command = buildTrimVideoCommand(inputPath, outputPath, options, videoProbe);
    } else {
      if (!audioProbe || !isOutputFormat(outputFormat)) throw new MediaProcessingError("INVALID_INPUT", "Choose a verified audio output format.");
      command = options.tool === "trim-audio"
        ? buildTrimCommand(inputPath, outputPath, options, audioProbe, outputFormat)
        : options.tool === "compress-audio"
          ? buildCompressCommand(inputPath, outputPath, options, outputFormat)
          : buildConverterCommand(inputPath, outputPath, options, audioProbe, outputFormat);
    }
    const exitCode = await loaded.ffmpeg.exec(command.args, -1, { signal });
    if (exitCode !== 0) throw new MediaProcessingError("PROCESSING_FAILED", "The browser could not create the requested media output.");
    const outputData = await loaded.ffmpeg.readFile(outputPath);
    if (typeof outputData === "string") throw new MediaProcessingError("PROCESSING_FAILED", "The browser returned an unreadable media output.");
    const outputBytes = new Uint8Array(outputData);
    const outputProbe = await probeExistingPath(loaded, outputPath, probePath, outputBytes, outputName, getOutputMime(outputFormat), signal);
    if (options.tool === "video-to-mp3") {
      validateVideoToMp3Output(outputBytes, outputName, outputProbe);
    } else if (options.tool === "mov-to-mp4" && videoProbe) {
      validateMovToMp4Output(outputBytes, outputName, videoProbe, outputProbe);
    } else if (options.tool === "video-converter" && videoProbe) {
      if (outputFormat !== "mp4" && outputFormat !== "mov") throw new MediaProcessingError("INVALID_INPUT", "Choose MP4 or MOV for this verified video converter.");
      validateVideoConverterOutput(outputBytes, outputName, outputFormat, videoProbe, outputProbe);
    } else if (options.tool === "trim-video" && videoProbe) {
      validateTrimVideoOutput(outputBytes, outputName, videoProbe, outputProbe);
    } else if (audioProbe && isOutputFormat(outputFormat)) {
      validateAudioOutput(outputBytes, outputName, outputFormat, outputProbe);
    }
    validateDuration(outputProbe.durationSeconds, expectedDuration, "The generated media duration did not match the requested result.");
    if (options.tool === "compress-audio" && outputBytes.byteLength >= file.size * 0.95) {
      throw new MediaProcessingError("NOT_SMALLER", "This preset did not make the audio meaningfully smaller. Choose a smaller preset or a longer source file.");
    }
    const item: MediaProcessItem = {
      input: file,
      output: new File([outputBytes], outputName, { type: getOutputMime(outputFormat), lastModified: Date.now() }),
      format: outputFormat,
      mime: getOutputMime(outputFormat),
      codec: outputProbe.codec,
      strategy: command.strategy,
      inputBytes: file.size,
      outputBytes: outputBytes.byteLength,
      inputDurationSeconds: probe.durationSeconds,
      outputDurationSeconds: outputProbe.durationSeconds,
      inputAudioStreams: probe.audioStreams.length,
      outputAudioStreams: outputProbe.audioStreams.length,
      inputVideoStreams: probe.videoStreams.length,
      outputVideoStreams: outputProbe.videoStreams.length,
      videoWidth: outputProbe.videoStreams[0]?.width,
      videoHeight: outputProbe.videoStreams[0]?.height,
    };
    if (options.tool === "compress-audio") item.sizeReductionPercent = percentReduction(file.size, outputBytes.byteLength);
    return item;
  }, expectedDuration, onProgress);
};

const processMerge = async (files: File[], probes: AudioMediaProbe[], options: NormalizedMediaOptions, signal?: AbortSignal, onProgress?: MediaProgressCallback): Promise<MediaProcessItem> => {
  const jobId = ++jobCounter;
  const outputFormat = options.outputFormat && isOutputFormat(options.outputFormat) ? options.outputFormat : "mp3";
  const outputName = getMediaOutputName(files[0].name, "merge-audio", outputFormat);
  const inputPaths = probes.map((probe, index) => getInputPath(jobId, index, probe.signature.container === "unknown" ? probe.detectedFormat : probe.signature.container));
  const outputPath = `/job-${jobId}-output.${getOutputExtension(outputFormat)}`;
  const listPath = `/job-${jobId}-concat.txt`;
  const probePath = `/job-${jobId}-output-probe.json`;
  const inputBytes = await runExclusive(async () => {
    const loaded = await getRuntime(signal);
    return Promise.all(files.map(async (file) => new Uint8Array(await loaded.fetchFile(file))));
  });
  const totalDuration = probes.reduce((total, probe) => total + probe.durationSeconds, 0);
  return runJob(probes.map((_, index) => ({ path: inputPaths[index], file: files[index], data: inputBytes[index] })), [outputPath, listPath, probePath], signal, async (loaded) => {
    const command = buildMergeCommand(inputPaths, outputPath, listPath, options, probes, outputFormat);
    for (const supportFile of command.supportFiles ?? []) await loaded.ffmpeg.writeFile(supportFile.path, supportFile.contents);
    const exitCode = await loaded.ffmpeg.exec(command.args, -1, { signal });
    if (exitCode !== 0) throw new MediaProcessingError("PROCESSING_FAILED", "The browser could not join these audio files into one continuous track.");
    const outputData = await loaded.ffmpeg.readFile(outputPath);
    if (typeof outputData === "string") throw new MediaProcessingError("PROCESSING_FAILED", "The browser returned an unreadable merged audio file.");
    const outputBytes = new Uint8Array(outputData);
    const outputNameMime = getOutputMime(outputFormat);
    const outputProbe = await probeExistingPath(loaded, outputPath, probePath, outputBytes, outputName, outputNameMime, signal);
    validateAudioOutput(outputBytes, outputName, outputFormat, outputProbe);
    validateDuration(outputProbe.durationSeconds, totalDuration, "The merged audio duration did not match the selected tracks.");
    return {
      input: files[0],
      output: new File([outputBytes], outputName, { type: outputNameMime, lastModified: Date.now() }),
      format: outputFormat,
      mime: outputNameMime,
      codec: outputProbe.codec,
      strategy: command.strategy,
      inputBytes: files.reduce((total, file) => total + file.size, 0),
      outputBytes: outputBytes.byteLength,
      inputDurationSeconds: totalDuration,
      outputDurationSeconds: outputProbe.durationSeconds,
      inputAudioStreams: probes.reduce((total, probe) => total + probe.audioStreams.length, 0),
      outputAudioStreams: outputProbe.audioStreams.length,
      inputVideoStreams: 0,
      outputVideoStreams: outputProbe.videoStreams.length,
    } satisfies MediaProcessItem;
  }, totalDuration, onProgress);
};

const validateToolOptions = (tool: MediaToolSlug, options: NormalizedMediaOptions, probes: MediaProbe[]) => {
  if (tool === "trim-audio") validateTrimRange(options, assertAudioProbe(probes[0]).durationSeconds);
  if (tool === "merge-audio" && probes.length < 2) throw new MediaProcessingError("INVALID_INPUT", "Choose at least two audio files to merge.");
  if (tool === "compress-audio" && options.outputFormat && (!isOutputFormat(options.outputFormat) || !["mp3", "m4a"].includes(options.outputFormat))) {
    throw new MediaProcessingError("INVALID_INPUT", "Choose MP3 or M4A for compressed audio output.");
  }
  if (tool === "trim-video" || tool === "video-to-mp3" || tool === "mov-to-mp4" || tool === "video-converter") {
    probes.forEach((probe) => validateVideoProbe(probe, tool));
    if (tool === "trim-video") validateVideoTrimRange(options, validateVideoProbe(probes[0], tool));
    if (tool === "video-to-mp3" && options.outputFormat && options.outputFormat !== "mp3") {
      throw new MediaProcessingError("INVALID_INPUT", "This tool produces MP3 output only.");
    }
    if ((tool === "mov-to-mp4" || tool === "trim-video") && options.outputFormat && options.outputFormat !== "mp4") {
      throw new MediaProcessingError("INVALID_INPUT", "This tool produces MP4 output only.");
    }
    if (tool === "video-converter" && options.outputFormat && options.outputFormat !== "mp4" && options.outputFormat !== "mov") {
      throw new MediaProcessingError("INVALID_INPUT", "This converter supports MP4 and MOV output only.");
    }
  }
};

const modeForTool = (tool: MediaToolSlug): ProbeMode => tool === "trim-video" ? "video-trim" : isVideoTool(tool) ? "video" : "audio";

export class BrowserMediaEngine implements MediaEngine {
  async inspect(file: File, signal?: AbortSignal, tool?: MediaToolSlug) {
    return probeJob(file, signal, tool ? modeForTool(tool) : "any");
  }

  async validate(files: File[], options: MediaOptions): Promise<{ valid: boolean; issues: Array<{ file: File; code: MediaErrorCode; message: string }> }> {
    const issues: Array<{ file: File; code: MediaErrorCode; message: string }> = [];
    if (!files.length) return { valid: false, issues: [] };
    let normalized: NormalizedMediaOptions;
    try {
      normalized = normalizeMediaOptions(options.tool, options);
    } catch (error) {
      const typedError = asMediaProcessingError(error, "INVALID_INPUT");
      return { valid: false, issues: files.map((file) => ({ file, code: typedError.code, message: typedError.userMessage })) };
    }
    if (normalized.tool === "merge-audio" && files.length < 2) {
      issues.push({ file: files[0], code: "INVALID_INPUT", message: "Choose at least two audio files to merge." });
      return { valid: false, issues };
    }
    const probes: Array<MediaProbe | undefined> = [];
    for (const file of files) {
      try {
        const probe = await probeJob(file, undefined, modeForTool(normalized.tool));
        probes.push(probe);
        if (normalized.tool.startsWith("audio") || normalized.tool.endsWith("audio")) assertAudioProbe(probe);
      } catch (error) {
        const typedError = asMediaProcessingError(error);
        probes.push(undefined);
        issues.push({ file, code: typedError.code, message: typedError.userMessage });
      }
    }
    const validProbes = probes.filter((probe): probe is MediaProbe => Boolean(probe));
    if (issues.length || !validProbes.length) return { valid: false, issues };
    try {
      validateToolOptions(normalized.tool, normalized, validProbes);
    } catch (error) {
      const typedError = asMediaProcessingError(error, "INVALID_INPUT");
      return { valid: false, issues: files.map((file) => ({ file, code: typedError.code, message: typedError.userMessage })) };
    }
    return { valid: true, issues: [] };
  }

  async process(files: File[], options: MediaOptions, signal?: AbortSignal, onProgress?: MediaProgressCallback): Promise<MediaProcessResult> {
    if (!files.length) throw new MediaProcessingError("INVALID_INPUT", "Choose a supported media file to continue.");
    const normalized = normalizeMediaOptions(options.tool, options);
    const result: MediaProcessResult = { items: [], failures: [] };
    if (normalized.tool === "merge-audio" && files.length < 2) {
      return { items: [], failures: [{ input: files[0], error: new MediaProcessingError("INVALID_INPUT", "Choose at least two audio files to merge.") }] };
    }
    const probes: Array<MediaProbe | undefined> = [];
    for (const file of files) {
      try {
        const probe = await probeJob(file, signal, modeForTool(normalized.tool));
        probes.push(probe);
      } catch (error) {
        const typedError = asMediaProcessingError(error);
        if (typedError.code === "CANCELLED") throw typedError;
        probes.push(undefined);
        result.failures.push({ input: file, error: typedError });
      }
    }
    const validEntries = files.map((file, index) => ({ file, probe: probes[index] })).filter((entry): entry is { file: File; probe: MediaProbe } => Boolean(entry.probe));
    if (!validEntries.length) return result;
    try {
      validateToolOptions(normalized.tool, normalized, validEntries.map((entry) => entry.probe));
    } catch (error) {
      const typedError = asMediaProcessingError(error, "INVALID_INPUT");
      result.failures.push(...validEntries.map(({ file }) => ({ input: file, error: typedError })));
      return result;
    }

    if (normalized.tool === "merge-audio") {
      if (validEntries.length !== files.length) {
        const mergeError = new MediaProcessingError("INVALID_INPUT", "All selected audio files must be readable before they can be merged.");
        result.failures.push(...validEntries.map(({ file }) => ({ input: file, error: mergeError })));
        return result;
      }
      try {
        const audioEntries = validEntries.map(({ file, probe }) => ({ file, probe: assertAudioProbe(probe) }));
        const item = await processMerge(audioEntries.map((entry) => entry.file), audioEntries.map((entry) => entry.probe), normalized, signal, onProgress);
        result.items.push(item);
      } catch (error) {
        const typedError = asMediaProcessingError(error);
        if (typedError.code === "CANCELLED") throw typedError;
        result.failures.push({ input: validEntries[0].file, error: typedError });
      }
      return result;
    }

    for (const { file, probe } of validEntries) {
      try {
        throwIfAborted(signal);
        result.items.push(await processSingle(file, probe, normalized, signal, onProgress));
      } catch (error) {
        const typedError = asMediaProcessingError(error);
        if (typedError.code === "CANCELLED") throw typedError;
        result.failures.push({ input: file, error: typedError });
      }
    }
    return result;
  }

  dispose() {
    resetRuntime();
  }
}

export const mediaEngine = new BrowserMediaEngine();

export const isImplementedMediaTool = (slug: string): slug is MediaToolSlug => mediaToolSlugs.includes(slug as MediaToolSlug);

export { getMediaErrorMessage } from "./errors";
