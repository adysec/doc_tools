const { invoke } = window.__TAURI__.core;

function formatSize(bytes) {
  if (!Number.isFinite(bytes) || bytes < 0) return "-";
  const units = ["B", "KB", "MB", "GB"];
  let size = bytes;
  let unitIndex = 0;
  while (size >= 1024 && unitIndex < units.length - 1) {
    size /= 1024;
    unitIndex += 1;
  }
  return `${size.toFixed(size < 10 && unitIndex > 0 ? 2 : 1)} ${units[unitIndex]}`;
}

function toOutputPath(inputPath) {
  if (/\.docx$/i.test(inputPath)) {
    return inputPath.replace(/\.docx$/i, "_压缩后.docx");
  }
  return `${inputPath}_压缩后.docx`;
}

function toUnlockedPath(inputPath) {
  if (/\.docx$/i.test(inputPath)) {
    return inputPath.replace(/\.docx$/i, "_unlocked.docx");
  }
  return `${inputPath}_unlocked.docx`;
}

function toNoWatermarkPath(inputPath) {
  if (/\.docx$/i.test(inputPath)) {
    return inputPath.replace(/\.docx$/i, "_nowatermark.docx");
  }
  return `${inputPath}_nowatermark.docx`;
}

function toNoOutlinePath(inputPath) {
  if (/\.docx$/i.test(inputPath)) {
    return inputPath.replace(/\.docx$/i, "_no_outline.docx");
  }
  return `${inputPath}_no_outline.docx`;
}

function getDroppedPathFromDataTransfer(event) {
  const file = event.dataTransfer?.files?.[0];
  if (file && typeof file.path === "string" && file.path.length > 0) {
    return file.path;
  }

  const textUri = event.dataTransfer?.getData("text/uri-list");
  if (textUri && textUri.startsWith("file://")) {
    try {
      return decodeURIComponent(textUri.replace(/^file:\/\//, "")).trim();
    } catch {
      return textUri.replace(/^file:\/\//, "").trim();
    }
  }

  return "";
}

window.addEventListener("DOMContentLoaded", () => {
  const tabs = Array.from(document.querySelectorAll(".tool-tab"));
  const views = Array.from(document.querySelectorAll(".tool-view"));

  const switchTool = (toolName) => {
    tabs.forEach((tab) => {
      tab.classList.toggle("is-active", tab.dataset.tool === toolName);
    });
    views.forEach((view) => {
      view.classList.toggle("is-active", view.id === `tool-${toolName}`);
    });
  };

  tabs.forEach((tab) => {
    tab.addEventListener("click", () => switchTool(tab.dataset.tool));
  });

  const form = document.querySelector("#compress-form");
  const submitBtn = form.querySelector("button[type='submit']");
  const dropZoneEl = document.querySelector("#drop-zone");
  const unlockDropZoneEl = document.querySelector("#unlock-drop-zone");
  const watermarkDropZoneEl = document.querySelector("#watermark-drop-zone");
  const outlineDropZoneEl = document.querySelector("#outline-drop-zone");
  const inputPathEl = document.querySelector("#input-path");
  const outputPathEl = document.querySelector("#output-path");
  const qualityEl = document.querySelector("#quality");
  const maxWidthEl = document.querySelector("#max-width");
  const statusEl = document.querySelector("#status");
  const timerWrapEl = document.querySelector("#timer-wrap");
  const elapsedTimeEl = document.querySelector("#elapsed-time");
  const resultListEl = document.querySelector("#result-list");

  const originalEl = document.querySelector("#r-original");
  const compressedEl = document.querySelector("#r-compressed");
  const timeEl = document.querySelector("#r-time");
  let timerId = null;
  let startedAt = 0;

  const unlockForm = document.querySelector("#unlock-form");
  const unlockSubmitBtn = unlockForm.querySelector("button[type='submit']");
  const unlockInputPathEl = document.querySelector("#unlock-input-path");
  const unlockOutputPathEl = document.querySelector("#unlock-output-path");
  const unlockStatusEl = document.querySelector("#unlock-status");
  const unlockTimerWrapEl = document.querySelector("#unlock-timer-wrap");
  const unlockElapsedTimeEl = document.querySelector("#unlock-elapsed-time");
  const unlockResultListEl = document.querySelector("#unlock-result-list");
  const unlockOutputResultEl = document.querySelector("#u-output");
  const unlockRulesEl = document.querySelector("#u-rules");
  const unlockTimeEl = document.querySelector("#u-time");
  let unlockTimerId = null;
  let unlockStartedAt = 0;

  const watermarkForm = document.querySelector("#watermark-form");
  const watermarkSubmitBtn = watermarkForm.querySelector("button[type='submit']");
  const watermarkInputPathEl = document.querySelector("#watermark-input-path");
  const watermarkOutputPathEl = document.querySelector("#watermark-output-path");
  const watermarkStatusEl = document.querySelector("#watermark-status");
  const watermarkTimerWrapEl = document.querySelector("#watermark-timer-wrap");
  const watermarkElapsedTimeEl = document.querySelector("#watermark-elapsed-time");
  const watermarkResultListEl = document.querySelector("#watermark-result-list");
  const watermarkOutputResultEl = document.querySelector("#w-output");
  const watermarkHeadersEl = document.querySelector("#w-headers");
  const watermarkTimeEl = document.querySelector("#w-time");
  let watermarkTimerId = null;
  let watermarkStartedAt = 0;

  const outlineForm = document.querySelector("#outline-form");
  const outlineSubmitBtn = outlineForm.querySelector("button[type='submit']");
  const outlineInputPathEl = document.querySelector("#outline-input-path");
  const outlineOutputPathEl = document.querySelector("#outline-output-path");
  const outlineStatusEl = document.querySelector("#outline-status");
  const outlineTimerWrapEl = document.querySelector("#outline-timer-wrap");
  const outlineElapsedTimeEl = document.querySelector("#outline-elapsed-time");
  const outlineResultListEl = document.querySelector("#outline-result-list");
  const outlineOutputResultEl = document.querySelector("#o-output");
  const outlineHeadingStylesEl = document.querySelector("#o-heading-styles");
  const outlineTimeEl = document.querySelector("#o-time");
  let outlineTimerId = null;
  let outlineStartedAt = 0;

  const setRunning = (running) => {
    submitBtn.disabled = running;
    submitBtn.textContent = running ? "压缩中..." : "开始压缩";
  };

  const stopTimer = () => {
    if (timerId) {
      clearInterval(timerId);
      timerId = null;
    }
  };

  const startTimer = () => {
    startedAt = Date.now();
    elapsedTimeEl.textContent = "0.0 s";
    timerWrapEl.classList.remove("hidden");
    stopTimer();
    timerId = setInterval(() => {
      const elapsed = (Date.now() - startedAt) / 1000;
      elapsedTimeEl.textContent = `${elapsed.toFixed(1)} s`;
    }, 100);
  };

  const setUnlockRunning = (running) => {
    unlockSubmitBtn.disabled = running;
    unlockSubmitBtn.textContent = running ? "处理中..." : "开始解除保护";
  };

  const setWatermarkRunning = (running) => {
    watermarkSubmitBtn.disabled = running;
    watermarkSubmitBtn.textContent = running ? "处理中..." : "开始去水印";
  };

  const setOutlineRunning = (running) => {
    outlineSubmitBtn.disabled = running;
    outlineSubmitBtn.textContent = running ? "处理中..." : "开始去除标题大纲";
  };

  const stopUnlockTimer = () => {
    if (unlockTimerId) {
      clearInterval(unlockTimerId);
      unlockTimerId = null;
    }
  };

  const startUnlockTimer = () => {
    unlockStartedAt = Date.now();
    unlockElapsedTimeEl.textContent = "0.0 s";
    unlockTimerWrapEl.classList.remove("hidden");
    stopUnlockTimer();
    unlockTimerId = setInterval(() => {
      const elapsed = (Date.now() - unlockStartedAt) / 1000;
      unlockElapsedTimeEl.textContent = `${elapsed.toFixed(1)} s`;
    }, 100);
  };

  const stopWatermarkTimer = () => {
    if (watermarkTimerId) {
      clearInterval(watermarkTimerId);
      watermarkTimerId = null;
    }
  };

  const startWatermarkTimer = () => {
    watermarkStartedAt = Date.now();
    watermarkElapsedTimeEl.textContent = "0.0 s";
    watermarkTimerWrapEl.classList.remove("hidden");
    stopWatermarkTimer();
    watermarkTimerId = setInterval(() => {
      const elapsed = (Date.now() - watermarkStartedAt) / 1000;
      watermarkElapsedTimeEl.textContent = `${elapsed.toFixed(1)} s`;
    }, 100);
  };

  const stopOutlineTimer = () => {
    if (outlineTimerId) {
      clearInterval(outlineTimerId);
      outlineTimerId = null;
    }
  };

  const startOutlineTimer = () => {
    outlineStartedAt = Date.now();
    outlineElapsedTimeEl.textContent = "0.0 s";
    outlineTimerWrapEl.classList.remove("hidden");
    stopOutlineTimer();
    outlineTimerId = setInterval(() => {
      const elapsed = (Date.now() - outlineStartedAt) / 1000;
      outlineElapsedTimeEl.textContent = `${elapsed.toFixed(1)} s`;
    }, 100);
  };

  const applyDroppedPath = (path) => {
    if (!path) return;

    if (!/\.docx$/i.test(path)) {
      statusEl.textContent = "仅支持拖入 .docx 文件";
      return;
    }

    inputPathEl.value = path;
    if (!outputPathEl.value.trim()) {
      outputPathEl.value = toOutputPath(path);
    }
    statusEl.textContent = "已接收拖拽文件，参数确认后可开始压缩";
  };

  const applyUnlockDroppedPath = (path) => {
    if (!path) return;

    if (!/\.docx$/i.test(path)) {
      unlockStatusEl.textContent = "仅支持拖入 .docx 文件";
      return;
    }

    unlockInputPathEl.value = path;
    if (!unlockOutputPathEl.value.trim()) {
      unlockOutputPathEl.value = toUnlockedPath(path);
    }
    unlockStatusEl.textContent = "已接收拖拽文件，可开始解除保护";
  };

  const applyWatermarkDroppedPath = (path) => {
    if (!path) return;

    if (!/\.docx$/i.test(path)) {
      watermarkStatusEl.textContent = "仅支持拖入 .docx 文件";
      return;
    }

    watermarkInputPathEl.value = path;
    if (!watermarkOutputPathEl.value.trim()) {
      watermarkOutputPathEl.value = toNoWatermarkPath(path);
    }
    watermarkStatusEl.textContent = "已接收拖拽文件，可开始去水印";
  };

  const applyOutlineDroppedPath = (path) => {
    if (!path) return;

    if (!/\.docx$/i.test(path)) {
      outlineStatusEl.textContent = "仅支持拖入 .docx 文件";
      return;
    }

    outlineInputPathEl.value = path;
    if (!outlineOutputPathEl.value.trim()) {
      outlineOutputPathEl.value = toNoOutlinePath(path);
    }
    outlineStatusEl.textContent = "已接收拖拽文件，可开始去除标题大纲";
  };

  const activeToolName = () => {
    const activeTab = document.querySelector(".tool-tab.is-active");
    return activeTab?.dataset.tool || "compress";
  };

  inputPathEl.addEventListener("change", () => {
    const inputPath = inputPathEl.value.trim();
    if (inputPath && !outputPathEl.value.trim()) {
      outputPathEl.value = toOutputPath(inputPath);
    }
  });

  unlockInputPathEl.addEventListener("change", () => {
    const inputPath = unlockInputPathEl.value.trim();
    if (inputPath && !unlockOutputPathEl.value.trim()) {
      unlockOutputPathEl.value = toUnlockedPath(inputPath);
    }
  });

  watermarkInputPathEl.addEventListener("change", () => {
    const inputPath = watermarkInputPathEl.value.trim();
    if (inputPath && !watermarkOutputPathEl.value.trim()) {
      watermarkOutputPathEl.value = toNoWatermarkPath(inputPath);
    }
  });

  outlineInputPathEl.addEventListener("change", () => {
    const inputPath = outlineInputPathEl.value.trim();
    if (inputPath && !outlineOutputPathEl.value.trim()) {
      outlineOutputPathEl.value = toNoOutlinePath(inputPath);
    }
  });

  // Prevent the webview from navigating away when files are dropped.
  window.addEventListener("dragover", (event) => {
    event.preventDefault();
  });
  window.addEventListener("drop", (event) => {
    event.preventDefault();
    const path = getDroppedPathFromDataTransfer(event);
    if (activeToolName() === "unlock") {
      applyUnlockDroppedPath(path);
    } else if (activeToolName() === "watermark") {
      applyWatermarkDroppedPath(path);
    } else if (activeToolName() === "outline") {
      applyOutlineDroppedPath(path);
    } else {
      applyDroppedPath(path);
    }
  });

  dropZoneEl.addEventListener("dragover", (event) => {
    event.preventDefault();
    dropZoneEl.classList.add("is-dragover");
  });

  dropZoneEl.addEventListener("dragleave", () => {
    dropZoneEl.classList.remove("is-dragover");
  });

  dropZoneEl.addEventListener("drop", (event) => {
    event.preventDefault();
    dropZoneEl.classList.remove("is-dragover");
    applyDroppedPath(getDroppedPathFromDataTransfer(event));
  });

  unlockDropZoneEl.addEventListener("dragover", (event) => {
    event.preventDefault();
    unlockDropZoneEl.classList.add("is-dragover");
  });

  unlockDropZoneEl.addEventListener("dragleave", () => {
    unlockDropZoneEl.classList.remove("is-dragover");
  });

  unlockDropZoneEl.addEventListener("drop", (event) => {
    event.preventDefault();
    unlockDropZoneEl.classList.remove("is-dragover");
    applyUnlockDroppedPath(getDroppedPathFromDataTransfer(event));
  });

  watermarkDropZoneEl.addEventListener("dragover", (event) => {
    event.preventDefault();
    watermarkDropZoneEl.classList.add("is-dragover");
  });

  watermarkDropZoneEl.addEventListener("dragleave", () => {
    watermarkDropZoneEl.classList.remove("is-dragover");
  });

  watermarkDropZoneEl.addEventListener("drop", (event) => {
    event.preventDefault();
    watermarkDropZoneEl.classList.remove("is-dragover");
    applyWatermarkDroppedPath(getDroppedPathFromDataTransfer(event));
  });

  outlineDropZoneEl.addEventListener("dragover", (event) => {
    event.preventDefault();
    outlineDropZoneEl.classList.add("is-dragover");
  });

  outlineDropZoneEl.addEventListener("dragleave", () => {
    outlineDropZoneEl.classList.remove("is-dragover");
  });

  outlineDropZoneEl.addEventListener("drop", (event) => {
    event.preventDefault();
    outlineDropZoneEl.classList.remove("is-dragover");
    applyOutlineDroppedPath(getDroppedPathFromDataTransfer(event));
  });

  const currentWindow = window.__TAURI__.webviewWindow?.getCurrentWebviewWindow?.();
  if (currentWindow?.onDragDropEvent) {
    currentWindow.onDragDropEvent((event) => {
      if (event.payload?.type === "over") {
        if (activeToolName() === "unlock") {
          unlockDropZoneEl.classList.add("is-dragover");
        } else if (activeToolName() === "watermark") {
          watermarkDropZoneEl.classList.add("is-dragover");
        } else if (activeToolName() === "outline") {
          outlineDropZoneEl.classList.add("is-dragover");
        } else {
          dropZoneEl.classList.add("is-dragover");
        }
        return;
      }
      if (event.payload?.type === "cancel") {
        dropZoneEl.classList.remove("is-dragover");
        unlockDropZoneEl.classList.remove("is-dragover");
        watermarkDropZoneEl.classList.remove("is-dragover");
        outlineDropZoneEl.classList.remove("is-dragover");
        return;
      }
      if (event.payload?.type === "drop") {
        dropZoneEl.classList.remove("is-dragover");
        unlockDropZoneEl.classList.remove("is-dragover");
        watermarkDropZoneEl.classList.remove("is-dragover");
        outlineDropZoneEl.classList.remove("is-dragover");
        const droppedPath = Array.isArray(event.payload.paths) ? event.payload.paths[0] : "";
        if (activeToolName() === "unlock") {
          applyUnlockDroppedPath(droppedPath);
        } else if (activeToolName() === "watermark") {
          applyWatermarkDroppedPath(droppedPath);
        } else if (activeToolName() === "outline") {
          applyOutlineDroppedPath(droppedPath);
        } else {
          applyDroppedPath(droppedPath);
        }
      }
    });
  }

  form.addEventListener("submit", async (e) => {
    e.preventDefault();

    const inputPath = inputPathEl.value.trim();
    const outputPath = outputPathEl.value.trim();
    const quality = Number.parseInt(qualityEl.value, 10);
    const maxWidth = Number.parseInt(maxWidthEl.value, 10);

    statusEl.textContent = "正在压缩，请稍候...";
    startTimer();
    setRunning(true);
    resultListEl.classList.add("hidden");

    try {
      const result = await invoke("compress_docx", {
        inputPath,
        outputPath,
        quality,
        maxWidth,
      });

      originalEl.textContent = formatSize(result.original_size);
      compressedEl.textContent = formatSize(result.compressed_size);
      timeEl.textContent = `${result.elapsed_seconds.toFixed(2)} s`;

      if (result.original_size > 0) {
        const deltaRatio = ((result.compressed_size - result.original_size) / result.original_size) * 100;
        if (deltaRatio <= 0) {
          statusEl.textContent = `完成，体积减小 ${Math.abs(deltaRatio).toFixed(2)}%`;
        } else {
          statusEl.textContent = `完成，体积增大 ${deltaRatio.toFixed(2)}%`;
        }
      } else {
        statusEl.textContent = "完成";
      }
      const elapsed = (Date.now() - startedAt) / 1000;
      elapsedTimeEl.textContent = `${elapsed.toFixed(1)} s`;
      resultListEl.classList.remove("hidden");
    } catch (error) {
      statusEl.textContent = `压缩失败: ${error}`;
    } finally {
      stopTimer();
      setRunning(false);
    }
  });

  unlockForm.addEventListener("submit", async (e) => {
    e.preventDefault();

    const inputPath = unlockInputPathEl.value.trim();
    const outputPath = unlockOutputPathEl.value.trim();

    unlockStatusEl.textContent = "正在解除编辑保护，请稍候...";
    startUnlockTimer();
    setUnlockRunning(true);
    unlockResultListEl.classList.add("hidden");

    try {
      const result = await invoke("unlock_docx", {
        inputPath,
        outputPath,
      });

      unlockOutputResultEl.textContent = result.output_path;
      unlockRulesEl.textContent = `${result.removed_rules}`;
      unlockTimeEl.textContent = `${result.elapsed_seconds.toFixed(2)} s`;

      if (result.removed_rules > 0) {
        unlockStatusEl.textContent = "解除保护完成";
      } else {
        unlockStatusEl.textContent = "完成，未发现保护节点";
      }

      const elapsed = (Date.now() - unlockStartedAt) / 1000;
      unlockElapsedTimeEl.textContent = `${elapsed.toFixed(1)} s`;
      unlockResultListEl.classList.remove("hidden");
    } catch (error) {
      unlockStatusEl.textContent = `处理失败: ${error}`;
    } finally {
      stopUnlockTimer();
      setUnlockRunning(false);
    }
  });

  watermarkForm.addEventListener("submit", async (e) => {
    e.preventDefault();

    const inputPath = watermarkInputPathEl.value.trim();
    const outputPath = watermarkOutputPathEl.value.trim();

    watermarkStatusEl.textContent = "正在去除水印，请稍候...";
    startWatermarkTimer();
    setWatermarkRunning(true);
    watermarkResultListEl.classList.add("hidden");

    try {
      const result = await invoke("remove_docx_watermark", {
        inputPath,
        outputPath,
      });

      watermarkOutputResultEl.textContent = result.output_path;
      watermarkHeadersEl.textContent = `${result.cleared_headers}`;
      watermarkTimeEl.textContent = `${result.elapsed_seconds.toFixed(2)} s`;

      watermarkStatusEl.textContent = "去水印完成";

      const elapsed = (Date.now() - watermarkStartedAt) / 1000;
      watermarkElapsedTimeEl.textContent = `${elapsed.toFixed(1)} s`;
      watermarkResultListEl.classList.remove("hidden");
    } catch (error) {
      watermarkStatusEl.textContent = `处理失败: ${error}`;
    } finally {
      stopWatermarkTimer();
      setWatermarkRunning(false);
    }
  });

  outlineForm.addEventListener("submit", async (e) => {
    e.preventDefault();

    const inputPath = outlineInputPathEl.value.trim();
    const outputPath = outlineOutputPathEl.value.trim();

    outlineStatusEl.textContent = "正在去除标题大纲，请稍候...";
    startOutlineTimer();
    setOutlineRunning(true);
    outlineResultListEl.classList.add("hidden");

    try {
      const result = await invoke("remove_docx_outline", {
        inputPath,
        outputPath,
      });

      outlineOutputResultEl.textContent = result.output_path;
      outlineHeadingStylesEl.textContent = `${result.removed_heading_styles}`;
      outlineTimeEl.textContent = `${result.elapsed_seconds.toFixed(2)} s`;

      outlineStatusEl.textContent = "去标题大纲完成";

      const elapsed = (Date.now() - outlineStartedAt) / 1000;
      outlineElapsedTimeEl.textContent = `${elapsed.toFixed(1)} s`;
      outlineResultListEl.classList.remove("hidden");
    } catch (error) {
      outlineStatusEl.textContent = `处理失败: ${error}`;
    } finally {
      stopOutlineTimer();
      setOutlineRunning(false);
    }
  });
});
