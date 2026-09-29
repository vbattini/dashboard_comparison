const folders = ["spartnash", "trendence"];
const pendingUploads = new Set();

async function refreshList(folder) {
  const res = await fetch(`/api/images/${folder}`);
  const images = await res.json();
  const list = document.querySelector(`.file-list[data-folder="${folder}"]`);
  list.innerHTML = "";
  images.forEach((img) => {
    const li = document.createElement("li");
    li.innerHTML = `
      <img src="${img.url}" alt="${img.filename}" />
      <span class="name">${img.filename}</span>
      <button data-folder="${folder}" data-filename="${img.filename}">Remove</button>
    `;
    list.appendChild(li);
  });
  list.querySelectorAll("button").forEach((btn) => {
    btn.addEventListener("click", async () => {
      await fetch(`/api/images/${btn.dataset.folder}/${btn.dataset.filename}`, { method: "DELETE" });
      refreshList(btn.dataset.folder);
    });
  });
}

async function uploadFiles(folder, files) {
  if (!files.length) return;
  pendingUploads.add(folder);
  document.getElementById("compareBtn").disabled = true;
  document.getElementById("status").textContent = `Uploading ${files.length} ${folder} images...`;
  const formData = new FormData();
  Array.from(files).forEach((file) => formData.append("files", file));
  try {
    const res = await fetch(`/api/upload/${folder}`, { method: "POST", body: formData });
    const body = await res.json();
    if (!res.ok) throw new Error(body.detail || "Upload failed");
    const errors = Array.isArray(body?.errors) ? body.errors : [];
    if (errors.length) {
      alert(`Some files were skipped:\n${errors.join("\n")}`);
    }
  } catch (error) {
    document.getElementById("status").textContent = `Upload error: ${error.message}`;
  } finally {
    pendingUploads.delete(folder);
    if (!pendingUploads.size) document.getElementById("status").textContent = "Uploads complete. Ready to compare.";
    document.getElementById("compareBtn").disabled = pendingUploads.size > 0;
    refreshList(folder);
  }
}

folders.forEach((folder) => {
  const dropzone = document.querySelector(`.dropzone[data-folder="${folder}"]`);
  const input = document.querySelector(`input[data-folder="${folder}"]`);

  dropzone.addEventListener("click", () => input.click());
  input.addEventListener("change", () => uploadFiles(folder, input.files));

  dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("drag");
  });
  dropzone.addEventListener("dragleave", () => dropzone.classList.remove("drag"));
  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("drag");
    uploadFiles(folder, e.dataTransfer.files);
  });

  refreshList(folder);
});

function renderResults(data) {
  const wrap = document.getElementById("resultsWrap");
  const summaryGrid = document.getElementById("summaryGrid");
  const pairList = document.getElementById("pairList");
  const usagePanel = document.getElementById("usagePanel");
  if (!data || typeof data !== "object") {
    throw new Error("The server returned an empty results response");
  }
  const pairs = Array.isArray(data.pairs) ? data.pairs : [];
  const imageStats = Array.isArray(data.image_stats) ? data.image_stats : [];
  const downloadLink = document.getElementById("downloadLink");
  const downloadJsonLink = document.getElementById("downloadJsonLink");
  summaryGrid.innerHTML = "";
  pairList.innerHTML = "";
  usagePanel.innerHTML = "";
  downloadLink.href = data.download_url || "/api/download";
  downloadLink.style.display = "inline-block";
  downloadJsonLink.href = data.json_download_url || "/api/download-json";
  downloadJsonLink.style.display = "inline-block";
  if (data.workbook_available) wrap.style.display = "block";
  if (!pairs.length) {
    return;
  }

  const totals = pairs.reduce((result, pair) => {
    result.matches += pair.matches || 0;
    result.differences += pair.differences || 0;
    result.only += (pair.trendence_only || 0) + (pair.spartnash_only || 0);
    return result;
  }, { matches: 0, differences: 0, only: 0 });
  const totalItems = totals.matches + totals.differences + totals.only;
  const overallRate = totalItems ? (totals.matches / totalItems) * 100 : 0;
  const cards = [
    [pairs.length, "Dashboard pairs", ""],
    [`${overallRate.toFixed(1)}%`, "Overall match rate", "good"],
    [totals.matches, "Matched items", "good"],
    [totals.differences + totals.only, "Items to review", "accent"],
  ];
  cards.forEach(([value, label, className]) => {
    summaryGrid.insertAdjacentHTML("beforeend", `<div class="summary-card ${className}"><strong>${value}</strong><span>${label}</span></div>`);
  });
  document.getElementById("summaryNote").textContent = `${pairs.length} dashboard comparisons completed`;

  pairs.forEach((pair) => {
    const pct = pair.match_percentage;
    const rate = pct === null || pct === undefined ? 0 : pct * 100;
    const rateText = pct === null || pct === undefined ? "N/A" : `${rate.toFixed(1)}%`;
    const rateClass = rate >= 80 ? "good" : "";
    pairList.insertAdjacentHTML("beforeend", `
      <article class="pair-card">
        <div class="pair-top"><div><div class="pair-number">Pair ${pair.pair}</div><h3>${pair.trendence_title || "Untitled dashboard"}</h3><p>vs ${pair.spartnash_title || "Untitled dashboard"}</p></div><div class="pair-rate ${rateClass}">${rateText}</div></div>
        <div class="bar"><span class="${rateClass}" style="width:${Math.min(rate, 100)}%"></span></div>
        <div class="pair-stats"><span><b>${pair.matches}</b> matched</span><span><b>${pair.differences}</b> different</span><span><b>${(pair.trendence_only || 0) + (pair.spartnash_only || 0)}</b> unmatched</span></div>
      </article>
    `);
  });
  const formatTokens = (value) => value === null || value === undefined ? "Unknown" : value.toLocaleString();
  const formatUsd = (value) => value === null || value === undefined ? "Unknown" : `$${(Math.round((value + Number.EPSILON) * 100000) / 100000).toFixed(5)}`;
  const formatInr = (value) => value === null || value === undefined ? "Unknown" : `₹${value.toFixed(2)}`;
  usagePanel.innerHTML = `
    <h3>Image and LLM usage</h3>
    <p>${data.total_llm_calls} LLM calls recorded · ${formatTokens(data.total_tokens)} total tokens</p>
    <table class="usage-table"><thead><tr><th>Image</th><th>Prompt Tokens</th><th>Image Tokens</th><th>Total Input Tokens</th><th>Output Tokens</th><th>Total Tokens</th><th>Input Cost ($)</th><th>Output Cost ($)</th><th>Total Cost ($)</th><th>Total Cost (₹)</th></tr></thead><tbody>
      ${imageStats.map((item) => `<tr><td>${item.folder} ${item.image}</td><td>${formatTokens(item.prompt_tokens)}</td><td>${formatTokens(item.image_input_tokens)}</td><td>${formatTokens(item.total_input_tokens)}</td><td>${formatTokens(item.output_tokens)}</td><td>${formatTokens(item.total_tokens)}</td><td>${formatUsd(item.input_cost_usd)}</td><td>${formatUsd(item.output_cost_usd)}</td><td>${formatUsd(item.total_cost_usd)}</td><td>${formatInr(item.total_cost_inr)}</td></tr>`).join("")}
      <tr><td><strong>Total</strong></td><td><strong>${formatTokens(data.total_prompt_tokens)}</strong></td><td><strong>${formatTokens(data.total_image_tokens)}</strong></td><td><strong>${formatTokens(data.total_input_tokens)}</strong></td><td><strong>${formatTokens(data.total_output_tokens)}</strong></td><td><strong>${formatTokens(data.total_tokens)}</strong></td><td><strong>${formatUsd(data.total_input_cost_usd)}</strong></td><td><strong>${formatUsd(data.total_output_cost_usd)}</strong></td><td><strong>${formatUsd(data.total_cost_usd)}</strong></td><td><strong>${formatInr(data.total_cost_inr)}</strong></td></tr>
    </tbody></table>`;
  wrap.style.display = "block";
}

document.getElementById("compareBtn").addEventListener("click", async () => {
  const btn = document.getElementById("compareBtn");
  const status = document.getElementById("status");
  btn.disabled = true;
  if (pendingUploads.size) {
    status.textContent = "Please wait for uploads to finish.";
    btn.disabled = false;
    return;
  }
  btn.textContent = "Working...";
  try {
    status.textContent = "Extracting images and comparing data... this can take a while.";
    const comparisonRes = await fetch("/api/compare", { method: "POST" });
    const responseText = await comparisonRes.text();
    let job;
    try {
      job = JSON.parse(responseText);
    } catch {
      throw new Error(`Server returned HTTP ${comparisonRes.status}: ${responseText.slice(0, 240)}`);
    }
    if (!comparisonRes.ok) throw new Error(job.detail || "Comparison failed");
    const jobId = job.job_id;

    while (job.status === "running") {
      status.textContent = "Comparison is running in the background...";
      await new Promise((resolve) => setTimeout(resolve, 3000));
      const pollRes = await fetch(`/api/compare/${jobId}`);
      const pollText = await pollRes.text();
      try {
        job = JSON.parse(pollText);
      } catch {
        throw new Error(`Server returned HTTP ${pollRes.status}: ${pollText.slice(0, 240)}`);
      }
      if (!pollRes.ok) {
        if (pollRes.status === 404) {
          const recoveryRes = await fetch("/api/results");
          const recoveryText = await recoveryRes.text();
          try {
            job = JSON.parse(recoveryText);
          } catch {
            throw new Error(`Server returned HTTP ${recoveryRes.status}: ${recoveryText.slice(0, 240)}`);
          }
          if (!recoveryRes.ok || !job.workbook_available) {
            throw new Error(job.detail || "Comparison job was lost before a report was generated");
          }
          job.status = "completed";
        } else {
          throw new Error(job.detail || "Comparison failed");
        }
      }
    }
    status.textContent = "Comparison complete.";
    renderResults(job);
  } catch (err) {
    status.textContent = `Error: ${err.message}`;
  } finally {
    btn.disabled = false;
    btn.textContent = "compare";
  }
});

// Load any existing results on page load.
fetch("/api/results")
  .then((res) => res.json())
  .then(renderResults)
  .catch(() => {});
