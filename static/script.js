(() => {
  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("file-input");
  const browseBtn = document.getElementById("browse-btn");
  const fileListEl = document.getElementById("file-list");
  const processBtn = document.getElementById("process-btn");
  const statusText = document.getElementById("status-text");

  const errorPanel = document.getElementById("error-panel");
  const errorList = document.getElementById("error-list");

  const results = document.getElementById("results");
  const monthLabel = document.getElementById("month-label");
  const portfolioTotal = document.getElementById("portfolio-total");
  const tenantCount = document.getElementById("tenant-count");
  const downloadBtn = document.getElementById("download-btn");
  const summaryBody = document.getElementById("summary-body");
  const tenantGrid = document.getElementById("tenant-grid");

  let selectedFiles = [];
  const charts = [];

  const currency = (n) =>
    "Rp " + Math.round(n).toLocaleString("id-ID");

  // ---------- File selection ----------

  function addFiles(fileArray) {
    for (const f of fileArray) {
      if (!/\.xlsx$|\.xls$/i.test(f.name)) continue;
      if (selectedFiles.some((existing) => existing.name === f.name && existing.size === f.size)) continue;
      selectedFiles.push(f);
    }
    renderFileList();
  }

  function renderFileList() {
    fileListEl.innerHTML = "";
    fileListEl.hidden = selectedFiles.length === 0;
    selectedFiles.forEach((f, idx) => {
      const li = document.createElement("li");
      li.innerHTML = `<span>${f.name}</span>`;
      const removeBtn = document.createElement("button");
      removeBtn.className = "remove-file";
      removeBtn.type = "button";
      removeBtn.textContent = "×";
      removeBtn.addEventListener("click", () => {
        selectedFiles.splice(idx, 1);
        renderFileList();
      });
      li.appendChild(removeBtn);
      fileListEl.appendChild(li);
    });
    processBtn.disabled = selectedFiles.length === 0;
  }

  browseBtn.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", (e) => addFiles([...e.target.files]));

  ["dragenter", "dragover"].forEach((evt) =>
    dropzone.addEventListener(evt, (e) => {
      e.preventDefault();
      dropzone.classList.add("is-dragover");
    })
  );
  ["dragleave", "drop"].forEach((evt) =>
    dropzone.addEventListener(evt, (e) => {
      e.preventDefault();
      dropzone.classList.remove("is-dragover");
    })
  );
  dropzone.addEventListener("drop", (e) => {
    addFiles([...e.dataTransfer.files]);
  });

  // ---------- Processing ----------

  processBtn.addEventListener("click", async () => {
    if (selectedFiles.length === 0) return;

    processBtn.disabled = true;
    statusText.classList.remove("is-error");
    statusText.textContent = `Processing ${selectedFiles.length} file(s)...`;
    errorPanel.hidden = true;
    results.hidden = true;

    const formData = new FormData();
    selectedFiles.forEach((f) => formData.append("files", f));

    try {
      const res = await fetch("/api/process", { method: "POST", body: formData });
      const data = await res.json();

      if (!res.ok) {
        statusText.classList.add("is-error");
        statusText.textContent = data.error || "Processing failed.";
        if (data.file_errors && data.file_errors.length) showFileErrors(data.file_errors);
        return;
      }

      statusText.textContent = "Done.";
      if (data.file_errors && data.file_errors.length) showFileErrors(data.file_errors);
      renderResults(data);
    } catch (err) {
      statusText.classList.add("is-error");
      statusText.textContent = "Could not reach the server. Is the app still running?";
    } finally {
      processBtn.disabled = false;
    }
  });

  function showFileErrors(fileErrors) {
    errorList.innerHTML = "";
    fileErrors.forEach((e) => {
      const li = document.createElement("li");
      li.innerHTML = `<code>${e.file}</code> — ${e.message}`;
      errorList.appendChild(li);
    });
    errorPanel.hidden = false;
  }

  // ---------- Rendering ----------

  function renderResults(data) {
    monthLabel.textContent = data.month_label;
    const total = data.summary.reduce((sum, t) => sum + t.total, 0);
    portfolioTotal.textContent = currency(total);
    tenantCount.textContent = `${data.summary.length} tenant${data.summary.length === 1 ? "" : "s"} processed`;
    downloadBtn.href = `/api/download/${data.download_token}`;

    // Summary table
    summaryBody.innerHTML = "";
    data.summary.forEach((t) => {
      const uplift = t.weekday_avg > 0 ? ((t.weekend_avg / t.weekday_avg - 1) * 100) : 0;
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${t.tenant}</td>
        <td>${currency(t.total)}</td>
        <td>${currency(t.weekday_avg)}</td>
        <td>${currency(t.weekend_avg)}</td>
        <td><span class="uplift-pill">${uplift >= 0 ? "+" : ""}${uplift.toFixed(0)}%</span></td>
      `;
      summaryBody.appendChild(tr);
    });

    // Tenant cards
    charts.forEach((c) => c.destroy());
    charts.length = 0;
    tenantGrid.innerHTML = "";

    Object.entries(data.series).forEach(([tenantName, s]) => {
      const card = document.createElement("div");
      card.className = "tenant-card";

      const totalSales = s.sales.reduce((a, b) => a + b, 0);

      card.innerHTML = `
        <div class="tenant-card__head">
          <h3 class="tenant-card__name">${tenantName}</h3>
          <span class="tenant-card__total">${currency(totalSales)}</span>
        </div>
        <div class="heartbeat">${s.day_types.map((t) => `<span class="heartbeat__tick" data-type="${t}"></span>`).join("")}</div>
        <canvas></canvas>
      `;
      tenantGrid.appendChild(card);

      const ctx = card.querySelector("canvas").getContext("2d");
      const pointColors = s.day_types.map((t) => (t === "Weekend" ? "#4c7a5d" : "#1f3864"));

      const chart = new Chart(ctx, {
        type: "line",
        data: {
          labels: s.dates.map((d) => d.slice(8, 10)),
          datasets: [
            {
              data: s.sales,
              borderColor: "#1f3864",
              borderWidth: 1.75,
              pointBackgroundColor: pointColors,
              pointRadius: 2.5,
              pointHoverRadius: 4,
              tension: 0.15,
              fill: {
                target: "origin",
                above: "rgba(31, 56, 100, 0.06)",
              },
            },
          ],
        },
        options: {
          responsive: true,
          plugins: { legend: { display: false } },
          scales: {
            x: { grid: { display: false }, ticks: { maxTicksLimit: 10, font: { family: "Consolas, Menlo, monospace", size: 10 } } },
            y: { grid: { color: "#eee" }, ticks: { font: { family: "Consolas, Menlo, monospace", size: 10 }, callback: (v) => (v >= 1e6 ? (v / 1e6).toFixed(0) + "M" : v) } },
          },
        },
      });
      charts.push(chart);
    });

    results.hidden = false;
    results.scrollIntoView({ behavior: "smooth", block: "start" });
  }
})();
