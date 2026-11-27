window.PCDT = window.PCDT || {};

(function (PCDT) {
  const DEFAULT_ND30_SETTINGS = {
    fontSize: 13,
    indent: 1.27,
    lineSpacing: 1.15,
    spaceBeforeAfter: 6,
    marginTop: 20,
    marginBottom: 20,
    marginLeft: 30,
    marginRight: 15
  };

  let sessionIgnoredKeys = new Set();

  function saveApiKey() {
    const key = document.getElementById("apiKey").value.trim();
    if (!key) return alert("Vui lòng nhập API Key.");
    localStorage.setItem("gemini_api_key", key);
    PCDT.setStatus("Đã lưu API Key vào trình duyệt.");
  }

  function updateRulesStatusUI() {
    const savedRules = localStorage.getItem("custom_rules_text");
    const savedName = localStorage.getItem("custom_rules_name");
    const statusEl = document.getElementById("rulesStatus");
    const clearBtn = document.getElementById("clearRulesBtn");

    if (savedRules) {
      statusEl.textContent = `✅ Đang áp dụng: ${savedName || "rules_pcdt.txt"} (${savedRules.length} ký tự)`;
      statusEl.style.color = "#0f766e";
      clearBtn.style.display = "inline-block";
    } else {
      statusEl.textContent = "Chưa nạp quy định riêng (chỉ dùng chuẩn NĐ 30/2020/NĐ-CP).";
      statusEl.style.color = "#64748b";
      clearBtn.style.display = "none";
    }
  }

  function handleRulesFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      const content = (e.target.result || "").trim();
      if (!content) return alert("File quy định đang trống.");
      localStorage.setItem("custom_rules_text", content);
      localStorage.setItem("custom_rules_name", file.name);
      updateRulesStatusUI();
      PCDT.setStatus(`Đã lưu quy định nội bộ từ file ${file.name} vào máy tính.`);
    };
    reader.readAsText(file, "UTF-8");
  }

  function clearCustomRules() {
    localStorage.removeItem("custom_rules_text");
    localStorage.removeItem("custom_rules_name");
    const input = document.getElementById("rulesFileInput");
    if (input) input.value = "";
    updateRulesStatusUI();
  }

  // --- Cơ chế Bỏ qua 2 cấp độ ---
  function getIssueKey(item) {
    const type = item.checkType || item.category || "general";
    const quoteNorm = PCDT.normalizeText(item.quote || "");
    if (quoteNorm) return `${type}::${quoteNorm}`;
    return `${type}::${PCDT.normalizeText(item.issue || "").slice(0, 60)}`;
  }

  function getGlobalIgnoredList() {
    try {
      return JSON.parse(localStorage.getItem("global_ignored_rules") || "[]");
    } catch (e) {
      return [];
    }
  }

  function saveGlobalIgnoredList(list) {
    localStorage.setItem("global_ignored_rules", JSON.stringify(list));
    renderGlobalIgnoreManager();
  }

  function addGlobalIgnoredRule(item, reason) {
    const list = getGlobalIgnoredList();
    const key = getIssueKey(item);
    if (!list.some((r) => r.key === key)) {
      list.push({
        id: Date.now().toString() + Math.random().toString(36).slice(2, 6),
        key,
        quote: item.quote || "",
        issue: item.issue || "",
        reason: reason.trim()
      });
      saveGlobalIgnoredList(list);
    }
  }

  function addSessionIgnoredKey(key) {
    sessionIgnoredKeys.add(key);
    updateRestoreBarUI();
  }

  function isIssueIgnored(item) {
    const key = getIssueKey(item);
    if (sessionIgnoredKeys.has(key)) return true;
    const globalList = getGlobalIgnoredList();
    const quoteNorm = PCDT.normalizeText(item.quote || "");
    return globalList.some(
      (r) => r.key === key || (quoteNorm && r.quote && PCDT.normalizeText(r.quote) === quoteNorm)
    );
  }

  function updateRestoreBarUI() {
    const bar = document.getElementById("restoreBar");
    const textEl = document.getElementById("restoreText");
    if (!bar || !textEl) return;
    const count = sessionIgnoredKeys.size;
    if (count > 0) {
      bar.style.display = "flex";
      textEl.textContent = `Đã bỏ qua ${count} lỗi ở văn bản này.`;
    } else {
      bar.style.display = "none";
    }
  }

  function restoreSessionIgnored() {
    sessionIgnoredKeys.clear();
    updateRestoreBarUI();
    if (PCDT.lastRawGroups && PCDT.renderFilteredGroups) {
      PCDT.renderFilteredGroups(PCDT.lastRawGroups);
      PCDT.setStatus("Đã khôi phục toàn bộ các lỗi đã bỏ qua ở văn bản này.");
    }
  }

  function renderGlobalIgnoreManager() {
    const list = getGlobalIgnoredList();
    const toggleBtn = document.getElementById("toggleGlobalIgnoreBtn");
    const listDiv = document.getElementById("globalIgnoreList");
    const exportBtn = document.getElementById("exportGlobalIgnoreBtn");
    if (!toggleBtn || !listDiv || !exportBtn) return;

    toggleBtn.textContent = `📋 Quản lý danh sách bỏ qua chung (${list.length})`;
    listDiv.innerHTML = "";

    if (list.length === 0) {
      listDiv.innerHTML = `<div class="hint" style="margin-bottom:6px;">Chưa có mục bỏ qua chung nào.</div>`;
      exportBtn.disabled = true;
      exportBtn.style.opacity = "0.5";
      return;
    }

    exportBtn.disabled = false;
    exportBtn.style.opacity = "1";
    list.forEach((rule) => {
      const row = document.createElement("div");
      row.className = "global-item";

      const header = document.createElement("div");
      header.className = "global-item-header";

      const info = document.createElement("div");
      const labelText = rule.quote ? `"${rule.quote}"` : (rule.issue || "").slice(0, 60);
      info.innerHTML = `<strong>${PCDT.escapeHtml(labelText)}</strong><br/><span>Lý do: ${PCDT.escapeHtml(rule.reason)}</span>`;

      const delBtn = document.createElement("button");
      delBtn.className = "global-item-del";
      delBtn.textContent = "🗑️ Xóa";
      delBtn.onclick = () => {
        const updated = getGlobalIgnoredList().filter((r) => r.id !== rule.id);
        saveGlobalIgnoredList(updated);
        if (PCDT.lastRawGroups && PCDT.renderFilteredGroups) {
          PCDT.renderFilteredGroups(PCDT.lastRawGroups);
        }
      };

      header.appendChild(info);
      header.appendChild(delBtn);
      row.appendChild(header);
      listDiv.appendChild(row);
    });
  }

  function exportGlobalIgnoreTxt() {
    const list = getGlobalIgnoredList();
    if (list.length === 0) return alert("Danh sách bỏ qua chung đang trống.");

    const lines = [
      "# DANH SÁCH LƯU Ý BỎ QUA CHUNG (DÙNG CHO WORD ADD-IN RÀ SOÁT VĂN BẢN - PC ĐỒNG THÁP)",
      ""
    ];
    list.forEach((r, idx) => {
      const target = r.quote ? `Cụm từ "${r.quote}"` : `Lỗi "${r.issue}"`;
      lines.push(`${idx + 1}. Không báo lỗi đối với ${target} — Lý do: ${r.reason}`);
    });

    const blob = new Blob([lines.join("\n")], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "luu_y_bo_qua_chung.txt";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function handleImportGlobalIgnoreTxt(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      const text = e.target.result || "";
      const lines = text.split(/\r?\n/).map((l) => l.trim()).filter((l) => l && !l.startsWith("#"));
      const currentList = getGlobalIgnoredList();
      let addedCount = 0;

      lines.forEach((line) => {
        let quote = "";
        let issue = "";
        let reason = "";

        const mQuote = line.match(/Cụm từ\s+"([^"]+)"\s*[—\-]+\s*Lý do:\s*(.+)$/i);
        const mIssue = line.match(/Lỗi\s+"([^"]+)"\s*[—\-]+\s*Lý do:\s*(.+)$/i);

        if (mQuote) {
          quote = mQuote[1].trim();
          reason = mQuote[2].trim();
        } else if (mIssue) {
          issue = mIssue[1].trim();
          reason = mIssue[2].trim();
        } else {
          const parts = line.replace(/^\d+\.\s*/, "").split(/\s*[—\-:]\s*/);
          if (parts.length >= 2) {
            quote = parts[0].replace(/^["']|["']$/g, "").trim();
            reason = parts.slice(1).join(" - ").trim();
          } else {
            issue = line.replace(/^\d+\.\s*/, "").trim();
            reason = "Nhập từ file chia sẻ";
          }
        }

        if (!quote && !issue) return;

        const tempItem = { category: "imported", quote, issue };
        const key = getIssueKey(tempItem);
        const quoteNorm = PCDT.normalizeText(quote);

        const isDuplicate = currentList.some(
          (r) => r.key === key || (quoteNorm && r.quote && PCDT.normalizeText(r.quote) === quoteNorm)
        );

        if (!isDuplicate) {
          currentList.push({
            id: Date.now().toString() + Math.random().toString(36).slice(2, 7),
            key,
            quote,
            issue,
            reason
          });
          addedCount++;
        }
      });

      if (addedCount > 0) {
        saveGlobalIgnoredList(currentList);
        if (PCDT.lastRawGroups && PCDT.renderFilteredGroups) {
          PCDT.renderFilteredGroups(PCDT.lastRawGroups);
        }
        PCDT.setStatus(`Đã nhập và gộp thêm ${addedCount} lưu ý mới (đã tự động lọc trùng).`);
      } else {
        PCDT.setStatus("Không có mục mới nào được thêm (các mục trong file đã tồn tại).");
      }
      event.target.value = "";
    };
    reader.readAsText(file, "UTF-8");
  }

  function buildGlobalIgnorePromptSection() {
    const list = getGlobalIgnoredList();
    if (list.length === 0) return "";
    const lines = list.map((r, i) => {
      const target = r.quote ? `"${r.quote}"` : r.issue;
      return `${i + 1}. BỎ QUA (KHÔNG báo lỗi) đối với ${target}. Lý do: ${r.reason}`;
    });
    return `\nCÁC LƯU Ý NGOẠI LỆ DO NGƯỜI DÙNG QUY ĐỊNH (TUYỆT ĐỐI KHÔNG BÁO LỖI TƯƠNG TỰ):\n${lines.join("\n")}\n`;
  }

  // --- Từ điển riêng cho kiểm tra chính tả Offline ---
  function getCustomDict() {
    try {
      const arr = JSON.parse(localStorage.getItem("custom_spell_dict") || "[]");
      return Array.isArray(arr) ? arr : [];
    } catch (e) {
      return [];
    }
  }

  function saveCustomDict(words) {
    localStorage.setItem("custom_spell_dict", JSON.stringify(words));
    renderCustomDictManager();
  }

  function addCustomDictWord(word) {
    const clean = (word || "").trim();
    if (!clean) return false;
    const list = getCustomDict();
    const exists = list.some((w) => w.toLowerCase() === clean.toLowerCase());
    if (exists) return false;
    list.push(clean);
    saveCustomDict(list);
    return true;
  }

  function removeCustomDictWord(word) {
    const list = getCustomDict().filter((w) => w.toLowerCase() !== (word || "").toLowerCase());
    saveCustomDict(list);
  }

  function renderCustomDictManager() {
    const toggleBtn = document.getElementById("toggleCustomDictBtn");
    const listDiv = document.getElementById("customDictList");
    if (!toggleBtn || !listDiv) return;

    const list = getCustomDict();
    toggleBtn.textContent = `📖 Từ điển riêng Offline (${list.length})`;
    listDiv.innerHTML = "";

    if (list.length === 0) {
      listDiv.innerHTML = `<div class="hint" style="margin-bottom:4px;">Chưa có từ riêng nào. Thêm tên riêng, viết tắt đơn vị để bỏ qua khi kiểm tra Offline.</div>`;
      return;
    }

    list.forEach((word) => {
      const chip = document.createElement("span");
      chip.className = "dict-chip";
      chip.innerHTML = `<span>${PCDT.escapeHtml(word)}</span>`;

      const delBtn = document.createElement("button");
      delBtn.className = "dict-chip-del";
      delBtn.title = "Xóa khỏi từ điển riêng";
      delBtn.textContent = "×";
      delBtn.onclick = () => removeCustomDictWord(word);

      chip.appendChild(delBtn);
      listDiv.appendChild(chip);
    });
  }

  // --- 8 thông số mặc định NĐ 30 ---
  function getND30Defaults() {
    try {
      const saved = JSON.parse(localStorage.getItem("nd30_default_settings") || "null");
      if (!saved || typeof saved !== "object") return { ...DEFAULT_ND30_SETTINGS };
      return { ...DEFAULT_ND30_SETTINGS, ...saved };
    } catch (e) {
      return { ...DEFAULT_ND30_SETTINGS };
    }
  }

  function saveND30Defaults(settings) {
    localStorage.setItem("nd30_default_settings", JSON.stringify(settings));
  }

  PCDT.DEFAULT_ND30_SETTINGS = DEFAULT_ND30_SETTINGS;
  PCDT.saveApiKey = saveApiKey;
  PCDT.updateRulesStatusUI = updateRulesStatusUI;
  PCDT.handleRulesFileUpload = handleRulesFileUpload;
  PCDT.clearCustomRules = clearCustomRules;
  PCDT.getIssueKey = getIssueKey;
  PCDT.getGlobalIgnoredList = getGlobalIgnoredList;
  PCDT.saveGlobalIgnoredList = saveGlobalIgnoredList;
  PCDT.addGlobalIgnoredRule = addGlobalIgnoredRule;
  PCDT.addSessionIgnoredKey = addSessionIgnoredKey;
  PCDT.isIssueIgnored = isIssueIgnored;
  PCDT.updateRestoreBarUI = updateRestoreBarUI;
  PCDT.restoreSessionIgnored = restoreSessionIgnored;
  PCDT.renderGlobalIgnoreManager = renderGlobalIgnoreManager;
  PCDT.exportGlobalIgnoreTxt = exportGlobalIgnoreTxt;
  PCDT.handleImportGlobalIgnoreTxt = handleImportGlobalIgnoreTxt;
  PCDT.buildGlobalIgnorePromptSection = buildGlobalIgnorePromptSection;
  PCDT.getCustomDict = getCustomDict;
  PCDT.addCustomDictWord = addCustomDictWord;
  PCDT.removeCustomDictWord = removeCustomDictWord;
  PCDT.renderCustomDictManager = renderCustomDictManager;
  PCDT.getND30Defaults = getND30Defaults;
  PCDT.saveND30Defaults = saveND30Defaults;
})(window.PCDT);
