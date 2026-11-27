window.PCDT = window.PCDT || {};

(function (PCDT) {
  PCDT.lastRawGroups = null;

  function updateConnectionBadge() {
    const badge = document.getElementById("connBadge");
    if (!badge) return;
    if (navigator.onLine) {
      badge.textContent = "Online";
      badge.className = "conn-badge conn-online";
    } else {
      badge.textContent = "Offline";
      badge.className = "conn-badge conn-offline";
    }
  }

  function updateNetworkModeBadge(isOfflineMode, reasonText) {
    const badge = document.getElementById("modeBadge");
    if (!badge) return;
    if (isOfflineMode) {
      badge.style.display = "block";
      badge.className = "mode-badge mode-offline";
      badge.textContent = reasonText || "📴 Chế độ Offline: Đang dùng bộ kiểm tra chính tả 3 lớp cục bộ.";
    } else {
      badge.style.display = "none";
    }
  }

  function renderFilteredGroups(rawGroups, docTypeLabel, statusOverride) {
    const filteredGroups = rawGroups.map((g) => ({
      ...g,
      items: (g.items || []).filter((item) => !PCDT.isIssueIgnored(item))
    }));

    const totalVisible = filteredGroups.reduce((sum, g) => sum + g.items.length, 0);
    PCDT.updateRestoreBarUI();

    if (totalVisible === 0) {
      document.getElementById("results").innerHTML = "";
      PCDT.setStatus(
        statusOverride ||
          (docTypeLabel
            ? `Hoàn tất 3 bước (${docTypeLabel})! Không có lỗi mới cần rà soát.`
            : "Không có lỗi nào cần rà soát.")
      );
    } else {
      PCDT.setStatus(
        statusOverride ||
          (docTypeLabel
            ? `Hoàn tất 3 bước (${docTypeLabel})! Còn ${totalVisible} lỗi cần rà soát:`
            : `Còn ${totalVisible} lỗi cần rà soát:`)
      );
      renderGroupedIssues(filteredGroups);
    }
  }

  function renderGroupedIssues(groups) {
    const container = document.getElementById("results");
    container.innerHTML = "";

    function updateAllCounts() {
      const remainingTotal = container.querySelectorAll(".issue-item:not(.issue-fixed)").length;
      if (remainingTotal === 0) {
        PCDT.setStatus("Tuyệt vời! Toàn bộ các lỗi đã được sửa hoặc bỏ qua.");
      } else {
        PCDT.setStatus(`Còn lại ${remainingTotal} lỗi chưa xử lý:`);
      }
    }

    groups.forEach((group) => {
      if (!group.items || group.items.length === 0) return;

      const groupWrapper = document.createElement("div");
      groupWrapper.className = "result-group";

      const header = document.createElement("div");
      header.className = `group-header ${group.headerClass}`;

      const titleSpan = document.createElement("span");
      titleSpan.textContent = group.title;

      const countBadge = document.createElement("span");
      countBadge.className = "group-count";
      countBadge.textContent = `${group.items.length} lỗi`;

      header.appendChild(titleSpan);
      header.appendChild(countBadge);
      groupWrapper.appendChild(header);

      const listContainer = document.createElement("div");
      header.onclick = () => {
        listContainer.style.display = listContainer.style.display === "none" ? "block" : "none";
      };

      function updateGroupCount() {
        const unfixed = listContainer.querySelectorAll(".issue-item:not(.issue-fixed)").length;
        const totalCards = listContainer.querySelectorAll(".issue-item").length;
        if (totalCards === 0) {
          groupWrapper.remove();
        } else if (unfixed === 0) {
          countBadge.textContent = "✅ Đã xong";
        } else {
          countBadge.textContent = `${unfixed} lỗi`;
        }
        updateAllCounts();
      }

      group.items.forEach((item) => {
        const card = document.createElement("div");
        card.className = "card issue-item";

        const quoteEl = document.createElement("span");
        quoteEl.className = "quote";
        quoteEl.textContent = `"${item.quote || "Cấu hình trang / Toàn văn bản"}"`;

        const issueEl = document.createElement("div");
        issueEl.innerHTML = `<strong>Lỗi:</strong> `;
        issueEl.appendChild(document.createTextNode(item.issue));

        const sugEl = document.createElement("div");
        sugEl.innerHTML = `<strong>Đề xuất:</strong> `;
        sugEl.appendChild(document.createTextNode(item.suggestion));

        const recheckMsg = document.createElement("div");
        recheckMsg.className = "recheck-msg";
        recheckMsg.style.display = "none";

        card.appendChild(quoteEl);
        card.appendChild(issueEl);
        card.appendChild(sugEl);

        const btnGroup = document.createElement("div");
        btnGroup.className = "btn-group";

        function markCardAsFixed(msgText) {
          card.classList.add("issue-fixed");
          recheckMsg.textContent = `✅ Đã sửa: ${msgText}`;
          recheckMsg.style.color = "#15803d";
          recheckMsg.style.display = "block";
          updateGroupCount();
        }

        if (item.quote) {
          const locateBtn = document.createElement("button");
          locateBtn.className = "btn-locate";
          locateBtn.textContent = "📍 Đến vị trí";
          locateBtn.onclick = () => PCDT.navigateToIssue(item.quote);
          btnGroup.appendChild(locateBtn);
        }

        if (item.quote || item.checkType) {
          const recheckBtn = document.createElement("button");
          recheckBtn.className = "btn-recheck";
          recheckBtn.textContent = "🔄 Kiểm tra lại";
          recheckBtn.onclick = async () => {
            recheckBtn.textContent = "⏳ Đang kiểm tra...";
            recheckBtn.disabled = true;
            try {
              const res = await PCDT.recheckSingleIssue(item);
              if (res.fixed) {
                recheckBtn.textContent = "✅ Đã sửa";
                recheckBtn.style.background = "#16a34a";
                markCardAsFixed(res.message);
              } else {
                recheckBtn.textContent = "🔄 Kiểm tra lại";
                recheckBtn.disabled = false;
                recheckMsg.textContent = `⚠️ Chưa xong: ${res.message}`;
                recheckMsg.style.color = "#b45309";
                recheckMsg.style.display = "block";
              }
            } catch (err) {
              recheckBtn.textContent = "🔄 Kiểm tra lại";
              recheckBtn.disabled = false;
              recheckMsg.textContent = `Lỗi kiểm tra: ${err.message}`;
              recheckMsg.style.color = "#991b1b";
              recheckMsg.style.display = "block";
            }
          };
          btnGroup.appendChild(recheckBtn);
        }

        if (item.quote) {
          const commentBtn = document.createElement("button");
          commentBtn.className = "btn-comment";
          commentBtn.textContent = "💬 Gắn Comment";
          commentBtn.onclick = () =>
            PCDT.insertWordComment(item.quote, `[${group.title}] ${item.issue}\n👉 Đề xuất: ${item.suggestion}`);
          btnGroup.appendChild(commentBtn);

          if (item.category === "Chính tả" && item.replacement && item.replacement !== item.quote) {
            const fixBtn = document.createElement("button");
            fixBtn.className = "btn-fix";
            fixBtn.textContent = "✏️ Sửa tự động";
            fixBtn.onclick = () =>
              PCDT.applyAutoFix(item.quote, item.replacement, fixBtn, () =>
                markCardAsFixed("Đã sửa tự động trực tiếp trong văn bản.")
              );
            btnGroup.appendChild(fixBtn);
          }
        }

        const ignoreBox = document.createElement("div");
        ignoreBox.className = "ignore-box";
        ignoreBox.style.display = "none";

        const ignoreTitle = document.createElement("div");
        ignoreTitle.className = "ignore-box-title";
        ignoreTitle.textContent = "Chọn phạm vi bỏ qua:";
        ignoreBox.appendChild(ignoreTitle);

        const ignoreOpts = document.createElement("div");
        ignoreOpts.className = "btn-group";

        const thisDocBtn = document.createElement("button");
        thisDocBtn.className = "btn-ignore";
        thisDocBtn.textContent = "📄 Bỏ qua cho văn bản này";
        thisDocBtn.onclick = () => {
          PCDT.addSessionIgnoredKey(PCDT.getIssueKey(item));
          card.remove();
          updateGroupCount();
        };

        const allDocBtn = document.createElement("button");
        allDocBtn.className = "btn-fix";
        allDocBtn.textContent = "🌐 Bỏ qua cho các văn bản khác";

        const reasonArea = document.createElement("div");
        reasonArea.className = "ignore-reason-area";
        reasonArea.style.display = "none";

        const reasonInput = document.createElement("input");
        reasonInput.type = "text";
        reasonInput.placeholder = "Nhập lý do (VD: Tên viết tắt ngành điện, không phải lỗi)...";

        const saveGlobalBtn = document.createElement("button");
        saveGlobalBtn.className = "btn-primary";
        saveGlobalBtn.style.padding = "6px";
        saveGlobalBtn.style.fontSize = "12px";
        saveGlobalBtn.textContent = "💾 Lưu lý do & Bỏ qua cho các văn bản sau";
        saveGlobalBtn.onclick = () => {
          const reason = reasonInput.value.trim();
          if (!reason) return alert("Vui lòng nhập lý do để AI hiểu và áp dụng cho các văn bản sau.");
          PCDT.addGlobalIgnoredRule(item, reason);
          card.remove();
          updateGroupCount();
        };

        allDocBtn.onclick = () => {
          reasonArea.style.display = reasonArea.style.display === "block" ? "none" : "block";
          if (reasonArea.style.display === "block") reasonInput.focus();
        };

        reasonArea.appendChild(reasonInput);
        reasonArea.appendChild(saveGlobalBtn);

        ignoreOpts.appendChild(thisDocBtn);
        ignoreOpts.appendChild(allDocBtn);
        ignoreBox.appendChild(ignoreOpts);
        ignoreBox.appendChild(reasonArea);

        const ignoreBtn = document.createElement("button");
        ignoreBtn.className = "btn-ignore";
        ignoreBtn.textContent = "✖ Bỏ qua";
        ignoreBtn.onclick = () => {
          ignoreBox.style.display = ignoreBox.style.display === "block" ? "none" : "block";
        };
        btnGroup.appendChild(ignoreBtn);

        card.appendChild(btnGroup);
        card.appendChild(ignoreBox);
        card.appendChild(recheckMsg);
        listContainer.appendChild(card);
      });

      groupWrapper.appendChild(listContainer);
      container.appendChild(groupWrapper);
    });
  }

  async function runSequentialCheck() {
    const apiKey = (localStorage.getItem("gemini_api_key") || "").trim();
    const runBtn = document.getElementById("runCheckBtn");
    const resultsDiv = document.getElementById("results");
    const docTypeSelect = document.getElementById("docTypeSelect");
    const docTypeHint = document.getElementById("docTypeHint");

    runBtn.disabled = true;
    resultsDiv.innerHTML = "";
    PCDT.resetStepsUI();
    updateNetworkModeBadge(false);

    try {
      PCDT.setStatus("Đang đọc nội dung văn bản...");
      const rawText = await Word.run(async (context) => {
        return await PCDT.extractCleanDocumentText(context);
      });

      if (!rawText.trim()) {
        PCDT.setStatus("Văn bản đang trống.");
        runBtn.disabled = false;
        return;
      }

      let docType = docTypeSelect.value;
      if (docType === "auto") {
        docType = await PCDT.detectDocumentType(apiKey, rawText);
        docTypeSelect.value = docType;
        docTypeHint.innerHTML = `✅ Đã tự động nhận diện: <strong>${PCDT.escapeHtml(docType)}</strong>. Nếu chưa đúng, hãy chọn lại ở ô trên rồi bấm kiểm tra lại.`;
      } else {
        docTypeHint.innerHTML = `Đang kiểm tra theo loại đã chọn: <strong>${PCDT.escapeHtml(docType)}</strong>.`;
      }

      // BƯỚC 1: Kỹ thuật & Thể thức cứng (100% cục bộ)
      PCDT.setStepState(1, "active");
      PCDT.setStatus("Bước 1/3: Đang kiểm tra kỹ thuật (Lề, Font, Cỡ chữ, Quy tắc cứng)...");
      const { fullText, step1Issues } = await PCDT.runStep1TechnicalAndHardRules(docType);
      PCDT.setStepState(1, "done");

      // BƯỚC 2: Chính tả (Online Gemini AI hoặc tự động dự phòng Offline 3 lớp)
      PCDT.setStepState(2, "active");
      let step2Issues = [];
      let usedOfflineStep2 = false;
      let offlineReason = "";

      if (!navigator.onLine) {
        usedOfflineStep2 = true;
        offlineReason = "📴 Đang mất kết nối mạng: Bước 2 tự động chạy Kiểm tra chính tả Offline (3 lớp).";
      } else if (!apiKey) {
        usedOfflineStep2 = true;
        offlineReason = "🔑 Chưa nhập Gemini API Key: Bước 2 tự động chạy Kiểm tra chính tả Offline (3 lớp).";
      } else {
        try {
          step2Issues = await PCDT.runStep2Spelling(apiKey, fullText);
        } catch (aiErr) {
          usedOfflineStep2 = true;
          offlineReason = `⚠️ Không kết nối được Gemini AI (${aiErr.message}). Đã tự động chuyển sang Kiểm tra chính tả Offline (3 lớp).`;
        }
      }

      if (usedOfflineStep2) {
        updateNetworkModeBadge(true, offlineReason);
        step2Issues = await PCDT.runStep2OfflineSpelling(fullText);
      }
      PCDT.setStepState(2, "done");

      // BƯỚC 3: Thể thức & Hợp lệ (Gemini AI — bỏ qua êm khi mất mạng hoặc thiếu API Key)
      let step3Issues = [];
      let skippedStep3 = false;

      if (!navigator.onLine || !apiKey) {
        skippedStep3 = true;
        const step3Pill = document.getElementById("step3Pill");
        if (step3Pill) step3Pill.textContent = "B3. Bỏ qua (Offline)";
      } else {
        PCDT.setStepState(3, "active");
        try {
          step3Issues = await PCDT.runStep3FormatAndValidity(apiKey, fullText, docType);
          PCDT.setStepState(3, "done");
        } catch (step3Err) {
          skippedStep3 = true;
          updateNetworkModeBadge(
            true,
            `⚠️ Bước 3 tạm bỏ qua do lỗi kết nối/API (${step3Err.message}). Đã hoàn tất Bước 1 & Bước 2.`
          );
        }
      }

      PCDT.lastRawGroups = [
        {
          id: "group-tech",
          title: "📐 Nhóm 1: Kỹ thuật & Thể thức cứng",
          headerClass: "group-tech",
          items: step1Issues
        },
        {
          id: "group-spell",
          title: usedOfflineStep2 ? "✍️ Nhóm 2: Chính tả & Ngôn ngữ (Offline 3 lớp)" : "✍️ Nhóm 2: Chính tả & Ngôn ngữ",
          headerClass: "group-spell",
          items: step2Issues
        },
        {
          id: "group-valid",
          title: `⚖️ Nhóm 3: Thể thức & Tính hợp lệ (${docType})`,
          headerClass: "group-valid",
          items: step3Issues
        }
      ];

      const summaryNote = skippedStep3
        ? `Hoàn tất Bước 1 & Bước 2 Offline (${docType}) — Bước 3 tạm bỏ qua do chưa có mạng/API Key.`
        : null;

      renderFilteredGroups(PCDT.lastRawGroups, docType, summaryNote);
    } catch (err) {
      PCDT.setStatus("Có lỗi xảy ra: " + err.message);
    } finally {
      runBtn.disabled = false;
    }
  }

  Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
      const savedKey = localStorage.getItem("gemini_api_key");
      if (savedKey) document.getElementById("apiKey").value = savedKey;

      PCDT.updateRulesStatusUI();
      PCDT.renderGlobalIgnoreManager();
      PCDT.renderCustomDictManager();
      PCDT.loadND30DefaultsToUI();
      updateConnectionBadge();
      window.addEventListener("online", updateConnectionBadge);
      window.addEventListener("offline", updateConnectionBadge);

      document.getElementById("saveKeyBtn").onclick = PCDT.saveApiKey;
      document.getElementById("rulesFileInput").onchange = PCDT.handleRulesFileUpload;
      document.getElementById("clearRulesBtn").onclick = PCDT.clearCustomRules;
      document.getElementById("runCheckBtn").onclick = runSequentialCheck;

      // Kiểm tra & Chuẩn hóa NĐ 30 (2 bước + Lưu mặc định)
      document.getElementById("autoFormatBtn").onclick = PCDT.scanFormatND30;
      document.getElementById("confirmFormatBtn").onclick = PCDT.applySelectedFormatND30;
      document.getElementById("saveFormatDefaultsBtn").onclick = PCDT.saveND30DefaultsFromUI;
      document.getElementById("closeFormatPanelBtn").onclick = () => {
        document.getElementById("formatReviewPanel").style.display = "none";
      };

      // Bỏ qua 2 cấp độ
      document.getElementById("restoreSessionBtn").onclick = PCDT.restoreSessionIgnored;
      document.getElementById("toggleGlobalIgnoreBtn").onclick = () => {
        const panel = document.getElementById("globalIgnorePanel");
        panel.style.display = panel.style.display === "block" ? "none" : "block";
      };
      document.getElementById("exportGlobalIgnoreBtn").onclick = PCDT.exportGlobalIgnoreTxt;
      document.getElementById("importGlobalIgnoreBtn").onclick = () => {
        document.getElementById("importGlobalIgnoreInput").click();
      };
      document.getElementById("importGlobalIgnoreInput").onchange = PCDT.handleImportGlobalIgnoreTxt;

      // Từ điển riêng Offline
      document.getElementById("toggleCustomDictBtn").onclick = () => {
        const panel = document.getElementById("customDictPanel");
        panel.style.display = panel.style.display === "block" ? "none" : "block";
      };
      document.getElementById("addCustomDictBtn").onclick = () => {
        const input = document.getElementById("customDictInput");
        if (PCDT.addCustomDictWord(input.value)) {
          input.value = "";
          PCDT.setStatus("Đã thêm từ vào từ điển riêng Offline.");
        } else {
          alert("Vui lòng nhập từ hợp lệ (hoặc từ đã tồn tại).");
        }
      };
    }
  });

  PCDT.renderFilteredGroups = renderFilteredGroups;
})(window.PCDT);
