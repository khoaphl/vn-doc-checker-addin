window.PCDT = window.PCDT || {};

(function (PCDT) {
  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  let lastScannedCfg = null;

  function setFormatProgress(percent, text) {
    let box = document.getElementById("formatProgressBox");
    const confirmBtn = document.getElementById("confirmFormatBtn");
    const actionRow = confirmBtn ? confirmBtn.closest(".io-row") : null;
    const anchor = actionRow || confirmBtn;
    if (!box && anchor && anchor.parentNode) {
      box = document.createElement("div");
      box.id = "formatProgressBox";
      box.style.cssText =
        "width:100%; box-sizing:border-box; margin-top:10px; margin-bottom:2px; padding:8px 10px; background:#eff6ff; border:1px solid #bfdbfe; border-radius:6px;";
      box.innerHTML = `
        <div style="display:flex; justify-content:space-between; font-size:11.5px; font-weight:600; color:#1e40af; margin-bottom:5px;">
          <span id="formatProgressLabel">Đang chuẩn hóa...</span>
          <span id="formatProgressPct">0%</span>
        </div>
        <div style="width:100%; height:7px; background:#dbeafe; border-radius:999px; overflow:hidden;">
          <div id="formatProgressBar" style="width:0%; height:100%; background:#2563eb; transition:width 0.15s ease;"></div>
        </div>
      `;
      anchor.parentNode.insertBefore(box, anchor);
    }
    if (box) {
      box.style.display = "block";
      const labelEl = document.getElementById("formatProgressLabel");
      const pctEl = document.getElementById("formatProgressPct");
      const barEl = document.getElementById("formatProgressBar");
      const clamped = Math.max(0, Math.min(100, Math.round(percent)));
      if (labelEl && text) labelEl.textContent = text;
      if (pctEl) pctEl.textContent = `${clamped}%`;
      if (barEl) barEl.style.width = `${clamped}%`;
    }
    if (confirmBtn) {
      confirmBtn.disabled = true;
      confirmBtn.textContent = `⏳ Đang chuẩn hóa (${Math.round(percent)}%)...`;
    }
    if (text) PCDT.setStatus(text);
  }

  function clearFormatProgress() {
    const box = document.getElementById("formatProgressBox");
    if (box) box.style.display = "none";
    const confirmBtn = document.getElementById("confirmFormatBtn");
    if (confirmBtn) confirmBtn.disabled = false;
  }

  function isBulletTextPrefix(text) {
    return /^\s*[-–+•](?:\t|\s+)/.test(text || "");
  }

  function cmToTwips(cm) {
    return Math.round(cm * 566.929);
  }

  function getTabOffsetCmForListString(listStr) {
    const len = (listStr || "").trim().length;
    if (len >= 4) return 1.0;
    if (len >= 3) return 0.75;
    return 0.5;
  }

  function updateParagraphTabStopOoxml(ooxmlStr, indentCm, isListItem, tabOffsetCm) {
    const offset = typeof tabOffsetCm === "number" ? tabOffsetCm : 0.5;
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(ooxmlStr, "application/xml");
    if (xmlDoc.getElementsByTagName("parsererror").length > 0) return null;

    const pNodes = xmlDoc.getElementsByTagNameNS(W_NS, "p");
    if (!pNodes || pNodes.length === 0) return null;
    const pNode = pNodes[0];

    let pPr = null;
    for (let i = 0; i < pNode.childNodes.length; i++) {
      if (pNode.childNodes[i].localName === "pPr" && pNode.childNodes[i].namespaceURI === W_NS) {
        pPr = pNode.childNodes[i];
        break;
      }
    }
    if (!pPr) {
      pPr = xmlDoc.createElementNS(W_NS, "w:pPr");
      pNode.insertBefore(pPr, pNode.firstChild);
    }

    // Đặt Tab stop = (indentCm + offset) cm (0.5 / 0.75 / 1.0 cm tùy độ dài ký hiệu)
    const oldTabs = pPr.getElementsByTagNameNS(W_NS, "tabs");
    while (oldTabs.length > 0) {
      oldTabs[0].parentNode.removeChild(oldTabs[0]);
    }
    const tabsEl = xmlDoc.createElementNS(W_NS, "w:tabs");
    const tabEl = xmlDoc.createElementNS(W_NS, "w:tab");
    tabEl.setAttributeNS(W_NS, "w:val", "left");
    tabEl.setAttributeNS(W_NS, "w:pos", String(cmToTwips(indentCm + offset)));
    tabsEl.appendChild(tabEl);
    pPr.appendChild(tabsEl);

    if (!isListItem) {
      // Nếu là dòng gõ tay "- Nội dung" (không phải Word List): đặt Left=0, First line=indentCm và chuẩn hóa "- " thành "-<tab>"
      const oldInd = pPr.getElementsByTagNameNS(W_NS, "ind");
      while (oldInd.length > 0) {
        oldInd[0].parentNode.removeChild(oldInd[0]);
      }
      const indEl = xmlDoc.createElementNS(W_NS, "w:ind");
      indEl.setAttributeNS(W_NS, "w:left", "0");
      indEl.setAttributeNS(W_NS, "w:firstLine", String(cmToTwips(indentCm)));
      pPr.appendChild(indEl);

      const tNodes = pNode.getElementsByTagNameNS(W_NS, "t");
      if (tNodes.length > 0) {
        const firstT = tNodes[0];
        const rawVal = firstT.textContent || "";
        const m = rawVal.match(/^\s*([-–+•])(\s+)(.*)$/);
        if (m) {
          firstT.textContent = m[1];
          const parentRun = firstT.parentNode;
          const rPr = parentRun.getElementsByTagNameNS(W_NS, "rPr")[0];

          const tabRun = xmlDoc.createElementNS(W_NS, "w:r");
          if (rPr) tabRun.appendChild(rPr.cloneNode(true));
          tabRun.appendChild(xmlDoc.createElementNS(W_NS, "w:tab"));

          const restRun = xmlDoc.createElementNS(W_NS, "w:r");
          if (rPr) restRun.appendChild(rPr.cloneNode(true));
          const restT = xmlDoc.createElementNS(W_NS, "w:t");
          restT.setAttribute("xml:space", "preserve");
          restT.textContent = m[3];
          restRun.appendChild(restT);

          const nextSib = parentRun.nextSibling;
          pNode.insertBefore(tabRun, nextSib);
          if (m[3]) pNode.insertBefore(restRun, nextSib);
        }
      }
    }

    return new XMLSerializer().serializeToString(xmlDoc);
  }

  function isBodyContentParagraph(p, inTable) {
    if (inTable) return false;
    const text = (p.text || "").replace(/[\v\u000b\r]/g, " ").trim();
    if (!text) return false;
    // Bỏ qua các dòng tiêu đề/ký tên đang canh giữa hoặc canh phải
    if (p.alignment === Word.Alignment.centered || p.alignment === Word.Alignment.right) return false;
    // Bỏ qua các dòng thể thức ngắn đặc thù (Số, Nơi nhận, tên loại văn bản) để không làm sai cỡ chữ/thụt dòng
    if (/^(Số\s*:|Nơi nhận\s*:|-\s*(Như\s+|Lưu\s*:))/i.test(text)) return false;
    if (/^(QUYẾT ĐỊNH|TỜ TRÌNH|BÁO CÁO|THÔNG BÁO|KẾ HOẠCH|BIÊN BẢN|CÔNG VĂN|PHỤ LỤC)\s*:?$/i.test(text)) return false;
    return true;
  }

  function loadND30DefaultsToUI() {
    const cfg = PCDT.getND30Defaults();
    const setVal = (id, val) => {
      const el = document.getElementById(id);
      if (!el || val === undefined || val === null) return;
      const numVal = Number(val);
      const matchedOpt = Array.from(el.options).find(
        (opt) =>
          opt.value === String(val) ||
          (!Number.isNaN(numVal) && Math.abs(parseFloat(opt.value) - numVal) < 0.001)
      );
      if (matchedOpt) el.value = matchedOpt.value;
    };
    setVal("optFontSize", cfg.fontSize);
    setVal("optIndent", cfg.indent);
    setVal("optLineSpacing", cfg.lineSpacing);
    setVal("optSpaceAfter", cfg.spaceBeforeAfter);
    setVal("optMarginTop", cfg.marginTop);
    setVal("optMarginBottom", cfg.marginBottom);
    setVal("optMarginLeft", cfg.marginLeft);
    setVal("optMarginRight", cfg.marginRight);
  }

  function setOptionMismatchUI(selectId, isMismatch, currentLabel, defaultLabel) {
    const selectEl = document.getElementById(selectId);
    if (!selectEl) return;
    const field = selectEl.closest(".opt-field");
    if (!field) return;

    let hintEl = field.querySelector(".opt-mismatch-hint");
    if (isMismatch) {
      field.classList.add("opt-mismatch");
      if (!hintEl) {
        hintEl = document.createElement("div");
        hintEl.className = "opt-mismatch-hint";
        field.appendChild(hintEl);
      }
      hintEl.textContent = `Hiện tại: ${currentLabel} → Mặc định: ${defaultLabel}`;
    } else {
      field.classList.remove("opt-mismatch");
      if (hintEl) hintEl.remove();
    }
  }

  function clearAllOptionMismatches() {
    [
      "optFontSize",
      "optIndent",
      "optLineSpacing",
      "optSpaceAfter",
      "optMarginTop",
      "optMarginBottom",
      "optMarginLeft",
      "optMarginRight"
    ].forEach((id) => setOptionMismatchUI(id, false));
  }

  function pickMostFrequentLabel(countsMap) {
    let bestKey = "";
    let bestCount = -1;
    countsMap.forEach((count, key) => {
      if (count > bestCount) {
        bestCount = count;
        bestKey = key;
      }
    });
    return bestKey;
  }

  function readND30SettingsFromUI() {
    return {
      fontSize: parseFloat(document.getElementById("optFontSize").value) || 13,
      indent: parseFloat(document.getElementById("optIndent").value) || 1.27,
      lineSpacing: parseFloat(document.getElementById("optLineSpacing").value) || 1.15,
      spaceBeforeAfter: parseFloat(document.getElementById("optSpaceAfter").value) || 6,
      marginTop: parseFloat(document.getElementById("optMarginTop").value) || 20,
      marginBottom: parseFloat(document.getElementById("optMarginBottom").value) || 20,
      marginLeft: parseFloat(document.getElementById("optMarginLeft").value) || 30,
      marginRight: parseFloat(document.getElementById("optMarginRight").value) || 15
    };
  }

  function saveND30DefaultsFromUI() {
    const settings = readND30SettingsFromUI();
    PCDT.saveND30Defaults(settings);
    PCDT.setStatus(
      `💾 Đã lưu mặc định 8 thông số NĐ 30 (Cách đoạn Before & After = ${settings.spaceBeforeAfter} pt).`
    );
  }

  async function scanFormatND30() {
    PCDT.setStatus("Đang quét thông số kỹ thuật theo NĐ 30/2020/NĐ-CP...");
    const panel = document.getElementById("formatReviewPanel");
    const listDiv = document.getElementById("formatIssuesList");
    const optionsBox = document.getElementById("formatOptionsBox");
    const confirmBtn = document.getElementById("confirmFormatBtn");

    const targetCfg = readND30SettingsFromUI();
    lastScannedCfg = { ...targetCfg };
    const targetSpace = targetCfg.spaceBeforeAfter;

    try {
      await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        const sections = context.document.sections;

        paragraphs.load(
          "items/text, items/font/name, items/font/size, items/font/color, items/alignment, items/firstLineIndent, items/leftIndent, items/isListItem, items/spaceAfter, items/spaceBefore, items/lineSpacing, items/parentTableOrNullObject/isNullObject"
        );

        const supportsPageSetup = Office.context.requirements.isSetSupported("WordApi", "1.5");
        if (supportsPageSetup) {
          sections.load(
            "items/pageSetup/topMargin, items/pageSetup/bottomMargin, items/pageSetup/leftMargin, items/pageSetup/rightMargin, items/pageSetup/pageWidth, items/pageSetup/pageHeight"
          );
        }

        await context.sync();

        const findings = [];
        const mismatchMeta = {
          marginTop: { mismatched: false, current: "" },
          marginBottom: { mismatched: false, current: "" },
          marginLeft: { mismatched: false, current: "" },
          marginRight: { mismatched: false, current: "" }
        };

        // 1. Khổ giấy A4 (210 x 297 mm, cho phép dọc hoặc ngang) & Căn lề theo mặc định
        if (supportsPageSetup) {
          let wrongPaperCount = 0;
          let wrongMarginCount = 0;
          let sampleMargin = "";

          sections.items.forEach((sec) => {
            const ps = sec.pageSetup;
            if (ps.pageWidth && ps.pageHeight) {
              const w = PCDT.ptToMm(ps.pageWidth);
              const h = PCDT.ptToMm(ps.pageHeight);
              const isPortraitA4 = Math.abs(w - 210) <= 3 && Math.abs(h - 297) <= 3;
              const isLandscapeA4 = Math.abs(w - 297) <= 3 && Math.abs(h - 210) <= 3;
              if (!isPortraitA4 && !isLandscapeA4) wrongPaperCount++;
            }

            const top = PCDT.ptToMm(ps.topMargin);
            const bottom = PCDT.ptToMm(ps.bottomMargin);
            const left = PCDT.ptToMm(ps.leftMargin);
            const right = PCDT.ptToMm(ps.rightMargin);
            const topDiff = Math.abs(top - targetCfg.marginTop) > 1;
            const bottomDiff = Math.abs(bottom - targetCfg.marginBottom) > 1;
            const leftDiff = Math.abs(left - targetCfg.marginLeft) > 1;
            const rightDiff = Math.abs(right - targetCfg.marginRight) > 1;

            if (topDiff && !mismatchMeta.marginTop.mismatched) {
              mismatchMeta.marginTop = { mismatched: true, current: `${top} mm` };
            }
            if (bottomDiff && !mismatchMeta.marginBottom.mismatched) {
              mismatchMeta.marginBottom = { mismatched: true, current: `${bottom} mm` };
            }
            if (leftDiff && !mismatchMeta.marginLeft.mismatched) {
              mismatchMeta.marginLeft = { mismatched: true, current: `${left} mm` };
            }
            if (rightDiff && !mismatchMeta.marginRight.mismatched) {
              mismatchMeta.marginRight = { mismatched: true, current: `${right} mm` };
            }

            if (topDiff || bottomDiff || leftDiff || rightDiff) {
              wrongMarginCount++;
              if (!sampleMargin) sampleMargin = `Hiện tại: Trên ${top}, Dưới ${bottom}, Trái ${left}, Phải ${right}mm`;
            }
          });

          if (wrongPaperCount > 0) {
            findings.push({
              id: "fmt_paper",
              label: `Khổ giấy chưa đúng chuẩn A4 (${wrongPaperCount} phần/section)`,
              detail: "Chuẩn NĐ 30: Khổ A4 (210 × 297 mm, giữ nguyên hướng ngang nếu có bảng rộng)."
            });
          }
          if (wrongMarginCount > 0) {
            findings.push({
              id: "fmt_margin",
              label: `Căn lề trang chưa khớp mặc định (${wrongMarginCount} phần/section)`,
              detail: `${sampleMargin}. Mặc định: Trên ${targetCfg.marginTop}, Dưới ${targetCfg.marginBottom}, Trái ${targetCfg.marginLeft}, Phải ${targetCfg.marginRight} mm.`
            });
          }
        }

        // 2. Quét các đoạn văn bản (phân biệt ngoài bảng và trong bảng)
        let wrongFontCount = 0;
        let wrongSizeCount = 0;
        let wrongAlignCount = 0;
        let wrongIndentCount = 0;
        let wrongSpacingCount = 0;

        const badSizeCounts = new Map();
        const badIndentCounts = new Map();
        const badLineCounts = new Map();
        const badSpaceCounts = new Map();
        const targetIndentPt = PCDT.cmToPt(targetCfg.indent);
        const targetLinePt = 12 * targetCfg.lineSpacing;

        paragraphs.items.forEach((p) => {
          const text = (p.text || "").replace(/[\v\u000b\r]/g, " ").trim();
          if (!text) return;

          const inTable = p.parentTableOrNullObject && !p.parentTableOrNullObject.isNullObject;

          // Phông chữ kiểm tra cả trong bảng và ngoài bảng
          if (p.font.name && p.font.name !== "Times New Roman") {
            wrongFontCount++;
          }

          // Các tiêu chí nội dung chỉ kiểm tra cho đoạn nội dung NGOÀI BẢNG
          if (isBodyContentParagraph(p, inTable)) {
            if (p.font.size && Math.abs(p.font.size - targetCfg.fontSize) > 0.5) {
              wrongSizeCount++;
              const lbl = `${Math.round(p.font.size * 10) / 10} pt`;
              badSizeCounts.set(lbl, (badSizeCounts.get(lbl) || 0) + 1);
            }
            if (p.alignment !== Word.Alignment.justified) {
              wrongAlignCount++;
            }
            const leftPt = p.leftIndent || 0;
            const indent = p.firstLineIndent || 0;
            const okLeft0 = Math.abs(leftPt) <= 2.5;
            const okTargetIndent = Math.abs(indent - targetIndentPt) <= 2.5;
            if (!okLeft0 || !okTargetIndent) {
              wrongIndentCount++;
              const indentCm = Math.round((indent / 28.3464567) * 100) / 100;
              const lbl = !okLeft0
                ? `Left ${Math.round((leftPt / 28.3464567) * 100) / 100} cm`
                : `${String(indentCm).replace(".", ",")} cm`;
              badIndentCounts.set(lbl, (badIndentCounts.get(lbl) || 0) + 1);
            }
            const spaceAfter = p.spaceAfter || 0;
            const spaceBefore = p.spaceBefore || 0;
            const lineSp = p.lineSpacing || 12;
            const beforeMismatch = Math.abs(spaceBefore - targetSpace) > 0.5;
            const afterMismatch = Math.abs(spaceAfter - targetSpace) > 0.5;
            const lineMismatch = Math.abs(lineSp - targetLinePt) > 1.0;
            if (beforeMismatch || afterMismatch) {
              const bVal = Math.round(spaceBefore * 10) / 10;
              const aVal = Math.round(spaceAfter * 10) / 10;
              const lbl = bVal === aVal ? `${bVal} pt` : `${bVal}/${aVal} pt`;
              badSpaceCounts.set(lbl, (badSpaceCounts.get(lbl) || 0) + 1);
            }
            if (lineMismatch) {
              const lineNum = Math.round((lineSp / 12) * 100) / 100;
              const lbl = `${String(lineNum).replace(".", ",")} lines`;
              badLineCounts.set(lbl, (badLineCounts.get(lbl) || 0) + 1);
            }
            if (beforeMismatch || afterMismatch || lineMismatch) {
              wrongSpacingCount++;
            }
          }
        });

        setOptionMismatchUI("optFontSize", badSizeCounts.size > 0, pickMostFrequentLabel(badSizeCounts), `${targetCfg.fontSize} pt`);
        setOptionMismatchUI("optIndent", badIndentCounts.size > 0, pickMostFrequentLabel(badIndentCounts), `${String(targetCfg.indent).replace(".", ",")} cm`);
        setOptionMismatchUI("optLineSpacing", badLineCounts.size > 0, pickMostFrequentLabel(badLineCounts), `${String(targetCfg.lineSpacing).replace(".", ",")} lines`);
        setOptionMismatchUI("optSpaceAfter", badSpaceCounts.size > 0, pickMostFrequentLabel(badSpaceCounts), `${targetSpace} pt`);
        setOptionMismatchUI("optMarginTop", mismatchMeta.marginTop.mismatched, mismatchMeta.marginTop.current, `${targetCfg.marginTop} mm`);
        setOptionMismatchUI("optMarginBottom", mismatchMeta.marginBottom.mismatched, mismatchMeta.marginBottom.current, `${targetCfg.marginBottom} mm`);
        setOptionMismatchUI("optMarginLeft", mismatchMeta.marginLeft.mismatched, mismatchMeta.marginLeft.current, `${targetCfg.marginLeft} mm`);
        setOptionMismatchUI("optMarginRight", mismatchMeta.marginRight.mismatched, mismatchMeta.marginRight.current, `${targetCfg.marginRight} mm`);

        if (wrongFontCount > 0) {
          findings.push({
            id: "fmt_font",
            label: `Phông chữ khác Times New Roman (${wrongFontCount} đoạn, kể cả trong bảng)`,
            detail: "Đổi toàn bộ sang phông Times New Roman chuẩn Unicode (và màu chữ đen #000000 cho nội dung)."
          });
        }
        if (wrongSizeCount > 0) {
          findings.push({
            id: "fmt_size",
            label: `Cỡ chữ nội dung chưa khớp ${targetCfg.fontSize} pt (${wrongSizeCount} đoạn ngoài bảng)`,
            detail: `Chuẩn hóa về cỡ chữ mặc định ${targetCfg.fontSize} pt (đề xuất bên dưới).`
          });
        }
        if (wrongAlignCount > 0) {
          findings.push({
            id: "fmt_align",
            label: `Đoạn nội dung chưa canh đều 2 lề (${wrongAlignCount} đoạn ngoài bảng)`,
            detail: "Chuyển sang canh đều 2 lề (Justify)."
          });
        }
        if (wrongIndentCount > 0) {
          findings.push({
            id: "fmt_indent",
            label: `Thụt đầu dòng chưa khớp ${String(targetCfg.indent).replace(".", ",")} cm (${wrongIndentCount} đoạn ngoài bảng)`,
            detail: `Chuẩn hóa thụt đầu dòng về mức mặc định ${String(targetCfg.indent).replace(".", ",")} cm.`
          });
        }
        if (wrongSpacingCount > 0) {
          findings.push({
            id: "fmt_spacing",
            label: `Cách đoạn (${targetSpace} pt) hoặc Giãn dòng (${String(targetCfg.lineSpacing).replace(".", ",")}) chưa khớp (${wrongSpacingCount} đoạn)`,
            detail: `Đặt đồng bộ Before = ${targetSpace} pt, After = ${targetSpace} pt và giãn dòng ${String(targetCfg.lineSpacing).replace(".", ",")} lines.`
          });
        }

        listDiv.innerHTML = "";
        panel.style.display = "block";

        if (findings.length === 0) {
          listDiv.innerHTML = `<div style="color:#15803d; font-weight:600; padding:6px 0;">✅ Tuyệt vời! Văn bản đã đạt đầy đủ các tiêu chuẩn kỹ thuật NĐ 30/2020/NĐ-CP (Cách đoạn Before & After = ${targetSpace} pt).</div>`;
          optionsBox.style.display = "block";
          confirmBtn.textContent = "🔄 Áp dụng theo thông số đang chọn";
          confirmBtn.style.display = "block";
          PCDT.setStatus("Kiểm tra kỹ thuật hoàn tất: Văn bản đã đạt chuẩn NĐ 30.");
        } else {
          findings.forEach((f) => {
            const itemDiv = document.createElement("label");
            itemDiv.className = "format-check-item";
            itemDiv.innerHTML = `
              <input type="checkbox" class="fmt-checkbox" value="${f.id}" checked />
              <div>
                <strong>${PCDT.escapeHtml(f.label)}</strong>
                <div class="format-check-detail">${PCDT.escapeHtml(f.detail)}</div>
              </div>
            `;
            listDiv.appendChild(itemDiv);
          });
          optionsBox.style.display = "block";
          confirmBtn.textContent = "✅ Xác nhận chuẩn hóa";
          confirmBtn.style.display = "block";
          PCDT.setStatus(`Phát hiện ${findings.length} thành phần cần chuẩn hóa theo NĐ 30. Hãy tích chọn và bấm Xác nhận.`);
        }
      });
    } catch (err) {
      PCDT.setStatus("Lỗi khi quét định dạng NĐ 30: " + err.message);
    }
  }

  async function applySelectedFormatND30() {
    const allCheckboxes = document.querySelectorAll(".fmt-checkbox");
    const cfg = readND30SettingsFromUI();
    let checkedIds = [];
    if (allCheckboxes.length === 0) {
      // Văn bản đã đạt chuẩn nhưng người dùng thay đổi 8 thông số và bấm "Áp dụng theo thông số đang chọn"
      checkedIds = ["fmt_paper", "fmt_margin", "fmt_font", "fmt_size", "fmt_align", "fmt_indent", "fmt_spacing"];
    } else {
      checkedIds = Array.from(document.querySelectorAll(".fmt-checkbox:checked")).map((cb) => cb.value);
      // Nếu người dùng thay đổi thông số đích ngay trong lần áp dụng đầu tiên, tự động áp dụng mục tương ứng
      if (lastScannedCfg) {
        if (cfg.fontSize !== lastScannedCfg.fontSize && !checkedIds.includes("fmt_size")) {
          checkedIds.push("fmt_size");
        }
        if (cfg.indent !== lastScannedCfg.indent && !checkedIds.includes("fmt_indent")) {
          checkedIds.push("fmt_indent");
        }
        if (
          (cfg.lineSpacing !== lastScannedCfg.lineSpacing || cfg.spaceBeforeAfter !== lastScannedCfg.spaceBeforeAfter) &&
          !checkedIds.includes("fmt_spacing")
        ) {
          checkedIds.push("fmt_spacing");
        }
        if (
          (cfg.marginTop !== lastScannedCfg.marginTop ||
            cfg.marginBottom !== lastScannedCfg.marginBottom ||
            cfg.marginLeft !== lastScannedCfg.marginLeft ||
            cfg.marginRight !== lastScannedCfg.marginRight) &&
          !checkedIds.includes("fmt_margin")
        ) {
          checkedIds.push("fmt_margin");
        }
      }
      if (checkedIds.length === 0) return alert("Vui lòng tích chọn ít nhất 1 thành phần để chuẩn hóa.");
    }

    setFormatProgress(5, "Đang khởi tạo quá trình chuẩn hóa NĐ 30...");

    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        const needPaperOrMargin = checkedIds.includes("fmt_paper") || checkedIds.includes("fmt_margin");
        const needParaUpdate =
          checkedIds.includes("fmt_font") ||
          checkedIds.includes("fmt_size") ||
          checkedIds.includes("fmt_align") ||
          checkedIds.includes("fmt_indent") ||
          checkedIds.includes("fmt_spacing");

        const supportsPageSetup = Office.context.requirements.isSetSupported("WordApi", "1.5");

        // 1. Khổ giấy A4 & Căn lề (tách riêng sections.load("items") để tránh lỗi items unavailable)
        if (supportsPageSetup && needPaperOrMargin) {
          setFormatProgress(10, "Bước 1/4: Đang chuẩn hóa khổ giấy A4 & căn lề trang...");
          const sections = context.document.sections;
          sections.load("items");
          await context.sync();

          if (checkedIds.includes("fmt_paper")) {
            sections.items.forEach((sec) => {
              sec.pageSetup.load("pageWidth, pageHeight");
            });
            await context.sync();
          }

          sections.items.forEach((sec) => {
            if (checkedIds.includes("fmt_paper")) {
              // Giữ nguyên hướng trang ngang nếu section đang đặt Landscape (theo chuẩn NĐ 30 cho bảng rộng)
              const w = sec.pageSetup.pageWidth || 0;
              const h = sec.pageSetup.pageHeight || 0;
              const isLandscape = w > h;
              sec.pageSetup.pageWidth = PCDT.mmToPt(isLandscape ? 297 : 210);
              sec.pageSetup.pageHeight = PCDT.mmToPt(isLandscape ? 210 : 297);
            }
            if (checkedIds.includes("fmt_margin")) {
              sec.pageSetup.topMargin = PCDT.mmToPt(cfg.marginTop);
              sec.pageSetup.bottomMargin = PCDT.mmToPt(cfg.marginBottom);
              sec.pageSetup.leftMargin = PCDT.mmToPt(cfg.marginLeft);
              sec.pageSetup.rightMargin = PCDT.mmToPt(cfg.marginRight);
            }
          });
          await context.sync();
        }

        // 2. Phông chữ (cả trong bảng và ngoài bảng)
        if (checkedIds.includes("fmt_font")) {
          setFormatProgress(18, "Bước 2/4: Đang chuyển phông chữ Times New Roman...");
          body.font.name = "Times New Roman";
          await context.sync();
        }

        // 3. Các định dạng nội dung NGOÀI BẢNG (chia theo lô để hiển thị tiến trình rõ ràng trên văn bản dài)
        if (needParaUpdate) {
          setFormatProgress(22, "Bước 3/4: Đang đọc danh sách đoạn văn bản...");
          const paragraphs = body.paragraphs;
          paragraphs.load("items");
          await context.sync();

          const totalParas = paragraphs.items.length;
          const supportsLists = Office.context.requirements.isSetSupported("WordApi", "1.3");
          const loadProps = supportsLists
            ? "text, alignment, isListItem, listItemOrNullObject/level, listItemOrNullObject/listString, listOrNullObject/id, parentTableOrNullObject/isNullObject"
            : "text, alignment, isListItem, parentTableOrNullObject/isNullObject";

          // Trong Word (Office.js), chế độ giãn dòng Multiple quy ước 1.0 line = 12 pt
          const lineSpacingPt = Math.round(12 * cfg.lineSpacing * 10) / 10;
          const targetIndentPt = PCDT.cmToPt(cfg.indent);
          const tabStopQueue = [];
          const updatedListLevels = new Set();

          const PARA_CHUNK = 80;
          for (let start = 0; start < totalParas; start += PARA_CHUNK) {
            const slice = paragraphs.items.slice(start, start + PARA_CHUNK);
            slice.forEach((p) => p.load(loadProps));
            await context.sync();

            slice.forEach((p) => {
              const inTable = p.parentTableOrNullObject && !p.parentTableOrNullObject.isNullObject;
              if (!isBodyContentParagraph(p, inTable)) return;

              if (checkedIds.includes("fmt_font")) {
                p.font.color = "#000000";
              }
              if (checkedIds.includes("fmt_size")) {
                p.font.size = cfg.fontSize;
              }
              if (checkedIds.includes("fmt_align")) {
                p.alignment = Word.Alignment.justified;
              }
              if (checkedIds.includes("fmt_spacing")) {
                p.spaceBefore = cfg.spaceBeforeAfter;
                p.spaceAfter = cfg.spaceBeforeAfter;
                p.lineSpacing = lineSpacingPt;
              }
              if (checkedIds.includes("fmt_indent")) {
                if (p.isListItem) {
                  // Giữ nguyên 100% Bullets & Numbering: dùng API gốc setLevelIndents(level, 0, targetIndentPt)
                  if (
                    supportsLists &&
                    p.listOrNullObject &&
                    !p.listOrNullObject.isNullObject &&
                    p.listItemOrNullObject &&
                    !p.listItemOrNullObject.isNullObject
                  ) {
                    const listId = p.listOrNullObject.id;
                    const lvl = p.listItemOrNullObject.level || 0;
                    const key = `${listId}::${lvl}`;
                    if (!updatedListLevels.has(key)) {
                      updatedListLevels.add(key);
                      p.listOrNullObject.setLevelIndents(lvl, 0, targetIndentPt);
                    }
                  }
                  const listStr =
                    supportsLists && p.listItemOrNullObject && !p.listItemOrNullObject.isNullObject
                      ? p.listItemOrNullObject.listString
                      : "";
                  p.leftIndent = 0;
                  p.firstLineIndent = targetIndentPt;
                  tabStopQueue.push({
                    p,
                    isListItem: true,
                    tabOffsetCm: getTabOffsetCmForListString(listStr)
                  });
                } else if (isBulletTextPrefix(p.text)) {
                  p.leftIndent = 0;
                  p.firstLineIndent = targetIndentPt;
                  tabStopQueue.push({ p, isListItem: false, tabOffsetCm: 0.5 });
                } else {
                  p.leftIndent = 0;
                  p.firstLineIndent = targetIndentPt;
                }
              }
            });

            await context.sync();
            const doneParas = Math.min(start + PARA_CHUNK, totalParas);
            const pct = totalParas > 0 ? 25 + (doneParas / totalParas) * 40 : 65;
            setFormatProgress(pct, `Bước 3/4: Đang chuẩn hóa đoạn văn bản (${doneParas}/${totalParas})...`);
          }

          // Đặt Tab stop = (cfg.indent + 0.5 / 0.75 / 1.0) cm qua OOXML cho Bullets & Numbering (chia lô)
          if (tabStopQueue.length > 0) {
            const TAB_CHUNK = 25;
            const totalTabs = tabStopQueue.length;
            for (let start = 0; start < totalTabs; start += TAB_CHUNK) {
              const batch = tabStopQueue.slice(start, start + TAB_CHUNK);
              const ooxmlResults = batch.map((item) => ({
                item,
                ooxml: item.p.getRange().getOoxml()
              }));
              await context.sync();

              ooxmlResults.forEach(({ item, ooxml }) => {
                try {
                  const updated = updateParagraphTabStopOoxml(
                    ooxml.value,
                    cfg.indent,
                    item.isListItem,
                    item.tabOffsetCm
                  );
                  if (updated) {
                    item.p.getRange().insertOoxml(updated, Word.InsertLocation.replace);
                  }
                } catch (e) {}
              });
              await context.sync();
              const doneTabs = Math.min(start + TAB_CHUNK, totalTabs);
              const pct = 65 + (doneTabs / totalTabs) * 30;
              setFormatProgress(pct, `Bước 4/4: Đang căn chỉnh Tab Bullets/Numbering (${doneTabs}/${totalTabs})...`);
            }
          }
        }

        setFormatProgress(100, "Hoàn tất chuẩn hóa!");
        clearFormatProgress();
        clearAllOptionMismatches();
        document.getElementById("formatReviewPanel").style.display = "none";
        PCDT.setStatus(`✅ Đã chuẩn hóa thành công ${checkedIds.length} thành phần theo NĐ 30!`);
      });
    } catch (err) {
      clearFormatProgress();
      PCDT.setStatus("Lỗi khi chuẩn hóa định dạng: " + err.message);
    }
  }

  PCDT.isBodyContentParagraph = isBodyContentParagraph;
  PCDT.loadND30DefaultsToUI = loadND30DefaultsToUI;
  PCDT.saveND30DefaultsFromUI = saveND30DefaultsFromUI;
  PCDT.scanFormatND30 = scanFormatND30;
  PCDT.applySelectedFormatND30 = applySelectedFormatND30;
})(window.PCDT);
