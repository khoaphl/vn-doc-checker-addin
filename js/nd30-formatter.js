window.PCDT = window.PCDT || {};

(function (PCDT) {
  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  function isBulletTextPrefix(text) {
    return /^\s*[-–+•](?:\t|\s+)/.test(text || "");
  }

  function isNumberedListString(listStr) {
    return /[0-9a-zA-Z]/.test(listStr || "");
  }

  function cmToTwips(cm) {
    return Math.round(cm * 566.929);
  }

  function updateParagraphOoxmlForBullet(ooxmlStr, indentCm, addDashTabPrefix) {
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

    // Xóa numPr nếu còn sót
    const numPrList = pPr.getElementsByTagNameNS(W_NS, "numPr");
    while (numPrList.length > 0) {
      numPrList[0].parentNode.removeChild(numPrList[0]);
    }

    // Đặt Left = 0 cm, First line = indentCm (1 cm = 567 twips, 1.27 cm = 720 twips)
    const oldInd = pPr.getElementsByTagNameNS(W_NS, "ind");
    while (oldInd.length > 0) {
      oldInd[0].parentNode.removeChild(oldInd[0]);
    }
    const indEl = xmlDoc.createElementNS(W_NS, "w:ind");
    indEl.setAttributeNS(W_NS, "w:left", "0");
    indEl.setAttributeNS(W_NS, "w:firstLine", String(cmToTwips(indentCm)));
    pPr.appendChild(indEl);

    // Đặt Tab stop = (indentCm + 0.5) cm (VD: 1 cm -> 1.5 cm = 850 twips)
    const oldTabs = pPr.getElementsByTagNameNS(W_NS, "tabs");
    while (oldTabs.length > 0) {
      oldTabs[0].parentNode.removeChild(oldTabs[0]);
    }
    const tabsEl = xmlDoc.createElementNS(W_NS, "w:tabs");
    const tabEl = xmlDoc.createElementNS(W_NS, "w:tab");
    tabEl.setAttributeNS(W_NS, "w:val", "left");
    tabEl.setAttributeNS(W_NS, "w:pos", String(cmToTwips(indentCm + 0.5)));
    tabsEl.appendChild(tabEl);
    pPr.appendChild(tabsEl);

    // Chuẩn hóa đầu dòng thành "-<tab>" nếu vừa tách từ Word Auto Bullet hoặc đang dùng "-   "
    if (addDashTabPrefix) {
      const rNodes = pNode.getElementsByTagNameNS(W_NS, "r");
      const firstRPr = rNodes.length > 0 ? rNodes[0].getElementsByTagNameNS(W_NS, "rPr")[0] : null;

      const dashRun = xmlDoc.createElementNS(W_NS, "w:r");
      if (firstRPr) dashRun.appendChild(firstRPr.cloneNode(true));
      const dashText = xmlDoc.createElementNS(W_NS, "w:t");
      dashText.textContent = "-";
      dashRun.appendChild(dashText);

      const tabRun = xmlDoc.createElementNS(W_NS, "w:r");
      if (firstRPr) tabRun.appendChild(firstRPr.cloneNode(true));
      tabRun.appendChild(xmlDoc.createElementNS(W_NS, "w:tab"));

      const refNode = pPr.nextSibling;
      pNode.insertBefore(tabRun, refNode);
      pNode.insertBefore(dashRun, tabRun);
    } else {
      // Nếu là dòng gõ tay "- Nội dung" (dùng space thay vì tab), chuyển khoảng trắng đầu sau dấu "-" thành <w:tab/>
      const tNodes = pNode.getElementsByTagNameNS(W_NS, "t");
      if (tNodes.length > 0) {
        const firstT = tNodes[0];
        const rawVal = firstT.textContent || "";
        const m = rawVal.match(/^\s*([-–+•])(\s+)(.*)$/);
        if (m) {
          firstT.textContent = "-";
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
    if (text.length <= 40) return false;
    // Bỏ qua các dòng tiêu đề đang canh giữa (VD: trích yếu dài, tên chương)
    if (p.alignment === Word.Alignment.centered) return false;
    return true;
  }

  function loadND30DefaultsToUI() {
    const cfg = PCDT.getND30Defaults();
    const setVal = (id, val) => {
      const el = document.getElementById(id);
      if (el && val !== undefined && val !== null) el.value = String(val);
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

        // 1. Khổ giấy A4 (210 x 297 mm, cho phép dọc hoặc ngang) & Căn lề
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
            if (top < 20 || top > 25 || bottom < 20 || bottom > 25 || left < 30 || left > 35 || right < 15 || right > 20) {
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
              label: `Căn lề trang chưa đúng NĐ 30 (${wrongMarginCount} phần/section)`,
              detail: `${sampleMargin}. Chuẩn: Trên/Dưới 20–25mm, Trái 30–35mm, Phải 15–20mm.`
            });
          }
        }

        // 2. Quét các đoạn văn bản (phân biệt ngoài bảng và trong bảng)
        let wrongFontCount = 0;
        let wrongSizeCount = 0;
        let wrongAlignCount = 0;
        let wrongIndentCount = 0;
        let wrongSpacingCount = 0;

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
            if (p.font.size && (p.font.size < 13 || p.font.size > 14)) {
              wrongSizeCount++;
            }
            if (p.alignment !== Word.Alignment.justified) {
              wrongAlignCount++;
            }
            // Thụt đầu dòng chuẩn 1cm (~28.35pt) hoặc 1.27cm (~36pt), dung sai ±2.5pt
            // Theo mẫu PCĐT: cả đoạn thường và đoạn Bullets đều dùng Left = 0 cm, First line = 1 cm (hoặc 1,27 cm)
            if (p.isListItem) {
              wrongIndentCount++;
            } else {
              const leftPt = p.leftIndent || 0;
              const indent = p.firstLineIndent || 0;
              const okLeft0 = Math.abs(leftPt) <= 2.5;
              const ok1cm = Math.abs(indent - 28.35) <= 2.5;
              const ok127cm = Math.abs(indent - 36.0) <= 2.5;
              if (!okLeft0 || (!ok1cm && !ok127cm)) {
                wrongIndentCount++;
              }
            }
            // Cách đoạn đồng bộ Before & After với mức chuẩn đang lưu/chọn & giãn dòng 1.0 - 1.5 lines
            const spaceAfter = p.spaceAfter || 0;
            const spaceBefore = p.spaceBefore || 0;
            const lineSp = p.lineSpacing || 12;
            const beforeMismatch = Math.abs(spaceBefore - targetSpace) > 0.5;
            const afterMismatch = Math.abs(spaceAfter - targetSpace) > 0.5;
            if (beforeMismatch || afterMismatch || lineSp < 11.5 || lineSp > 22) {
              wrongSpacingCount++;
            }
          }
        });

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
            label: `Cỡ chữ nội dung ngoài khoảng 13–14 pt (${wrongSizeCount} đoạn ngoài bảng)`,
            detail: "Chuẩn hóa về cỡ chữ nội dung bạn chọn bên dưới."
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
            label: `Thụt đầu dòng chưa đúng 1 cm / 1,27 cm (${wrongIndentCount} đoạn ngoài bảng)`,
            detail: "Chuẩn hóa thụt đầu dòng theo mức bạn chọn bên dưới."
          });
        }
        if (wrongSpacingCount > 0) {
          findings.push({
            id: "fmt_spacing",
            label: `Cách đoạn Before/After chưa đồng bộ ${targetSpace} pt hoặc Giãn dòng chưa đạt (${wrongSpacingCount} đoạn ngoài bảng)`,
            detail: `Đặt đồng bộ cả Before = ${targetSpace} pt và After = ${targetSpace} pt, giãn dòng 1,0 – 1,5 lines.`
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
    let checkedIds = [];
    if (allCheckboxes.length === 0) {
      // Văn bản đã đạt chuẩn nhưng người dùng thay đổi 8 thông số và bấm "Áp dụng theo thông số đang chọn"
      checkedIds = ["fmt_paper", "fmt_margin", "fmt_font", "fmt_size", "fmt_align", "fmt_indent", "fmt_spacing"];
    } else {
      checkedIds = Array.from(document.querySelectorAll(".fmt-checkbox:checked")).map((cb) => cb.value);
      if (checkedIds.length === 0) return alert("Vui lòng tích chọn ít nhất 1 thành phần để chuẩn hóa.");
    }

    const cfg = readND30SettingsFromUI();
    PCDT.setStatus("Đang áp dụng chuẩn hóa các mục đã chọn...");

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
        }

        // 2. Phông chữ (cả trong bảng và ngoài bảng)
        if (checkedIds.includes("fmt_font")) {
          body.font.name = "Times New Roman";
        }

        // 3. Các định dạng nội dung NGOÀI BẢNG (tách riêng paragraphs.load("items") + sync)
        if (needParaUpdate) {
          const paragraphs = body.paragraphs;
          paragraphs.load("items");
          await context.sync();

          paragraphs.items.forEach((p) => {
            p.load("text, alignment, isListItem, listItemOrNullObject/listString, parentTableOrNullObject/isNullObject");
          });
          await context.sync();

          // Trong Word (Office.js), chế độ giãn dòng Multiple quy ước 1.0 line = 12 pt
          const lineSpacingPt = Math.round(12 * cfg.lineSpacing * 10) / 10;
          const targetIndentPt = PCDT.cmToPt(cfg.indent);
          const bulletUpdateQueue = [];

          paragraphs.items.forEach((p) => {
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
              const listStr =
                p.listItemOrNullObject && !p.listItemOrNullObject.isNullObject
                  ? p.listItemOrNullObject.listString
                  : "";
              const isAutoBullet = p.isListItem && !isNumberedListString(listStr);
              const isManualBullet = !p.isListItem && isBulletTextPrefix(p.text);

              if (isAutoBullet) {
                if (typeof p.detachFromList === "function") p.detachFromList();
                p.leftIndent = 0;
                p.firstLineIndent = targetIndentPt;
                bulletUpdateQueue.push({ p, addDashTabPrefix: true });
              } else if (isManualBullet) {
                p.leftIndent = 0;
                p.firstLineIndent = targetIndentPt;
                bulletUpdateQueue.push({ p, addDashTabPrefix: false });
              } else if (!p.isListItem) {
                p.leftIndent = 0;
                p.firstLineIndent = targetIndentPt;
              }
            }
          });

          await context.sync();

          // Chèn Tab stop = (cfg.indent + 0.5) cm qua OOXML cho các đoạn có Bullets (giữ nguyên định dạng chữ)
          if (bulletUpdateQueue.length > 0) {
            const ooxmlResults = bulletUpdateQueue.map((item) => ({
              item,
              ooxml: item.p.getRange().getOoxml()
            }));
            await context.sync();

            ooxmlResults.forEach(({ item, ooxml }) => {
              try {
                const updated = updateParagraphOoxmlForBullet(
                  ooxml.value,
                  cfg.indent,
                  item.addDashTabPrefix
                );
                if (updated) {
                  item.p.getRange().insertOoxml(updated, Word.InsertLocation.replace);
                }
              } catch (e) {}
            });
          }
        }

        await context.sync();
        document.getElementById("formatReviewPanel").style.display = "none";
        PCDT.setStatus(`✅ Đã chuẩn hóa thành công ${checkedIds.length} thành phần theo NĐ 30!`);
      });
    } catch (err) {
      PCDT.setStatus("Lỗi khi chuẩn hóa định dạng: " + err.message);
    }
  }

  PCDT.isBodyContentParagraph = isBodyContentParagraph;
  PCDT.loadND30DefaultsToUI = loadND30DefaultsToUI;
  PCDT.saveND30DefaultsFromUI = saveND30DefaultsFromUI;
  PCDT.scanFormatND30 = scanFormatND30;
  PCDT.applySelectedFormatND30 = applySelectedFormatND30;
})(window.PCDT);
