window.PCDT = window.PCDT || {};

(function (PCDT) {
  function setStatus(msg) {
    const el = document.getElementById("status");
    if (el) el.textContent = msg;
  }

  function setStepState(stepNumber, state) {
    const bar = document.getElementById("stepsBar");
    if (bar) bar.style.display = "flex";
    const pill = document.getElementById(`step${stepNumber}Pill`);
    if (!pill) return;
    pill.classList.remove("active", "done");
    if (state === "active") pill.classList.add("active");
    if (state === "done") pill.classList.add("done");
  }

  function resetStepsUI() {
    [1, 2, 3].forEach((n) => setStepState(n, ""));
    const step3Pill = document.getElementById("step3Pill");
    if (step3Pill) step3Pill.textContent = "B3. Hợp lệ";
  }

  function escapeHtml(str) {
    return (str || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function ptToMm(pt) {
    return Math.round((pt / 2.83465) * 10) / 10;
  }

  function mmToPt(mm) {
    return mm * 2.83465;
  }

  function cmToPt(cm) {
    return cm * 28.3465;
  }

  function normalizeText(str) {
    return (str || "").replace(/[\s\u00A0\u200B\v\f\r\n]+/g, " ").trim().toLowerCase();
  }

  /**
   * Đọc văn bản sạch từ Word, chuyển mọi ngắt dòng mềm (Shift+Enter, \v, \r)
   * thành \n để tránh lỗi dính chữ trong bảng hoặc tiêu đề.
   */
  async function extractCleanDocumentText(context) {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/text");
    await context.sync();

    const lines = paragraphs.items.map((p) =>
      (p.text || "").replace(/[\v\u000b\r]/g, "\n").trim()
    );
    return lines.filter((l) => l.length > 0).join("\n");
  }

  /**
   * Tìm kiếm thông minh trong Word (hỗ trợ cả ngoài bảng lẫn sâu trong ô Table).
   */
  async function findTargetInWord(context, quoteText) {
    if (!quoteText || !quoteText.trim()) return null;

    const parts = quoteText
      .split(/[\r\n\t\v\f\u000b]+/)
      .map((s) => s.trim())
      .filter((s) => s.length >= 2)
      .sort((a, b) => b.length - a.length);

    if (parts.length === 0) return null;

    const mainPart = parts[0];
    const candidates = [mainPart.slice(0, 80)];
    if (mainPart.length > 35) candidates.push(mainPart.slice(0, 35).trim());
    if (mainPart.length > 18) candidates.push(mainPart.slice(0, 18).trim());

    for (const cand of candidates) {
      if (!cand) continue;
      try {
        const results = context.document.body.search(cand, {
          matchCase: false,
          matchWholeWord: false
        });
        results.load("items");
        await context.sync();
        if (results.items.length > 0) return results.items[0];
      } catch (e) {}
    }

    const targetNorm = normalizeText(mainPart.slice(0, 40));
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();

    for (const table of tables.items) {
      for (const cand of candidates) {
        if (!cand) continue;
        try {
          const tRes = table.search(cand, { matchCase: false, matchWholeWord: false });
          tRes.load("items");
          await context.sync();
          if (tRes.items.length > 0) return tRes.items[0];
        } catch (e) {}
      }

      const rows = table.rows;
      rows.load("items");
      await context.sync();
      for (const row of rows.items) {
        const cells = row.cells;
        cells.load("items/body/text");
        await context.sync();
        for (const cell of cells.items) {
          if (normalizeText(cell.body.text).includes(targetNorm)) {
            return cell.body;
          }
        }
      }
    }

    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/text");
    await context.sync();
    for (const p of paragraphs.items) {
      if (normalizeText(p.text).includes(targetNorm)) {
        return p;
      }
    }

    return null;
  }

  async function navigateToIssue(quoteText) {
    try {
      await Word.run(async (context) => {
        const target = await findTargetInWord(context, quoteText);
        if (target) {
          target.select();
          await context.sync();
        } else {
          alert("Không tìm thấy vị trí đoạn trích dẫn này (có thể bạn đã sửa hoặc xóa nội dung này).");
        }
      });
    } catch (err) {
      alert("Lỗi khi di chuyển đến vị trí lỗi: " + err.message);
    }
  }

  async function insertWordComment(quoteText, commentBody) {
    try {
      await Word.run(async (context) => {
        const target = await findTargetInWord(context, quoteText);
        if (target) {
          const range = typeof target.getRange === "function" ? target.getRange() : target;
          range.insertComment(commentBody);
          target.select();
          await context.sync();
        } else {
          alert("Không tìm thấy chính xác đoạn trích dẫn này trong văn bản.");
        }
      });
    } catch (err) {
      alert("Lỗi khi chèn Comment: " + err.message);
    }
  }

  async function applyAutoFix(quoteText, replacementText, btnElement, onFixedCallback) {
    try {
      await Word.run(async (context) => {
        const target = await findTargetInWord(context, quoteText);
        if (target && typeof target.insertText === "function") {
          target.insertText(replacementText, Word.InsertLocation.replace);
          target.select();
          await context.sync();

          btnElement.textContent = "✅ Đã sửa";
          btnElement.disabled = true;
          btnElement.style.background = "#16a34a";
          if (onFixedCallback) onFixedCallback();
        } else {
          alert("Không tìm thấy chính xác cụm từ lỗi để thay thế.");
        }
      });
    } catch (err) {
      alert("Lỗi khi sửa tự động: " + err.message);
    }
  }

  PCDT.setStatus = setStatus;
  PCDT.setStepState = setStepState;
  PCDT.resetStepsUI = resetStepsUI;
  PCDT.escapeHtml = escapeHtml;
  PCDT.ptToMm = ptToMm;
  PCDT.mmToPt = mmToPt;
  PCDT.cmToPt = cmToPt;
  PCDT.normalizeText = normalizeText;
  PCDT.extractCleanDocumentText = extractCleanDocumentText;
  PCDT.findTargetInWord = findTargetInWord;
  PCDT.navigateToIssue = navigateToIssue;
  PCDT.insertWordComment = insertWordComment;
  PCDT.applyAutoFix = applyAutoFix;
})(window.PCDT);
