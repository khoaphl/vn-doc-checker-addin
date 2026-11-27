window.PCDT = window.PCDT || {};

(function (PCDT) {
  let dictCache = null;

  const TECH_ALLOWLIST = new Set([
    "http", "https", "www", "com", "vn", "org", "net", "edu", "gov",
    "pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx", "txt", "json", "xml", "html", "css", "js",
    "file", "email", "online", "offline", "server", "local", "localhost", "port",
    "add", "addin", "word", "excel", "powerpoint", "outlook", "teams", "zoom", "zalo",
    "google", "gemini", "office", "windows", "microsoft", "android", "ios",
    "recloser", "scada", "dms", "oms", "pmis", "cmis", "erp", "hrms", "doffice", "eoffice",
    "internet", "website", "software", "hardware", "usb", "lan", "wan", "vpn", "wifi", "ip",
    "iso", "tcvn", "qcvn", "iec", "ieee", "kpi", "saidi", "saifi", "maifi",
    "pt", "mm", "cm", "dm", "km", "mg", "kg", "kv", "kw", "mw", "mva", "mvar", "kva", "kvar", "kwh", "hz", "ma"
  ]);

  const VN_DIACRITIC_RE = /[àáảãạăằắẳẵặâầấẩẫậđèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵ]/i;

  function isWordChar(ch) {
    if (!ch) return false;
    return /[\p{L}\p{N}_]/u.test(ch);
  }

  async function loadOfflineDictionaries() {
    if (dictCache) return dictCache;

    const [syllablesRes, typosRes, evnRes] = await Promise.all([
      fetch("data/vn-syllables.json").then((r) => (r.ok ? r.json() : [])).catch(() => []),
      fetch("data/common-typos.json").then((r) => (r.ok ? r.json() : { typos: [] })).catch(() => ({ typos: [] })),
      fetch("data/evn-dict.json").then((r) => (r.ok ? r.json() : { terms: [] })).catch(() => ({ terms: [] }))
    ]);

    const syllableSet = new Set((Array.isArray(syllablesRes) ? syllablesRes : []).map((s) => s.toLowerCase()));
    const typoList = (typosRes && Array.isArray(typosRes.typos) ? typosRes.typos : []).filter((t) => !t.skip && t.wrong);

    const evnTermSet = new Set();
    const evnTokenSet = new Set();
    const terms = evnRes && Array.isArray(evnRes.terms) ? evnRes.terms : [];
    terms.forEach((term) => {
      const clean = (term || "").trim();
      if (!clean) return;
      evnTermSet.add(clean.toLowerCase());
      clean.split(/[\s\-\/]+/).forEach((part) => {
        if (part) evnTokenSet.add(part.toLowerCase());
      });
    });

    dictCache = { syllableSet, typoList, evnTermSet, evnTokenSet };
    return dictCache;
  }

  function isExcludedToken(token, customDictSet) {
    if (!token || token.length < 2) return true;
    const lower = token.toLowerCase();

    // Bỏ qua số thuần, mã có chữ số (35kV, 110kV, 05, 2026)
    if (/^\d+$/.test(token) || /\d/.test(token)) return true;

    // Bỏ qua từ viết hoa toàn bộ (EVN, PCĐT, UBND, HĐTV...)
    if (token === token.toUpperCase() && token !== token.toLowerCase()) return true;

    // Bỏ qua từ điển EVN, allowlist kỹ thuật và từ điển riêng người dùng
    if (dictCache.evnTermSet.has(lower) || dictCache.evnTokenSet.has(lower)) return true;
    if (TECH_ALLOWLIST.has(lower)) return true;
    if (customDictSet.has(lower)) return true;

    return false;
  }

  function preserveCase(source, target) {
    if (!source || !target) return target;
    if (source[0] === source[0].toUpperCase() && source[0] !== source[0].toLowerCase()) {
      return target[0].toUpperCase() + target.slice(1);
    }
    return target;
  }

  /**
   * Kiểm tra chính tả Offline 3 lớp:
   * - Lớp 1: Âm tiết không tồn tại trong vn-syllables.json
   * - Lớp 2: Cặp từ hay nhầm, khẩu ngữ/biệt ngữ, lỗi OCR trong common-typos.json
   * - Lớp 3: Loại trừ evn-dict.json, Global Ignore và Từ điển riêng của người dùng
   */
  async function runStep2OfflineSpelling(documentText) {
    PCDT.setStatus("Bước 2/3 (Offline): Đang kiểm tra chính tả 3 lớp cục bộ...");
    await loadOfflineDictionaries();

    const customDictSet = new Set(PCDT.getCustomDict().map((w) => w.toLowerCase()));
    const lines = (documentText || "").split(/\r?\n/);
    const issues = [];
    const seenQuotes = new Set();

    for (const rawLine of lines) {
      const line = rawLine.trim();
      if (!line) continue;

      // Bỏ qua dòng URL/email thuần
      const sanitizedLine = line
        .replace(/https?:\/\/\S+/gi, " ")
        .replace(/\S+@\S+\.\S+/gi, " ");

      // LỚP 2: Kiểm tra cặp từ hay nhầm & khẩu ngữ trong common-typos.json
      const lineLower = sanitizedLine.toLowerCase();
      for (const item of dictCache.typoList) {
        const needle = item.wrong.toLowerCase();
        let searchIdx = 0;
        while (searchIdx < lineLower.length) {
          const foundIdx = lineLower.indexOf(needle, searchIdx);
          if (foundIdx === -1) break;

          const beforeChar = foundIdx > 0 ? sanitizedLine[foundIdx - 1] : "";
          const afterChar = foundIdx + needle.length < sanitizedLine.length ? sanitizedLine[foundIdx + needle.length] : "";
          const leftOk = !isWordChar(beforeChar);
          const rightOk = !isWordChar(afterChar);

          if (leftOk && rightOk) {
            const exactQuote = sanitizedLine.slice(foundIdx, foundIdx + needle.length);
            if (!customDictSet.has(exactQuote.toLowerCase()) && !seenQuotes.has(`L2::${exactQuote.toLowerCase()}`)) {
              seenQuotes.add(`L2::${exactQuote.toLowerCase()}`);
              const replacement = preserveCase(exactQuote, item.replacement);
              issues.push({
                category: "Chính tả",
                checkType: "spelling_offline_typo",
                quote: exactQuote,
                issue: `[Offline - Lớp 2] ${item.reason || `Từ/cụm từ "${exactQuote}" có thể bị nhầm hoặc là khẩu ngữ.`}`,
                suggestion: `Đề xuất sửa thành "${replacement}".`,
                replacement
              });
            }
          }
          searchIdx = foundIdx + needle.length;
        }
      }

      // LỚP 1 & LỚP 3: Tách từng âm tiết tiếng Việt kiểm tra với vn-syllables.json
      const tokenMatches = sanitizedLine.match(/[\p{L}\p{N}]+/gu) || [];
      for (const token of tokenMatches) {
        if (isExcludedToken(token, customDictSet)) continue;

        const lower = token.toLowerCase();
        // Kiểm tra nếu từ có dấu tiếng Việt HOẶC là từ viết thường hoàn toàn
        const hasVnDiacritic = VN_DIACRITIC_RE.test(lower);
        const isAllLower = token === lower;

        if (!hasVnDiacritic && !isAllLower) {
          // Từ viết hoa chữ đầu không dấu (VD: Microsoft, Word, Gemini...) -> bỏ qua để tránh báo nhầm tên riêng/thuật ngữ ngoại
          continue;
        }

        if (dictCache.syllableSet.size > 0 && !dictCache.syllableSet.has(lower)) {
          const key = `L1::${lower}`;
          if (seenQuotes.has(key)) continue;
          seenQuotes.add(key);

          issues.push({
            category: "Chính tả",
            checkType: "spelling_offline_syllable",
            quote: token,
            issue: `[Offline - Lớp 1] Âm tiết "${token}" không có trong từ điển âm tiết tiếng Việt chuẩn.`,
            suggestion: "Kiểm tra lại lỗi gõ phím/bỏ dấu tiếng Việt, hoặc thêm vào 'Từ điển riêng Offline' nếu là tên riêng/thuật ngữ đúng.",
            replacement: ""
          });
        }
      }
    }

    return issues;
  }

  PCDT.runStep2OfflineSpelling = runStep2OfflineSpelling;
})(window.PCDT);
