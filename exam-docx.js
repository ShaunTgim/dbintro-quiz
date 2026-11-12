// 考卷 Word（.docx）產生器：版面比照 exam-pdf.js。
// 每頁頁首 Name: / Ch:標籤 / S:；題目「1.(　) 題幹 - 原題號」，選項接下一行；區塊圖放題幹下方、行內圖在原位置；
// 一題不拆頁；最後是答案格子頁（NO | ANS，每欄 25 列，每 5 列交錯淺綠底）。
// 需要瀏覽器全域 docx（docx 9.x iife）與 ExamPDF.toRuns（exam-pdf.js）。
(function (root) {
  const PAPER = { B4: [728.5, 1031.8], A4: [595.3, 841.9] };   // pt
  const TW = pt => Math.round(pt * 20);                          // pt → twip
  const PX = pt => Math.round(pt * 96 / 72);                     // pt → px（ImageRun 的單位）
  const FONT = { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft JhengHei", cs: "Arial" };

  async function build(o) {
    const { docx, items, chLabel, paper = "B4", title = "考卷", toRuns } = o;
    const { Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, Header, WidthType,
      TabStopType, ShadingType, BorderStyle, PageBreak, VerticalAlign } = docx;
    const [PW, PH] = PAPER[paper] || PAPER.B4;
    const k = PW / PAPER.B4[0];
    const FS = 11 * Math.max(k, 0.9);                // 本文字級 pt
    const ML = 43 * k, MR = 43 * k, TOPM = 64 * k, BOTM = 44 * k;
    const IND = 54 * k;                              // 題幹左縮排（題號欄寬）
    const BODY_W = PW - ML - MR - IND;               // 題幹可用寬度 pt
    const sz = pt => Math.round(pt * 2);             // pt → 半點
    const run = (text, extra = {}) => new TextRun({ text, font: FONT, size: sz(FS), ...extra });

    // HTML（<sub>、<sup>、行內圖、\n）→ TextRun / ImageRun
    function runsOf(html, it) {
      const out = [];
      for (const r of toRuns(html)) {
        if (r.img) {
          const src = it.inl[r.img]; if (!src) continue;
          const s0 = Math.min(k, BODY_W / r.w);
          out.push(new ImageRun({ type: "jpg", data: src.bytes, transformation: { width: PX(r.w * s0), height: PX(r.h * s0) } }));
          continue;
        }
        const parts = r.t.split("\n");
        parts.forEach((p, i) => {
          const opt = { break: i ? 1 : 0 };
          if (r.sc === "sub") opt.subScript = true;
          if (r.sc === "sup") opt.superScript = true;
          if (p || i) out.push(run(p, opt));
        });
      }
      return out;
    }

    const header = new Header({
      children: [new Paragraph({
        tabStops: [{ type: TabStopType.LEFT, position: TW((263 - 48) * k) }, { type: TabStopType.LEFT, position: TW((481 - 48) * k) }],
        indent: { left: TW(5 * k) },
        children: [run("Name:", { size: sz(10.5 * k) }), run("\tCh:" + chLabel, { size: sz(10.5 * k) }), run("\tS:", { size: sz(10.5 * k) })],
      })],
    });

    const body = [];
    const spacing = { before: 0, after: Math.round(FS * 0.42 * 20), line: 300 };
    items.forEach((it, n) => {
      const paras = [];
      // 題幹：題號「n.(」、括號內留空、「)」，以定位點對齊；第二行起縮排
      paras.push({
        tabStops: [{ type: TabStopType.LEFT, position: TW(38 * k) }, { type: TabStopType.LEFT, position: TW(IND) }],
        indent: { left: TW(IND), hanging: TW(IND) },
        children: [run(`${n + 1}.(`), run("\t)\t"), ...runsOf(it.stem, it), run(" - " + it.ref)],
      });
      for (const im of it.imgs || []) {
        const s0 = Math.min(k, BODY_W / im.w, (340 * k) / im.h);
        paras.push({ img: true, indent: { left: TW(IND) }, children: [new ImageRun({ type: "jpg", data: im.bytes, transformation: { width: PX(im.w * s0), height: PX(im.h * s0) } })] });
      }
      if (!it.imgOpts && it.opts.length) {
        const ch = [];
        it.opts.forEach((op, i) => { if (i) ch.push(run("   ")); ch.push(run("(" + "ABCDE"[i] + ")")); ch.push(...runsOf(op, it)); });
        paras.push({ indent: { left: TW(IND) }, children: ch });
      }
      paras.forEach(({ img, ...p }, i) => {
        const sp = { ...spacing, after: i === paras.length - 1 ? spacing.after * 2 : spacing.after };
        if (img || p.children.some(c => c instanceof ImageRun)) delete sp.line;   // 含圖段落用單行自動行距，圖片才不會被裁掉
        body.push(new Paragraph({ ...p, keepLines: true, keepNext: i < paras.length - 1, spacing: sp }));   // 一題不拆在兩頁
      });
    });

    // ---- 答案頁 ----
    const N = items.length;
    const rows = Math.min(25, Math.ceil(N / 2));
    const pairsTotal = Math.ceil(N / rows);
    const PER_PAGE = 4;
    const colW = TW(Math.min(128.6 * k, (PW - ML - MR) / (Math.min(PER_PAGE, pairsTotal) * 2)));
    const line = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
    const borders = { top: line, bottom: line, left: line, right: line };
    const cell = (text, shade, bold) => new TableCell({
      width: { size: colW, type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.CENTER,
      margins: { left: TW(4 * k), right: TW(2 * k), top: 0, bottom: 0 },
      shading: shade ? { type: ShadingType.CLEAR, color: "auto", fill: "D9F2D0" } : undefined,
      children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [run(text, { size: sz((bold ? 10.5 : 9.5) * k) })] })],
    });
    for (let p0 = 0; p0 < pairsTotal; p0 += PER_PAGE) {
      const pairs = Math.min(PER_PAGE, pairsTotal - p0);
      const trs = [new TableRow({ height: { value: TW(18.4 * k), rule: "exact" }, children: [].concat(...Array.from({ length: pairs }, () => [cell("NO", false, true), cell("ANS", false, true)])) })];
      for (let r = 0; r < rows; r++) {
        const shade = Math.floor(r / 5) % 2 === 0, cs = [];
        for (let pi = 0; pi < pairs; pi++) {
          const idx = (p0 + pi) * rows + r;
          cs.push(cell(idx < N ? String(idx + 1) : "", shade), cell(idx < N ? items[idx].ans : "", false));
        }
        trs.push(new TableRow({ height: { value: TW(18.4 * k), rule: "exact" }, children: cs }));
      }
      body.push(new Paragraph({ children: [new PageBreak()] }));
      body.push(new Table({ width: { size: colW * pairs * 2, type: WidthType.DXA }, columnWidths: Array(pairs * 2).fill(colW), rows: trs }));
    }

    const doc = new Document({
      title: title + " " + chLabel, creator: title,
      styles: { default: { document: { run: { font: FONT, size: sz(FS) } } } },
      sections: [{
        properties: { page: { size: { width: TW(PW), height: TW(PH) }, margin: { top: TW(TOPM), bottom: TW(BOTM), left: TW(ML), right: TW(MR), header: TW(25 * k), footer: TW(20 * k) } } },
        headers: { default: header },
        children: body,
      }],
    });
    return await Packer.toBlob(doc);
  }

  root.ExamDOCX = { build, PAPER };
})(typeof window !== "undefined" ? window : globalThis);
