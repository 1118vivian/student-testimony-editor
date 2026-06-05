(function () {
  const fields = [
    "groupStart",
    "groupEnd",
    "name",
    "week",
    "startDate",
    "endDate",
    "topic",
    "trainingLife",
    "churchLife",
  ];

  const form = document.getElementById("testimonyForm");
  const fileNamePreview = document.getElementById("fileNamePreview");
  const preview = document.getElementById("documentPreview");

  const data = () =>
    Object.fromEntries(fields.map((field) => [field, document.getElementById(field).value.trim()]));

  const escapeText = (value) =>
    value.replace(/[&<>"']/g, (match) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" })[match]);

  const dateText = (value) => {
    if (!value) return "";
    const [year, month, day] = value.split("-");
    return `${year}年${Number(month)}月${Number(day)}日`;
  };

  const fileName = (values, extension) => {
    const name = values.name || "未命名";
    const week = values.week ? `第${values.week}週` : "學員蒙恩見證稿";
    return `${name}_${week}_學員蒙恩見證稿.${extension}`;
  };

  const syncPreview = () => {
    const values = data();
    const previewValues = {
      ...values,
      startDateText: dateText(values.startDate),
      endDateText: dateText(values.endDate),
    };

    document.querySelectorAll("[data-preview]").forEach((node) => {
      const key = node.getAttribute("data-preview");
      node.innerHTML = escapeText(previewValues[key] || "");
    });

    fileNamePreview.textContent = fileName(values, "docx");
  };

  const lineBreaks = (text, size = 22) => {
    const docx = window.docx;
    const lines = text ? text.split(/\r?\n/) : [""];
    return lines.flatMap((line, index) => {
      const runs = [];
      if (index > 0) runs.push(new docx.TextRun({ break: 1 }));
      runs.push(new docx.TextRun({ text: line, font: "Microsoft JhengHei", size }));
      return runs;
    });
  };

  const paragraph = (children, options = {}) => {
    const docx = window.docx;
    return new docx.Paragraph({
      alignment: options.alignment,
      spacing: options.spacing,
      indent: options.indent,
      children,
    });
  };

  const textRun = (text, options = {}) => {
    const docx = window.docx;
    return new docx.TextRun({
      text,
      bold: options.bold,
      font: "Microsoft JhengHei",
      size: options.size || 22,
    });
  };

  const contentParagraph = (text) =>
    paragraph(lineBreaks(text), {
      spacing: { line: 360, after: 160 },
    });

  const promptParagraph = (text) =>
    paragraph([textRun(text)], {
      indent: { left: 420 },
      spacing: { line: 320 },
    });

  const buildDocx = async () => {
    if (!window.docx) throw new Error("DOCX 套件尚未載入，請確認檔案路徑後重試。");
    const docx = window.docx;
    const values = data();

    const start = dateText(values.startDate) || "20     年      月      日";
    const end = dateText(values.endDate) || "20     年      月      日";
    const groupStart = values.groupStart || "　";
    const groupEnd = values.groupEnd || "　";
    const name = values.name || "　　　    　　　";
    const week = values.week || "　";
    const topic = values.topic || "將訓練所學應用於生活與事奉中";

    const doc = new docx.Document({
      sections: [
        {
          properties: {
            page: {
              margin: { top: 720, right: 720, bottom: 560, left: 720 },
            },
          },
          children: [
            paragraph([textRun("臺灣福音工作全時間訓練壯年成全班", { bold: true, size: 28 })], {
              alignment: docx.AlignmentType.CENTER,
            }),
            paragraph([textRun("學員蒙恩見證稿", { bold: true, size: 36 })], {
              alignment: docx.AlignmentType.CENTER,
            }),
            paragraph([textRun("(若本週改寫其他作業，則不必填寫本表)", { size: 20 })], {
              alignment: docx.AlignmentType.CENTER,
              spacing: { after: 220 },
            }),
            paragraph(
              [
                textRun(`第 ${groupStart} - ${groupEnd} 組   `),
                textRun(`姓名 ${name}   `),
                textRun(`第 ${week} 週   `),
                textRun(`${start}至${end}`),
              ],
              { spacing: { after: 160 } }
            ),
            paragraph([textRun("主題：", { bold: true }), textRun(topic, { bold: true })], {
              spacing: { after: 160 },
            }),
            paragraph([textRun("一、在訓練生活中：", { bold: true, size: 24 })], {
              spacing: { after: 80 },
            }),
            promptParagraph("1.對於訓練課程的心得及深刻得著的啟示與光照(略舉課程要點)"),
            promptParagraph("2.在訓練生活中的操練或專項服事的配搭上，所學到的生命功課(舉例述說)"),
            contentParagraph(values.trainingLife),
            paragraph([textRun("二、在當地召會生活中：", { bold: true, size: 24 })], {
              spacing: { before: 160, after: 80 },
            }),
            promptParagraph("1.如何將訓練所學的功課應用於個人及家庭生活中(列舉實例)"),
            promptParagraph("2.應用於召會開展之配搭服事(包括對人與事的方面)中的經歷(列舉實例)"),
            contentParagraph(values.churchLife),
          ],
        },
      ],
    });

    const blob = await docx.Packer.toBlob(doc);
    saveBlob(blob, fileName(values, "docx"));
  };

  const buildPdf = async () => {
    if (!window.html2pdf) throw new Error("PDF 套件尚未載入，請確認檔案路徑後重試。");
    const values = data();
    await window.html2pdf()
      .set({
        margin: 8,
        filename: fileName(values, "pdf"),
        image: { type: "jpeg", quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true, backgroundColor: "#fffdfa" },
        jsPDF: { unit: "mm", format: "a4", orientation: "portrait" },
        pagebreak: { mode: ["avoid-all"] },
      })
      .from(preview)
      .save();
  };

  const saveBlob = (blob, name) => {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = name;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  };

  const runAction = async (button, action) => {
    const original = button.textContent;
    button.disabled = true;
    button.textContent = "產生中...";
    try {
      await action();
    } catch (error) {
      alert(error.message || "產生檔案時發生錯誤。");
    } finally {
      button.disabled = false;
      button.textContent = original;
    }
  };

  form.addEventListener("input", syncPreview);
  document.getElementById("clearForm").addEventListener("click", () => {
    form.reset();
    document.getElementById("topic").value = "將訓練所學應用於生活與事奉中";
    syncPreview();
  });
  document.getElementById("downloadDocx").addEventListener("click", (event) => runAction(event.currentTarget, buildDocx));
  document.getElementById("downloadPdf").addEventListener("click", (event) => runAction(event.currentTarget, buildPdf));

  syncPreview();
})();
