(function () {
  const fields = [
    "group",
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

  const captureA4Preview = async (scale = 2) => {
    if (!window.html2canvas) throw new Error("版面產生套件尚未載入，請重新整理頁面後再試。");
    const exportPreview = preview.cloneNode(true);
    exportPreview.id = "documentPreviewExport";
    exportPreview.classList.add("pdf-export-page");
    document.body.appendChild(exportPreview);

    try {
      return await window.html2canvas(exportPreview, {
        scale,
        useCORS: true,
        backgroundColor: "#ffffff",
      });
    } finally {
      exportPreview.remove();
    }
  };

  async function createDocxBlob() {
    if (!window.docx) throw new Error("Word 套件尚未載入，請重新整理後再試。");
    const d = window.docx;
    const data = Object.fromEntries(fields.map(key => [key, document.getElementById(key).value.trim()]));
    const font = { ascii: "Microsoft JhengHei", hAnsi: "Microsoft JhengHei", eastAsia: "Microsoft JhengHei", cs: "Microsoft JhengHei" };
    const paragraph = (text, options = {}) => new d.Paragraph({
      alignment: options.alignment || d.AlignmentType.LEFT,
      spacing: { before: 0, after: options.after ?? 0, line: options.line || 340, lineRule: d.LineRuleType.EXACT },
      keepNext: false,
      keepLines: false,
      children: text.split(/\r?\n/).map((line, index) => new d.TextRun({
        text: line, break: index ? 1 : undefined, font,
        size: options.size || 26, bold: options.bold || false, snapToGrid: false,
      })),
    });
    const tableWidth = 10706;
    const border = { style: d.BorderStyle.SINGLE, size: 6, color: "C9D1D9" };
    const contentBox = (text, height) => new d.Table({
      width: { size: tableWidth, type: d.WidthType.DXA }, columnWidths: [tableWidth], layout: d.TableLayoutType.FIXED,
      margins: { top: 100, bottom: 100, left: 140, right: 140 },
      borders: { top: border, bottom: border, left: border, right: border },
      rows: [new d.TableRow({ cantSplit: false, height: { value: height, rule: d.HeightRule.ATLEAST }, children: [new d.TableCell({ width: { size: tableWidth, type: d.WidthType.DXA }, verticalAlign: d.VerticalAlign.TOP, children: (text || "").split(/\r?\n/).map(line => paragraph(line)) })] })],
    });
    const doc = new d.Document({ title: "學員蒙恩見證稿",
      styles: { default: { document: { run: { font, size: 26, color: "000000" }, paragraph: { spacing: { after: 0 } } } } },
      sections: [{
        properties: { verticalAlign: d.VerticalAlign.TOP, page: { size: { width: 11906, height: 16838 }, margin: { top: 600, right: 600, bottom: 600, left: 600 } } },
        children: [
          paragraph("臺灣福音工作全時間訓練壯年成全班", { bold: true, size: 28, alignment: d.AlignmentType.CENTER, line: 360, after: 40 }),
          paragraph("學員蒙恩見證稿", { bold: true, size: 38, alignment: d.AlignmentType.CENTER, line: 460, after: 100 }),
          paragraph("（若本週改寫其他作業，則不必填寫本表）", { size: 20, alignment: d.AlignmentType.CENTER, line: 260, after: 80 }),
          paragraph(`第 ${data.group} 組    姓名 ${data.name}    第 ${data.week} 週`, { size: 24, line: 300, after: 20 }),
          paragraph(`${dateText(data.startDate)} 至 ${dateText(data.endDate)}`, { size: 24, line: 300, after: 80 }),
          paragraph(`主題：${data.topic}`, { bold: true, size: 26, line: 340, after: 100 }),
          ...Array.from(preview.querySelectorAll(".doc-section"), (section, index) => [
            paragraph(section.querySelector("h3").textContent, { bold: true, alignment: d.AlignmentType.CENTER, line: 340, after: 40 }),
            ...Array.from(section.querySelectorAll("li"), (node, i) => paragraph(`${i + 1}. ${node.textContent}`, { size: 22, line: 280 })),
            contentBox(index === 0 ? data.trainingLife : data.churchLife, 3200),
          ]).flat(),
        ],
      }],
    });
    return d.Packer.toBlob(doc);
  }

  const buildDocx = async () => {
    const values = data();
    const blob = await createDocxBlob();
    saveBlob(blob, fileName(values, "docx"));
  };

  const shareDocxToLine = async () => {
    const values = data();
    const name = fileName(values, "docx");
    const blob = await createDocxBlob();
    const file = new File([blob], name, {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });

    if (navigator.canShare?.({ files: [file] }) && navigator.share) {
      await navigator.share({
        title: "學員蒙恩見證稿",
        text: "學員蒙恩見證稿 Word 檔",
        files: [file],
      });
      return;
    }

    saveBlob(blob, name);
    alert("這個瀏覽器不支援直接分享 Word 檔。已先下載 Word，請到 LINE 選擇檔案傳送。");
  };

  const buildPdf = async () => {
    if (!window.html2canvas || !window.jspdf?.jsPDF) throw new Error("PDF 單頁產生套件尚未載入，請重新整理頁面後再試。");
    const values = data();
    const canvas = await captureA4Preview(2);
    const pdf = new window.jspdf.jsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
    const pageWidth = 210;
    const pageHeight = 297;
    const margin = 6;
    const maxWidth = pageWidth - margin * 2;
    const maxHeight = pageHeight - margin * 2;
    const imageRatio = canvas.width / canvas.height;
    let imageWidth = maxWidth;
    let imageHeight = imageWidth / imageRatio;

    if (imageHeight > maxHeight) {
      imageHeight = maxHeight;
      imageWidth = imageHeight * imageRatio;
    }

    const x = (pageWidth - imageWidth) / 2;
    const y = (pageHeight - imageHeight) / 2;
    pdf.addImage(canvas.toDataURL("image/jpeg", 0.98), "JPEG", x, y, imageWidth, imageHeight);
    pdf.save(fileName(values, "pdf"));
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
  document.getElementById("shareDocx").addEventListener("click", (event) => runAction(event.currentTarget, shareDocxToLine));

  syncPreview();
})();
