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
  const A4_WIDTH_DXA = 11906;
  const A4_HEIGHT_DXA = 16838;

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

  const createDocxBlob = async () => {
    if (!window.docx) throw new Error("DOCX 套件尚未載入，請重新整理頁面後再試。");
    if (!window.html2canvas) throw new Error("手機版 Word 版面套件尚未載入，請重新整理頁面後再試。");
    const docx = window.docx;
    const canvas = await captureA4Preview(1);

    const doc = new docx.Document({
      sections: [
        {
          properties: {
            page: {
              size: { width: A4_WIDTH_DXA, height: A4_HEIGHT_DXA },
              margin: { top: 0, right: 0, bottom: 0, left: 0 },
            },
          },
          children: [
            new docx.Paragraph({
              spacing: { before: 0, after: 0 },
              children: [
                new docx.ImageRun({
                  data: canvas.toDataURL("image/png"),
                  transformation: { width: 794, height: 1123 },
                  altText: {
                    title: "學員蒙恩見證稿",
                    description: "由學員蒙恩見證稿編輯器產生的固定 A4 版面",
                    name: "學員蒙恩見證稿",
                  },
                }),
              ],
            }),
          ],
        },
      ],
    });

    return docx.Packer.toBlob(doc);
  };

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
