/* global document, Office, Word */
import MarkdownIt from 'markdown-it';

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // 绑定 "Run" 按钮的事件
    document.getElementById("run").onclick = run;
    // 绑定 "Convert to Word" 按钮事件
    document.getElementById("convertToWord").onclick = convertMarkdownToWord;
  }
});

/**
 * 1) 将 Markdown 文本转为 HTML
 * 2) 使用 DOMParser（或 createElement）解析 HTML
 * 3) 递归处理各类标签：<h1/h2/h3>, <ul/ol>, <table>, <p> 等
 */
async function convertMarkdownToWord() {
  const md = new MarkdownIt();
  const markdownText = document.getElementById("markdownInput").value;

  // 1) 整块渲染
  const renderedHTML = md.render(markdownText);

  await Word.run(async (context) => {
    const body = context.document.body;

    // 2) 构造一个容器并解析 HTML
    const container = document.createElement("div");
    container.innerHTML = renderedHTML;

    // 3) 遍历 container 的子节点，依次插入到 Word
    for (let node of container.childNodes) {
      await processNode(node, body, context);
    }

    await context.sync();
  });

  console.log("Markdown successfully inserted into Word!");
}

/**
 * 递归处理节点（支持：<h1/h2/h3>, <ul/ol>, <table>, <p> 等常见块级元素）
 */
async function processNode(node, body, context, level = 0) {
  // 如果是文本节点（多半是空白），可直接忽略
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent.trim();
    if (text) {
      body.insertParagraph(text, Word.InsertLocation.end);
    }
    return;
  }

  // 如果是元素节点
  if (node.nodeType === Node.ELEMENT_NODE) {
    const tagName = node.nodeName.toLowerCase();

    // 1) 处理标题 <h1/h2/h3>
    if (["h1", "h2", "h3"].includes(tagName)) {
      const headingText = node.textContent.trim();
      if (!headingText) return;

      const paragraph = body.insertParagraph(headingText, Word.InsertLocation.end);
      if (tagName === "h1") {
        paragraph.styleBuiltIn = Word.Style.heading1;
      } else if (tagName === "h2") {
        paragraph.styleBuiltIn = Word.Style.heading2;
      } else {
        paragraph.styleBuiltIn = Word.Style.heading3;
      }
      return;
    }

    // 2) 处理列表 <ul> / <ol>
    if (tagName === "ul" || tagName === "ol") {
      let isOrdered = tagName === "ol";
      let bulletChar = isOrdered ? "1." : "•";

      let items = Array.from(node.children).filter(li => li.nodeName.toLowerCase() === "li");
      if (items.length === 0) return;

      for (let i = 0; i < items.length; i++) {
        let li = items[i];
        let text = li.firstChild?.textContent.trim() || "";

        let listPrefix = isOrdered ? `${i + 1}.` : bulletChar;

        let paragraph = body.insertParagraph(" ".repeat(level * 4) + listPrefix + " " + text, Word.InsertLocation.end);
        paragraph.styleBuiltIn = Word.Style.listParagraph;

        // 递归处理嵌套列表
        let nestedList = li.querySelector("ul, ol");
        if (nestedList) {
          await processNode(nestedList, body, context, level + 1);
        }
      }
      return;
    }

    // 3) 处理表格 <table>
    if (tagName === "table") {
      const tableData = parseHTMLTable(node);
      if (tableData) {
        insertWordTable(body, tableData, context);
      }
      return;
    }

    // 4) 处理段落 <p>
    if (tagName === "p") {
      const htmlContent = node.innerHTML.trim();
      if (htmlContent) {
        const paragraph = body.insertParagraph("", Word.InsertLocation.end);
        paragraph.insertHtml(htmlContent, Word.InsertLocation.replace);
      }
      return;
    }

    // 5) 递归处理 <div>, <section> 等容器
    if (["div", "section", "article"].includes(tagName)) {
      for (let child of node.childNodes) {
        await processNode(child, body, context, level);
      }
      return;
    }

    // 6) 其他未知情况，直接插入 textContent
    const fallbackText = node.textContent.trim();
    if (fallbackText) {
      body.insertParagraph(fallbackText, Word.InsertLocation.end);
    }
  }
}

/**
 * 从 <table> DOM 元素中解析出二维数组 [ [cell, cell], [cell, cell] ... ]
 */
function parseHTMLTable(tableElem) {
  const rows = tableElem.querySelectorAll("tr");
  if (!rows || rows.length < 1) return null;

  const tableData = [];
  for (let rowElem of rows) {
    const cells = rowElem.querySelectorAll("td, th");
    if (!cells || cells.length === 0) continue;

    let rowData = [];
    for (let cellElem of cells) {
      // 去除多余空白
      rowData.push(cellElem.textContent.trim() || " ");
    }
    tableData.push(rowData);
  }

  // 检查行数
  if (tableData.length < 1) return null;
  // 检查列数是否一致
  const colCount = tableData[0].length;
  if (!tableData.every(r => r.length === colCount)) {
    console.error("Invalid table: inconsistent column counts", tableData);
    return null;
  }

  return tableData;
}

/**
 * 将二维数组插入为 Word 原生表格，并优化表格样式
 * 现在增加 context 参数，确保可以加载表格内容
 */
async function insertWordTable(body, tableData, context) {
  // 确保每个单元格不为空
  for (let r = 0; r < tableData.length; r++) {
    for (let c = 0; c < tableData[r].length; c++) {
      if (!tableData[r][c]) {
        tableData[r][c] = " ";
      }
    }
  }
  const rows = tableData.length;
  const cols = tableData[0].length;
  console.log(`Inserting table with ${rows} rows and ${cols} columns`);

  // 插入表格
  const table = body.insertTable(rows, cols, Word.InsertLocation.end, tableData);
  table.styleBuiltIn = Word.Style.tableGrid; // 使用细线表格

  // 加载表格相关属性
  context.load(table, "rows/items/cells/items");
  await context.sync();

  // 为第一行（表头）逐个单元格添加灰色背景
  if (table.rows.items.length > 0) {
    let headerRow = table.rows.items[0];
    headerRow.cells.items.forEach(cell => {
      if (cell && cell.shading) {
        cell.shading.backgroundColor = "F2F2F2";
      }
    });
  }

  // 设置表格前后空行：先获取范围并判断 paragraphFormat 是否存在
  let range = table.getRange();
  if (range && range.paragraphFormat) {
    range.paragraphFormat.spaceBefore = 0;
    range.paragraphFormat.spaceAfter = 0;
  } else {
    console.warn("table.getRange().paragraphFormat is undefined, skipping space adjustments.");
  }

  // 设置表格边框颜色
  table.getBorder(Word.BorderType.insideHorizontal).color = "#000000";
  table.getBorder(Word.BorderType.insideVertical).color = "#000000";
  table.getBorder(Word.BorderType.top).color = "#000000";
  table.getBorder(Word.BorderType.bottom).color = "#000000";
  table.getBorder(Word.BorderType.left).color = "#000000";
  table.getBorder(Word.BorderType.right).color = "#000000";

  await context.sync();
}

/**
 * 测试按钮：插入一个 "Hello World" 段落
 */
async function run() {
  return Word.run(async (context) => {
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
    paragraph.font.color = "blue";
    await context.sync();
  });
}
