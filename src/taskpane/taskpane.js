/* global document, Office, Word */
import MarkdownIt from 'markdown-it';

/**
 * 将日志信息输出到侧边栏的 debugArea 中
 * 而不是依赖 console.log() / alert()
 */
function debugLog(msg) {
  const debugElem = document.getElementById("debugArea");
  if (debugElem) {
    debugElem.value += msg + "\n";
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // 绑定 "Run" 按钮
    document.getElementById("run").onclick = run;
    // 绑定 "Convert to Word" 按钮
    document.getElementById("convertToWord").onclick = convertMarkdownToWord;
  }
});

/**
 * 将 Markdown 转为 HTML，遍历 DOM 并插入 Word
 */
async function convertMarkdownToWord() {
  debugLog("convertMarkdownToWord() start...");

  const md = new MarkdownIt();
  const markdownText = document.getElementById("markdownInput").value;
  debugLog("Raw Markdown Input:\n" + markdownText);

  // 1) 整块渲染
  const renderedHTML = md.render(markdownText);
  debugLog("Converted HTML:\n" + renderedHTML);

  await Word.run(async (context) => {
    const body = context.document.body;

    // 2) 构造容器解析 HTML
    const container = document.createElement("div");
    container.innerHTML = renderedHTML;
    debugLog("Parsed Container InnerHTML:\n" + container.innerHTML);

    // 3) 遍历子节点插入 Word
    for (let node of container.childNodes) {
      await processNode(node, body, context);
    }
    await context.sync();
  });

  debugLog("Markdown successfully inserted into Word!");
}

/**
 * 递归处理常见标签：<h1/h2/h3>, <ul/ol>, <table>, <p>, etc.
 */
async function processNode(node, body, context, level = 0) {
  // 文本节点
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent.trim();
    if (text) {
      debugLog("Processing TEXT_NODE: " + text);
      body.insertParagraph(text, Word.InsertLocation.end);
    }
    return;
  }

  // 元素节点
  if (node.nodeType === Node.ELEMENT_NODE) {
    const tagName = node.nodeName.toLowerCase();
    debugLog("Processing ELEMENT_NODE: <" + tagName + "> ... </" + tagName + ">");

    // 1) 标题
    if (["h1", "h2", "h3"].includes(tagName)) {
      const headingText = node.textContent.trim();
      if (!headingText) {
        debugLog("Skipping empty heading");
        return;
      }
      debugLog("Inserting heading: " + headingText);
      const paragraph = body.insertParagraph(headingText, Word.InsertLocation.end);
      if (tagName === "h1") {
        paragraph.styleBuiltIn = Word.Style.heading1;
      } else if (tagName === "h2") {
        paragraph.styleBuiltIn = Word.Style.heading2;
      } else {
        paragraph.styleBuiltIn = Word.Style.heading3;
      }

      // 若环境支持，去掉标题段落 spaceAfter
      try {
        context.load(paragraph, "paragraphFormat");
        await context.sync();
        if (paragraph.paragraphFormat) {
          paragraph.paragraphFormat.spaceAfter = 0;
        }
        await context.sync();
      } catch (err) {
        debugLog("Failed to adjust heading spaceAfter=0: " + err);
      }
      return;
    }

    // 2) 列表 <ul>/<ol>
    if (tagName === "ul" || tagName === "ol") {
      let isOrdered = (tagName === "ol");
      let bulletChar = isOrdered ? "1." : "•";
      let items = Array.from(node.children).filter(li => li.nodeName.toLowerCase() === "li");
      if (items.length === 0) {
        debugLog("Skipping empty <ul>/<ol>");
        return;
      }
      for (let i = 0; i < items.length; i++) {
        let li = items[i];
        let text = li.firstChild?.textContent.trim() || "";
        let listPrefix = isOrdered ? `${i + 1}.` : bulletChar;
        debugLog("Inserting list item: " + listPrefix + " " + text);

        let paragraph = body.insertParagraph(" ".repeat(level * 4) + listPrefix + " " + text, Word.InsertLocation.end);
        paragraph.styleBuiltIn = Word.Style.listParagraph;

        // 嵌套列表
        let nestedList = li.querySelector("ul, ol");
        if (nestedList) {
          await processNode(nestedList, body, context, level + 1);
        }
      }
      return;
    }

    // 3) 表格 <table>
    if (tagName === "table") {
      const tableData = parseHTMLTable(node);
      if (tableData) {
        debugLog("Inserting table with row=" + tableData.length);
        await insertWordTable(body, tableData, context);
      } else {
        debugLog("Skipping invalid table");
      }
      return;
    }

    // 4) 段落 <p>
    if (tagName === "p") {
      const rawHtml = node.innerHTML;
      // 去掉 <br> / &nbsp; / 空白
      let stripped = rawHtml.replace(/<br\s*\/?>/gi, "").replace(/&nbsp;/gi, "").trim();
      if (!stripped) {
        debugLog("Skipping empty paragraph <p>");
        return;
      }
      debugLog("Inserting paragraph <p>: " + stripped);

      const paragraph = body.insertParagraph("", Word.InsertLocation.end);
      paragraph.insertHtml(rawHtml, Word.InsertLocation.replace);
      return;
    }

    // 5) 容器标签 <div>/<section>/<article>
    if (["div", "section", "article"].includes(tagName)) {
      debugLog("Recursively processing container <" + tagName + ">");
      for (let child of node.childNodes) {
        await processNode(child, body, context, level);
      }
      return;
    }

    // 6) 其他未知情况
    const fallbackText = node.textContent.trim();
    if (fallbackText) {
      debugLog("Inserting fallback text for <" + tagName + ">: " + fallbackText);
      body.insertParagraph(fallbackText, Word.InsertLocation.end);
    } else {
      debugLog("Skipping unknown <" + tagName + "> with no text");
    }
  }
}

/**
 * 从 <table> DOM 中解析出二维数组
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
      rowData.push(cellElem.textContent.trim() || " ");
    }
    tableData.push(rowData);
  }

  // 检查行列一致
  const colCount = tableData[0].length;
  if (!tableData.every(r => r.length === colCount)) {
    debugLog("Invalid table: inconsistent column counts");
    return null;
  }
  return tableData;
}

/**
 * 插入表格并手动设置边框
 */
async function insertWordTable(body, tableData, context) {
  for (let r = 0; r < tableData.length; r++) {
    for (let c = 0; c < tableData[r].length; c++) {
      if (!tableData[r][c]) {
        tableData[r][c] = " ";
      }
    }
  }
  const rows = tableData.length;
  const cols = tableData[0].length;
  debugLog(`Inserting table with ${rows} rows and ${cols} columns`);

  const table = body.insertTable(rows, cols, Word.InsertLocation.end, tableData);

  // 不使用内置样式
  try {
    if (typeof table.clearFormats === "function") {
      table.clearFormats();
    }
  } catch (e) {
    debugLog("table.clearFormats() not supported, skip clearing formats");
  }

  // 设置边框颜色
  table.getBorder(Word.BorderType.insideHorizontal).color = "#000000";
  table.getBorder(Word.BorderType.insideVertical).color = "#000000";
  table.getBorder(Word.BorderType.top).color = "#000000";
  table.getBorder(Word.BorderType.bottom).color = "#000000";
  table.getBorder(Word.BorderType.left).color = "#000000";
  table.getBorder(Word.BorderType.right).color = "#000000";

  // 尝试去除表格段前后空行
  context.load(table, "rows/items/cells/items");
  await context.sync();

  let range = table.getRange();
  if (range && range.paragraphFormat) {
    range.paragraphFormat.spaceBefore = 0;
    range.paragraphFormat.spaceAfter = 0;
  } else {
    debugLog("table.getRange().paragraphFormat is undefined, skip spacing adjustments.");
  }

  await context.sync();
  debugLog("Table inserted successfully.");
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
