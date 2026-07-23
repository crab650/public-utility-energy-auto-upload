const table = document.querySelector("#fill-table");
const tbody = table.querySelector("tbody");
const message = document.querySelector("#message");
const columnCount = 6;

for (const row of tbody.rows) {
    for (let column = 0; column < columnCount; column += 1) {
        const cell = row.insertCell();
        const input = document.createElement("input");
        input.type = "text";
        input.inputMode = "decimal";
        input.setAttribute("aria-label", `${row.cells[0].innerText.replace(/\n/g, " ")} 第 ${column + 1} 欄`);
        cell.appendChild(input);
    }
}

function showMessage(text, type = "info") {
    message.textContent = text;
    message.className = `message ${type} show`;
}

function normalizeNumber(rawValue) {
    let value = rawValue.trim().replace(/\s/g, "");
    if (value === "") return { valid: true, value: "" };
    const negative = /^\(.*\)$/.test(value);
    if (negative) value = value.slice(1, -1);
    value = value.replace(/,/g, "").replace(/NT\$|[$￥¥]/gi, "");
    if (negative) value = `-${value}`;
    return { valid: /^[-+]?\d+(\.\d+)?$/.test(value), value };
}

function getInput(rowIndex, columnIndex) {
    return tbody.rows[rowIndex]?.cells[columnIndex + 1]?.querySelector("input") || null;
}

table.addEventListener("paste", (event) => {
    const startInput = event.target.closest("input");
    if (!startInput) return;
    event.preventDefault();

    const startCell = startInput.closest("td");
    const startRow = startCell.parentElement.rowIndex - table.tHead.rows.length;
    const startColumn = startCell.cellIndex - 1;
    const clipboardText = event.clipboardData.getData("text/plain").replace(/\r/g, "").replace(/\n$/, "");
    const rows = clipboardText.split("\n").map((row) => row.split("\t"));
    let success = 0;
    let invalid = 0;
    let overflow = 0;

    table.querySelectorAll("input").forEach((input) => input.classList.remove("pasted", "invalid"));
    rows.forEach((values, rowOffset) => {
        values.forEach((rawValue, columnOffset) => {
            const input = getInput(startRow + rowOffset, startColumn + columnOffset);
            if (!input) {
                overflow += 1;
                return;
            }
            const result = normalizeNumber(rawValue);
            input.value = result.valid ? result.value : rawValue.trim();
            input.classList.add(result.valid ? "pasted" : "invalid");
            if (result.valid) success += 1;
            else invalid += 1;
        });
    });

    if (invalid || overflow) {
        showMessage(`已貼上 ${success} 格；${invalid} 格不是有效數字；${overflow} 格超出表格範圍。`, "warning");
    } else {
        showMessage(`貼上成功，共填入 ${success} 格（${rows.length} 列）。`, "success");
    }
});

document.querySelector("#clear-button").addEventListener("click", () => {
    table.querySelectorAll("input").forEach((input) => {
        input.value = "";
        input.classList.remove("pasted", "invalid");
    });
    showMessage("已清除全部測試資料。", "info");
});

document.querySelector("#sample-button").addEventListener("click", () => {
    const samples = [
        ["12500", "36800", "23100", "70400", "9800", "29600"],
        ["320", "4480", "610", "8540", "275", "3850"],
    ];
    samples.forEach((row, rowIndex) => row.forEach((value, columnIndex) => {
        const input = getInput(rowIndex, columnIndex);
        input.value = value;
        input.classList.add("pasted");
    }));
    showMessage("已填入兩列範例資料；按「全部清除」後可自行測試 Excel 貼上。", "success");
});

document.querySelector("#submit-button").addEventListener("click", () => {
    const invalid = table.querySelectorAll("input.invalid").length;
    showMessage(invalid ? `尚有 ${invalid} 格格式錯誤，請修正後再送出。` : "模擬送出成功（測試頁不會儲存資料）。", invalid ? "warning" : "success");
});
