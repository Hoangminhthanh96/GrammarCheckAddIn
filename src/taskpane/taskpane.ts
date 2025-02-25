/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

interface GrammarError {
  position: number;
  word: string;
  suggestions: string[];
}

interface CheckResult {
  status: string;
  errors: GrammarError[];
}

let stats = {
  totalErrors: 0,
  fixedErrors: 0,
  remainingErrors: 0,
  currentErrors: [] as GrammarError[],
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("check-text").onclick = checkText;
  }
});

async function checkText() {
  try {
    /*
    // Reset tất cả thông tin từ lần kiểm tra trước
    stats = {
      totalErrors: 0,
      fixedErrors: 0,
      remainingErrors: 0,
      currentErrors: [],
    };
    
    // Xóa nội dung các container hiển thị
    document.getElementById("error-list").innerHTML = "";
    document.getElementById("positions-list").innerHTML = "";
    const rangeInfo = document.getElementById("range-info");
    if (rangeInfo) {
      rangeInfo.remove();
    }
    */
    // Hiển thị biểu tượng loading
    document.getElementById("loading").style.display = "block";

    await Word.run(async (context) => {
      // Kiểm tra xem có đoạn văn bản nào được chọn không
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      if (!selection.text || selection.text.trim() === "") {
        // Hiển thị thông báo nếu không có text được chọn
        document.getElementById("error-list").innerHTML =
          '<div class="error-card" style="color: #d13438">Vui lòng chọn đoạn văn bản cần kiểm tra</div>';
        return;
      }

      const result = await grammarCheck(selection.text);
      /*const result = {
        status: "success",
        has_errors: true,
        errors: [
          {
            word: "đona",
            suggestions: ["đoạn"],
            position: 25,
          },
          {
            word: "ban",
            suggestions: ["bản"],
            position: 34,
          },
          {
            word: "loi",
            suggestions: ["loại", "sai"],
            position: 41,
          },
          {
            word: "xai",
            suggestions: ["loại", "sai"],
            position: 45,
          },
        ],
        corrected_text: "Hoa ban trắng nở rộ. Một đoạn văn bản về loại sai chính tả. Cần hoàn thiên đoạn này ngay. ",
        processing_time: 0.2884,
      };*/
      if (result.status === "success") {
        updateStats(result.errors.length, 0);
        await highlightErrors(result.errors, context);
        displayErrors(result.errors);
      }
    });
  } catch (error) {
    //console.error("Error: " + error);
    // Hiển thị thông báo lỗi cho người dùng
    document.getElementById("error-list").innerHTML =
      `<div class="error-card" style="color: #d13438">Đã xảy ra lỗi: ${error.message}</div>`;
  } finally {
    // Ẩn biểu tượng loading
    document.getElementById("loading").style.display = "none";
  }
}

async function grammarCheck(text: string): Promise<CheckResult> {
  const apiUrl = "https://spellcheck.vcntt.tech/spellcheck";

  try {
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ text }),
    });

    if (!response.ok) {
      throw new Error(`API error: ${response.status} ${response.statusText}`);
    }

    const result: CheckResult = await response.json();
    return result;
  } catch (error) {
    // Thay thế console.error
    document.getElementById("error-list").innerHTML =
      `<div class="error-card" style="color: #d13438">Lỗi khi gọi API kiểm tra ngữ pháp: ${error.message}</div>`;
    return {
      status: "error",
      errors: [],
    };
  }
}

async function highlightErrors(errors: GrammarError[], context: Word.RequestContext) {
  try {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    for (const error of errors) {
      const text = selection.text;
      const regex = new RegExp(error.word, "g");
      const matches = [...text.matchAll(regex)];

      const positions = matches.map((match, index) => ({
        startIndex: match.index!,
        text: match[0], // Lấy từ đã khớp
        ordnumber: index, // Số thứ tự bắt đầu từ 1
      }));
      let position = positions.find((p) => p.startIndex == error.position);
      const range = selection.search(error.word, {
        matchCase: true,
        matchWholeWord: true,
      });
      range.load("text");
      await context.sync();

      /*if (range && range.items && range.items.length > 0) {
        range.items[0].font.color = "yellow";
      }*/
      if (range && range.items.length > 0 && position && range.items[position.ordnumber] != null) {
        range.items[position.ordnumber].font.color = "yellow";
      }
    }
    await context.sync();
  } catch (error) {
    // Thay thế console.error
    document.getElementById("error-list").innerHTML =
      `<div class="error-card" style="color: #d13438">Lỗi khi đánh dấu văn bản: ${error.message}</div>`;
    throw error;
  }
}

function updateStats(total: number, fixed: number) {
  stats.totalErrors = total;
  stats.fixedErrors = fixed;
  stats.remainingErrors = total - fixed;

  document.querySelector("#total-errors .stat-value").textContent = total.toString();
  document.querySelector("#fixed-errors .stat-value").textContent = fixed.toString();
  document.querySelector("#remaining-errors .stat-value").textContent = stats.remainingErrors.toString();
}

function displayErrors(errors: GrammarError[]) {
  const container = document.getElementById("error-list");
  container.innerHTML = "";

  // Lưu lại các lỗi hiện tại
  stats.currentErrors = errors;

  errors.forEach((error, index) => {
    const card = document.createElement("div");
    card.className = "error-card";
    card.innerHTML = `
      <div class="wrong-word">
        <span class="error-label">Lỗi:</span>
        ${error.word}
      </div>
      <div class="suggestions">
        <span class="suggestion-label">Gợi ý:</span>
        <div class="suggestion-chips">
          ${error.suggestions
            .map(
              (word) => `<span class="suggestion-chip" data-error-index="${index}" data-word="${word}">${word}</span>`
            )
            .join("")}
        </div>
      </div>
    `;
    container.appendChild(card);
  });

  // Thêm event listeners cho các suggestion
  document.querySelectorAll(".suggestion-chip").forEach((chip) => {
    chip.addEventListener("click", applySuggestion);
  });
}

async function applySuggestion(event: Event) {
  const chip = event.target as HTMLElement;
  const card = chip.closest(".error-card");

  // Kiểm tra nếu card đã được sửa thì không làm gì cả
  if (card.classList.contains("fixed")) {
    return;
  }

  const wordFixed = chip.getAttribute("data-word");
  const errorIndex = parseInt(chip.getAttribute("data-error-index"));

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      let text = selection.text;
      const word = stats.currentErrors[errorIndex].word;
      const regex = new RegExp(word, "g");
      const matches = [...text.matchAll(regex)];

      const positions = matches.map((match, index) => ({
        startIndex: match.index!,
        text: match[0], // Lấy từ đã khớp
        ordnumber: index, // Số thứ tự bắt đầu từ 1
      }));

      // Hiển thị nội dung của positions ra UI
      //const positionsContainer = document.getElementById("positions-list");
      /*positionsContainer.innerHTML = positions
        .map((pos) => `<div> ${pos.text} (Vị trí: ${pos.startIndex}, Thứ tự: ${pos.ordnumber})</div>`)
        .join("");*/

      let position = positions.find((p) => p.startIndex == stats.currentErrors[errorIndex].position);
      //positionsContainer.innerHTML = `<div> ${position.text} (Vị trí: ${position.startIndex}, Thứ tự: ${position.ordnumber + 1})</div>`;
      const range = selection.search(word, {
        matchCase: true,
        matchWholeWord: true,
      });
      range.load("text");
      await context.sync();
      /*
      // Thêm div mới để hiển thị thông tin về range
      const rangeInfo = document.createElement("div");
      rangeInfo.id = "range-info";
      positionsContainer.appendChild(rangeInfo);
      rangeInfo.innerHTML = `
        <div>Thông tin về Range:</div>
        <div>Tổng số kết quả tìm thấy: ${range.items.length}</div>
        ${range.items
          .map(
            (item, index) => `
          <div>Kết quả #${index}:
            <ul>
              <li>Văn bản: ${item.text}</li>
            </ul>
          </div>
        `
          )
          .join("")}
      `;*/

      // Kiểm tra xem có tồn tại range.items[position.ordnumber] hay không
      if (range && range.items.length > 0 && position && range.items[position.ordnumber] != null) {
        range.items[position.ordnumber].insertText(wordFixed, Word.InsertLocation.replace);
        range.items[position.ordnumber].font.color = "green";
        await context.sync();

        // Tính toán độ chênh lệch độ dài giữa từ cũ và từ mới
        const lengthDiff = wordFixed.length - word.length;

        // Cập nhật position cho các lỗi còn lại
        stats.currentErrors.forEach((error, idx) => {
          if (idx > errorIndex && error.position > position.startIndex) {
            error.position += lengthDiff;
          }
        });

        // Cập nhật UI và stats như cũ
        stats.fixedErrors++;
        stats.remainingErrors--;
        updateStats(stats.totalErrors, stats.fixedErrors);

        // Đánh dấu card đã sửa
        card.classList.add("fixed");

        // Disable tất cả các suggestion chips trong card này
        card.querySelectorAll(".suggestion-chip").forEach((chip) => {
          (chip as HTMLElement).style.pointerEvents = "none";
        });
      }
    });
  } catch (error) {
    document.getElementById("error-list").innerHTML =
      `<div class="error-card" style="color: #d13438">Lỗi khi áp dụng gợi ý: ${error.message}</div>`;
  }
}
