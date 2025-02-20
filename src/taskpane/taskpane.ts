/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

interface GrammarError {
  position: number;
  wrongWord: string;
  suggestionWords: string[];
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

      // Phần code kiểm tra lỗi chính tả
      const result = {
        status: "success",
        errors: [
          {
            position: 0,
            wrongWord: "Loi",
            suggestionWords: ["Lỗi", "Lời"],
          },
          {
            position: 4,
            wrongWord: "xai",
            suggestionWords: ["sai", "xài", "xa"],
          },
          {
            position: 8,
            wrongWord: "chinh",
            suggestionWords: ["chính", "chình"],
          },
          {
            position: 14,
            wrongWord: "ta",
            suggestionWords: ["tả", "tà"],
          },
        ],
      };

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
  }
}

/*async function mockGrammarCheck(text: string): Promise<CheckResult> {
  // Giả lập API call - thay thế bằng API thật sau này
  return new Promise((resolve) => {
    setTimeout(() => {
      resolve({
        status: "success",
        errors: [
          {
            position: 0,
            wrongWord: "Loi",
            suggestionWords: ["Lỗi"],
          },
          {
            position: 4,
            wrongWord: "xai",
            suggestionWords: ["sai"],
          },
          {
            position: 8,
            wrongWord: "chinh",
            suggestionWords: ["chính"],
          },
          {
            position: 14,
            wrongWord: "ta",
            suggestionWords: ["tả"],
          },
        ],
      });
    }, 1000);
  });
}*/

async function highlightErrors(errors: GrammarError[], context: Word.RequestContext) {
  try {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    for (const error of errors) {
      const range = selection.search(error.wrongWord, {
        matchCase: true,
        matchWholeWord: true,
      });
      range.load("text");
      await context.sync();

      // Kiểm tra kỹ hơn trước khi highlight
      if (range && range.items && range.items.length > 0) {
        range.items[0].font.color = "yellow";
      }
    }
    await context.sync();
  } catch (error) {
    console.error("Error highlighting text:", error);
    throw error; // Ném lỗi để hàm gọi có thể xử lý
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
        ${error.wrongWord}
      </div>
      <div class="suggestions">
        <span class="suggestion-label">Gợi ý:</span>
        <div class="suggestion-chips">
          ${error.suggestionWords
            .map(
              (word) =>
                `<span class="suggestion-chip" data-error-index="${index}" data-word="${word}">${word}</span>`
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

  const word = chip.getAttribute("data-word");
  const errorIndex = parseInt(chip.getAttribute("data-error-index"));

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const wrongWord = stats.currentErrors[errorIndex].wrongWord;
      const range = selection.search(wrongWord, {
        matchCase: true,
        matchWholeWord: true,
      });
      range.load("text");
      await context.sync();

      if (range && range.items.length > 0) {
        range.items[0].insertText(word, Word.InsertLocation.replace);
        range.items[0].font.color = "green";
      }
      await context.sync();

      // Cập nhật UI
      stats.fixedErrors++;
      stats.remainingErrors--;
      updateStats(stats.totalErrors, stats.fixedErrors);

      // Đánh dấu card đã sửa
      card.classList.add("fixed");
      
      // Disable tất cả các suggestion chips trong card này
      card.querySelectorAll(".suggestion-chip").forEach((chip) => {
        (chip as HTMLElement).style.pointerEvents = "none";
      });
    });
  } catch (error) {
    console.error("Error: " + error);
  }
}
