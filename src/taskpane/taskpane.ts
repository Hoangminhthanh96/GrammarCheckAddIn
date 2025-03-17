/* eslint-disable prettier/prettier */
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
};
let errorList = [] as GrammarError[];
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    try {
      document.getElementById("app-body").style.display = "flex";
      document.getElementById("check-text").onclick = checkText;
      document.getElementById("fix-all").onclick = fixAll;
      document.getElementById("accept").onclick = acceptChanges;
    } catch (error) {
      console.error("Lỗi khi khởi tạo add-in:", error);
      // Hiển thị thông báo lỗi cho người dùng
      document.body.innerHTML = `
        <div style="color: #d13438; padding: 20px; text-align: center;">
          <h2>Đã xảy ra lỗi khi khởi tạo add-in</h2>
          <p>Vui lòng thử tải lại add-in hoặc liên hệ hỗ trợ.</p>
          <p>Chi tiết lỗi: ${error.message}</p>
        </div>
      `;
    }
  }
});

async function checkText() {
  try {
    // Reset tất cả thông tin từ lần kiểm tra trước
    stats = {
      totalErrors: 0,
      fixedErrors: 0,
      remainingErrors: 0,
    };
    
    // Xóa nội dung các container hiển thị
    document.getElementById("error-list").innerHTML = "";
    document.getElementById("fix-all").style.display = "none";
    
    // Hiển thị biểu tượng loading
    document.getElementById("loading").style.display = "block";

    await Word.run(async (context) => {
      // Kiểm tra xem có đoạn văn bản nào được chọn không
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      if (!selection.text || selection.text.trim() === "") {
        document.getElementById("error-list").innerHTML =
          '<div class="error-card" style="color: #d13438">Vui lòng chọn đoạn văn bản cần kiểm tra</div>';
        return;
      }

      const result = await grammarCheck(selection.text);
      if (result.status === "success") {
        errorList = result.errors;
        
        if (errorList.length === 0) {
          document.getElementById("error-list").innerHTML =
            '<div class="error-card" style="color: #107c10; border-color: #107c10">Không tìm thấy lỗi chính tả trong văn bản đã chọn.</div>';
        } else {
          updateStats(result.errors.length, 0);
          await highlightErrors(result.errors, context);
          displayErrors(result.errors);
          document.getElementById("fix-all").style.display = "block";
        }
      }
    });
  } catch (error) {
    document.getElementById("error-list").innerHTML =
      `<div class="error-card" style="color: #d13438">Đã xảy ra lỗi: ${error.message}</div>`;
  } finally {
    document.getElementById("loading").style.display = "none";
  }
}

async function grammarCheck(text: string): Promise<CheckResult> {
  const apiUrl = "https://spellcheck.vcntt.tech/spellcheck";
  //const apiUrl = "https://httpbin.org/delay/11"; //test timeout
  //const proxyUrl = "https://cors-anywhere.herokuapp.com/";   
  try {
    // Thêm timeout cho fetch
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 10000);
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      mode: "cors",
      body: JSON.stringify({ text: text }),
      signal: controller.signal
    });
    
    clearTimeout(timeoutId);
    
    if (!response.ok) {
      throw new Error(`API error: ${response.status} ${response.statusText}`);
    }

    return await response.json();
  } catch (error) {
    if (error.name === 'AbortError') {
      document.getElementById("error-list").innerHTML = 
        '<div class="error-card" style="color: #d13438">Kết nối đến API bị timeout. Vui lòng thử lại sau.</div>';
    } else {
      document.getElementById("error-list").innerHTML = 
        `<div class="error-card" style="color: #d13438">Lỗi khi gọi API: ${error.message}</div>`;
    }
    
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
      const regex = new RegExp(`(?<!\\w)${error.word}(?!\\w)`, "g");
      const matches = [...text.matchAll(regex)];

      const positions = matches.map((match, index) => ({
        startIndex: match.index!,
        text: match[0], // Lấy từ đã khớp
        ordnumber: index,
      }));

      let position = positions.find((p) => p.startIndex == error.position);
      const range = selection.search(error.word, {
        matchCase: true,
        matchWholeWord: true,
      });
      range.load("text");
      await context.sync();

      if (range && range.items.length > 0 && position && range.items[position.ordnumber] != null) {
        range.items[position.ordnumber].font.color = "red";
        
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

  errors.forEach((error, index) => {
    const card = document.createElement("div");
    card.className = "error-card";
    card.innerHTML = `
      <div style="display: flex; align-items: center; flex-grow: 1">
        <div class="wrong-word-container">
          <div class="wrong-word" style="color: #d13438; font-weight: bold; border: 1px solid #d13438; padding: 4px 8px; border-radius: 12px;">
            ${error.word}
          </div>
        </div>
        <div class="suggestions" style="display: flex; gap: 8px; flex-wrap: wrap; flex-grow: 1;">
          ${error.suggestions
            .map(
              (word) =>
                `<span class="suggestion-chip" data-error-index="${index}" data-word="${word}" style="color: #107c10; border: 1px solid #107c10; padding: 4px 8px; border-radius: 12px;">${word}</span>`
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
      const word = errorList[errorIndex].word;
      const regex = new RegExp(`(?<!\\w)${word}(?!\\w)`, "g");
      const matches = [...text.matchAll(regex)];

      const positions = matches.map((match, index) => ({
        startIndex: match.index!,
        text: match[0], // Lấy từ đã khớp
        ordnumber: index,
      }));

      let position = positions.find((p) => p.startIndex == errorList[errorIndex].position);
      const range = selection.search(word, {
        matchCase: true,
        matchWholeWord: true,
      });
      range.load("text");
      await context.sync();

      // Kiểm tra xem có tồn tại range.items[position.ordnumber] hay không
      if (range && range.items.length > 0 && position && range.items[position.ordnumber] != null) {
        range.items[position.ordnumber].insertText(wordFixed, Word.InsertLocation.replace);
        range.items[position.ordnumber].font.color = "green";
        await context.sync();

        // Tính toán độ chênh lệch độ dài giữa từ cũ và từ mới
        const lengthDiff = wordFixed.length - word.length;

        // Cập nhật position cho các lỗi còn lại
        errorList.forEach((error, idx) => {
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

async function fixAll() {
  if (!errorList.length) return; // Nếu không có lỗi, không làm gì cả

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      for (let i = 0; i < errorList.length; i++) {
        const wordFixed = errorList[i].suggestions[0]; // Lấy từ gợi ý đầu tiên
        const word = errorList[i].word;
        
        const range = selection.search(word, {
          matchCase: true,
          matchWholeWord: true,
        });
        range.load("text");
        await context.sync();

        let text = selection.text;
        const regex = new RegExp(`(?<!\\w)${word}(?!\\w)`, "g");
        const matches = [...text.matchAll(regex)];

        const positions = matches.map((match, index) => ({
          startIndex: match.index!,
          text: match[0], // Lấy từ đã khớp
          ordnumber: index,
        }));

        let position = positions.find((p) => p.startIndex == errorList[i].position);
        if (range && range.items.length > 0 && position && range.items[position.ordnumber] != null) {
          range.items[position.ordnumber].insertText(wordFixed, Word.InsertLocation.replace);
          range.items[position.ordnumber].font.color = "green";
          await context.sync();
          
          //Không cần tính độ dài từ cũ và từ mới. Nếu thêm vào sẽ bị lệch vị trí
          /*
          // Tính toán độ chênh lệch độ dài giữa từ cũ và từ mới
          const lengthDiff = wordFixed.length - word.length;

          // Cập nhật position cho các lỗi còn lại
          //Có thể lấy j từ i+1 vì error sẽ sắp xếp theo vị trí từ nhỏ đến lớn
          for (let j: number = i + 1; j < errorList.length; j++) {
            //kiểm tra cho an toàn
            if (errorList[j].position > position.startIndex) { 
              errorList[j].position += lengthDiff;
            }
          }*/

          // Cập nhật UI và stats
          stats.fixedErrors++;
          stats.remainingErrors--;
          updateStats(stats.totalErrors, stats.fixedErrors);
        }
      }
      await context.sync();
    });
  } catch (error) {
    document.getElementById("error-list").innerHTML =
      `<div class="error-card" style="color: #d13438">Lỗi khi sửa tất cả: ${error.message}</div>`;
  }
}

function acceptChanges() {
  // Bỏ bôi màu các từ đã sửa trong Word
  Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const fixedWords = errorList.map((error) => error.suggestions[0]);
    const range = selection.search(fixedWords.join("|"), {
      matchCase: true,
      matchWholeWord: true,
    });
    range.load("text");
    await context.sync();

    for (const item of range.items) {
      item.font.color = ""; // Bỏ màu trong Word
    }
  });
  document.getElementById("fix-all").style.display = "none"; // Ẩn nút sửa tất cả
  document.getElementById("accept").style.display = "none"; // Ẩn nút xác nhận
}
