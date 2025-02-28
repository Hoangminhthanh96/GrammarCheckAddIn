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
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("check-text").onclick = checkText;
    document.getElementById("fix-all").onclick = fixAll;
    document.getElementById("accept").onclick = acceptChanges;
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
            word: "ẫn",
            suggestions: ["dẫn"],
            position: 343,
          },
          {
            word: "chãy",
            suggestions: ["chảy"],
            position: 502,
          },
          {
            word: "vậc",
            suggestions: ["vật", "vã", "trở", "về", "vội", "vã", "ra", "đi"],
            position: 671,
          },
          {
            word: "củ",
            suggestions: ["cũ"],
            position: 1082,
          },
        ],
        corrected_text:
          "Trong một ngôi làng nhỏ nằm giữa những ngọn đồi trập trùng, có một cậu bé tên là Nam. Cậu sống cùng ông bà trong một căn nhà gỗ nhỏ bé nhưng ấm cúng. Hằng ngày, Nam thường dậy từ sớm để phụ giúp ông bà làm việc nhà, sau đó cậu lại chạy ra đồng chơi đùa cùng lũ bạn. Một hôm, khi đang chơi gần bìa rừng, Nam vô tình phát hiên một con đường nhỏ dẫn sau những bụi cây rậm rạp. Vì tò mò, cậu quyết định đi theo con đường ấy mà không hề báo với ai. \nĐi được một đoạn, Nam bắt gặp một con suối trong vắt, nước chảy róc rách nghe thật vui tai. Bên cạnh con suối là một cái cây cổ thụ to lớn, trên cành cây có một cái tổ chim với mấy chú chim non đang đợi mẹ mang thức ăn về. Cảnh vật xung quanh thật đẹp làm Nam mê mẫng, cậu cứ đứng đó ngắm nhìn mà quên mất thời gian. Đến khi mặt trời bắt đầu khuất sau những rặng cây, Nam mới giật mình nhớ ra rằng mình đã đi quá xa. \nCậu vội vã quay trở về, nhưng càng đi càng thấy mọi thứ xung quanh trở nên lạ lấm. Nam hoang mang, không biết phải làm sao. Đúng lúc đó, cậu nghe thấy tiếng gọi quen thuộc của ông. Nhờ vào tiếng gọi ấy, Nam lần theo đường cũ và tìm được lối ra khỏi rừng. Khi trở về nhà, ông bà trách mắng Nam một hồi, nhưng sau cùng vẫn ôm cậu vào lòng, dặn dò cậu rằng lần sau không được tự ý đi vào rừng nửa. Nam gật đầu lia lịa, tự hứa với mình rằng sẽ không bao giờ để bản thân rơi vào tình huống như vậy lần nào nửa. ",
        processing_time: 1.7545,
      };*/
      if (result.status === "success") {
        // Lưu lại các lỗi hiện tại
        errorList = result.errors;
        updateStats(result.errors.length, 0);
        await highlightErrors(result.errors, context);
        displayErrors(result.errors);
        document.getElementById("fix-all").style.display = "block"; // Hiển thị nút sửa tất cả
        //document.getElementById("accept").style.display = "block"; // Hiển thị nút xác nhận
      }
    });
  } catch (error) {
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
  const proxyUrl = "https://cors-anywhere.herokuapp.com/";
  const debugContainer = document.getElementById("debug-container");

  try {
    /*// Hiển thị thông tin request
    debugContainer.innerHTML = `
      <div class="error-card">
        <div>Đang gọi API với text:</div>
        <div style="word-break: break-all;">${text}</div>
      </div>
    `;*/

    const response = await fetch(proxyUrl + apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      mode: "cors",
      body: JSON.stringify({ text: text }),
    });

    /*// Hiển thị thông tin response status
    debugContainer.innerHTML += `
      <div class="error-card">
        <div>Trạng thái phản hồi API:</div>
        <div>Status: ${response.status}</div>
        <div>StatusText: ${response.statusText}</div>
      </div>
    `;*/

    if (!response.ok) {
      const errorText = await response.text();
      /*// Hiển thị chi tiết lỗi
      debugContainer.innerHTML += `
        <div class="error-card" style="color: #d13438">
          <div>Lỗi từ API:</div>
          <div>${errorText}</div>
        </div>
      `;*/
      throw new Error(`API error: ${response.status} ${response.statusText} - ${errorText}`);
    }

    const result = await response.json();
    /*// Hiển thị kết quả API
    debugContainer.innerHTML += `
      <div class="error-card">
        <div>Dữ liệu phản hồi từ API:</div>
        <pre style="white-space: pre-wrap;">${JSON.stringify(result, null, 2)}</pre>
      </div>
    `;*/
    return result;
  } catch (error) {
    // Hiển thị lỗi nếu có
    debugContainer.innerHTML += `
      <div class="error-card" style="color: #d13438">
        <div>Lỗi khi gọi API kiểm tra ngữ pháp:</div>
        <div>${error.message}</div>
      </div>
    `;
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
      /*
      // Hiển thị nội dung của positions ra UI
      const positionsContainer = document.getElementById("debug-container");
      positionsContainer.innerHTML = positions
        .map((pos) => `<div> ${pos.text} (Vị trí từ: ${pos.startIndex}, Thứ tự: ${pos.ordnumber})</div>`)
        .join("");*/

      let position = positions.find((p) => p.startIndex == error.position);
      /*positionsContainer.innerHTML += `Vị trí lỗi sai: <div> ${position.text} (Vị trí: ${position.startIndex}, Thứ tự: ${position.ordnumber})</div>`;*/
      const range = selection.search(error.word, {
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
      `;
      */
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
