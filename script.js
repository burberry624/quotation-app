let index = 1;

// 四捨五入
function round(num, decimal = 2) {
  return Number(Number(num).toFixed(decimal));
}

// 新增列
function addRow() {
  const row = document.createElement("tr");

  row.innerHTML = `
    <td>${index++}</td>
    <td><input></td>
    <td><input type="number" value="0"></td>
    <td><input type="number" class="totalW" value="0"></td>
    <td><input type="number" class="tareW" value="0"></td>
    <td class="netW">0</td>
    <td><input type="number" class="price" value="0"></td>
    <td class="amount">0</td>
  `;

  document.getElementById("items").appendChild(row);

  row.querySelectorAll("input").forEach(i => {
    i.addEventListener("input", () => {
      calc();
      saveData();
    });
  });
}

// 計算
function calc() {
  let total = 0;

  document.querySelectorAll("#items tr").forEach(row => {

    const totalW = Number(row.querySelector(".totalW")?.value) || 0;
    const tareW = Number(row.querySelector(".tareW")?.value) || 0;
    const price = Number(row.querySelector(".price")?.value) || 0;

    let net = totalW - tareW;
    net = round(net, 2);

    let amount = net * price;
    amount = Math.round(amount);

    row.querySelector(".netW").innerText = net;
    row.querySelector(".amount").innerText = amount.toLocaleString();

    total += amount;
  });

  document.getElementById("total").innerText = total.toLocaleString();
}

// 預設列
window.onload = () => {
  for (let i = 0; i < 6; i++) addRow();

  const logo = localStorage.getItem("logo");
  if (logo) logoPreview.src = logo;

  loadData();
};

// PDF
function downloadPDF() {
  const pdf = document.getElementById("pdf");

  let html = `<h1 style="text-align:center;">報價單</h1>`;
  html += `<p>客戶：${clientName.value}</p>`;
  html += `<p>電話：${phone.value}</p>`;
  html += `<p>地址：${address.value}</p>`;
  html += `<table border="1" width="100%">`;

  html += `<tr>
    <th>品項</th><th>總重</th><th>空重</th><th>淨重</th><th>單價</th><th>金額</th>
  </tr>`;

  document.querySelectorAll("#items tr").forEach(row => {
    html += `<tr>
      <td>${row.children[1].querySelector("input").value}</td>
      <td>${row.children[3].querySelector("input").value}</td>
      <td>${row.children[4].querySelector("input").value}</td>
      <td>${row.children[5].innerText}</td>
      <td>${row.children[6].querySelector("input").value}</td>
      <td>${row.children[7].innerText}</td>
    </tr>`;
  });

  html += `</table><h2>總計 ${total.innerText}</h2>`;

  pdf.innerHTML = html;
  pdf.style.display = "block";

  html2pdf().from(pdf).save("報價單.pdf");

  pdf.style.display = "none";
}

// Excel
function downloadExcel() {
  const data = [["品項","總重","空重","淨重","單價","金額"]];

  document.querySelectorAll("#items tr").forEach(row => {
    data.push([
      row.children[1].querySelector("input").value,
      row.children[3].querySelector("input").value,
      row.children[4].querySelector("input").value,
      row.children[5].innerText,
      row.children[6].querySelector("input").value,
      row.children[7].innerText
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "報價單");
  XLSX.writeFile(wb, "報價單.xlsx");
}

// 自動儲存
function saveData() {
  const data = [];

  document.querySelectorAll("#items tr").forEach(row => {
    data.push({
      name: row.children[1].querySelector("input").value,
      totalW: row.children[3].querySelector("input").value,
      tareW: row.children[4].querySelector("input").value,
      price: row.children[6].querySelector("input").value
    });
  });

  localStorage.setItem("data", JSON.stringify(data));
}

function loadData() {
  const data = JSON.parse(localStorage.getItem("data"));
  if (!data) return;

  document.getElementById("items").innerHTML = "";
  index = 1;

  data.forEach(d => {
    addRow();
    const row = document.querySelectorAll("#items tr")[index-2];

    row.children[1].querySelector("input").value = d.name;
    row.children[3].querySelector("input").value = d.totalW;
    row.children[4].querySelector("input").value = d.tareW;
    row.children[6].querySelector("input").value = d.price;
  });

  calc();
}

// Logo
logoInput.onchange = e => {
  const reader = new FileReader();
  reader.onload = () => {
    logoPreview.src = reader.result;
    localStorage.setItem("logo", reader.result);
  };
  reader.readAsDataURL(e.target.files[0]);
};

// Google Sheet
function sendToGoogleSheet() {
  const data = [];

  document.querySelectorAll("#items tr").forEach(row => {
    data.push({
      name: row.children[1].querySelector("input").value,
      totalW: row.children[3].querySelector("input").value,
      tareW: row.children[4].querySelector("input").value,
      net: row.children[5].innerText,
      price: row.children[6].querySelector("input").value,
      amount: row.children[7].innerText
    });
  });

  fetch("https://docs.google.com/spreadsheets/d/19WP5qaib6mGfaZZxPDm0uUdZpU6ftnhKjuyrQ4xHDto/edit?gid=0#gid=0", {
    method: "POST",
    body: JSON.stringify(data)
  });

  alert("已同步");
}
