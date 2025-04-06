const works = [
  { name: "Шпаклювання", unit: "м²", qty: 0, priceClient: 320, priceWorker: 280 },
  { name: "Плитка", unit: "м²", qty: 0, priceClient: 650, priceWorker: 550 },
  { name: "Електрика (штроба)", unit: "м.п.", qty: 0, priceClient: 70, priceWorker: 55 },
  { name: "Фундамент", unit: "м.п./м³", qty: 0, priceClient: 1000, priceWorker: 2500 }
];

function renderTable() {
  const table = document.getElementById("works-table");
  table.innerHTML = "";
  works.forEach((w, i) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${w.name}</td>
      <td>${w.unit}</td>
      <td><input type='number' value='${w.qty}' onchange='updateQty(${i}, this.value)' /></td>
      <td>${w.priceClient}</td>
      <td>${w.priceWorker}</td>
      <td id='sumClient${i}'>0</td>
      <td id='sumWorker${i}'>0</td>
    `;
    table.appendChild(row);
  });
}

function updateQty(index, value) {
  works[index].qty = parseFloat(value) || 0;
  document.getElementById(`sumClient${index}`).textContent = works[index].qty * works[index].priceClient;
  document.getElementById(`sumWorker${index}`).textContent = works[index].qty * works[index].priceWorker;
}

function downloadExcel() {
  const data = works.map(w => ({
    "Вид робіт": w.name,
    "Одиниця": w.unit,
    "Обсяг": w.qty,
    "Ціна для замовника": w.priceClient,
    "Ціна для працівника": w.priceWorker,
    "Сума для замовника": w.qty * w.priceClient,
    "Сума для працівника": w.qty * w.priceWorker
  }));
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Кошторис");
  XLSX.writeFile(workbook, "K_Estimator.xlsx");
}

renderTable();
