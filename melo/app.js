const fmt = new Intl.NumberFormat("zh-CN");
const pct = (part, total) => total ? `${(part / total * 100).toFixed(1)}%` : "0%";
const ACCESS_HASH = "8b871155d3003ffd714cd01e64b9557df781b3e67241edd17bfe47eee622acd9";
const AUTH_KEY = "melo_access_granted";

function shortRoute(name) {
  return name
    .replace(" 单品单件快道", "")
    .replace(" 单品多件工作站", "")
    .replace(" 播种墙", "");
}

function metric(label, value, note) {
  return `<div class="metric"><span>${label}</span><strong>${value}</strong><em>${note}</em></div>`;
}

function bar(label, value, total, color) {
  const width = total ? Math.max(2, value / total * 100) : 0;
  return `
    <div class="bar-row">
      <div class="bar-label">${label}</div>
      <div class="track"><div class="fill" style="width:${width}%;background:${color}"></div></div>
      <div class="bar-value">${pct(value, total)}</div>
    </div>
  `;
}

function render(data) {
  const orders = data.orders;
  const inventory = data.inventory;
  const inbound = data.inbound;
  const totalOrders = orders.totals.orders;
  const routeColors = ["#1f7a66", "#2b8fb3", "#b7822b"];

  document.getElementById("metricGrid").innerHTML = [
    metric("30天出库订单", fmt.format(totalOrders), `${fmt.format(orders.totals.units)} 件`),
    metric("日均订单", fmt.format(Math.round(totalOrders / 30)), "按30天口径"),
    metric("规划面积", "12,000㎡", "Melo 专属区域"),
    metric("Alpha + Beta", pct((orders.route_orders["Alpha 单品单件快道"] || 0) + (orders.route_orders["Beta 单品多件工作站"] || 0), totalOrders), "前置区主流量"),
    metric("库存件数", fmt.format(inventory.totals.qty), `${fmt.format(Math.round(inventory.totals.pallets))} 托`),
    metric("SKU 数", fmt.format(inventory.totals.skus), `${fmt.format(Math.round(inventory.totals.volume))} m³`),
    metric("30天入库", fmt.format(inbound.totals.received_qty), `${fmt.format(inbound.totals.containers)} 个柜/跟踪号`),
    metric("P90 出库时效", `${orders.stats.outbound_hours_p90}h`, `中位数 ${orders.stats.outbound_hours_median}h`)
  ].join("");

  document.getElementById("diagnosis").innerHTML = data.design.diagnosis
    .map((text, i) => `<div class="diagnosis-item"><b>${String(i + 1).padStart(2, "0")}</b><span>${text}</span></div>`)
    .join("");

  document.getElementById("routeBars").innerHTML = Object.entries(orders.route_orders)
    .map(([name, value], i) => bar(shortRoute(name), value, totalOrders, routeColors[i % routeColors.length]))
    .join("");

  document.getElementById("zoneCards").innerHTML = data.design.zones
    .map(z => `<div class="zone-card"><strong>${z.name}</strong><p>${z.size}</p><p>${z.logic}</p></div>`)
    .join("");

  document.getElementById("processFlow").innerHTML = data.design.process
    .map(p => `<div class="process-step"><b>${p.step}</b><span>${p.text}</span></div>`)
    .join("");

  document.getElementById("staffing").innerHTML = data.design.staffing
    .map(s => `<div class="staff-card"><b>${s.area}</b><strong>${s.staff}</strong><span>${s.target}</span></div>`)
    .join("");

  document.getElementById("inboundCards").innerHTML = [
    ["入库行数", fmt.format(inbound.totals.lines)],
    ["收货件数", fmt.format(inbound.totals.received_qty)],
    ["上架件数", fmt.format(inbound.totals.putaway_qty)],
    ["到仓到上架 P90", `${inbound.stats.arrival_to_putaway_p90_h}h`],
    ["日均收货", fmt.format(inbound.stats.avg_daily_received_qty)],
    ["预报体积", `${fmt.format(Math.round(inbound.totals.forecast_volume))} m³`],
    ["收货完成", inbound.status_receiving[0]?.[0] || "未知"],
    ["上架完成", inbound.status_putaway[0]?.[0] || "未知"]
  ].map(([label, value]) => `<div class="inbound-card"><span>${label}</span><strong>${value}</strong></div>`).join("");

  document.getElementById("inboundSkuRows").innerHTML = inbound.top_skus.slice(0, 12).map(s => `
    <tr>
      <td>${s.sku}</td>
      <td class="num">${fmt.format(s.received_qty)}</td>
      <td class="num">${fmt.format(s.lines)}</td>
    </tr>
  `).join("");

  const maxInbound = Math.max(...inbound.daily_received.map(d => d.qty), 1);
  document.getElementById("inboundChart").innerHTML = inbound.daily_received.map(d => {
    const height = Math.max(4, d.qty / maxInbound * 165);
    return `<div class="daily-bar ${d.qty > maxInbound * 0.85 ? "peak" : ""}" style="height:${height}px" data-label="${d.date} · ${fmt.format(d.qty)}件"></div>`;
  }).join("");

  document.getElementById("topSkuRows").innerHTML = orders.top_skus.slice(0, 12).map(s => `
    <tr>
      <td>${s.sku}</td>
      <td class="num">${fmt.format(s.units_30d)}</td>
      <td class="num">${fmt.format(Math.round(s.avg_daily_units))}</td>
      <td class="num">${s.stock_days}</td>
      <td>${s.family}</td>
    </tr>
  `).join("");

  const familyTotal = Object.values(inventory.family_counter).reduce((a, b) => a + b, 0);
  const familyColors = ["#1f7a66", "#b7822b", "#2b8fb3", "#b65345"];
  document.getElementById("familyBars").innerHTML = Object.entries(inventory.family_counter)
    .sort((a, b) => b[1] - a[1])
    .map(([name, value], i) => bar(name, value, familyTotal, familyColors[i % familyColors.length]))
    .join("");

  document.getElementById("comboList").innerHTML = orders.top_combinations.slice(0, 6)
    .map(([combo, count]) => `<div class="combo"><b>${combo}</b><span>${fmt.format(count)} 单</span></div>`)
    .join("");

  const maxDaily = Math.max(...orders.daily.map(d => d.orders));
  document.getElementById("dailyChart").innerHTML = orders.daily.map(d => {
    const height = Math.max(4, d.orders / maxDaily * 230);
    return `<div class="daily-bar ${d.orders > 9000 ? "peak" : ""}" style="height:${height}px" data-label="${d.date} · ${fmt.format(d.orders)}单"></div>`;
  }).join("");

  const maxStructure = Math.max(...orders.daily_structure.map(d => d.total), 1);
  document.getElementById("structureStack").innerHTML = orders.daily_structure.map(d => {
    const barHeight = Math.max(4, d.total / maxStructure * 230);
    const one = d.single_one / d.total * 100;
    const two = d.single_multi / d.total * 100;
    const three = d.multi_sku / d.total * 100;
    const label = `${d.date}\n单品单件 ${fmt.format(d.single_one)}\n单品多件 ${fmt.format(d.single_multi)}\n多品多件 ${fmt.format(d.multi_sku)}`;
    return `
      <div class="stack-bar" style="height:${barHeight}px" data-label="${label}">
        <div class="stack-seg seg-one" style="height:${one}%"></div>
        <div class="stack-seg seg-two" style="height:${two}%"></div>
        <div class="stack-seg seg-three" style="height:${three}%"></div>
      </div>
    `;
  }).join("");

  document.getElementById("openQuestions").innerHTML = data.design.open_questions
    .map(q => `<div class="question">${q}</div>`)
    .join("");
}

async function sha256(text) {
  const bytes = new TextEncoder().encode(text);
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(digest))
    .map(byte => byte.toString(16).padStart(2, "0"))
    .join("");
}

function unlock() {
  document.body.classList.remove("locked");
  loadDashboard();
}

function loadDashboard() {
  fetch("./melo_analysis.json")
    .then(res => res.json())
    .then(render)
    .catch(err => {
      document.body.insertAdjacentHTML("beforeend", `<pre>${err.message}</pre>`);
    });
}

function initAuth() {
  const form = document.getElementById("authForm");
  const input = document.getElementById("accessPassword");
  const error = document.getElementById("authError");

  if (sessionStorage.getItem(AUTH_KEY) === "true") {
    unlock();
    return;
  }

  input?.focus();
  form?.addEventListener("submit", async event => {
    event.preventDefault();
    error.textContent = "";
    const hash = await sha256(input.value);
    if (hash === ACCESS_HASH) {
      sessionStorage.setItem(AUTH_KEY, "true");
      input.value = "";
      unlock();
      return;
    }
    error.textContent = "密码不正确，请重新输入。";
    input.select();
  });
}

initAuth();
