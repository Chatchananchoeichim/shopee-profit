function renderDashboard() {
  const isDark = state.theme === 'dark';
  const textColor = isDark ? '#94a3b8' : '#64748b';
  const gridColor = isDark ? '#334155' : '#e2e8f0';

  // Clear existing ApexCharts instances if they exist
  if(state.charts.marginTier) state.charts.marginTier.destroy();
  if(state.charts.timeSeries) state.charts.timeSeries.destroy();
  if(state.charts.bcg) state.charts.bcg.destroy();
  if(state.charts.topProfit) state.charts.topProfit.destroy();
  if(state.charts.salesQty) state.charts.salesQty.destroy();
  if(state.charts.costBreakdown) state.charts.costBreakdown.destroy();

  // Calculation logic
  let totalRevenue = state.summary.reduce((a,b)=>a+b.revenue, 0);
  let totalCost = state.summary.reduce((a,b)=>a+b.cost, 0);
  let totalNet = totalRevenue - totalCost;
  let overallMargin = totalRevenue > 0 ? (totalNet / totalRevenue) * 100 : 0;
  let totalAdSpend = state.adsData.reduce((a, b) => a + b.adSpend, 0);
  let netAfterAds = totalNet - totalAdSpend;

  let tierA = 0, tierB = 0, tierC = 0, lossCount = 0;
  let skuProfitCount = 0, starCount = 0;
  let sumRev = 0, sumProfit = 0;
  const bcgData = [];

  state.summary.forEach(s => {
    const profit = s.revenue - s.cost;
    const marginPct = s.revenue > 0 ? (profit / s.revenue) * 100 : 0;
    if(marginPct >= 30) tierA++; else if(marginPct >= 15) tierB++; else if(profit > 0) tierC++;
    if(profit > 0) skuProfitCount++; else if(profit < 0) lossCount++;
    if(s.qty > 0) {
      sumRev += s.revenue; sumProfit += profit;
      bcgData.push({ x: Math.round(s.revenue), y: Math.round(profit), z: Math.max(8, Math.min(s.qty * 1.5, 30)), label: s.sku || s.title.substring(0, 15) });
    }
  });

  const avgRev =  bcgData.length ? sumRev / bcgData.length : 0;
  const avgProfit = bcgData.length ? sumProfit / bcgData.length : 0;
  bcgData.forEach(item => { if(item.x >= avgRev && item.y >= avgProfit) starCount++; });

  // Update DOM KPI
  const setVal = (id, txt) => { const el = document.getElementById(id); if(el) el.innerText = txt; };
  setVal('d-margin', overallMargin.toFixed(1) + '%');
  setVal('d-star-count', starCount + ' SKUs');
  setVal('d-profit-skus', skuProfitCount + ' SKUs');
  setVal('d-loss-skus', lossCount + ' SKUs');
  setVal('d-total-ads', '฿' + Math.round(totalAdSpend).toLocaleString());
  setVal('d-net-after-ads', '฿' + Math.round(netAfterAds).toLocaleString());

  const adsSection = document.getElementById('ads-summary-dashboard');
  if (adsSection) adsSection.style.display = totalAdSpend > 0 ? 'grid' : 'none';

  // Shop Stats aggregation
  const statsSection = document.getElementById('shop-stats-dashboard');
  if (statsSection && state.shopStats && state.shopStats.length > 0) {
    statsSection.style.display = 'grid';
    
    // Aggregated stats (Prioritize Shopee's Summary Row if available)
    const ss = state.shopStatsSummary;
    const totalVisitors = ss ? ss.visitors : state.shopStats.reduce((a,b) => a + (b.visitors || 0), 0);
    const avgCR = ss ? ss.cr : state.shopStats.reduce((a,b) => a + (b.cr || 0), 0) / state.shopStats.length;
    const avgAOV = ss ? ss.aov : state.shopStats.reduce((a,b) => a + (b.aov || 0), 0) / state.shopStats.length;
    const totalRepeat = ss ? ss.repeatRate : (state.shopStats.reduce((a,b) => a + (b.repeatRate || 0), 0) / state.shopStats.length);
    
    setVal('d-visitors', totalVisitors.toLocaleString());
    setVal('d-cr', avgCR.toFixed(2) + '%');
    setVal('d-aov', '฿' + Math.round(avgAOV).toLocaleString());
    setVal('d-repeat', totalRepeat > 0 ? totalRepeat.toFixed(2) + '%' : '-');
  } else if (statsSection) {
    statsSection.style.display = 'none';
  }

  // 1. Cost & Fee Breakdown
  let sumComm = 0, sumServ = 0, sumTrans = 0;
  state.results.forEach(r => { if (r.isFirst && !r.isCancelled) { sumComm += Math.abs(r.commFee); sumServ += Math.abs(r.servFee); sumTrans += Math.abs(r.transFee); } });

  const costBreakdownOptions = {
    series: [Math.round(totalCost), Math.round(sumComm), Math.round(sumServ), Math.round(sumTrans)],
    chart: { type: 'pie', height: 270, fontFamily: 'Sarabun, sans-serif' },
    labels: ['ต้นทุนสินค้า', 'ค่าคอมมิชชั่น', 'ค่าบริการ', 'ค่าธุรกรรม'],
    colors: ['#2b8a3e', '#e67700', '#f03e3e', '#1c7ed6'],
    legend: { position: 'bottom', labels: { colors: textColor } },
    stroke: { show: false }
  };
  const costBreakdownEl = document.querySelector("#costBreakdownChart");
  if (costBreakdownEl) {
    state.charts.costBreakdown = new ApexCharts(costBreakdownEl, costBreakdownOptions);
    state.charts.costBreakdown.render();
  }

  // 2. Margin Tier
  const marginTierOptions = {
    series: [tierA, tierB, tierC, lossCount],
    chart: { type: 'donut', height: 270, fontFamily: 'Sarabun, sans-serif' },
    labels: ['Tier A (>30%)', 'Tier B (15-30%)', 'Tier C (<15%)', 'ขาดทุน'],
    colors: ['#2b8a3e', '#74b816', '#fab005', '#e03131'],
    legend: { position: 'bottom', labels: { colors: textColor } },
    stroke: { show: false },
    plotOptions: { pie: { donut: { labels: { show: true, total: { show: true, color: textColor } } } } }
  };
  const marginTierEl = document.querySelector("#marginTierChart");
  if (marginTierEl) {
    state.charts.marginTier = new ApexCharts(marginTierEl, marginTierOptions);
    state.charts.marginTier.render();
  }

  // 3. Time Series
  const timeSeriesOptions = {
    series: [
      { name: 'ยอดขาย (Gross)', data: state.timeSeries.map(d => Math.round(d.statRevenue || 0)) },
      { name: 'รายรับ (Income)', data: state.timeSeries.map(d => Math.round(d.revenue || 0)) },
      { name: 'กำไร (Net)', data: state.timeSeries.map(d => Math.round(d.profit || 0)) }
    ],
    chart: { type: 'area', height: 350, toolbar: { show: true }, zoom: { enabled: true }, animations: { enabled: true }, fontFamily: 'Sarabun, sans-serif' },
    dataLabels: { enabled: false },
    stroke: { curve: 'smooth', width: [3, 3, 4] },
    colors: ['#3b82f6', '#f59e0b', '#10b981'],
    fill: { type: 'gradient', gradient: { shadeIntensity: 1, opacityFrom: 0.45, opacityTo: 0.05 } },
    xaxis: { 
        categories: state.timeSeries.map(d => d.date),
        labels: { style: { colors: textColor } }
    },
    yaxis: { labels: { style: { colors: textColor } } },
    legend: { position: 'top', horizontalAlign: 'left', offsetX: 40, labels: { colors: textColor } }
  };
  const timeSeriesEl = document.querySelector("#chart-timeseries");
  if (timeSeriesEl) {
    state.charts.timeSeries = new ApexCharts(timeSeriesEl, timeSeriesOptions);
    state.charts.timeSeries.render();
  }

  // 4. BCG Matrix
  const bcgOptions = {
    series: [
      { name: 'Star', data: bcgData.filter(d => d.x >= avgRev && d.y >= avgProfit) },
      { name: 'Cash Cow', data: bcgData.filter(d => d.x >= avgRev && d.y < avgProfit) },
      { name: 'Question Mark', data: bcgData.filter(d => d.x < avgRev && d.y >= avgProfit) },
      { name: 'Dog', data: bcgData.filter(d => d.x < avgRev && d.y < avgProfit) }
    ],
    chart: { type: 'bubble', height: 380, fontFamily: 'Sarabun, sans-serif' },
    grid: { borderColor: gridColor },
    xaxis: { labels: { style: { colors: textColor } } },
    yaxis: { labels: { style: { colors: textColor } } },
    colors: ['#fab005', '#1c7ed6', '#be4bdb', '#868e96'],
    legend: { labels: { colors: textColor } }
  };
  const bcgEl = document.querySelector("#bcgMatrixChart");
  if (bcgEl) {
    state.charts.bcg = new ApexCharts(bcgEl, bcgOptions);
    state.charts.bcg.render();
  }

  // 5. Horizontal Bars (Top Profit)
  const sortedByProfit = [...state.summary].sort((a,b) => (b.revenue-b.cost) - (a.revenue-a.cost)).slice(0, 5);
  const topProfitOptions = {
    series: [{ name: 'กำไร', data: sortedByProfit.map(r => Math.round(r.revenue - r.cost)) }],
    chart: { type: 'bar', height: 300, fontFamily: 'Sarabun, sans-serif' },
    plotOptions: { bar: { horizontal: true, borderRadius: 4, barHeight: '60%' } },
    xaxis: { categories: sortedByProfit.map(r => r.sku || r.title.substring(0,18)+'...'), labels: { style: { colors: textColor } } },
    yaxis: { labels: { style: { colors: textColor } } },
    colors: ['#2b8a3e']
  };
  const topProfitEl = document.querySelector("#topProfitChart");
  if (topProfitEl) {
    state.charts.topProfit = new ApexCharts(topProfitEl, topProfitOptions);
    state.charts.topProfit.render();
  }

  // 6. Vertical Bars (Top Qty)
  const sortedByQty = [...state.summary].sort((a,b) => b.qty - a.qty).slice(0, 10);
  const salesQtyOptions = {
    series: [{ name: 'จำนวนชิ้น', data: sortedByQty.map(r => r.qty) }],
    chart: { type: 'bar', height: 300, fontFamily: 'Sarabun, sans-serif' },
    xaxis: { categories: sortedByQty.map(r => r.sku || r.title.substring(0,18)+'...'), labels: { show: false } },
    yaxis: { labels: { style: { colors: textColor } } },
    colors: ['#1c7ed6']
  };
  const salesQtyEl = document.querySelector("#salesQtyChart");
  if (salesQtyEl) {
    state.charts.salesQty = new ApexCharts(salesQtyEl, salesQtyOptions);
    state.charts.salesQty.render();
  }

  renderSparklines();
  renderDailyTable();
}

function renderDailyTable() {
  const tbody = document.getElementById('daily-performance-body');
  if(!tbody) return;
  tbody.innerHTML = '';
  
  if(!state.timeSeries || state.timeSeries.length === 0) return;
  
  const sortedDays = [...state.timeSeries].sort((a,b) => b.date.localeCompare(a.date));
  
  sortedDays.forEach(d => {
    const margin = d.revenue > 0 ? (d.profit / d.revenue * 100).toFixed(1) : '0.0';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="font-weight:600; font-family:monospace;">${d.date}</td>
      <td style="text-align:right">${(d.visitors || 0).toLocaleString()}</td>
      <td style="text-align:right">${(d.statCR || 0).toFixed(2)}%</td>
      <td style="text-align:right">${(d.orderCount || 0).toLocaleString()}</td>
      <td style="text-align:right">฿${Math.round(d.statRevenue || 0).toLocaleString()}</td>
      <td style="text-align:right">฿${Math.round(d.revenue || 0).toLocaleString()}</td>
      <td style="text-align:right; color:var(--green); font-weight:700">฿${Math.round(d.profit || 0).toLocaleString()}</td>
      <td style="text-align:right; font-weight:600">${margin}%</td>
    `;
    tbody.appendChild(tr);
  });
}

function renderSparklines() {
  const isDark = state.theme === 'dark';
  const sparklineData = {
    'spark-total-orders': (state.timeSeries || []).map(d => state.results.filter(r => r.isFirst && !r.isCancelled && r.orderId.includes(d.date)).length || 5), // simplified
    'spark-revenue': (state.timeSeries || []).map(d => Math.round(d.revenue)),
    'spark-orders-success': (state.timeSeries || []).map(d => state.results.filter(r => r.isFirst && !r.isCancelled && r.orderId.includes(d.date)).length),
    'spark-income': (state.timeSeries || []).map(d => Math.round(d.revenue)),
    'spark-cost': (state.timeSeries || []).map(d => Math.round(d.revenue * 0.6)), // estimated curve
    'spark-profit': (state.timeSeries || []).map(d => Math.round(d.profit))
  };

  const colors = {
    'spark-total-orders': '#64748b',
    'spark-revenue': '#f59e0b',
    'spark-orders-success': '#3b82f6',
    'spark-income': '#10b981',
    'spark-cost': '#ef4444',
    'spark-profit': '#10b981'
  };

  Object.keys(sparklineData).forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    el.innerHTML = '';
    const options = {
      series: [{ data: sparklineData[id] }],
      chart: { type: 'area', height: 40, width: 80, sparkline: { enabled: true }, animations: { enabled: false } },
      stroke: { curve: 'smooth', width: 2 },
      fill: { opacity: 0.3 },
      colors: [colors[id] || '#cbd5e1'],
      tooltip: { enabled: false }
    };
    new ApexCharts(el, options).render();
  });
}
