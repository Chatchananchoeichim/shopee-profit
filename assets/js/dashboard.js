function renderDashboard() {
  if(state.charts.marginTier) state.charts.marginTier.destroy();
  if(state.charts.timeSeries) state.charts.timeSeries.destroy();
  if(state.charts.bcg) state.charts.bcg.destroy();
  if(state.charts.topProfit) state.charts.topProfit.destroy();
  if(state.charts.salesQty) state.charts.salesQty.destroy();
  if(state.charts.costBreakdown) state.charts.costBreakdown.destroy();

  let totalRevenue = state.summary.reduce((a,b)=>a+b.revenue, 0);
  let totalCost = state.summary.reduce((a,b)=>a+b.cost, 0);
  let totalNet = totalRevenue - totalCost;
  let overallMargin = totalRevenue > 0 ? (totalNet / totalRevenue) * 100 : 0;

  // Ad Spend Calculation
  let totalAdSpend = state.adsData.reduce((a, b) => a + b.adSpend, 0);
  let netAfterAds = totalNet - totalAdSpend;
  let marginAfterAds = totalRevenue > 0 ? (netAfterAds / totalRevenue) * 100 : 0;

  // Margin Tiers & BCG Data
  let tierA = 0, tierB = 0, tierC = 0, lossCount = 0;
  let skuProfitCount = 0;
  let starCount = 0;
  let sumRev = 0, sumProfit = 0;
  const bcgData = [];

  state.summary.forEach(s => {
    const profit = s.revenue - s.cost;
    const marginPct = s.revenue > 0 ? (profit / s.revenue) * 100 : 0;
    
    if(marginPct >= 30) tierA++;
    else if(marginPct >= 15) tierB++;
    else if(profit > 0) tierC++;
    
    if(profit > 0) skuProfitCount++;
    if(profit < 0) lossCount++;
    
    if(s.qty > 0) {
      sumRev += s.revenue;
      sumProfit += profit;
      bcgData.push({
        x: s.revenue, y: profit, r: Math.max(8, Math.min(s.qty * 1.5, 30)),
        label: s.sku || s.title.substring(0, 15)
      });
    }
  });

  const avgRev =  bcgData.length ? sumRev / bcgData.length : 0;
  const avgProfit = bcgData.length ? sumProfit / bcgData.length : 0;

  bcgData.forEach(item => { if(item.x >= avgRev && item.y >= avgProfit) starCount++; });

  document.getElementById('d-margin').innerText = overallMargin.toFixed(1) + '%';
  document.getElementById('d-star-count').innerText = starCount + ' SKUs';
  document.getElementById('d-profit-skus').innerText = skuProfitCount + ' SKUs';
  document.getElementById('d-loss-skus').innerText = lossCount + ' SKUs';

  const adsDash = document.getElementById('ads-summary-dashboard');
  if (adsDash) {
    if (totalAdSpend > 0) {
      document.getElementById('d-total-ads').innerText = '฿' + Math.round(totalAdSpend).toLocaleString();
      document.getElementById('d-net-after-ads').innerText = '฿' + Math.round(netAfterAds).toLocaleString();
      adsDash.style.display = 'grid';
    } else {
      adsDash.style.display = 'none';
    }
  }

  const shopDash = document.getElementById('shop-stats-dashboard');
  if (shopDash) {
    if (state.shopStats) {
      document.getElementById('d-visitors').innerText = parseInt(state.shopStats.visitors).toLocaleString();
      document.getElementById('d-cr').innerText       = state.shopStats.cr;
      document.getElementById('d-aov').innerText      = '฿' + Math.round(parseFloat(state.shopStats.aov)).toLocaleString();
      document.getElementById('d-repeat').innerText   = state.shopStats.repeat;
      shopDash.style.display = 'grid';
    } else {
      shopDash.style.display = 'none';
    }
  }

  // 1. Cost & Fee Breakdown Chart
  let sumComm = 0, sumCommAMS = 0, sumServ = 0, sumPlat = 0, sumTrans = 0, sumShip = 0;
  state.results.forEach(r => {
    if (r.isFirst && !r.isCancelled) {
       sumComm    += Math.abs(r.commFee    || 0);
       sumCommAMS += Math.abs(r.commAMSFee || 0);
       sumServ    += Math.abs(r.servFee    || 0);
       sumPlat    += Math.abs(r.platFee    || 0);
       sumTrans   += Math.abs(r.transFee   || 0);
       sumShip    += Math.abs(r.shipDeduct || 0);
    }
  });

  const ctxCost = document.getElementById('costBreakdownChart').getContext('2d');
  state.charts.costBreakdown = new Chart(ctxCost, {
    type: 'pie',
    data: {
      labels: [
        'ต้นทุนสินค้า',
        'ค่าคอมมิชชั่น',
        'ค่าบริการ',
        'ค่าธุรกรรม (ชำระเงิน)'
      ],
      datasets: [{
        data: [totalCost, sumComm, sumServ, sumTrans],
        backgroundColor: ['#2b8a3e','#e67700','#f03e3e','#1c7ed6'],
        borderWidth: 2, borderColor: '#fff'
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { 
        legend: { position: 'right', labels: { font: { size: 10 }, boxWidth: 14 } },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.label}: ฿${Math.round(ctx.raw).toLocaleString()}`
          }
        }
      }
    }
  });

  // 2. Margin Tier Chart
  const ctxTier = document.getElementById('marginTierChart').getContext('2d');
  state.charts.marginTier = new Chart(ctxTier, {
    type: 'doughnut',
    data: {
      labels: ['Tier A (>30%)', 'Tier B (15-30%)', 'Tier C (<15%)', 'ขาดทุน'],
      datasets: [{
        data: [tierA, tierB, tierC, lossCount],
        backgroundColor: ['#2b8a3e', '#74b816', '#fab005', '#e03131'],
        borderWidth: 2, borderColor: '#fff'
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'right' } }
    }
  });

  // 3. Seasonality Chart (Time Series)
  const ctxTime = document.getElementById('chart-timeseries').getContext('2d');
  state.charts.timeSeries = new Chart(ctxTime, {
    type: 'line',
    data: {
      labels: (state.timeSeries || []).map(d => d.date),
      datasets: [
        {
          type: 'bar', label: 'ยอดขาย (Revenue)',
          data: (state.timeSeries || []).map(d => d.revenue),
          backgroundColor: '#a5d8ff', borderRadius: 4
        },
        {
          type: 'line', label: 'กำไร (Profit)',
          data: state.timeSeries.map(d=>d.profit),
          borderColor: '#2b8a3e',
          backgroundColor: '#2b8a3e',
          yAxisID: 'y1',
          fill: false,
          tension: 0.3
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { grid: { display: false } },
        y: {
          beginAtZero: true,
          position: 'left',
          title: { display: true, text: 'ยอดขาย (THB)' }
        },
        y1: {
          beginAtZero: true,
          position: 'right',
          grid: { drawOnChartArea: false },
          title: { display: true, text: 'กำไร (THB)' }
        }
      },
      plugins: {
        legend: { position: 'top' },
        tooltip: {
          mode: 'index',
          intersect: false
        }
      }
    }
  });

  // Populate Daily Table
  const dailyBody = document.getElementById('daily-performance-body');
  if (dailyBody) {
    dailyBody.innerHTML = state.timeSeries.map(d => {
      const margin = d.revenue > 0 ? (d.profit / d.revenue * 100).toFixed(1) : 0;
      return `
        <tr>
          <td style="font-family:monospace">${d.date}</td>
          <td class="text-right">฿${Math.round(d.revenue).toLocaleString()}</td>
          <td class="text-right ${d.profit >= 0 ? 'green' : 'red'}">฿${Math.round(d.profit).toLocaleString()}</td>
          <td class="text-right">${margin}%</td>
        </tr>
      `;
    }).reverse().join(''); // Show latest first
  }

  // 3. BCG Matrix
  const scatterColors = bcgData.map(d => {
    if(d.x >= avgRev && d.y >= avgProfit) return 'rgba(250, 176, 5, 0.8)'; // Star
    if(d.x >= avgRev && d.y < avgProfit) return 'rgba(28, 126, 214, 0.8)'; // Cash Cow
    if(d.x < avgRev && d.y >= avgProfit) return 'rgba(190, 75, 219, 0.8)'; // Question Mark
    return 'rgba(134, 142, 150, 0.7)'; // Dog
  });

  const ctxBCG = document.getElementById('bcgMatrixChart').getContext('2d');
  state.charts.bcg = new Chart(ctxBCG, {
    type: 'bubble',
    data: {
      datasets: [{
        label: 'สินค้า', data: bcgData,
        backgroundColor: scatterColors,
        borderColor: '#fff', borderWidth: 1
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: function(ctx) {
              const d = ctx.raw;
              return `${d.label} | ขาย: ฿${Math.round(d.x).toLocaleString()} | กำไร: ฿${Math.round(d.y).toLocaleString()}`;
            }
          }
        }
      },
      scales: {
        x: { title: { display: true, text: 'ยอดขายสุทธิ' }, grid: { color: (ctx) => ctx.tick.value === 0 ? '#333' : '#eee' } },
        y: { title: { display: true, text: 'กำไร (Profit)' }, grid: { color: (ctx) => ctx.tick.value === 0 ? '#333' : '#eee' } }
      }
    }
  });

  // 4. Top Profit
  const sortedByProfit = [...state.summary].sort((a,b) => (b.revenue-b.cost) - (a.revenue-a.cost)).slice(0, 5);
  const ctxTopProfit = document.getElementById('topProfitChart').getContext('2d');
  state.charts.topProfit = new Chart(ctxTopProfit, {
    type: 'bar',
    data: {
      labels: sortedByProfit.map(r => r.sku || r.title.substring(0,18)+'..'),
      datasets: [{
        data: sortedByProfit.map(r => r.revenue - r.cost),
        backgroundColor: '#2b8a3e', borderRadius: 4
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { y: { beginAtZero: true } }
    }
  });

  // 5. Top Qty
  const sortedByQty = [...state.summary].sort((a,b) => b.qty - a.qty).slice(0, 10);
  const ctxSalesQty = document.getElementById('salesQtyChart').getContext('2d');
  state.charts.salesQty = new Chart(ctxSalesQty, {
    type: 'bar',
    data: {
      labels: sortedByQty.map(r => r.sku || r.title.substring(0,18)+'..'),
      datasets: [{
        data: sortedByQty.map(r => r.qty),
        backgroundColor: '#1c7ed6', borderRadius: 4
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { y: { beginAtZero: true } }
    }
  });
}
