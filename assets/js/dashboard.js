function renderDashboard() {
  // Clear existing ApexCharts instances if they exist
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
        x: Math.round(s.revenue), 
        y: Math.round(profit), 
        z: Math.max(8, Math.min(s.qty * 1.5, 30)),
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

  // 1. Cost & Fee Breakdown Chart (ApexCharts Pie)
  let sumComm = 0, sumServ = 0, sumTrans = 0;
  state.results.forEach(r => {
    if (r.isFirst && !r.isCancelled) {
       sumComm    += Math.abs(r.commFee    || 0);
       sumServ    += Math.abs(r.servFee    || 0);
       sumTrans   += Math.abs(r.transFee   || 0);
    }
  });

  const costBreakdownOptions = {
    series: [Math.round(totalCost), Math.round(sumComm), Math.round(sumServ), Math.round(sumTrans)],
    chart: { type: 'pie', height: 270, fontFamily: 'Sarabun, sans-serif' },
    labels: ['ต้นทุนสินค้า', 'ค่าคอมมิชชั่น', 'ค่าบริการ', 'ค่าธุรกรรม'],
    colors: ['#2b8a3e', '#e67700', '#f03e3e', '#1c7ed6'],
    legend: { position: 'bottom' },
    dataLabels: { enabled: true, formatter: val => val.toFixed(1) + "%" },
    tooltip: { y: { formatter: val => '฿' + val.toLocaleString() } }
  };
  state.charts.costBreakdown = new ApexCharts(document.querySelector("#costBreakdownChart"), costBreakdownOptions);
  state.charts.costBreakdown.render();

  // 2. Margin Tier Chart (ApexCharts Donut)
  const marginTierOptions = {
    series: [tierA, tierB, tierC, lossCount],
    chart: { type: 'donut', height: 270, fontFamily: 'Sarabun, sans-serif' },
    labels: ['Tier A (>30%)', 'Tier B (15-30%)', 'Tier C (<15%)', 'ขาดทุน'],
    colors: ['#2b8a3e', '#74b816', '#fab005', '#e03131'],
    legend: { position: 'bottom' },
    plotOptions: { pie: { donut: { labels: { show: true, total: { show: true, label: 'สินค้าทั้งหมด' } } } } }
  };
  state.charts.marginTier = new ApexCharts(document.querySelector("#marginTierChart"), marginTierOptions);
  state.charts.marginTier.render();

  // 3. Seasonality Chart (ApexCharts Mixed Line/Area)
  const timeSeriesOptions = {
    series: [{
      name: 'ยอดขาย (Revenue)',
      type: 'area',
      data: (state.timeSeries || []).map(d => Math.round(d.revenue))
    }, {
      name: 'กำไร (Profit)',
      type: 'line',
      data: (state.timeSeries || []).map(d => Math.round(d.profit))
    }],
    chart: { height: 350, type: 'line', stacked: false, fontFamily: 'Sarabun, sans-serif', toolbar: { show: false } },
    stroke: { width: [0, 3], curve: 'smooth' },
    fill: { opacity: [0.35, 1], gradient: { inverseColors: false, shade: 'light', type: "vertical", opacityFrom: 0.85, opacityTo: 0.55, stops: [0, 100, 100, 100] } },
    labels: (state.timeSeries || []).map(d => d.date),
    markers: { size: 0 },
    yaxis: [
      { title: { text: 'ยอดขาย (THB)' }, labels: { formatter: val => val.toLocaleString() } },
      { opposite: true, title: { text: 'กำไร (THB)' }, labels: { formatter: val => val.toLocaleString() } }
    ],
    colors: ['#3b82f6', '#10b981'],
    tooltip: { shared: true, intersect: false, y: { formatter: val => '฿' + val.toLocaleString() } }
  };
  state.charts.timeSeries = new ApexCharts(document.querySelector("#chart-timeseries"), timeSeriesOptions);
  state.charts.timeSeries.render();

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
    }).reverse().join('');
  }

  // 4. BCG Matrix (ApexCharts Bubble)
  const bcgOptions = {
    series: [
      { name: 'Star', data: bcgData.filter(d => d.x >= avgRev && d.y >= avgProfit) },
      { name: 'Cash Cow', data: bcgData.filter(d => d.x >= avgRev && d.y < avgProfit) },
      { name: 'Question Mark', data: bcgData.filter(d => d.x < avgRev && d.y >= avgProfit) },
      { name: 'Dog', data: bcgData.filter(d => d.x < avgRev && d.y < avgProfit) }
    ],
    chart: { type: 'bubble', height: 380, fontFamily: 'Sarabun, sans-serif' },
    dataLabels: { enabled: false },
    fill: { opacity: 0.7 },
    xaxis: { title: { text: 'ยอดขายสุทธิ' }, labels: { formatter: val => val.toLocaleString() } },
    yaxis: { title: { text: 'กำไร (Profit)' }, labels: { formatter: val => val.toLocaleString() } },
    colors: ['#fab005', '#1c7ed6', '#be4bdb', '#868e96'],
    tooltip: {
      custom: function({series, seriesIndex, dataPointIndex, w}) {
        const item = w.config.series[seriesIndex].data[dataPointIndex];
        return '<div class="apexcharts-tooltip-custom" style="padding:10px; font-size:12px; background:#fff; border:1px solid #ccc;">' +
               '<b>' + item.label + '</b><br/>' +
               'ยอดขาย: ฿' + item.x.toLocaleString() + '<br/>' +
               'กำไร: ฿' + item.y.toLocaleString() +
               '</div>';
      }
    }
  };
  state.charts.bcg = new ApexCharts(document.querySelector("#bcgMatrixChart"), bcgOptions);
  state.charts.bcg.render();

  // 5. Top Profit (ApexCharts Horizontal Bar)
  const sortedByProfit = [...state.summary].sort((a,b) => (b.revenue-b.cost) - (a.revenue-a.cost)).slice(0, 5);
  const topProfitOptions = {
    series: [{ name: 'กำไร', data: sortedByProfit.map(r => Math.round(r.revenue - r.cost)) }],
    chart: { type: 'bar', height: 300, fontFamily: 'Sarabun, sans-serif' },
    plotOptions: { bar: { horizontal: true, borderRadius: 4, barHeight: '60%' } },
    colors: ['#2b8a3e'],
    xaxis: { categories: sortedByProfit.map(r => r.sku || r.title.substring(0,18)+'...'), labels: { formatter: val => val.toLocaleString() } },
    tooltip: { y: { formatter: val => '฿' + val.toLocaleString() } }
  };
  state.charts.topProfit = new ApexCharts(document.querySelector("#topProfitChart"), topProfitOptions);
  state.charts.topProfit.render();

  // 6. Top Qty (ApexCharts Bar)
  const sortedByQty = [...state.summary].sort((a,b) => b.qty - a.qty).slice(0, 10);
  const salesQtyOptions = {
    series: [{ name: 'จำนวนชิ้น', data: sortedByQty.map(r => r.qty) }],
    chart: { type: 'bar', height: 300, fontFamily: 'Sarabun, sans-serif' },
    plotOptions: { bar: { borderRadius: 4, columnWidth: '50%' } },
    colors: ['#1c7ed6'],
    xaxis: { categories: sortedByQty.map(r => r.sku || r.title.substring(0,18)+'...'), labels: { show: false } },
    tooltip: { y: { formatter: val => val + ' ชิ้น' } }
  };
  state.charts.salesQty = new ApexCharts(document.querySelector("#salesQtyChart"), salesQtyOptions);
  state.charts.salesQty.render();

  // Initialize Tippy Tooltips
  tippy('[title]', {
    animation: 'shift-away',
    theme: 'light-border',
  });
}
