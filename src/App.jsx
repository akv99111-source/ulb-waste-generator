import React, { useState } from 'react';
import { Building2, Download, CreditCard, BarChart3, RefreshCw, Calculator } from 'lucide-react';

export default function App() {
  const [ulbName, setUlbName] = useState('Nagar Palika Parishad');
  const [population, setPopulation] = useState(150000);
  const [perCapita, setPerCapita] = useState('450');
  const [fixedTons, setFixedTons] = useState('');
  const [year, setYear] = useState(2026);
  const [month, setMonth] = useState(10);
  
  const [displayUnit, setDisplayUnit] = useState('Tons');
  const [generatedData, setGeneratedData] = useState(null);
  const [calcBasis, setCalcBasis] = useState('');
  const [isPaid, setIsPaid] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);

  const loadRazorpayScript = () => {
    return new Promise((resolve) => {
      const script = document.createElement('script');
      script.src = 'https://checkout.razorpay.com/v1/checkout.js';
      script.onload = () => resolve(true);
      script.onerror = () => resolve(false);
      document.body.appendChild(script);
    });
  };

  const handleGenerate = (e) => {
    e.preventDefault();
    setIsPaid(false);

    const daysInMonth = [1, 3, 5, 7, 8, 10, 12].includes(Number(month)) ? 31 : (Number(month) === 2 ? 28 : 30);
    
    let targetDailyTons = 0;
    let modeText = '';

    if (fixedTons && parseFloat(fixedTons) > 0) {
      targetDailyTons = parseFloat(fixedTons);
      modeText = `Fixed Output Override (${targetDailyTons.toFixed(2)} Tons/Day)`;
    } else {
      const perCapitaKg = parseFloat(perCapita) / 1000.0;
      targetDailyTons = (population * perCapitaKg) / 1000.0;
      modeText = `Population (${population.toLocaleString()}) × ${perCapita}g/capita/day = ${targetDailyTons.toFixed(2)} Tons/Day`;
    }

    setCalcBasis(modeText);

    const baseFractions = [0.50, 0.10, 0.04, 0.08, 0.02, 0.26];
    const logs = [];

    for (let day = 1; day <= daysInMonth; day++) {
      const currentDate = new Date(year, month - 1, day);
      const dayName = currentDate.toLocaleDateString('en-US', { weekday: 'short' });
      const isWeekend = currentDate.getDay() === 0 || currentDate.getDay() === 6;

      let noise = 0.95 + Math.random() * 0.10;
      if (isWeekend) noise *= 1.05;
      const dailyTotalTons = targetDailyTons * noise;

      let rawRatios = baseFractions.map(r => r * (0.85 + Math.random() * 0.30));
      const sumRatios = rawRatios.reduce((a, b) => a + b, 0);
      const normRatios = rawRatios.map(r => r / sumRatios);

      const wet = parseFloat((dailyTotalTons * normRatios[0]).toFixed(2));
      const plastic = parseFloat((dailyTotalTons * normRatios[1]).toFixed(2));
      const cloth = parseFloat((dailyTotalTons * normRatios[2]).toFixed(2));
      const cd = parseFloat((dailyTotalTons * normRatios[3]).toFixed(2));
      const haz = parseFloat((dailyTotalTons * normRatios[4]).toFixed(2));
      const inerts = parseFloat((dailyTotalTons * normRatios[5]).toFixed(2));
      const total = parseFloat((wet + plastic + cloth + cd + haz + inerts).toFixed(2));

      logs.push({
        date: currentDate.toISOString().split('T')[0],
        dayName, wet, plastic, cloth, cd, haz, inerts, total
      });
    }

    setGeneratedData(logs);
  };

  const handlePayment = async () => {
    setIsProcessing(true);
    const sdkLoaded = await loadRazorpayScript();

    if (!sdkLoaded) {
      alert('Razorpay SDK failed to load.');
      setIsProcessing(false);
      return;
    }

    try {
      const orderRes = await fetch('/api/create-order', { method: 'POST' });
      const orderData = await orderRes.json();

      const options = {
        key: import.meta.env.VITE_RAZORPAY_KEY_ID,
        amount: orderData.amount,
        currency: orderData.currency,
        name: 'ULB Logbook Generator',
        description: `Waste Dataset Export: ${ulbName}`,
        order_id: orderData.id,
        handler: function () {
          setIsPaid(true);
          setIsProcessing(false);
          downloadCSV();
        },
        theme: { color: '#059669' },
      };

      const paymentObject = new window.Razorpay(options);
      paymentObject.open();
    } catch (err) {
      alert('Payment Failed: ' + err.message);
      setIsProcessing(false);
    }
  };

  const formatVal = (tonsVal) => {
    return displayUnit === 'kg' ? Math.round(tonsVal * 1000) : tonsVal.toFixed(2);
  };

  const downloadCSV = () => {
    if (!generatedData) return;
    const unitLabel = displayUnit === 'kg' ? 'kg' : 'Tons';
    const headers = ["Date", "Day", `Wet (${unitLabel})`, `Plastic (${unitLabel})`, `Textile (${unitLabel})`, `C&D (${unitLabel})`, `Hazardous (${unitLabel})`, `Inerts (${unitLabel})`, `Total (${unitLabel})` ];

    const rows = generatedData.map(r => [
      r.date, r.dayName, 
      formatVal(r.wet), formatVal(r.plastic), formatVal(r.cloth), 
      formatVal(r.cd), formatVal(r.haz), formatVal(r.inerts), formatVal(r.total)
    ]);

    const csvContent = "data:text/csv;charset=utf-8," + [headers.join(","), ...rows.map(e => e.join(","))].join("\n");
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `${ulbName.replace(/\s+/g, '_')}_Logbook_${year}_${month}_${displayUnit}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div style={{ padding: '20px', fontFamily: 'sans-serif', maxWidth: '1000px', margin: '0 auto' }}>
      <h2><Building2 /> ULB Waste Logbook Generator</h2>
      
      <form onSubmit={handleGenerate} style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '10px', marginBottom: '20px' }}>
        <div><label>ULB Name</label><input style={{ width: '100%' }} type="text" required value={ulbName} onChange={(e) => setUlbName(e.target.value)} /></div>
        <div><label>Population</label><input style={{ width: '100%' }} type="number" value={population} onChange={(e) => setPopulation(Number(e.target.value))} /></div>
        <div>
          <label>Per Capita Rate</label>
          <select style={{ width: '100%' }} value={perCapita} onChange={(e) => setPerCapita(e.target.value)}>
            <option value="300">300 g/capita/day</option>
            <option value="350">350 g/capita/day</option>
            <option value="400">400 g/capita/day</option>
            <option value="450">450 g/capita/day</option>
            <option value="500">500 g/capita/day</option>
          </select>
        </div>
        <div><label>Fixed Daily Tons Override</label><input style={{ width: '100%' }} type="number" step="0.01" value={fixedTons} onChange={(e) => setFixedTons(e.target.value)} /></div>
        <div><label>Year</label><input style={{ width: '100%' }} type="number" value={year} onChange={(e) => setYear(Number(e.target.value))} /></div>
        <div><label>Month</label><input style={{ width: '100%' }} type="number" min="1" max="12" value={month} onChange={(e) => setMonth(Number(e.target.value))} /></div>
        <button type="submit" style={{ gridColumn: 'span 3', padding: '10px', background: '#059669', color: '#fff', border: 'none', borderRadius: '5px', cursor: 'pointer' }}>Generate Dataset</button>
      </form>

      {generatedData && (
        <div>
          <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '10px' }}>
            <div><strong>{ulbName}</strong> — <small>{calcBasis}</small></div>
            <div>
              <button onClick={() => setDisplayUnit(displayUnit === 'Tons' ? 'kg' : 'Tons')} style={{ marginRight: '10px' }}>
                Unit: {displayUnit}
              </button>
              {isPaid ? (
                <button onClick={downloadCSV}><Download /> Export CSV</button>
              ) : (
                <button onClick={handlePayment} disabled={isProcessing}>
                  {isProcessing ? 'Processing...' : 'Pay ₹299 to Download'}
                </button>
              )}
            </div>
          </div>

          <table border="1" cellPadding="6" style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
            <thead>
              <tr>
                <th>Date</th><th>Day</th><th>Wet ({displayUnit})</th><th>Plastic ({displayUnit})</th>
                <th>Textile ({displayUnit})</th><th>C&D ({displayUnit})</th><th>Hazardous ({displayUnit})</th>
                <th>Inerts ({displayUnit})</th><th>Total ({displayUnit})</th>
              </tr>
            </thead>
            <tbody>
              {generatedData.map((row, idx) => (
                <tr key={idx}>
                  <td>{row.date}</td><td>{row.dayName}</td>
                  <td>{formatVal(row.wet)}</td><td>{formatVal(row.plastic)}</td>
                  <td>{formatVal(row.cloth)}</td><td>{formatVal(row.cd)}</td>
                  <td>{formatVal(row.haz)}</td><td>{formatVal(row.inerts)}</td>
                  <td><strong>{formatVal(row.total)}</strong></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}