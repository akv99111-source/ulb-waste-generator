import React, { useState } from 'react';
import { Building2, Download, Lock, Globe, Check } from 'lucide-react';
import * as XLSX from 'xlsx';

const MONTH_NAMES = [
  { id: 1, shortEn: 'Jan', shortHi: 'जनवरी', fullEn: 'January' },
  { id: 2, shortEn: 'Feb', shortHi: 'फरवरी', fullEn: 'February' },
  { id: 3, shortEn: 'Mar', shortHi: 'मार्च', fullEn: 'March' },
  { id: 4, shortEn: 'Apr', shortHi: 'अप्रैल', fullEn: 'April' },
  { id: 5, shortEn: 'May', shortHi: 'मई', fullEn: 'May' },
  { id: 6, shortEn: 'Jun', shortHi: 'जून', fullEn: 'June' },
  { id: 7, shortEn: 'Jul', shortHi: 'जुलाई', fullEn: 'July' },
  { id: 8, shortEn: 'Aug', shortHi: 'अगस्त', fullEn: 'August' },
  { id: 9, shortEn: 'Sep', shortHi: 'सितंबर', fullEn: 'September' },
  { id: 10, shortEn: 'Oct', shortHi: 'अक्टूबर', fullEn: 'October' },
  { id: 11, shortEn: 'Nov', shortHi: 'नवंबर', fullEn: 'November' },
  { id: 12, shortEn: 'Dec', shortHi: 'दिसंबर', fullEn: 'December' }
];

export default function App() {
  const [lang, setLang] = useState('hi');
  const [facilityType, setFacilityType] = useState('ULB');
  const [name, setName] = useState('Nagar Palika Parishad');
  
  const [population, setPopulation] = useState(150000);
  const [perCapita, setPerCapita] = useState('450');
  const [ulbFixedTons, setUlbFixedTons] = useState('');
  
  const [mrfDailyDryTons, setMrfDailyDryTons] = useState(15);
  const [mrfMaxCapacityTons, setMrfMaxCapacityTons] = useState(25);

  const [startYear, setStartYear] = useState(2026);
  const [selectedMonths, setSelectedMonths] = useState([10]);

  const [displayUnit, setDisplayUnit] = useState('Tons');
  
  // Data stored per month: { [monthId]: [rows] }
  const [generatedMonthlyData, setGeneratedMonthlyData] = useState(null);
  const [activeTabMonth, setActiveTabMonth] = useState(null);

  const [isPaid, setIsPaid] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);

  const toggleMonth = (monthId) => {
    if (selectedMonths.includes(monthId)) {
      if (selectedMonths.length === 1) return;
      setSelectedMonths(selectedMonths.filter(m => m !== monthId));
    } else {
      if (selectedMonths.length >= 3) {
        alert(lang === 'hi' ? 'आप एक बार में अधिकतम 3 माह चुन सकते हैं।' : 'You can select a maximum of 3 months at a time.');
        return;
      }
      setSelectedMonths([...selectedMonths, monthId].sort((a, b) => a - b));
    }
  };

  const getPrice = () => {
    const count = selectedMonths.length;
    if (count === 1) return 50;
    if (count === 2) return 100;
    if (count === 3) return 125;
    return 50;
  };

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

    let monthlyDataMap = {};

    selectedMonths.forEach((calcMonth) => {
      let daysInMonth = [1, 3, 5, 7, 8, 10, 12].includes(calcMonth) ? 31 : (calcMonth === 2 ? 28 : 30);
      let targetDailyTons = 0;

      if (facilityType === 'ULB') {
        if (ulbFixedTons && parseFloat(ulbFixedTons) > 0) {
          targetDailyTons = parseFloat(ulbFixedTons);
        } else {
          targetDailyTons = (population * (parseFloat(perCapita) / 1000.0)) / 1000.0;
        }
      } else {
        targetDailyTons = parseFloat(mrfDailyDryTons);
      }

      let monthLogs = [];
      for (let day = 1; day <= daysInMonth; day++) {
        const currentDate = new Date(startYear, calcMonth - 1, day);
        const dayName = currentDate.toLocaleDateString('en-US', { weekday: 'short' });
        const isWeekend = currentDate.getDay() === 0 || currentDate.getDay() === 6;

        let noise = 0.95 + Math.random() * 0.10;
        if (isWeekend) noise *= 1.05;
        const dailyTotal = targetDailyTons * noise;

        if (facilityType === 'ULB') {
          const baseFractions = [0.50, 0.10, 0.04, 0.08, 0.02, 0.26];
          let rawRatios = baseFractions.map(r => r * (0.85 + Math.random() * 0.30));
          const sumRatios = rawRatios.reduce((a, b) => a + b, 0);
          const norm = rawRatios.map(r => r / sumRatios);

          monthLogs.push({
            date: currentDate.toISOString().split('T')[0],
            dayName,
            c1: parseFloat((dailyTotal * norm[0]).toFixed(2)),
            c2: parseFloat((dailyTotal * norm[1]).toFixed(2)),
            c3: parseFloat((dailyTotal * norm[2]).toFixed(2)),
            c4: parseFloat((dailyTotal * norm[3]).toFixed(2)),
            c5: parseFloat((dailyTotal * norm[4]).toFixed(2)),
            c6: parseFloat((dailyTotal * norm[5]).toFixed(2)),
            total: parseFloat(dailyTotal.toFixed(2))
          });
        } else {
          const baseFractions = [0.20, 0.15, 0.25, 0.20, 0.10, 0.10];
          let rawRatios = baseFractions.map(r => r * (0.85 + Math.random() * 0.30));
          const sumRatios = rawRatios.reduce((a, b) => a + b, 0);
          const norm = rawRatios.map(r => r / sumRatios);

          monthLogs.push({
            date: currentDate.toISOString().split('T')[0],
            dayName,
            c1: parseFloat((dailyTotal * norm[0]).toFixed(2)),
            c2: parseFloat((dailyTotal * norm[1]).toFixed(2)),
            c3: parseFloat((dailyTotal * norm[2]).toFixed(2)),
            c4: parseFloat((dailyTotal * norm[3]).toFixed(2)),
            c5: parseFloat((dailyTotal * norm[4]).toFixed(2)),
            c6: parseFloat((dailyTotal * norm[5]).toFixed(2)),
            total: parseFloat(dailyTotal.toFixed(2))
          });
        }
      }
      monthlyDataMap[calcMonth] = monthLogs;
    });

    setGeneratedMonthlyData(monthlyDataMap);
    setActiveTabMonth(selectedMonths[0]);
  };

  const handlePayment = async () => {
    setIsProcessing(true);
    const sdkLoaded = await loadRazorpayScript();

    if (!sdkLoaded) {
      alert(lang === 'hi' ? 'रेजरपे लोड नहीं हो सका।' : 'Razorpay SDK failed to load.');
      setIsProcessing(false);
      return;
    }

    try {
      const orderRes = await fetch('/api/create-order', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ amount: getPrice() * 100 })
      });
      const orderData = await orderRes.json();

      const options = {
        key: import.meta.env.VITE_RAZORPAY_KEY_ID,
        amount: orderData.amount,
        currency: orderData.currency,
        name: 'ULB / MRF Logbook Engine',
        description: `${facilityType} Multi-Tab Excel (${selectedMonths.length} Month/s) - ${name}`,
        order_id: orderData.id,
        handler: function () {
          setIsPaid(true);
          setIsProcessing(false);
          downloadMultiSheetExcel();
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

  const formatVal = (val) => displayUnit === 'kg' ? Math.round(val * 1000) : val.toFixed(2);

  const downloadMultiSheetExcel = () => {
    if (!generatedMonthlyData) return;
    const unitLabel = displayUnit === 'kg' ? 'kg' : 'Tons';

    let headers = [];
    if (facilityType === 'ULB') {
      headers = ["Date", "Day", `Wet (${unitLabel})`, `Plastic (${unitLabel})`, `Textile (${unitLabel})`, `C&D (${unitLabel})`, `Hazardous (${unitLabel})`, `Inerts (${unitLabel})`, `Total (${unitLabel})` ];
    } else {
      headers = ["Date", "Day", `PET Plastics (${unitLabel})`, `HDPE/Hard Plastic (${unitLabel})`, `Cardboard/Paper (${unitLabel})`, `RDF/Combustibles (${unitLabel})`, `Glass/Metal (${unitLabel})`, `Rejects (${unitLabel})`, `Total Dry Waste (${unitLabel})` ];
    }

    const workbook = XLSX.utils.book_new();

    selectedMonths.forEach((mId) => {
      const monthName = MONTH_NAMES.find(m => m.id === mId)?.fullEn || `Month_${mId}`;
      const monthRows = generatedMonthlyData[mId] || [];

      const sheetData = [
        headers,
        ...monthRows.map(r => [
          r.date, r.dayName, formatVal(r.c1), formatVal(r.c2), formatVal(r.c3), formatVal(r.c4), formatVal(r.c5), formatVal(r.c6), formatVal(r.total)
        ])
      ];

      const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
      XLSX.utils.book_append_sheet(workbook, worksheet, monthName);
    });

    XLSX.writeFile(workbook, `${name.replace(/\s+/g, '_')}_Logbook_${selectedMonths.length}Months_${displayUnit}.xlsx`);
  };

  const activeRows = (generatedMonthlyData && activeTabMonth) ? generatedMonthlyData[activeTabMonth] : [];
  const visibleRows = isPaid ? activeRows : activeRows.slice(0, 5);

  return (
    <div style={{ padding: '20px', fontFamily: 'sans-serif', maxWidth: '1050px', margin: '0 auto' }}>
      
      {/* Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
        <h2><Building2 /> {lang === 'hi' ? 'यूएलबी एवं एमआरएफ अपशिष्ट लोगबुक जनरेटर' : 'ULB & MRF Waste Logbook Generator'}</h2>
        <button onClick={() => setLang(lang === 'hi' ? 'en' : 'hi')} style={{ padding: '6px 12px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '5px' }}>
          <Globe size={16} /> {lang === 'hi' ? 'English' : 'हिंदी'}
        </button>
      </div>

      {/* Form */}
      <form onSubmit={handleGenerate} style={{ background: '#f8fafc', padding: '20px', borderRadius: '8px', border: '1px solid #e2e8f0', marginBottom: '20px' }}>
        <div style={{ marginBottom: '15px' }}>
          <label style={{ fontWeight: 'bold', marginRight: '15px' }}>{lang === 'hi' ? 'सुविधा का प्रकार:' : 'Facility Type:'}</label>
          <label style={{ marginRight: '15px', cursor: 'pointer' }}>
            <input type="radio" name="facility" value="ULB" checked={facilityType === 'ULB'} onChange={() => setFacilityType('ULB')} /> {lang === 'hi' ? 'नगर निकाय (ULB)' : 'Urban Local Body (ULB)'}
          </label>
          <label style={{ cursor: 'pointer' }}>
            <input type="radio" name="facility" value="MRF" checked={facilityType === 'MRF'} onChange={() => setFacilityType('MRF')} /> {lang === 'hi' ? 'एमआरएफ केंद्र (MRF Centre)' : 'MRF Centre'}
          </label>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '15px', marginBottom: '15px' }}>
          <div>
            <label>{lang === 'hi' ? 'निकाय / एमआरएफ का नाम' : 'ULB / MRF Name'}</label>
            <input style={{ width: '100%', padding: '6px' }} type="text" required value={name} onChange={(e) => setName(e.target.value)} />
          </div>

          {facilityType === 'ULB' ? (
            <>
              <div>
                <label>{lang === 'hi' ? 'जनसंख्या (Population)' : 'Population'}</label>
                <input style={{ width: '100%', padding: '6px', background: ulbFixedTons ? '#e2e8f0' : '#fff' }} type="number" disabled={!!ulbFixedTons} value={population} onChange={(e) => setPopulation(Number(e.target.value))} />
              </div>

              <div>
                <label>{lang === 'hi' ? 'प्रतिव्यक्ति अपशिष्ट (g/दिन)' : 'Per Capita Rate'}</label>
                <select style={{ width: '100%', padding: '6px', background: ulbFixedTons ? '#e2e8f0' : '#fff' }} disabled={!!ulbFixedTons} value={perCapita} onChange={(e) => setPerCapita(e.target.value)}>
                  <option value="300">300 g/capita/day</option>
                  <option value="350">350 g/capita/day</option>
                  <option value="400">400 g/capita/day</option>
                  <option value="450">450 g/capita/day</option>
                  <option value="500">500 g/capita/day</option>
                </select>
              </div>

              <div style={{ gridColumn: 'span 3' }}>
                <label>{lang === 'hi' ? 'या सीधा कुल टन प्रति दिन दर्ज करें (ऑप्शनल Override)' : 'OR Enter Direct Total Tons/Day (Override)'}</label>
                <input style={{ width: '100%', padding: '6px' }} type="number" step="0.01" placeholder={lang === 'hi' ? 'जनसंख्या का उपयोग करने के लिए इसे खाली छोड़ें' : 'Leave blank to use population calculation'} value={ulbFixedTons} onChange={(e) => setUlbFixedTons(e.target.value)} />
              </div>
            </>
          ) : (
            <>
              <div>
                <label>{lang === 'hi' ? 'औसत सूखा कचरा आवक (टन/दिन)' : 'Avg Daily Dry Waste (Tons/Day)'}</label>
                <input style={{ width: '100%', padding: '6px' }} type="number" step="0.01" required value={mrfDailyDryTons} onChange={(e) => setMrfDailyDryTons(e.target.value)} />
              </div>

              <div>
                <label>{lang === 'hi' ? 'एमआरएफ की कुल क्षमता (टन/दिन)' : 'MRF Max Capacity (Tons/Day)'}</label>
                <input style={{ width: '100%', padding: '6px' }} type="number" step="0.01" required value={mrfMaxCapacityTons} onChange={(e) => setMrfMaxCapacityTons(e.target.value)} />
              </div>
            </>
          )}

          <div>
            <label>{lang === 'hi' ? 'वर्ष (Year)' : 'Year'}</label>
            <input style={{ width: '100%', padding: '6px' }} type="number" value={startYear} onChange={(e) => setStartYear(Number(e.target.value))} />
          </div>
        </div>

        {/* Month Picker Buttons */}
        <div style={{ marginBottom: '15px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
            <label style={{ fontWeight: 'bold' }}>{lang === 'hi' ? 'माह चुनें (अधिकतम 3 माह तक टिक करें):' : 'Select Months (Tick up to 3):'}</label>
            <span style={{ fontSize: '13px', fontWeight: 'bold', color: '#059669' }}>
              {selectedMonths.length} {lang === 'hi' ? 'माह चयनित' : 'Month/s Selected'} — {lang === 'hi' ? `शुल्क: ₹${getPrice()}` : `Total: ₹${getPrice()}`} 
              {selectedMonths.length === 3 && (lang === 'hi' ? ' (₹25 छूट लागू)' : ' (₹25 Discount Included)')}
            </span>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(6, 1fr)', gap: '8px' }}>
            {MONTH_NAMES.map((m) => {
              const isSelected = selectedMonths.includes(m.id);
              return (
                <button
                  key={m.id}
                  type="button"
                  onClick={() => toggleMonth(m.id)}
                  style={{
                    padding: '8px 4px', borderRadius: '6px',
                    border: isSelected ? '2px solid #059669' : '1px solid #cbd5e1',
                    background: isSelected ? '#ecfdf5' : '#ffffff',
                    color: isSelected ? '#065f46' : '#334155',
                    fontWeight: isSelected ? 'bold' : 'normal',
                    cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '4px', fontSize: '13px'
                  }}
                >
                  {isSelected && <Check size={14} style={{ color: '#059669' }} />}
                  {lang === 'hi' ? m.shortHi : m.shortEn}
                </button>
              );
            })}
          </div>
        </div>

        <button type="submit" style={{ width: '100%', padding: '12px', background: '#059669', color: '#fff', border: 'none', borderRadius: '5px', cursor: 'pointer', fontWeight: 'bold', fontSize: '15px' }}>
          {lang === 'hi' ? 'लोगबुक जनरेट करें' : 'Generate Dataset'}
        </button>
      </form>

      {/* Dataset Output & Tab Navigation */}
      {generatedMonthlyData && (
        <div>
          {/* Month Tab Headers */}
          <div style={{ display: 'flex', borderBottom: '2px solid #cbd5e1', marginBottom: '15px' }}>
            {selectedMonths.map((mId) => {
              const monthObj = MONTH_NAMES.find(m => m.id === mId);
              const isActive = activeTabMonth === mId;
              return (
                <button
                  key={mId}
                  onClick={() => setActiveTabMonth(mId)}
                  style={{
                    padding: '10px 20px',
                    border: 'none',
                    borderBottom: isActive ? '3px solid #059669' : 'none',
                    background: isActive ? '#ecfdf5' : 'transparent',
                    fontWeight: isActive ? 'bold' : 'normal',
                    color: isActive ? '#065f46' : '#64748b',
                    cursor: 'pointer',
                    fontSize: '14px'
                  }}
                >
                  {lang === 'hi' ? monthObj?.shortHi : monthObj?.fullEn}
                </button>
              );
            })}
          </div>

          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
            <div>
              <strong>{name}</strong> ({facilityType}) — <small>{MONTH_NAMES.find(m => m.id === activeTabMonth)?.fullEn} {startYear}</small>
            </div>
            <div>
              <button onClick={() => setDisplayUnit(displayUnit === 'Tons' ? 'kg' : 'Tons')} style={{ marginRight: '10px', padding: '6px 10px' }}>
                {lang === 'hi' ? 'इकाई:' : 'Unit:'} {displayUnit}
              </button>
              {isPaid ? (
                <button onClick={downloadMultiSheetExcel} style={{ padding: '8px 15px', background: '#0f172a', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>
                  <Download size={14} /> {lang === 'hi' ? 'मल्टी-शीट एक्सल (.xlsx) एक्सपोर्ट करें' : 'Export Multi-Sheet Excel (.xlsx)'}
                </button>
              ) : (
                <button onClick={handlePayment} disabled={isProcessing} style={{ padding: '8px 15px', background: '#059669', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>
                  {isProcessing ? 'Wait...' : `${lang === 'hi' ? `₹${getPrice()} भुगतान करें एवं एक्सल डाउनलोड करें` : `Pay ₹${getPrice()} to Unlock & Download Excel`}`}
                </button>
              )}
            </div>
          </div>

          <table border="1" cellPadding="6" style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13px' }}>
            <thead>
              <tr style={{ background: '#f1f5f9' }}>
                <th>Date</th><th>Day</th>
                {facilityType === 'ULB' ? (
                  <>
                    <th>Wet ({displayUnit})</th><th>Plastic ({displayUnit})</th><th>Textile ({displayUnit})</th>
                    <th>C&D ({displayUnit})</th><th>Hazardous ({displayUnit})</th><th>Inerts ({displayUnit})</th><th>Total ({displayUnit})</th>
                  </>
                ) : (
                  <>
                    <th>PET ({displayUnit})</th><th>HDPE ({displayUnit})</th><th>Cardboard ({displayUnit})</th>
                    <th>RDF ({displayUnit})</th><th>Glass/Metal ({displayUnit})</th><th>Rejects ({displayUnit})</th><th>Total Dry ({displayUnit})</th>
                  </>
                )}
              </tr>
            </thead>
            <tbody>
              {visibleRows.map((row, idx) => (
                <tr key={idx}>
                  <td>{row.date}</td><td>{row.dayName}</td>
                  <td>{formatVal(row.c1)}</td><td>{formatVal(row.c2)}</td>
                  <td>{formatVal(row.c3)}</td><td>{formatVal(row.c4)}</td>
                  <td>{formatVal(row.c5)}</td><td>{formatVal(row.c6)}</td>
                  <td><strong>{formatVal(row.total)}</strong></td>
                </tr>
              ))}
            </tbody>
          </table>

          {!isPaid && (
            <div style={{ border: '2px dashed #059669', background: '#ecfdf5', padding: '20px', textAlign: 'center', marginTop: '10px', borderRadius: '6px' }}>
              <Lock style={{ color: '#059669', marginBottom: '5px' }} />
              <h3 style={{ margin: '5px 0', color: '#065f46' }}>
                {lang === 'hi' ? 'पूर्वावलोकन लॉक है (केवल शुरुआती 5 दिन दिख रहे हैं)' : 'Preview Locked (Showing Days 1 to 5)'}
              </h3>
              <p style={{ margin: '5px 0 15px 0', color: '#047857', fontSize: '14px' }}>
                {lang === 'hi' ? `चयनित ${selectedMonths.length} माह की अलग-अलग शीट वाली एक्सल (.xlsx) फाइल डाउनलोड करने के लिए ₹${getPrice()} का भुगतान करें।` : `Pay ₹${getPrice()} to unlock all months and export a multi-sheet Excel file (.xlsx) with separate tabs.`}
              </p>
              <button onClick={handlePayment} disabled={isProcessing} style={{ padding: '10px 20px', background: '#059669', color: '#fff', border: 'none', borderRadius: '5px', fontSize: '15px', fontWeight: 'bold', cursor: 'pointer' }}>
                {isProcessing ? 'Connecting...' : `${lang === 'hi' ? `₹${getPrice()} का भुगतान करें` : `Pay ₹${getPrice()} & Download Excel File`}`}
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
