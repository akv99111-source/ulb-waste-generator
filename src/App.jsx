import React, { useState, useRef } from 'react';
import { Building2, Download, Lock, Globe, Check, Info, ShieldCheck } from 'lucide-react';
import * as XLSX from 'xlsx';

const MONTHS = [
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

const getSeasonalFractions = (m, type) => {
  if (type === 'ULB') {
    if ([5, 6, 7].includes(m)) return [0.58, 0.18, 0.04, 0.02, 0.05, 0.13]; // High Wet (Fruits)
    if ([8, 9].includes(m)) return [0.56, 0.19, 0.04, 0.02, 0.05, 0.14];    // Monsoon
    if ([11, 12, 1, 2].includes(m)) return [0.50, 0.23, 0.05, 0.02, 0.06, 0.14]; // Winter Packaging
    return [0.54, 0.20, 0.04, 0.02, 0.05, 0.15];
  }
  return [5, 6, 7].includes(m) ? [0.25, 0.15, 0.20, 0.22, 0.08, 0.10] : [0.20, 0.15, 0.25, 0.20, 0.10, 0.10];
};

const inputStyle = { width: '100%', padding: '9px', borderRadius: '6px', border: '1px solid #cbd5e1', marginTop: '4px', boxSizing: 'border-box' };

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
  const [selectedMonths, setSelectedMonths] = useState([1, 12]);
  const [displayUnit, setDisplayUnit] = useState('Tons');
  const [generatedMonthlyData, setGeneratedMonthlyData] = useState(null);
  const [activeTabMonth, setActiveTabMonth] = useState(null);
  const [isPaid, setIsPaid] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);

  const resultsRef = useRef(null);

  const toggleMonth = (mId) => {
    if (selectedMonths.includes(mId)) {
      if (selectedMonths.length > 1) setSelectedMonths(selectedMonths.filter(m => m !== mId));
    } else {
      if (selectedMonths.length >= 3) return alert(lang === 'hi' ? 'अधिकतम 3 माह चुन सकते हैं।' : 'Max 3 months allowed.');
      setSelectedMonths([...selectedMonths, mId].sort((a, b) => a - b));
    }
  };

  const getPrice = () => [0, 50, 100, 125][selectedMonths.length] || 50;

  const handleGenerate = (e) => {
    e.preventDefault();
    setIsPaid(false);
    let monthlyDataMap = {};

    selectedMonths.forEach((m) => {
      const days = [1, 3, 5, 7, 8, 10, 12].includes(m) ? 31 : (m === 2 ? 28 : 30);
      const targetTons = facilityType === 'ULB' 
        ? (ulbFixedTons > 0 ? parseFloat(ulbFixedTons) : (population * (parseFloat(perCapita) / 1000)) / 1000)
        : parseFloat(mrfDailyDryTons);

      const baseFractions = getSeasonalFractions(m, facilityType);
      let logs = [];

      for (let day = 1; day <= days; day++) {
        const dateObj = new Date(startYear, m - 1, day);
        let dailyTotal = targetTons * (0.95 + Math.random() * 0.10) * ([0, 6].includes(dateObj.getDay()) ? 1.05 : 1);
        let raw = baseFractions.map(r => r * (0.88 + Math.random() * 0.24));
        let sum = raw.reduce((a, b) => a + b, 0);
        let norm = raw.map(r => r / sum);

        logs.push({
          date: dateObj.toISOString().split('T')[0],
          dayName: dateObj.toLocaleDateString('en-US', { weekday: 'short' }),
          c1: parseFloat((dailyTotal * norm[0]).toFixed(2)),
          c2: parseFloat((dailyTotal * norm[1]).toFixed(2)),
          c3: parseFloat((dailyTotal * norm[2]).toFixed(2)),
          c4: parseFloat((dailyTotal * norm[3]).toFixed(2)),
          c5: parseFloat((dailyTotal * norm[4]).toFixed(2)),
          c6: parseFloat((dailyTotal * norm[5]).toFixed(2)),
          total: parseFloat(dailyTotal.toFixed(2))
        });
      }
      monthlyDataMap[m] = logs;
    });

    setGeneratedMonthlyData(monthlyDataMap);
    setActiveTabMonth(selectedMonths[0]);
    setTimeout(() => resultsRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 100);
  };

  const handlePayment = async () => {
    setIsProcessing(true);
    if (!window.Razorpay) {
      await new Promise((res) => {
        const s = document.createElement('script');
        s.src = 'https://checkout.razorpay.com/v1/checkout.js';
        s.onload = () => res(true);
        document.body.appendChild(s);
      });
    }

    try {
      const res = await fetch('/api/create-order', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ amount: getPrice() * 100 })
      });
      const order = await res.json();

      new window.Razorpay({
        key: import.meta.env.VITE_RAZORPAY_KEY_ID,
        amount: order.amount,
        currency: order.currency,
        name: 'SWM 2024 Logbook Engine',
        description: `${facilityType} (${selectedMonths.length} Months) - ${name}`,
        order_id: order.id,
        handler: () => { setIsPaid(true); setIsProcessing(false); downloadMultiSheetExcel(); },
        theme: { color: '#059669' }
      }).open();
    } catch (err) {
      alert('Payment Error: ' + err.message);
      setIsProcessing(false);
    }
  };

  const formatVal = (v) => displayUnit === 'kg' ? Math.round(v * 1000) : v.toFixed(2);

  const downloadMultiSheetExcel = () => {
    if (!generatedMonthlyData) return;
    const u = displayUnit === 'kg' ? 'kg' : 'Tons';
    const headers = facilityType === 'ULB'
      ? ["Date", "Day", `Wet (${u})`, `Dry (${u})`, `Sanitary (${u})`, `Special Care/Hazardous (${u})`, `C&D (${u})`, `Inerts (${u})`, `Total (${u})`]
      : ["Date", "Day", `PET (${u})`, `HDPE (${u})`, `Paper/Cardboard (${u})`, `RDF/SCF (${u})`, `Glass & Metal (${u})`, `Rejects (${u})`, `Total Dry (${u})` ];

    const wb = XLSX.utils.book_new();
    selectedMonths.forEach((mId) => {
      const sheetData = [headers, ...generatedMonthlyData[mId].map(r => [
        r.date, r.dayName, formatVal(r.c1), formatVal(r.c2), formatVal(r.c3), formatVal(r.c4), formatVal(r.c5), formatVal(r.c6), formatVal(r.total)
      ])];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sheetData), MONTHS.find(m => m.id === mId)?.fullEn);
    });
    XLSX.writeFile(wb, `${name.replace(/\s+/g, '_')}_SWM2024_${selectedMonths.length}M.xlsx`);
  };

  const activeRows = generatedMonthlyData?.[activeTabMonth] || [];
  const visibleRows = isPaid ? activeRows : activeRows.slice(0, 5);

  return (
    <div style={{ fontFamily: 'sans-serif', background: '#f8fafc', minHeight: '100vh' }}>
      
      {/* HEADER BANNER */}
      <div style={{ background: 'linear-gradient(135deg, #064e3b 0%, #047857 100%)', color: '#fff', padding: '20px 15px' }}>
        <div style={{ maxWidth: '1000px', margin: '0 auto', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
          <div>
            <span style={{ background: 'rgba(255,255,255,0.2)', padding: '3px 8px', borderRadius: '12px', fontSize: '11px', fontWeight: 'bold' }}>
              <ShieldCheck size={12} style={{ verticalAlign: 'middle' }} /> SWM 2024 COMPLIANT
            </span>
            <h1 style={{ fontSize: '22px', margin: '6px 0 2px 0', fontWeight: '800' }}>
              <Building2 size={22} style={{ verticalAlign: 'middle', marginRight: '6px' }} />
              {lang === 'hi' ? 'यूएलबी एवं एमआरएफ लोगबुक जनरेटर (SWM 2024)' : 'ULB & MRF Waste Logbook Generator (SWM 2024)'}
            </h1>
            <p style={{ fontSize: '13px', margin: 0, color: '#a7f3d0' }}>
              {lang === 'hi' ? 'उत्तर भारत एवं यूपी के लिए 4-स्ट्रीम अपशिष्ट पृथक्कीकरण टूल' : 'Automated 4-Stream Logbook Engine aligned with SWM 2024'}
            </p>
          </div>
          <button onClick={() => setLang(lang === 'hi' ? 'en' : 'hi')} style={{ padding: '6px 12px', background: '#fff', color: '#047857', border: 'none', borderRadius: '5px', fontWeight: 'bold', cursor: 'pointer', fontSize: '13px' }}>
            <Globe size={15} style={{ verticalAlign: 'middle', marginRight: '4px' }} /> {lang === 'hi' ? 'English' : 'हिंदी'}
          </button>
        </div>
      </div>

      {/* FORM SECTION */}
      <div style={{ padding: '15px', maxWidth: '1000px', margin: '0 auto' }}>
        <form onSubmit={handleGenerate} style={{ background: '#fff', padding: '16px', borderRadius: '8px', border: '1px solid #cbd5e1', marginBottom: '20px' }}>
          
          <div style={{ marginBottom: '14px', display: 'flex', gap: '15px', alignItems: 'center', fontSize: '14px' }}>
            <strong>{lang === 'hi' ? 'सुविधा प्रकार:' : 'Facility:'}</strong>
            <label><input type="radio" value="ULB" checked={facilityType === 'ULB'} onChange={() => setFacilityType('ULB')} /> ULB</label>
            <label><input type="radio" value="MRF" checked={facilityType === 'MRF'} onChange={() => setFacilityType('MRF')} /> MRF Centre</label>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '12px', marginBottom: '14px' }}>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'निकाय / एमआरएफ का नाम' : 'ULB / MRF Name'}</label>
              <input style={inputStyle} type="text" required value={name} onChange={(e) => setName(e.target.value)} />
            </div>

            {facilityType === 'ULB' ? (
              <>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'जनसंख्या' : 'Population'}</label>
                  <input style={{ ...inputStyle, background: ulbFixedTons ? '#e2e8f0' : '#fff' }} type="number" disabled={!!ulbFixedTons} value={population} onChange={(e) => setPopulation(Number(e.target.value))} />
                </div>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'प्रतिव्यक्ति अपशिष्ट' : 'Per Capita Rate'}</label>
                  <select style={{ ...inputStyle, background: ulbFixedTons ? '#e2e8f0' : '#fff' }} disabled={!!ulbFixedTons} value={perCapita} onChange={(e) => setPerCapita(e.target.value)}>
                    <option value="350">350 g/capita/day</option>
                    <option value="400">400 g/capita/day</option>
                    <option value="450">450 g/capita/day</option>
                    <option value="500">500 g/capita/day</option>
                  </select>
                </div>
              </>
            ) : (
              <>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'सूखा कचरा आवक (टन/दिन)' : 'Daily Dry Input (Tons)'}</label>
                  <input style={inputStyle} type="number" step="0.01" required value={mrfDailyDryTons} onChange={(e) => setMrfDailyDryTons(e.target.value)} />
                </div>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'एमआरएफ क्षमता (टन/दिन)' : 'MRF Capacity (Tons)'}</label>
                  <input style={inputStyle} type="number" step="0.01" required value={mrfMaxCapacityTons} onChange={(e) => setMrfMaxCapacityTons(e.target.value)} />
                </div>
              </>
            )}

            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'वर्ष' : 'Year'}</label>
              <input style={inputStyle} type="number" value={startYear} onChange={(e) => setStartYear(Number(e.target.value))} />
            </div>
          </div>

          {/* MONTH SELECTION */}
          <div style={{ marginBottom: '14px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '6px', fontSize: '13px' }}>
              <strong>{lang === 'hi' ? 'माह चुनें (अधिकतम 3):' : 'Select Months (Max 3):'}</strong>
              <span style={{ color: '#059669', fontWeight: 'bold' }}>
                {selectedMonths.length} {lang === 'hi' ? 'माह' : 'Month/s'} — ₹{getPrice()}
              </span>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(75px, 1fr))', gap: '6px' }}>
              {MONTHS.map((m) => {
                const active = selectedMonths.includes(m.id);
                return (
                  <button key={m.id} type="button" onClick={() => toggleMonth(m.id)} style={{
                    padding: '8px 2px', borderRadius: '5px', border: active ? '2px solid #059669' : '1px solid #cbd5e1',
                    background: active ? '#ecfdf5' : '#fff', color: active ? '#065f46' : '#334155', fontWeight: active ? 'bold' : 'normal', cursor: 'pointer', fontSize: '12px'
                  }}>
                    {active && <Check size={11} />} {lang === 'hi' ? m.shortHi : m.shortEn}
                  </button>
                );
              })}
            </div>
          </div>

          <button type="submit" style={{ width: '100%', padding: '12px', background: '#059669', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '15px' }}>
            {lang === 'hi' ? 'SWM 2024 लोगबुक जनरेट करें →' : 'Generate SWM 2024 Dataset →'}
          </button>
        </form>

        {/* RESULTS SECTION */}
        {generatedMonthlyData && (
          <div ref={resultsRef} style={{ background: '#fff', padding: '15px', borderRadius: '8px', border: '1px solid #cbd5e1', scrollMarginTop: '15px' }}>
            
            <div style={{ background: '#eff6ff', border: '1px solid #bfdbfe', padding: '10px', borderRadius: '6px', marginBottom: '12px', fontSize: '12px', color: '#1e40af' }}>
              <strong><Info size={14} style={{ verticalAlign: 'middle' }} /> SWM 2024 Guidance:</strong> 4-Stream segregation (Wet, Dry, Sanitary, Special Care) applied. Seasonal variations (~55-60% Wet) scaled for UP/North India.
            </div>

            {/* TABS */}
            <div style={{ display: 'flex', borderBottom: '2px solid #e2e8f0', marginBottom: '12px', overflowX: 'auto' }}>
              {selectedMonths.map((mId) => (
                <button key={mId} onClick={() => setActiveTabMonth(mId)} style={{
                  padding: '8px 16px', border: 'none', borderBottom: activeTabMonth === mId ? '3px solid #059669' : 'none',
                  background: activeTabMonth === mId ? '#ecfdf5' : 'transparent', fontWeight: activeTabMonth === mId ? 'bold' : 'normal', cursor: 'pointer', fontSize: '13px'
                }}>
                  {MONTHS.find(m => m.id === mId)?.[lang === 'hi' ? 'shortHi' : 'fullEn']}
                </button>
              ))}
            </div>

            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px', flexWrap: 'wrap', gap: '8px' }}>
              <span style={{ fontSize: '13px', fontWeight: 'bold' }}>{name} — {MONTHS.find(m => m.id === activeTabMonth)?.fullEn} {startYear}</span>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button onClick={() => setDisplayUnit(displayUnit === 'Tons' ? 'kg' : 'Tons')} style={{ padding: '6px 10px', fontSize: '12px', borderRadius: '4px', border: '1px solid #cbd5e1' }}>
                  Unit: <strong>{displayUnit}</strong>
                </button>
                {isPaid ? (
                  <button onClick={downloadMultiSheetExcel} style={{ padding: '6px 12px', background: '#0f172a', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold' }}>
                    <Download size={13} /> Export Excel (.xlsx)
                  </button>
                ) : (
                  <button onClick={handlePayment} disabled={isProcessing} style={{ padding: '6px 12px', background: '#059669', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold' }}>
                    {isProcessing ? 'Wait...' : `Pay ₹${getPrice()} to Unlock`}
                  </button>
                )}
              </div>
            </div>

            {/* TABLE */}
            <div style={{ overflowX: 'auto', border: '1px solid #cbd5e1', borderRadius: '4px' }}>
              <table cellPadding="8" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', minWidth: '700px' }}>
                <thead>
                  <tr style={{ background: '#f1f5f9', borderBottom: '1px solid #cbd5e1' }}>
                    <th>Date</th><th>Day</th>
                    {facilityType === 'ULB' ? (
                      <>
                        <th style={{ background: '#ecfdf5' }}>Wet ({displayUnit})</th>
                        <th>Dry ({displayUnit})</th>
                        <th>Sanitary ({displayUnit})</th>
                        <th>Special Care ({displayUnit})</th>
                        <th>C&D ({displayUnit})</th>
                        <th>Inerts ({displayUnit})</th>
                        <th>Total ({displayUnit})</th>
                      </>
                    ) : (
                      <>
                        <th>PET</th><th>HDPE</th><th>Paper</th><th>RDF</th><th>Glass/Metal</th><th>Rejects</th><th>Total Dry</th>
                      </>
                    )}
                  </tr>
                </thead>
                <tbody>
                  {visibleRows.map((r, idx) => (
                    <tr key={idx} style={{ borderBottom: '1px solid #e2e8f0' }}>
                      <td>{r.date}</td><td>{r.dayName}</td>
                      <td style={{ background: facilityType === 'ULB' ? '#ecfdf5' : '#fff' }}>{formatVal(r.c1)}</td>
                      <td>{formatVal(r.c2)}</td><td>{formatVal(r.c3)}</td><td>{formatVal(r.c4)}</td>
                      <td>{formatVal(r.c5)}</td><td>{formatVal(r.c6)}</td><td><strong>{formatVal(r.total)}</strong></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {!isPaid && (
              <div style={{ border: '2px dashed #059669', background: '#ecfdf5', padding: '15px', textAlign: 'center', marginTop: '12px', borderRadius: '6px' }}>
                <Lock style={{ color: '#059669' }} size={18} />
                <h4 style={{ margin: '4px 0', color: '#065f46', fontSize: '15px' }}>Preview Locked (Days 1–5 Only)</h4>
                <p style={{ margin: '4px 0 10px 0', color: '#047857', fontSize: '13px' }}>Pay ₹{getPrice()} to unlock complete dataset and multi-tab Excel (.xlsx) file.</p>
                <button onClick={handlePayment} disabled={isProcessing} style={{ padding: '10px 20px', background: '#059669', color: '#fff', border: 'none', borderRadius: '6px', fontSize: '14px', fontWeight: 'bold', cursor: 'pointer' }}>
                  {isProcessing ? 'Connecting...' : `Pay ₹${getPrice()} & Download Complete File`}
                </button>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
