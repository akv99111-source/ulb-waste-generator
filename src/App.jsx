import React, { useState } from 'react';
import { Building2, Download, Lock, Globe, Check, Info, ShieldCheck } from 'lucide-react';
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

  // SWM 2024 Seasonal Fraction Distribution Logic
  const getSeasonalFractions = (monthId, type) => {
    if (type === 'ULB') {
      // 4-Stream Fractions: [Wet, Dry, Sanitary, Special Care/Hazardous, C&D, Inerts]
      if ([5, 6, 7].includes(monthId)) return [0.58, 0.18, 0.04, 0.02, 0.05, 0.13]; // High Wet Waste (Fruit/Mango Season)
      if ([8, 9].includes(monthId)) return [0.56, 0.19, 0.04, 0.02, 0.05, 0.14]; // Monsoon Moisture
      if ([11, 12, 1, 2].includes(monthId)) return [0.50, 0.23, 0.05, 0.02, 0.06, 0.14]; // Winter Dry Packaging
      return [0.54, 0.20, 0.04, 0.02, 0.05, 0.15]; // Standard baseline
    } else {
      // MRF Dry Waste Fractions: [PET, HDPE/Hard Plastic, Cardboard/Paper, RDF/SCF, Glass/Metal, Rejects]
      if ([5, 6, 7].includes(monthId)) return [0.25, 0.15, 0.20, 0.22, 0.08, 0.10];
      return [0.20, 0.15, 0.25, 0.20, 0.10, 0.10];
    }
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

      const baseFractions = getSeasonalFractions(calcMonth, facilityType);

      let monthLogs = [];
      for (let day = 1; day <= daysInMonth; day++) {
        const currentDate = new Date(startYear, calcMonth - 1, day);
        const dayName = currentDate.toLocaleDateString('en-US', { weekday: 'short' });
        const isWeekend = currentDate.getDay() === 0 || currentDate.getDay() === 6;

        let noise = 0.95 + Math.random() * 0.10;
        if (isWeekend) noise *= 1.05;
        const dailyTotal = targetDailyTons * noise;

        let rawRatios = baseFractions.map(r => r * (0.88 + Math.random() * 0.24));
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
        name: 'SWM 2024 Logbook Engine',
        description: `${facilityType} SWM 2024 Excel (${selectedMonths.length} Month/s) - ${name}`,
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
      headers = [
        "Date", "Day", 
        `Wet Waste (${unitLabel})`, 
        `Dry Waste (${unitLabel})`, 
        `Sanitary Waste (${unitLabel})`, 
        `Special Care / Domestic Hazardous (${unitLabel})`, 
        `C&D Waste (${unitLabel})`, 
        `Inerts / Residual (${unitLabel})`, 
        `Total Generated (${unitLabel})`
      ];
    } else {
      headers = [
        "Date", "Day", 
        `PET Plastics (${unitLabel})`, 
        `HDPE / Hard Plastics (${unitLabel})`, 
        `Paper / Cardboard (${unitLabel})`, 
        `RDF / SCF (${unitLabel})`, 
        `Glass & Metals (${unitLabel})`, 
        `Rejects (${unitLabel})`, 
        `Total Dry Waste (${unitLabel})`
      ];
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

    XLSX.writeFile(workbook, `${name.replace(/\s+/g, '_')}_SWM2024_Logbook_${selectedMonths.length}Months_${displayUnit}.xlsx`);
  };

  const activeRows = (generatedMonthlyData && activeTabMonth) ? generatedMonthlyData[activeTabMonth] : [];
  const visibleRows = isPaid ? activeRows : activeRows.slice(0, 5);

  return (
    <div style={{ fontFamily: 'sans-serif', margin: 0, background: '#f8fafc', minHeight: '100vh' }}>
      
      {/* SWM 2024 HERO BANNER SECTION */}
      <div style={{
        background: 'linear-gradient(135deg, #064e3b 0%, #047857 100%)',
        color: '#ffffff',
        padding: '25px 15px',
        borderBottom: '1px solid #065f46',
        textAlign: 'left'
      }}>
        <div style={{ maxWidth: '1050px', margin: '0 auto', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '15px' }}>
          
          <div>
            <div style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', background: 'rgba(255,255,255,0.2)', padding: '3px 10px', borderRadius: '20px', fontSize: '11px', fontWeight: 'bold', textTransform: 'uppercase', marginBottom: '8px', letterSpacing: '0.5px' }}>
              <ShieldCheck size={14} /> {lang === 'hi' ? 'SWM 2024 दिशानिर्देशों के अनुरूप (4-Stream Segregation)' : 'SWM 2024 Rules Compliant (4-Stream Segregation)'}
            </div>
            <h1 style={{ fontSize: '24px', margin: '0 0 6px 0', fontWeight: '800', display: 'flex', alignItems: 'center', gap: '10px' }}>
              <Building2 size={26} /> {lang === 'hi' ? 'यूएलबी एवं एमआरएफ लोगबुक जनरेटर (SWM 2024)' : 'ULB & MRF Waste Logbook Generator (SWM 2024)'}
            </h1>
            <p style={{ fontSize: '14px', margin: 0, color: '#a7f3d0' }}>
              {lang === 'hi' ? 'उत्तर भारत एवं यूपी के लिए 4-स्ट्रीम अपशिष्ट पृथक्कीकरण और सीजनल वेरिएशंस हेतु स्वचालित टूल' : 'Automated 4-Stream Waste Logbook Generation Engine aligned with SWM 2024 Standards'}
            </p>
          </div>

          <button onClick={() => setLang(lang === 'hi' ? 'en' : 'hi')} style={{
            padding: '8px 16px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px', 
            borderRadius: '6px', border: 'none', background: '#ffffff', color: '#047857', fontWeight: 'bold', fontSize: '14px'
          }}>
            <Globe size={18} /> {lang === 'hi' ? 'English Interface' : 'हिंदी इंटरफेस'}
          </button>

        </div>
      </div>

      {/* Main Container */}
      <div style={{ padding: '20px', maxWidth: '1050px', margin: '0 auto' }}>
        
        {/* Form Container */}
        <form onSubmit={handleGenerate} style={{ background: '#ffffff', padding: '18px', borderRadius: '10px', border: '1px solid #e2e8f0', marginBottom: '25px', boxShadow: '0 1px 3px rgba(0,0,0,0.1)' }}>
          
          <div style={{ marginBottom: '18px', display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
            <label style={{ fontWeight: 'bold', fontSize: '15px' }}>{lang === 'hi' ? 'सुविधा का प्रकार चुनें:' : 'Select Facility Type:'}</label>
            <label style={{ cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '5px' }}>
              <input type="radio" name="facility" value="ULB" checked={facilityType === 'ULB'} onChange={() => setFacilityType('ULB')} style={{ transform: 'scale(1.1)' }}/> {lang === 'hi' ? 'नगर निकाय (ULB)' : 'Urban Local Body (ULB)'}
            </label>
            <label style={{ cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '5px' }}>
              <input type="radio" name="facility" value="MRF" checked={facilityType === 'MRF'} onChange={() => setFacilityType('MRF')} style={{ transform: 'scale(1.1)' }}/> {lang === 'hi' ? 'एमआरएफ केंद्र (MRF Centre)' : 'MRF Centre'}
            </label>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '15px', marginBottom: '18px' }}>
            <div>
              <label style={{ fontSize: '13px', fontWeight: '600', color: '#475569' }}>{lang === 'hi' ? 'निकाय / एमआरएफ का नाम' : 'Name of Municipal Body / MRF Centre'}</label>
              <input style={{ width: '100%', padding: '10px', boxSizing: 'border-box', borderRadius: '6px', border: '1px solid #cbd5e1', marginTop: '4px' }} type="text" required value={name} onChange={(e) => setName(e.target.value)} />
            </div>

            {facilityType === 'ULB' ? (
              <>
                <div>
                  <label style={{ fontSize: '13px', fontWeight: '600', color: '#475569' }}>{lang === 'hi' ? 'अनुमानित जनसंख्या (Population)' : 'Population'}</label>
                  <input style={{ width: '100%', padding: '10px', boxSizing: 'border-box', borderRadius: '6px', border: '1px solid #cbd5e1', marginTop: '4px', background: ulbFixedTons ? '#e2e8f0' : '#fff' }} type="number" disabled={!!ulbFixedTons} value={population} onChange={(e) => setPopulation(Number(e.target.value))} />
                </div>

                <div>
                  <label style={{ fontSize: '13px', fontWeight: '600', color: '#475569' }}>{lang === 'hi' ? 'प्रतिव्यक्ति अपशिष्ट (g/दिन)' : 'Per Capita Waste Rate'}</label>
                  <select style={{ width: '100%', padding: '10px', boxSizing: 'border-box', borderRadius: '6px', border: '1px solid #cbd5e1', marginTop: '4px', background: ulbFixedTons ? '#e2e8f0' : '#fff' }} disabled={!!ulbFixedTons} value={perCapita} onChange={(e) => setPerCapita(e.target.value)}>
                    <option value="300">300 g/capita/day</option>
                    <option value="350">350 g/capita/day</option>
                    <option value="400">400 g/capita/day</option>
                    <option value="450">450 g/capita/day</option>
                    <option value="500">500 g/capita/day</option>
                  </select>
                </div>

                <div style={{ gridColumn: '1 / -1', background: '#f1f5f9', padding: '10px', borderRadius: '6px', border: '1px solid #e2e8f0' }}>
                  <label style={{ fontSize: '13px', fontWeight: '600', color: '#475569' }}>{lang === 'hi' ? 'या सीधा कुल टन प्रति दिन दर्ज करें (ऑप्शनल Override)' : 'OR Enter Direct Total Tons/Day (Override)'}</label>
                  <input style={{ width: '100%', padding: '10px', boxSizing: 'border-box', borderRadius: '6px', border: '1px solid #cbd5e1', marginTop: '4px' }} type="number" step="0.01" placeholder={lang === 'hi' ? 'जनसंख्या का उपयोग करने के लिए इसे खाली छोड़ें' : 'Leave blank to use population calculation'} value={ulbFixedTons} onChange={(e) => setUlbFixedTons(e.target.value)} />
                </div>
              </>
            ) : (
              <>
                <div>
                  <label style={{ fontSize: '13px', fontWeight: '600', color: '#475569' }}>{lang === 'hi' ? 'औसत सूखा कचरा आवक (टन/दिन)' : 'Avg Daily Dry Waste Input (Tons/Day)'}</label>
                  <input style={{ width: '100%', padding: '10px', boxSizing: 'border-box', borderRadius: '6px', border: '1px solid #cbd5e1', marginTop: '4px' }} type="number" step="0.01" required value={mrfDailyDryTons} onChange={(e) => setMrfDailyDryTons(e.target.value)} />
                </div>

                <div>
                  <label style={{ fontSize: '13px', fontWeight: '600', color: '#475569' }}>{lang === 'hi' ? 'एमआरएफ की कुल क्षमता (टन/दिन)' : 'MRF Max Capacity (Tons/Day)'}</label>
                  <input style={{ width: '100%', padding: '10px', boxSizing: 'border-box', borderRadius: '6px', border: '1px solid #cbd5e1', marginTop: '4px' }} type="number" step="0.01" required value={mrfMaxCapacityTons} onChange={(e) => setMrfMaxCapacityTons(e.target.value)} />
                </div>
              </>
            )}

            <div>
              <label style={{ fontSize: '13px', fontWeight: '600', color: '#475569' }}>{lang === 'hi' ? 'वर्ष चुनें (Year)' : 'Choose Year'}</label>
              <input style={{ width: '100%', padding: '10px', boxSizing: 'border-box', borderRadius: '6px', border: '1px solid #cbd5e1', marginTop: '4px' }} type="number" value={startYear} onChange={(e) => setStartYear(Number(e.target.value))} />
            </div>
          </div>

          <div style={{ marginBottom: '18px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px', flexWrap: 'wrap', gap: '5px' }}>
              <label style={{ fontWeight: 'bold', fontSize: '15px' }}>{lang === 'hi' ? 'माह चुनें (अधिकतम 3 माह तक टिक करें):' : 'Select Months (Tick up to 3):'}</label>
              <span style={{ fontSize: '14px', fontWeight: 'bold', color: '#059669', background: '#ecfdf5', padding: '4px 10px', borderRadius: '20px' }}>
                {selectedMonths.length} {lang === 'hi' ? 'माह चयनित' : 'Month/s Selected'} — {lang === 'hi' ? `शुल्क: ₹${getPrice()}` : `Total: ₹${getPrice()}`} 
                {selectedMonths.length === 3 && (lang === 'hi' ? ' (₹25 छूट लागू)' : ' (₹25 Discount Included)')}
              </span>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(85px, 1fr))', gap: '8px' }}>
              {MONTH_NAMES.map((m) => {
                const isSelected = selectedMonths.includes(m.id);
                return (
                  <button
                    key={m.id}
                    type="button"
                    onClick={() => toggleMonth(m.id)}
                    style={{
                      padding: '10px 4px', borderRadius: '8px',
                      border: isSelected ? '2px solid #059669' : '1px solid #cbd5e1',
                      background: isSelected ? '#ecfdf5' : '#ffffff',
                      color: isSelected ? '#065f46' : '#334155',
                      fontWeight: isSelected ? 'bold' : 'normal',
                      cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '3px', fontSize: '13px'
                    }}
                  >
                    {isSelected && <Check size={14} style={{ color: '#059669' }} />}
                    {lang === 'hi' ? m.shortHi : m.shortEn}
                  </button>
                );
              })}
            </div>
          </div>

          <button type="submit" style={{ width: '100%', padding: '14px', background: '#059669', color: '#fff', border: 'none', borderRadius: '8px', cursor: 'pointer', fontWeight: 'bold', fontSize: '16px', boxShadow: '0 2px 4px rgba(0,0,0,0.15)' }}>
            {lang === 'hi' ? 'SWM 2024 लोगबुक जनरेट करें →' : 'Generate SWM 2024 Dataset →'}
          </button>
        </form>

        {/* Dataset Output & SWM 2024 Guidance */}
        {generatedMonthlyData && (
          <div style={{ background: '#ffffff', padding: '18px', borderRadius: '10px', border: '1px solid #cbd5e1' }}>
            
            <div style={{ background: '#eff6ff', border: '1px solid #bfdbfe', padding: '15px', borderRadius: '8px', marginBottom: '18px', fontSize: '14px', color: '#1e40af', lineHeight: '1.5' }}>
              <div style={{ fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '6px' }}>
                <Info size={18} /> {lang === 'hi' ? 'SWM 2024 अनुपालन एवं सीजनल दिशानिर्देश:' : 'SWM 2024 Compliance & Seasonal Guidance Notes:'}
              </div>
              <ul style={{ margin: 0, paddingLeft: '20px' }}>
                <li>{lang === 'hi' ? 'SWM 2024 के तहत 4-स्ट्रीम पृथक्कीकरण (Wet, Dry, Sanitary, Special Care) अनिवार्य किया गया है।' : 'SWM 2024 framework mandates 4-stream waste segregation at source (Wet, Dry, Sanitary, Special Care/Hazardous).'}</li>
                <li>{lang === 'hi' ? 'उत्तर भारत एवं यूपी में मौसम (गर्मियों में आम/फलों के सीजन एवं मॉनसून) के अनुसार गीले कचरे (Wet Waste) का अनुपात 55-60% तक बदलता है।' : 'Wet waste fraction naturally varies up to 55-60% in North India based on seasonal fruit yields (Mango season) and monsoon moisture.'}</li>
                <li>{lang === 'hi' ? 'ऑडिटिंग एवं सीपीसीबी/एसपीसीबी (CPCB/SPCB) पोर्टल पर वार्षिक रिटर्न दाखिल करने के लिए हर महीने की अलग एक्सल शीट होना आवश्यक है।' : 'Separate monthly logbook tabs are required for environmental auditing and annual returns submission on CPCB/SPCB portals.'}</li>
              </ul>
            </div>

            {/* Month Tab Headers */}
            <div style={{ display: 'flex', borderBottom: '2px solid #e2e8f0', marginBottom: '20px', overflowX: 'auto', whiteSpace: 'nowrap' }}>
              {selectedMonths.map((mId) => {
                const monthObj = MONTH_NAMES.find(m => m.id === mId);
                const isActive = activeTabMonth === mId;
                return (
                  <button
                    key={mId}
                    onClick={() => setActiveTabMonth(mId)}
                    style={{
                      padding: '12px 20px',
                      border: 'none',
                      borderBottom: isActive ? '4px solid #059669' : 'none',
                      background: isActive ? '#ecfdf5' : 'transparent',
                      fontWeight: isActive ? 'bold' : 'normal',
                      color: isActive ? '#065f46' : '#64748b',
                      cursor: 'pointer',
                      fontSize: '15px'
                    }}
                  >
                    {lang === 'hi' ? monthObj?.shortHi : monthObj?.fullEn}
                  </button>
                );
              })}
            </div>

            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px', flexWrap: 'wrap', gap: '10px' }}>
              <div style={{ fontSize: '15px' }}>
                <strong style={{ color: '#0f172a' }}>{name}</strong> — <small style={{ color: '#475569' }}>{MONTH_NAMES.find(m => m.id === activeTabMonth)?.fullEn} {startYear}</small>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <button onClick={() => setDisplayUnit(displayUnit === 'Tons' ? 'kg' : 'Tons')} style={{ padding: '8px 12px', fontSize: '13px', borderRadius: '6px', border: '1px solid #cbd5e1', background: '#f1f5f9', cursor: 'pointer' }}>
                  {lang === 'hi' ? 'इकाई:' : 'Unit:'} <strong>{displayUnit}</strong>
                </button>
                {isPaid ? (
                  <button onClick={downloadMultiSheetExcel} style={{ padding: '10px 16px', background: '#0f172a', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontSize: '14px', fontWeight: 'bold' }}>
                    <Download size={16} /> {lang === 'hi' ? 'SWM 2024 एक्सल (.xlsx) डाउनलोड' : 'Export Excel (.xlsx)'}
                  </button>
                ) : (
                  <button onClick={handlePayment} disabled={isProcessing} style={{ padding: '10px 16px', background: '#059669', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '14px' }}>
                    {isProcessing ? 'Wait...' : `${lang === 'hi' ? `₹${getPrice()} भुगतान व अनलॉक` : `Pay ₹${getPrice()} & Unlock Now`}`}
                  </button>
                )}
              </div>
            </div>

            {/* Table Container */}
            <div style={{ overflowX: 'auto', border: '1px solid #cbd5e1', borderRadius: '6px', background: '#f8fafc' }}>
              <table cellPadding="10" style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13px', minWidth: '780px', background: '#ffffff' }}>
                <thead>
                  <tr style={{ background: '#f1f5f9', borderBottom: '2px solid #cbd5e1' }}>
                    <th>Date</th><th>Day</th>
                    {facilityType === 'ULB' ? (
                      <>
                        <th style={{ background: '#ecfdf5' }}>Wet ({displayUnit})</th>
                        <th>Dry ({displayUnit})</th>
                        <th>Sanitary ({displayUnit})</th>
                        <th>Special Care / Hazardous ({displayUnit})</th>
                        <th>C&D ({displayUnit})</th>
                        <th>Inerts ({displayUnit})</th>
                        <th style={{ background: '#f1f5f9' }}>Total ({displayUnit})</th>
                      </>
                    ) : (
                      <>
                        <th>PET ({displayUnit})</th>
                        <th>HDPE ({displayUnit})</th>
                        <th>Cardboard ({displayUnit})</th>
                        <th>RDF / SCF ({displayUnit})</th>
                        <th>Glass & Metal ({displayUnit})</th>
                        <th>Rejects ({displayUnit})</th>
                        <th>Total Dry ({displayUnit})</th>
                      </>
                    )}
                  </tr>
                </thead>
                <tbody>
                  {visibleRows.map((row, idx) => (
                    <tr key={idx} style={{ borderBottom: '1px solid #e2e8f0' }}>
                      <td>{row.date}</td><td>{row.dayName}</td>
                      <td style={{ background: facilityType === 'ULB' ? '#ecfdf5' : '#fff' }}>{formatVal(row.c1)}</td>
                      <td>{formatVal(row.c2)}</td>
                      <td>{formatVal(row.c3)}</td>
                      <td>{formatVal(row.c4)}</td>
                      <td>{formatVal(row.c5)}</td>
                      <td>{formatVal(row.c6)}</td>
                      <td><strong>{formatVal(row.total)}</strong></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {!isPaid && (
              <div style={{ border: '3px dashed #059669', background: '#ecfdf5', padding: '20px', textAlign: 'center', marginTop: '15px', borderRadius: '10px' }}>
                <Lock style={{ color: '#059669', marginBottom: '8px' }} size={20} />
                <h4 style={{ margin: '8px 0', color: '#065f46', fontSize: '16px' }}>
                  {lang === 'hi' ? 'पूर्वावलोकन लॉक है (केवल शुरुआती 5 दिन दिख रहे हैं)' : 'Preview Locked (Showing Days 1 to 5 only)'}
                </h4>
                <p style={{ margin: '8px 0 15px 0', color: '#047857', fontSize: '14px', lineHeight: '1.4' }}>
                  {lang === 'hi' ? `चयनित ${selectedMonths.length} माह की SWM 2024 प्रारूप वाली पूर्ण एक्सल (.xlsx) फाइल डाउनलोड करने के लिए ₹${getPrice()} का भुगतान करें।` : `Pay ₹${getPrice()} to unlock the full SWM 2024 Multi-Sheet Excel file (.xlsx) with separate tabs for all ${selectedMonths.length} selected months.`}
                </p>
                <button onClick={handlePayment} disabled={isProcessing} style={{ padding: '12px 24px', background: '#059669', color: '#fff', border: 'none', borderRadius: '8px', fontSize: '15px', fontWeight: 'bold', cursor: 'pointer', boxShadow: '0 2px 4px rgba(0,0,0,0.15)' }}>
                  {isProcessing ? 'Connecting...' : `${lang === 'hi' ? `₹${getPrice()} भुगतान व पूर्ण फाइल डाउनलोड करें` : `Pay ₹${getPrice()} & Download Complete File`}`}
                </button>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
