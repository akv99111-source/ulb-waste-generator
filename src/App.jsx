import React, { useState, useRef } from 'react';
import { Building2, Download, Lock, Globe, Check, Info, ShieldCheck, MapPin } from 'lucide-react';
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

const REGION_PROFILES = {
  north_plains: { wetBase: 0.54, dryBase: 0.20, nameEn: 'North Plains', nameHi: 'उत्तरी मैदानी क्षेत्र' },
  coastal_south: { wetBase: 0.62, dryBase: 0.16, nameEn: 'South & Coastal', nameHi: 'दक्षिण एवं तटीय क्षेत्र' },
  western_central: { wetBase: 0.48, dryBase: 0.26, nameEn: 'West & Central', nameHi: 'पश्चिम एवं मध्य भारत' },
  eastern_states: { wetBase: 0.56, dryBase: 0.18, nameEn: 'East India', nameHi: 'पूर्वी भारत' },
  hilly_ne: { wetBase: 0.45, dryBase: 0.28, nameEn: 'Hilly & North-East', nameHi: 'पहाड़ी व पूर्वोत्तर क्षेत्र' },
  national_avg: { wetBase: 0.52, dryBase: 0.22, nameEn: 'Pan-India Standard', nameHi: 'राष्ट्रीय औसत' }
};

const STATES_LIST = [
  { nameEn: 'Andaman & Nicobar Islands', nameHi: 'अंडमान और निकोबार', region: 'coastal_south' },
  { nameEn: 'Andhra Pradesh', nameHi: 'आंध्र प्रदेश', region: 'coastal_south' },
  { nameEn: 'Arunachal Pradesh', nameHi: 'अरुणाचल प्रदेश', region: 'hilly_ne' },
  { nameEn: 'Assam', nameHi: 'असम', region: 'hilly_ne' },
  { nameEn: 'Bihar', nameHi: 'बिहार', region: 'north_plains' },
  { nameEn: 'Chandigarh', nameHi: 'चंडीगढ़', region: 'north_plains' },
  { nameEn: 'Chhattisgarh', nameHi: 'छत्तीसगढ़', region: 'western_central' },
  { nameEn: 'Dadra and Nagar Haveli and Daman and Diu', nameHi: 'दादरा नगर हवेली एवं दमन दीव', region: 'western_central' },
  { nameEn: 'Delhi (NCR)', nameHi: 'दिल्ली (एनसीआर)', region: 'north_plains' },
  { nameEn: 'Goa', nameHi: 'गोवा', region: 'coastal_south' },
  { nameEn: 'Gujarat', nameHi: 'गुजरात', region: 'western_central' },
  { nameEn: 'Haryana', nameHi: 'हरियाणा', region: 'north_plains' },
  { nameEn: 'Himachal Pradesh', nameHi: 'हिमाचल प्रदेश', region: 'hilly_ne' },
  { nameEn: 'Jammu and Kashmir', nameHi: 'जम्मू और कश्मीर', region: 'hilly_ne' },
  { nameEn: 'Jharkhand', nameHi: 'झारखंड', region: 'eastern_states' },
  { nameEn: 'Karnataka', nameHi: 'कर्नाटक', region: 'coastal_south' },
  { nameEn: 'Kerala', nameHi: 'केरल', region: 'coastal_south' },
  { nameEn: 'Ladakh', nameHi: 'लद्दाख', region: 'hilly_ne' },
  { nameEn: 'Lakshadweep', nameHi: 'लक्षद्वीप', region: 'coastal_south' },
  { nameEn: 'Madhya Pradesh', nameHi: 'मध्य प्रदेश', region: 'western_central' },
  { nameEn: 'Maharashtra', nameHi: 'महाराष्ट्र', region: 'western_central' },
  { nameEn: 'Manipur', nameHi: 'मणिपुर', region: 'hilly_ne' },
  { nameEn: 'Meghalaya', nameHi: 'मेघालय', region: 'hilly_ne' },
  { nameEn: 'Mizoram', nameHi: 'मिजोरम', region: 'hilly_ne' },
  { nameEn: 'Nagaland', nameHi: 'नागालैंड', region: 'hilly_ne' },
  { nameEn: 'Odisha', nameHi: 'ओडिशा', region: 'eastern_states' },
  { nameEn: 'Puducherry', nameHi: 'पुडुचेरी', region: 'coastal_south' },
  { nameEn: 'Punjab', nameHi: 'पंजाब', region: 'north_plains' },
  { nameEn: 'Rajasthan', nameHi: 'राजस्थान', region: 'western_central' },
  { nameEn: 'Sikkim', nameHi: 'सिक्किम', region: 'hilly_ne' },
  { nameEn: 'Tamil Nadu', nameHi: 'तमिलनाडु', region: 'coastal_south' },
  { nameEn: 'Telangana', nameHi: 'तेलंगाना', region: 'coastal_south' },
  { nameEn: 'Tripura', nameHi: 'त्रिपुरा', region: 'hilly_ne' },
  { nameEn: 'Uttar Pradesh', nameHi: 'उत्तर प्रदेश', region: 'north_plains' },
  { nameEn: 'Uttarakhand', nameHi: 'उत्तराखंड', region: 'hilly_ne' },
  { nameEn: 'West Bengal', nameHi: 'पश्चिम बंगाल', region: 'eastern_states' },
  { nameEn: 'Other / Pan-India Standard', nameHi: 'अन्य / राष्ट्रीय मानक', region: 'national_avg' }
];

const getSeasonalFractions = (m, type, regionKey) => {
  const profile = REGION_PROFILES[regionKey] || REGION_PROFILES.north_plains;
  if (type === 'ULB') {
    if ([5, 6, 7].includes(m)) return [profile.wetBase + 0.05, profile.dryBase - 0.02, 0.04, 0.02, 0.05, 0.12];
    if ([8, 9].includes(m)) return [profile.wetBase + 0.03, profile.dryBase - 0.01, 0.04, 0.02, 0.05, 0.13];
    return [profile.wetBase, profile.dryBase, 0.04, 0.02, 0.05, 0.15];
  }
  return [5, 6, 7].includes(m) ? [0.25, 0.15, 0.20, 0.22, 0.08, 0.10] : [0.20, 0.15, 0.25, 0.20, 0.10, 0.10];
};

// Deterministic PRNG helpers for reproducible datasets
const cyrb128 = (str) => {
  let h1 = 1779033703, h2 = 3144134277, h3 = 1013904242, h4 = 2773480762;
  for (let i = 0; i < str.length; i++) {
    let k = str.charCodeAt(i);
    h1 = h2 ^ Math.imul(h1 ^ k, 597399067);
    h2 = h3 ^ Math.imul(h2 ^ k, 2869860233);
    h3 = h4 ^ Math.imul(h3 ^ k, 951274213);
    h4 = h1 ^ Math.imul(h4 ^ k, 2716044179);
  }
  return (Math.imul(h3 ^ (h1 >>> 18), 597399067) ^ Math.imul(h4 ^ (h2 >>> 22), 2869860233) ^ Math.imul(h1 ^ (h3 >>> 17), 951274213) ^ Math.imul(h2 ^ (h4 >>> 19), 2716044179)) >>> 0;
};
const mulberry32 = (a) => {
  return function() {
    var t = a += 0x6D2B79F5;
    t = Math.imul(t ^ t >>> 15, t | 1);
    t ^= t + Math.imul(t ^ t >>> 7, t | 61);
    return ((t ^ t >>> 14) >>> 0) / 4294967296;
  }
};

const inputStyle = { width: '100%', padding: '9px', borderRadius: '6px', border: '1px solid #cbd5e1', marginTop: '4px', boxSizing: 'border-box' };

export default function App() {
  const [lang, setLang] = useState('hi');
  const [selectedState, setSelectedState] = useState('Uttar Pradesh');
  const [facilityType, setFacilityType] = useState('ULB');
  const [name, setName] = useState('Nagar Palika Parishad');
  
  // ULB specific state variables
  const [ulbCalculationMode, setUlbCalculationMode] = useState('population');
  const [population, setPopulation] = useState(150000);
  const [perCapitaOption, setPerCapitaOption] = useState('450');
  const [customPerCapita, setCustomPerCapita] = useState('');
  const [actualAverageTpd, setActualAverageTpd] = useState(35);
  const [referencePeriod, setReferencePeriod] = useState(30);

  // MRF specific state variables
  const [mrfDailyDryTons, setMrfDailyDryTons] = useState(15);
  const [mrfMaxCapacityTons, setMrfMaxCapacityTons] = useState(25);
  
  const [startYear, setStartYear] = useState(2026);
  const [selectedMonths, setSelectedMonths] = useState([1, 12]);
  const [displayUnit, setDisplayUnit] = useState('Tons');
  
  const [generatedMonthlyData, setGeneratedMonthlyData] = useState(null);
  const [activeTabMonth, setActiveTabMonth] = useState(null);
  const [isPaid, setIsPaid] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [showStateInfo, setShowStateInfo] = useState(false);

  const resultsRef = useRef(null);

  const currentStateObj = STATES_LIST.find(s => s.nameEn === selectedState) || STATES_LIST[33];
  const currentRegionKey = currentStateObj.region;
  const currentRegionObj = REGION_PROFILES[currentRegionKey];

  // Dynamic values for display calculations
  const parsedPerCapita = perCapitaOption === 'custom' ? Number(customPerCapita) : Number(perCapitaOption);
  const estimatedDailyWaste = ((Number(population) * parsedPerCapita) / 1000000).toFixed(2);
  const mrfCapacityUtilization = Number(mrfMaxCapacityTons) > 0 ? (Number(mrfDailyDryTons) / Number(mrfMaxCapacityTons)) * 100 : 0;

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
      // Accurate Days in Month calculation mapping leap years properly
      const days = new Date(startYear, m, 0).getDate();
      
      let targetTons = 0;
      if (facilityType === 'ULB') {
        if (ulbCalculationMode === 'population') {
          targetTons = (Number(population) * parsedPerCapita) / 1000000;
        } else {
          targetTons = Number(actualAverageTpd);
        }
      } else {
        targetTons = Number(mrfDailyDryTons);
      }

      const baseFractions = getSeasonalFractions(m, facilityType, currentRegionKey);
      
      // Seed deterministic generator based on inputs to ensure reproducibility
      const seedString = `${facilityType}-${selectedState}-${name}-${startYear}-${m}-${ulbCalculationMode}-${targetTons}`;
      const random = mulberry32(cyrb128(seedString));

      let logs = [];

      for (let day = 1; day <= days; day++) {
        // String conversion preserving exact YYYY-MM-DD avoiding UTC date shifts
        const dateStr = `${startYear}-${String(m).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        
        // For day names, we can safely use local date parsing
        const dateObj = new Date(startYear, m - 1, day);
        const dayName = dateObj.toLocaleDateString('en-US', { weekday: 'short' });
        const isWeekend = [0, 6].includes(dateObj.getDay());

        // Apply noise via seeded random
        let noise = 0.95 + random() * 0.10;
        if (isWeekend) noise *= 1.05;
        const dailyTotal = targetTons * noise;

        let raw = baseFractions.map(r => r * (0.88 + random() * 0.24));
        let sum = raw.reduce((a, b) => a + b, 0);
        let norm = raw.map(r => r / sum);

        // Calculate and round individual categories first
        let c1 = Number((dailyTotal * norm[0]).toFixed(2));
        let c2 = Number((dailyTotal * norm[1]).toFixed(2));
        let c3 = Number((dailyTotal * norm[2]).toFixed(2));
        let c4 = Number((dailyTotal * norm[3]).toFixed(2));
        let c5 = Number((dailyTotal * norm[4]).toFixed(2));
        let c6 = Number((dailyTotal * norm[5]).toFixed(2));
        
        // Ensure standard category sum aligns exactly with final daily total
        let exactTotal = Number((c1 + c2 + c3 + c4 + c5 + c6).toFixed(2));

        logs.push({
          date: dateStr,
          dayName: dayName,
          c1: c1,
          c2: c2,
          c3: c3,
          c4: c4,
          c5: c5,
          c6: c6,
          total: exactTotal
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
        if (document.querySelector('script[src*="checkout.razorpay.com"]')) return res(true);
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
        name: 'SWM Logbook Engine',
        description: `${facilityType} (${selectedMonths.length} Months) - ${name}`,
        order_id: order.id,
        handler: () => { 
          // NOTE: Server-side Razorpay signature verification is required here before production deployment.
          setIsPaid(true); 
          setIsProcessing(false); 
          downloadMultiSheetExcel(); 
        },
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
    XLSX.writeFile(wb, `${name.replace(/\s+/g, '_')}_SWM_${selectedMonths.length}M.xlsx`);
  };

  const activeRows = generatedMonthlyData?.[activeTabMonth] || [];
  const visibleRows = isPaid ? activeRows : activeRows.slice(0, 5);
  const activeMonthObj = MONTHS.find(m => m.id === activeTabMonth);

  return (
    <div style={{ fontFamily: 'sans-serif', background: '#f8fafc', minHeight: '100vh', display: 'flex', flexDirection: 'column', justifyContent: 'space-between' }}>
      
      <div>
        {/* HEADER BANNER */}
        <div style={{ background: 'linear-gradient(135deg, #064e3b 0%, #047857 100%)', color: '#fff', padding: '20px 15px' }}>
          <div style={{ maxWidth: '1000px', margin: '0 auto', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
            <div>
              <span style={{ background: 'rgba(255,255,255,0.2)', padding: '3px 8px', borderRadius: '12px', fontSize: '11px', fontWeight: 'bold' }}>
                <ShieldCheck size={12} style={{ verticalAlign: 'middle' }} /> SWM-ALIGNED ESTIMATION TOOL
              </span>
              <h1 style={{ fontSize: '22px', margin: '6px 0 2px 0', fontWeight: '800' }}>
                <Building2 size={22} style={{ verticalAlign: 'middle', marginRight: '6px' }} />
                {lang === 'hi' ? 'यूएलबी एवं एमआरएफ लोगबुक जनरेटर' : 'ULB & MRF Waste Logbook Generator'}
              </h1>
              <p style={{ fontSize: '13px', margin: 0, color: '#a7f3d0' }}>
                {lang === 'hi' ? '4-स्ट्रीम अपशिष्ट पृथक्कीकरण एवं अखिल भारतीय राज्यवार एडजस्टमेंट टूल' : 'Automated 4-Stream Logbook Engine with Pan-India State-Wise Calibration'}
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
            
            <div style={{ marginBottom: '14px', display: 'flex', gap: '15px', alignItems: 'center', fontSize: '14px', flexWrap: 'wrap' }}>
              <strong>{lang === 'hi' ? 'सुविधा प्रकार:' : 'Facility:'}</strong>
              <label><input type="radio" value="ULB" checked={facilityType === 'ULB'} onChange={() => setFacilityType('ULB')} /> ULB</label>
              <label><input type="radio" value="MRF" checked={facilityType === 'MRF'} onChange={() => setFacilityType('MRF')} /> MRF Centre</label>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '12px', marginBottom: '14px' }}>
              
              <div>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                  <label style={{ fontSize: '12px', fontWeight: '600', color: '#047857', display: 'flex', alignItems: 'center', gap: '4px' }}>
                    <MapPin size={13} /> {lang === 'hi' ? 'राज्य चुनें (State)' : 'Select State'}
                  </label>
                  <button 
                    type="button" 
                    onClick={() => setShowStateInfo(!showStateInfo)} 
                    style={{ background: 'none', border: 'none', color: '#0284c7', cursor: 'pointer', padding: '0 2px', display: 'flex', alignItems: 'center', gap: '2px', fontSize: '12px', fontWeight: 'bold' }}
                  >
                    <Info size={14} />
                  </button>
                </div>

                <select style={{ ...inputStyle, border: '1px solid #059669', background: '#f0fdf4' }} value={selectedState} onChange={(e) => setSelectedState(e.target.value)}>
                  {STATES_LIST.map((s) => (
                    <option key={s.nameEn} value={s.nameEn}>
                      {lang === 'hi' ? s.nameHi : s.nameEn}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'निकाय / एमआरएफ का नाम' : 'ULB / MRF Name'}</label>
                <input style={inputStyle} type="text" required value={name} onChange={(e) => setName(e.target.value)} />
              </div>

              {facilityType === 'ULB' ? (
                <>
                  <div style={{ gridColumn: '1 / -1', background: '#f1f5f9', padding: '10px', borderRadius: '6px', border: '1px solid #e2e8f0', marginTop: '6px' }}>
                    <strong style={{ fontSize: '13px' }}>{lang === 'hi' ? 'ULB गणना का आधार' : 'ULB Waste Calculation Basis'}</strong>
                    <div style={{ display: 'flex', gap: '15px', marginTop: '8px', fontSize: '13px', flexWrap: 'wrap' }}>
                      <label style={{ cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '4px' }}>
                        <input type="radio" checked={ulbCalculationMode === 'population'} onChange={() => setUlbCalculationMode('population')} /> 
                        {lang === 'hi' ? 'जनसंख्या आधारित अनुमान' : 'Population-based estimation'}
                      </label>
                      <label style={{ cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '4px' }}>
                        <input type="radio" checked={ulbCalculationMode === 'actual'} onChange={() => setUlbCalculationMode('actual')} /> 
                        {lang === 'hi' ? 'वास्तविक / देखा गया औसत (TPD)' : 'Actual / Observed average waste generation'}
                      </label>
                    </div>
                  </div>

                  {ulbCalculationMode === 'population' ? (
                    <>
                      <div>
                        <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'वर्तमान अनुमानित जनसंख्या' : 'Current Estimated Population'}</label>
                        <input style={inputStyle} type="number" required={ulbCalculationMode === 'population'} value={population} onChange={(e) => setPopulation(e.target.value)} />
                      </div>
                      <div>
                        <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'सांकेतिक प्रतिव्यक्ति दर' : 'Indicative per-capita generation rate'}</label>
                        <select style={inputStyle} value={perCapitaOption} onChange={(e) => setPerCapitaOption(e.target.value)}>
                          <option value="300">300 g/person/day</option>
                          <option value="350">350 g/person/day</option>
                          <option value="400">400 g/person/day</option>
                          <option value="450">450 g/person/day</option>
                          <option value="500">500 g/person/day</option>
                          <option value="custom">{lang === 'hi' ? 'कस्टम (Custom)' : 'Custom'}</option>
                        </select>
                        {perCapitaOption === 'custom' && (
                          <input style={{ ...inputStyle, marginTop: '8px' }} type="number" placeholder="g/person/day" required value={customPerCapita} onChange={(e) => setCustomPerCapita(e.target.value)} />
                        )}
                      </div>
                      <div style={{ gridColumn: '1 / -1', fontSize: '12px', color: '#475569', fontStyle: 'italic' }}>
                        Select a rate appropriate to local conditions. Where reliable local waste-generation records are available, the Actual / Observed TPD option should be preferred.
                      </div>
                      <div style={{ gridColumn: '1 / -1', fontSize: '13px', fontWeight: 'bold', color: '#047857' }}>
                        {lang === 'hi' ? 'अनुमानित दैनिक अपशिष्ट:' : 'Estimated Daily Waste:'} {estimatedDailyWaste} TPD
                      </div>
                    </>
                  ) : (
                    <>
                      <div>
                        <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'औसत वास्तविक/देखा गया अपशिष्ट (TPD)' : 'Average Actual / Observed Waste Generation (TPD)'}</label>
                        <input style={inputStyle} type="number" step="0.01" required={ulbCalculationMode === 'actual'} value={actualAverageTpd} onChange={(e) => setActualAverageTpd(e.target.value)} />
                      </div>
                      <div>
                        <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'संदर्भ अवधि (दिन)' : 'Reference Period (days)'}</label>
                        <input style={inputStyle} type="number" required={ulbCalculationMode === 'actual'} value={referencePeriod} onChange={(e) => setReferencePeriod(e.target.value)} />
                      </div>
                    </>
                  )}
                </>
              ) : (
                <>
                  <div>
                    <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'दैनिक सूखा कचरा आवक (TPD)' : 'Daily Dry Waste Input (TPD)'}</label>
                    <input style={inputStyle} type="number" step="0.01" required value={mrfDailyDryTons} onChange={(e) => setMrfDailyDryTons(e.target.value)} />
                  </div>
                  <div>
                    <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'एमआरएफ क्षमता (TPD)' : 'MRF Capacity (TPD)'}</label>
                    <input style={inputStyle} type="number" step="0.01" required value={mrfMaxCapacityTons} onChange={(e) => setMrfMaxCapacityTons(e.target.value)} />
                  </div>
                  <div style={{ gridColumn: '1 / -1', fontSize: '13px', fontWeight: 'bold', display: 'flex', alignItems: 'center', flexWrap: 'wrap' }}>
                    {lang === 'hi' ? 'क्षमता उपयोग (Capacity Utilization):' : 'Capacity Utilization:'} {mrfCapacityUtilization.toFixed(1)}%
                    {mrfCapacityUtilization > 100 && (
                      <span style={{ color: '#dc2626', marginLeft: '6px', fontSize: '12px' }}>
                        (Warning: Daily input exceeds MRF design capacity.)
                      </span>
                    )}
                  </div>
                </>
              )}

              <div>
                <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'वर्ष' : 'Year'}</label>
                <input style={inputStyle} type="number" required value={startYear} onChange={(e) => setStartYear(Number(e.target.value))} />
              </div>
            </div>

            {showStateInfo && (
              <div style={{ background: '#f0f9ff', border: '1px solid #7dd3fc', padding: '10px 14px', borderRadius: '6px', marginBottom: '14px', fontSize: '12px', color: '#0369a1', lineHeight: '1.4' }}>
                <div style={{ fontWeight: 'bold', display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px' }}>
                  <span>ℹ️ {lang === 'hi' ? 'राज्य चयन का महत्व:' : 'Relevance of Selecting Your State:'}</span>
                  <button type="button" onClick={() => setShowStateInfo(false)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#0369a1', fontWeight: 'bold' }}>✕</button>
                </div>
                <p style={{ margin: '0 0 4px 0' }}>
                  {lang === 'hi'
                    ? `आपके द्वारा चुने गए राज्य (${currentStateObj.nameHi}) को ${currentRegionObj.nameHi} क्षेत्र में वर्गीकृत किया गया है।`
                    : `Selected state (${currentStateObj.nameEn}) is mapped to the ${currentRegionObj.nameEn} zone.`}
                </p>
                <ul style={{ margin: 0, paddingLeft: '18px' }}>
                  <li>{lang === 'hi' ? 'तटीय/दक्षिणी राज्यों में नमी एवं जैविक कचरा अधिक रहता है (~62% Wet Base)।' : 'Coastal & Southern states experience higher humidity and organic fraction (~62% Wet Base).'}</li>
                  <li>{lang === 'hi' ? 'मेट्रो/पश्चिमी राज्यों में प्लास्टिक व पैकेजिंग अधिक होती है (~26% Dry Base)।' : 'Metro & Western states produce higher packaging/plastic waste (~26% Dry Base).'}</li>
                  <li>{lang === 'hi' ? 'उत्तरी व पूर्वोत्तर क्षेत्रों में मौसमी तापमान और फलों के उत्पादन के अनुसार गीला कचरा बदलता है।' : 'Northern & NE states scale wet waste dynamically with seasonal fruit harvests and temperature.'}</li>
                </ul>
              </div>
            )}

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
              {lang === 'hi' ? 'लोगबुक जनरेट करें →' : 'Generate Dataset →'}
            </button>
          </form>

          {generatedMonthlyData && (
            <div ref={resultsRef} style={{ background: '#fff', padding: '15px', borderRadius: '8px', border: '1px solid #cbd5e1', scrollMarginTop: '15px' }}>
              
              <div style={{ background: '#eff6ff', border: '1px solid #bfdbfe', padding: '10px', borderRadius: '6px', marginBottom: '12px', fontSize: '12px', color: '#1e40af' }}>
                <strong><Info size={14} style={{ verticalAlign: 'middle' }} /> Indicative regional modelling assumptions ({currentStateObj.nameEn} / {currentRegionObj.nameEn}):</strong> 4-Stream segregation applied. Baseline waste composition calibrated for chosen state profile.
              </div>

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
                <span style={{ fontSize: '13px', fontWeight: 'bold' }}>{name} ({currentStateObj.nameEn}) — {activeMonthObj?.fullEn} {startYear}</span>
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

      <div style={{ marginTop: '30px', padding: '15px 10px', borderTop: '1px solid #cbd5e1', textAlign: 'center', fontSize: '11px', color: '#64748b', lineHeight: '1.5', background: '#ffffff' }}>
        <p style={{ margin: '0 0 6px 0' }}>
          <strong>Disclaimer:</strong> This web tool is developed strictly for educational, research, and estimation purposes. Output datasets serve as decision-support models for solid waste management planning.
        </p>
        <p style={{ margin: 0 }}>
          Copyright © 2026 CRF | Engineered & Maintained by <strong>Team CRF</strong> — <a href="https://www.consilienceresearch.in/" target="_blank" rel="noopener noreferrer" style={{ color: '#059669', fontWeight: 'bold', textDecoration: 'none' }}>Consilience Research Foundation</a>, an Urban & Infrastructure Research Consultancy Institute.
        </p>
      </div>

    </div>
  );
}
