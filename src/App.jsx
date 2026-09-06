import React, { useState, useRef, useEffect } from 'react';
import { Building2, Download, Lock, Globe, ShieldCheck, Plus, Trash2, ArrowLeft, Settings2, FileSpreadsheet, BookOpen, CheckCircle2, AlertTriangle, Zap } from 'lucide-react';
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

const STATES_LIST = [
  { nameEn: 'Andaman & Nicobar Islands', nameHi: 'अंडमान और निकोबार' },
  { nameEn: 'Andhra Pradesh', nameHi: 'आंध्र प्रदेश' },
  { nameEn: 'Arunachal Pradesh', nameHi: 'अरुणाचल प्रदेश' },
  { nameEn: 'Assam', nameHi: 'असम' },
  { nameEn: 'Bihar', nameHi: 'बिहार' },
  { nameEn: 'Chandigarh', nameHi: 'चंडीगढ़' },
  { nameEn: 'Chhattisgarh', nameHi: 'छत्तीसगढ़' },
  { nameEn: 'Delhi (NCR)', nameHi: 'दिल्ली (एनसीआर)' },
  { nameEn: 'Goa', nameHi: 'गोवा' },
  { nameEn: 'Gujarat', nameHi: 'गुजरात' },
  { nameEn: 'Haryana', nameHi: 'हरियाणा' },
  { nameEn: 'Himachal Pradesh', nameHi: 'हिमाचल प्रदेश' },
  { nameEn: 'Jammu and Kashmir', nameHi: 'जम्मू और कश्मीर' },
  { nameEn: 'Jharkhand', nameHi: 'झारखंड' },
  { nameEn: 'Karnataka', nameHi: 'कर्नाटक' },
  { nameEn: 'Kerala', nameHi: 'केरल' },
  { nameEn: 'Madhya Pradesh', nameHi: 'मध्य प्रदेश' },
  { nameEn: 'Maharashtra', nameHi: 'महाराष्ट्र' },
  { nameEn: 'Odisha', nameHi: 'ओडिशा' },
  { nameEn: 'Punjab', nameHi: 'पंजाब' },
  { nameEn: 'Rajasthan', nameHi: 'राजस्थान' },
  { nameEn: 'Tamil Nadu', nameHi: 'तमिलनाडु' },
  { nameEn: 'Telangana', nameHi: 'तेलंगाना' },
  { nameEn: 'Uttar Pradesh', nameHi: 'उत्तर प्रदेश' },
  { nameEn: 'Uttarakhand', nameHi: 'उत्तराखंड' },
  { nameEn: 'West Bengal', nameHi: 'पश्चिम बंगाल' },
  { nameEn: 'Other / Pan-India Standard', nameHi: 'अन्य / राष्ट्रीय मानक' }
];

const DEFAULT_MRF_CONFIG = [
  { id: 'pet', label: 'PET Bottles', userWeight: 25 },
  { id: 'hdpe', label: 'HDPE / Plastics', userWeight: 20 },
  { id: 'paper', label: 'Cardboard & Paper', userWeight: 25 },
  { id: 'rdf', label: 'Combustible RDF', userWeight: 20 },
  { id: 'rejects', label: 'Inert Rejects', userWeight: 10 }
];

const DEFAULT_MIXED_CONFIG = [
  { id: 'fines', label: 'Organic Fines', userWeight: 45 },
  { id: 'coarse_rdf', label: 'Coarse RDF', userWeight: 35 },
  { id: 'heavy_inerts', label: 'Heavy Inerts', userWeight: 20 }
];

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
  };
};

const inputStyle = { width: '100%', padding: '9px', borderRadius: '6px', border: '1px solid #cbd5e1', marginTop: '4px', boxSizing: 'border-box' };

export default function App() {
  const [lang, setLang] = useState('hi');
  const [facilityType, setFacilityType] = useState('ULB'); // 'ULB', 'MRF', 'MIXED_PLANT'
  const [selectedState, setSelectedState] = useState('Uttar Pradesh');
  const [name, setName] = useState('Nagar Palika Parishad');
  const [phone, setPhone] = useState('');
  
  // ULB State
  const [ulbCalculationMode, setUlbCalculationMode] = useState('population');
  const [population, setPopulation] = useState(50000);
  const [perCapitaOption, setPerCapitaOption] = useState('450');
  const [actualAverageTpd, setActualAverageTpd] = useState(22.5);
  
  // MRF / Mixed State
  const [mrfDailyDryTons, setMrfDailyDryTons] = useState(10);
  const [mrfMaxCapacityTons, setMrfMaxCapacityTons] = useState(15);
  const [isAdvancedMode, setIsAdvancedMode] = useState(false);
  const [mrfStreamConfig, setMrfStreamConfig] = useState(DEFAULT_MRF_CONFIG);
  const [mixedStreamConfig, setMixedStreamConfig] = useState(DEFAULT_MIXED_CONFIG);
  
  // Global Settings
  const [startYear, setStartYear] = useState(2026);
  const [selectedMonths, setSelectedMonths] = useState([1]);
  const [displayUnit, setDisplayUnit] = useState('Tons');
  
  // App State
  const [generatedMonthlyData, setGeneratedMonthlyData] = useState(null);
  const [activeTabMonth, setActiveTabMonth] = useState(null);
  const [isPaid, setIsPaid] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [activePolicyModal, setActivePolicyModal] = useState(null);

  const resultsRef = useRef(null);

  const currentStreamConfig = facilityType === 'MRF' ? mrfStreamConfig : mixedStreamConfig;
  const setStreamConfig = facilityType === 'MRF' ? setMrfStreamConfig : setMixedStreamConfig;

  const totalPercentage = currentStreamConfig.reduce((sum, item) => sum + (Number(item.userWeight) || 0), 0);
  const isValidTotal = totalPercentage === 100;
  const generateDisabled = isAdvancedMode && !isValidTotal;

  const parsedPerCapita = Number(perCapitaOption);
  const calculatedTpdDisplay = ((Number(population) * parsedPerCapita) / 1000000).toFixed(2);
  const targetTotalTpd = ulbCalculationMode === 'population' ? Number(calculatedTpdDisplay) : Number(actualAverageTpd || 0);

  const updateStreamConfig = (id, field, value) => {
    setStreamConfig(currentStreamConfig.map(s => s.id === id ? { ...s, [field]: field === 'userWeight' ? Number(value) : value } : s));
  };

  const addCustomStream = () => {
    setStreamConfig([...currentStreamConfig, { id: `custom_${Date.now()}`, label: 'New Fraction', userWeight: 0 }]);
  };

  const removeCustomStream = (id) => {
    setStreamConfig(currentStreamConfig.filter(s => s.id !== id));
  };

  const getSessionKey = () => `crf_paid_standalone_${facilityType}_${name.trim().toLowerCase().replace(/\s+/g, '_')}_${selectedMonths.join('_')}_${startYear}`;

  useEffect(() => {
    const rawData = localStorage.getItem(getSessionKey());
    if (rawData) {
      try {
        const parsed = JSON.parse(rawData);
        if (parsed.paid && (Date.now() - parsed.timestamp < 12 * 60 * 60 * 1000)) {
          setIsPaid(true); return;
        }
      } catch (e) {
        if (rawData === 'true') { setIsPaid(true); return; }
      }
    }
    setIsPaid(false);
  }, [name, selectedMonths, startYear, facilityType]);

  const toggleMonth = (mId) => {
    if (selectedMonths.includes(mId)) {
      if (selectedMonths.length > 1) setSelectedMonths(selectedMonths.filter(m => m !== mId));
    } else {
      setSelectedMonths([...selectedMonths, mId].sort((a, b) => a - b));
    }
  };

  const getPricingDetails = () => {
    const count = selectedMonths.length;
    const freeMonths = Math.floor(count / 6);
    const billableMonths = count - freeMonths;
    let baseRate = 100;
    if (facilityType === 'MRF') baseRate = 150;
    if (facilityType === 'MIXED_PLANT') baseRate = 200;
    
    const baseTotal = billableMonths * baseRate;
    const finalTotalWithCharges = Math.round(baseTotal / (1 - 0.0236));
    return { count, freeMonths, billableMonths, baseTotal, total: finalTotalWithCharges };
  };

  const pricing = getPricingDetails();

  const handleGenerate = (e) => {
    e.preventDefault();
    if (isAdvancedMode && !isValidTotal) {
      alert(`Fractions must equal 100%. Currently at ${totalPercentage}%.`);
      return;
    }

    let monthlyDataMap = {};

    selectedMonths.forEach((m) => {
      const days = new Date(startYear, m, 0).getDate();
      const baseVal = facilityType === 'ULB' ? targetTotalTpd : Number(mrfDailyDryTons || 10);
      const seedString = `STANDALONE-V2-${facilityType}-${selectedState}-${name}-${startYear}-${m}-${baseVal}`;
      const random = mulberry32(cyrb128(seedString));
      
      let seasonalMultiplier = 1.0;
      if ([7, 8, 9].includes(m)) seasonalMultiplier = 1.05; 
      else if ([4, 5, 6].includes(m)) seasonalMultiplier = 0.95; 

      let logs = [];

      for (let day = 1; day <= days; day++) {
        const dateStr = `${startYear}-${String(m).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const dayName = new Date(startYear, m - 1, day).toLocaleDateString('en-US', { weekday: 'short' });

        let noise = 0.95 + random() * 0.10;
        const dailyIntake = Number((baseVal * noise * seasonalMultiplier).toFixed(3));
        
        let rowData = { date: dateStr, dayName, intake: dailyIntake };

        if (facilityType === 'ULB') {
          const mswRatio = 0.85 + random() * 0.10; // 85-95% is MSW
          const cndRatio = 0.02 + random() * 0.03; // 2-5% C&D
          const drainRatio = 0.01 + random() * 0.02; // 1-3% Drain Silt

          const domestic = Number((dailyIntake * mswRatio).toFixed(3));
          const commercial = Number((dailyIntake * (0.98 - mswRatio - cndRatio - drainRatio)).toFixed(3));
          const cnd = Number((dailyIntake * cndRatio).toFixed(3));
          const drain = Number((dailyIntake * drainRatio).toFixed(3));

          // Ensuring perfect 100% balance
          rowData = { ...rowData, domestic, commercial, cnd, drain: Number((dailyIntake - domestic - commercial - cnd).toFixed(3)) };
        
        } else if (facilityType === 'MRF' || facilityType === 'MIXED_PLANT') {
          let accumulated = 0;
          let breakdown = {};
          
          currentStreamConfig.forEach((frac, index) => {
            if (index === currentStreamConfig.length - 1) {
              breakdown[frac.id] = Number(Math.max(0, dailyIntake - accumulated).toFixed(3));
            } else {
              const fracNoise = 0.85 + random() * 0.30;
              const noisyWeight = Number((dailyIntake * (frac.userWeight / 100) * fracNoise).toFixed(3));
              breakdown[frac.id] = noisyWeight;
              accumulated += noisyWeight;
            }
          });
          rowData.fractions = breakdown;
        }

        logs.push(rowData);
      }
      monthlyDataMap[m] = logs;
    });

    const rawData = localStorage.getItem(getSessionKey());
    let verifiedPaid = false;
    if (rawData) {
      try {
        const parsed = JSON.parse(rawData);
        if (parsed.paid && (Date.now() - parsed.timestamp < 12 * 60 * 60 * 1000)) verifiedPaid = true;
      } catch (e) { if (rawData === 'true') verifiedPaid = true; }
    }
    setIsPaid(verifiedPaid);

    setGeneratedMonthlyData(monthlyDataMap);
    setActiveTabMonth(selectedMonths[0]);
    setTimeout(() => resultsRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 100);
  };

  const handlePayment = async () => {
    if (!phone || phone.length < 10) {
      alert(lang === 'hi' ? 'कृपया एक वैध 10-अंकों का मोबाइल नंबर दर्ज करें।' : 'Please enter a valid 10-digit mobile number.');
      return;
    }
    setIsProcessing(true);

    if (!window.Cashfree) {
      await new Promise((res) => {
        if (document.querySelector('script[src*="cashfree.com"]')) return res(true);
        const s = document.createElement('script');
        s.src = 'https://sdk.cashfree.com/js/v3/cashfree.js';
        s.onload = () => res(true);
        document.body.appendChild(s);
      });
    }

    try {
      const isProd = import.meta.env.VITE_CASHFREE_MODE === 'production';
      const res = await fetch('/api/create-order', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ amount: pricing.total, customerName: name, customerPhone: phone })
      });

      // Fix for Cashfree HTML error response
      const rawText = await res.text();
      let order;
      try {
        order = JSON.parse(rawText);
      } catch (err) {
        throw new Error(`Server returned non-JSON response. Check Vercel Logs. Response preview: ${rawText.substring(0, 50)}`);
      }

      if (!order.payment_session_id) throw new Error(order.message || 'Failed to initialize payment session.');

      const cashfree = window.Cashfree({ mode: isProd ? 'production' : 'sandbox' });

      cashfree.checkout({
        paymentSessionId: order.payment_session_id,
        redirectTarget: '_modal'
      }).then((result) => {
        if (result.error) {
          alert('Payment Failed: ' + result.error.message);
          setIsProcessing(false);
        } else if (result.paymentDetails) {
          setIsPaid(true);
          setIsProcessing(false);
          localStorage.setItem(getSessionKey(), JSON.stringify({ paid: true, timestamp: Date.now() }));
          downloadExcel();
        }
      });
    } catch (err) {
      alert('Payment Error: ' + err.message);
      setIsProcessing(false);
    }
  };

  const formatVal = (v) => displayUnit === 'kg' ? Math.round(Number(v || 0) * 1000) : Number(v || 0).toFixed(3);

  const downloadExcel = () => {
    if (!generatedMonthlyData) return;
    if (!isPaid) return alert('Payment verified flag missing. Please complete checkout to download full sheets.');

    try {
      const u = displayUnit === 'kg' ? 'kg' : 'Tons';

      selectedMonths.forEach((mId) => {
        const wb = XLSX.utils.book_new();
        const monthData = MONTHS.find(m => m.id === mId);
        const monthName = monthData?.shortEn || `M${mId}`;
        const fullMonthName = monthData?.fullEn || `Month${mId}`;

        let headers = [];
        let rows = [];

        if (facilityType === 'ULB') {
          headers = ["Date", "Day", `Total Collection (${u})`, `Domestic Waste (${u})`, `Commercial Waste (${u})`, `C&D Waste (${u})`, `Drain Silt (${u})`];
          rows = generatedMonthlyData[mId].map(r => [
            r.date, r.dayName, formatVal(r.intake), formatVal(r.domestic), formatVal(r.commercial), formatVal(r.cnd), formatVal(r.drain)
          ]);
        } else {
          headers = ["Date", "Day", `Total Intake (${u})`, ...currentStreamConfig.map(frac => `${frac.label} (${frac.userWeight}%)`)];
          rows = generatedMonthlyData[mId].map(r => [
            r.date, r.dayName, formatVal(r.intake), ...currentStreamConfig.map(frac => formatVal(r.fractions[frac.id]))
          ]);
        }

        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([headers, ...rows]), `${monthName}_Logbook`);
        
        const safeFileName = `Standalone_${facilityType}_Logbook_${name.replace(/\s+/g, '_')}_${fullMonthName}_${startYear}.xlsx`;
        XLSX.writeFile(wb, safeFileName);
      });

    } catch (err) {
      alert('Excel Generation Error: ' + err.message);
    }
  };

  const activeRows = generatedMonthlyData?.[activeTabMonth] || [];
  const visibleRows = isPaid ? activeRows : activeRows.slice(0, 5);

  return (
    <div style={{ fontFamily: 'sans-serif', background: '#f8fafc', minHeight: '100vh', padding: '15px' }}>
      <div style={{ maxWidth: '1000px', margin: '0 auto' }}>
        
        {/* HEADER */}
        <div style={{ background: 'linear-gradient(135deg, #0f172a 0%, #334155 100%)', color: '#fff', padding: '20px', borderRadius: '8px', marginBottom: '16px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap' }}>
            <div>
              <span style={{ background: 'rgba(255,255,255,0.2)', padding: '3px 8px', borderRadius: '12px', fontSize: '11px', fontWeight: 'bold', display: 'inline-flex', alignItems: 'center', gap: '4px' }}>
                <ShieldCheck size={12} /> STANDALONE LOGBOOK GENERATOR
              </span>
              <h1 style={{ fontSize: '22px', margin: '6px 0 2px 0', fontWeight: '800' }}>
                <Building2 size={22} style={{ verticalAlign: 'middle', marginRight: '6px' }} />
                {lang === 'hi' ? 'सिंगल-फैसिलिटी SWM लॉग-बुक टूल' : 'Standalone SWM Logbook Tool'}
              </h1>
            </div>
            <button type="button" onClick={() => setLang(lang === 'hi' ? 'en' : 'hi')} style={{ padding: '6px 12px', background: '#fff', color: '#0f172a', border: 'none', borderRadius: '5px', fontWeight: 'bold', cursor: 'pointer' }}>
              <Globe size={15} style={{ verticalAlign: 'middle' }} /> {lang === 'hi' ? 'English' : 'हिंदी'}
            </button>
          </div>
        </div>

        {/* BILINGUAL USER GUIDE CONTAINER */}
        <div style={{ background: '#f0f9ff', border: '1px solid #bae6fd', padding: '16px', borderRadius: '8px', marginBottom: '20px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px', borderBottom: '1px solid #cbd5e1', paddingBottom: '8px' }}>
            <h3 style={{ margin: 0, color: '#0369a1', fontSize: '15px', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px' }}>
              <BookOpen size={18} /> {lang === 'hi' ? 'लॉग-बुक जनरेटर उपयोग निर्देशिका (User Guide)' : 'User Guide & Operator Instructions'}
            </h3>
            <span style={{ fontSize: '11px', background: '#e0f2fe', color: '#0369a1', padding: '2px 8px', borderRadius: '12px', fontWeight: 'bold' }}>
              {lang === 'hi' ? 'सहायता गाइड' : 'Help Manual'}
            </span>
          </div>

          {lang === 'hi' ? (
            <div style={{ fontSize: '13px', color: '#0c4a6e', lineHeight: '1.6' }}>
              <p style={{ margin: '0 0 8px 0', fontWeight: 'bold', color: '#0284c7' }}>
                इस टूल का उपयोग करके 30-दिवसीय ऑडिट-रेडी SWM एक्सेल लॉग-बुक कैसे बनाएं:
              </p>
              <ol style={{ margin: 0, paddingLeft: '20px' }}>
                <li style={{ marginBottom: '6px' }}><strong>सुविधा चुनें:</strong> निकाय (ULB), MRF, या मिश्रित कचरा प्लांट में से चुनें।</li>
                <li style={{ marginBottom: '6px' }}><strong>विवरण दर्ज करें:</strong> राज्य, नाम, और मोबाइल नंबर भरें।</li>
                <li style={{ marginBottom: '6px' }}><strong>क्षमता सेट करें:</strong> TPD या जनसंख्या दर्ज करें। एडवांस्ड सेटिंग से स्ट्रीम प्रतिशत को कस्टमाइज़ करें (कुल 100% होना चाहिए)।</li>
                <li style={{ marginBottom: '6px' }}><strong>महीने चुनें:</strong> हर 6ठा महीना मुफ़्त है।</li>
                <li style={{ marginBottom: '6px' }}><strong>डाउनलोड करें:</strong> 5-दिन का प्रीव्यू देखें, फिर सुरक्षित भुगतान करके पूरी एक्सेल फ़ाइल डाउनलोड करें।</li>
              </ol>
            </div>
          ) : (
            <div style={{ fontSize: '13px', color: '#0c4a6e', lineHeight: '1.6' }}>
              <p style={{ margin: '0 0 8px 0', fontWeight: 'bold', color: '#0284c7' }}>
                How to generate your 30-day audit-ready SWM Excel logbooks step-by-step:
              </p>
              <ol style={{ margin: 0, paddingLeft: '20px' }}>
                <li style={{ marginBottom: '6px' }}><strong>Select Facility:</strong> Choose between ULB Collection, MRF, or Mixed Waste Plant.</li>
                <li style={{ marginBottom: '6px' }}><strong>Enter Details:</strong> Provide State, Name, and Mobile Number.</li>
                <li style={{ marginBottom: '6px' }}><strong>Set Capacity:</strong> Enter TPD or population. Use Advanced Settings to customize fraction percentages (must equal 100%).</li>
                <li style={{ marginBottom: '6px' }}><strong>Choose Months:</strong> Select months (every 6th month is free).</li>
                <li style={{ marginBottom: '6px' }}><strong>Download:</strong> Preview the first 5 days, then pay to download the full Excel workbook.</li>
              </ol>
            </div>
          )}
        </div>

        {/* MAIN FORM */}
        <form onSubmit={handleGenerate} style={{ background: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #cbd5e1', marginBottom: '20px' }}>
          
          <div style={{ marginBottom: '14px', display: 'flex', gap: '15px', alignItems: 'center', fontSize: '14px', flexWrap: 'wrap', background: '#f8fafc', padding: '10px', borderRadius: '6px', border: '1px solid #e2e8f0' }}>
            <strong>{lang === 'hi' ? 'सिंगल-फैसिलिटी चुनें:' : 'Select Standalone Logbook:'}</strong>
            <label style={{ cursor: 'pointer' }}><input type="radio" value="ULB" checked={facilityType === 'ULB'} onChange={() => { setFacilityType('ULB'); setGeneratedMonthlyData(null); }} /> {lang === 'hi' ? 'निकाय (ULB) (₹100)' : 'ULB Collection (₹100/mo)'}</label>
            <label style={{ cursor: 'pointer' }}><input type="radio" value="MRF" checked={facilityType === 'MRF'} onChange={() => { setFacilityType('MRF'); setGeneratedMonthlyData(null); }} /> {lang === 'hi' ? 'एमआरएफ (MRF) (₹150)' : 'MRF Centre (₹150/mo)'}</label>
            <label style={{ cursor: 'pointer' }}><input type="radio" value="MIXED_PLANT" checked={facilityType === 'MIXED_PLANT'} onChange={() => { setFacilityType('MIXED_PLANT'); setGeneratedMonthlyData(null); }} /> {lang === 'hi' ? 'मिश्रित कचरा (₹200)' : 'Mixed Waste Plant (₹200/mo)'}</label>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '12px', marginBottom: '14px' }}>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'राज्य चुनें' : 'Select State'}</label>
              <select style={inputStyle} value={selectedState} onChange={(e) => setSelectedState(e.target.value)}>
                {STATES_LIST.map((s) => <option key={s.nameEn} value={s.nameEn}>{lang === 'hi' ? s.nameHi : s.nameEn}</option>)}
              </select>
            </div>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'निकाय / प्लांट का नाम' : 'Facility Name'}</label>
              <input style={inputStyle} type="text" required value={name} onChange={(e) => setName(e.target.value)} />
            </div>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'मोबाइल नंबर' : 'Mobile Number'}</label>
              <input style={inputStyle} type="tel" maxLength={10} placeholder="9876543210" required value={phone} onChange={(e) => setPhone(e.target.value.replace(/\D/g, ''))} />
            </div>

            {facilityType === 'ULB' && (
              <>
                <div style={{ gridColumn: '1 / -1', background: '#f1f5f9', padding: '10px', borderRadius: '6px', border: '1px solid #e2e8f0' }}>
                  <strong style={{ fontSize: '13px' }}>{lang === 'hi' ? 'कचरा उत्पादन का आधार' : 'Estimation Basis'}</strong>
                  <div style={{ display: 'flex', gap: '15px', marginTop: '6px', fontSize: '13px' }}>
                    <label style={{ cursor: 'pointer' }}><input type="radio" checked={ulbCalculationMode === 'population'} onChange={() => setUlbCalculationMode('population')} /> {lang === 'hi' ? 'जनसंख्या आधारित' : 'Population Based'}</label>
                    <label style={{ cursor: 'pointer' }}><input type="radio" checked={ulbCalculationMode === 'actual'} onChange={() => setUlbCalculationMode('actual')} /> {lang === 'hi' ? 'वास्तविक TPD' : 'Actual TPD'}</label>
                  </div>
                </div>

                {ulbCalculationMode === 'population' ? (
                  <>
                    <div>
                      <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'जनसंख्या' : 'Population'}</label>
                      <input style={inputStyle} type="number" value={population} onChange={(e) => setPopulation(e.target.value)} />
                    </div>
                    <div>
                      <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'प्रति व्यक्ति दर' : 'Per Capita Rate'}</label>
                      <select style={inputStyle} value={perCapitaOption} onChange={(e) => setPerCapitaOption(e.target.value)}>
                        <option value="300">300 g/day</option>
                        <option value="450">450 g/day</option>
                        <option value="500">500 g/day</option>
                      </select>
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', justifyContent: 'center', background: '#ecfdf5', padding: '10px', borderRadius: '6px', border: '1px solid #a7f3d0' }}>
                      <span style={{ fontSize: '11px', color: '#065f46', fontWeight: 'bold' }}>{lang === 'hi' ? 'अनुमानित कचरा' : 'Calculated Waste'}</span>
                      <span style={{ fontSize: '18px', color: '#047857', fontWeight: '900' }}>{calculatedTpdDisplay} TPD</span>
                    </div>
                  </>
                ) : (
                  <div>
                    <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'वास्तविक TPD' : 'Observed TPD'}</label>
                    <input style={inputStyle} type="number" value={actualAverageTpd} onChange={(e) => setActualAverageTpd(e.target.value)} />
                  </div>
                )}
              </>
            )}

            {(facilityType === 'MRF' || facilityType === 'MIXED_PLANT') && (
              <>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'दैनिक आवक (TPD)' : 'Daily Intake (TPD)'}</label>
                  <input style={inputStyle} type="number" value={mrfDailyDryTons} onChange={(e) => setMrfDailyDryTons(e.target.value)} />
                </div>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'प्लांट क्षमता (TPD)' : 'Capacity (TPD)'}</label>
                  <input style={inputStyle} type="number" value={mrfMaxCapacityTons} onChange={(e) => setMrfMaxCapacityTons(e.target.value)} />
                </div>
              </>
            )}
          </div>

          {(facilityType === 'MRF' || facilityType === 'MIXED_PLANT') && (
            <div style={{ marginTop: '10px', background: isAdvancedMode ? '#fffbeb' : '#f8fafc', padding: '12px', borderRadius: '6px', border: '1px solid #cbd5e1', marginBottom: '14px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontWeight: 'bold', fontSize: '13px', cursor: 'pointer' }}>
                  <input type="checkbox" checked={isAdvancedMode} onChange={(e) => setIsAdvancedMode(e.target.checked)} />
                  {lang === 'hi' ? 'एडवांस्ड स्ट्रीम सेटिंग चालू करें' : 'Enable Advanced Stream Configuration'}
                </label>
                
                {isAdvancedMode && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <span style={{ fontSize: '13px', fontWeight: 'bold', color: isValidTotal ? '#166534' : '#dc2626', background: isValidTotal ? '#dcfce7' : '#fee2e2', padding: '4px 10px', borderRadius: '4px', border: `1px solid ${isValidTotal ? '#86efac' : '#fca5a5'}` }}>
                      Total: {totalPercentage}%
                    </span>
                  </div>
                )}
              </div>

              {isAdvancedMode && (
                <>
                  <div style={{ marginTop: '14px', display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', gap: '8px' }}>
                    {currentStreamConfig.map(s => (
                      <div key={s.id} style={{ background: '#fff', padding: '8px', border: '1px solid #e2e8f0', borderRadius: '6px' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                          {s.id.startsWith('custom_') ? (
                            <input type="text" value={s.label} onChange={(e) => updateStreamConfig(s.id, 'label', e.target.value)} style={{ fontSize: '11px', fontWeight: 'bold', width: '100%', padding: '4px' }} placeholder="Custom Name" />
                          ) : (
                            <span style={{ fontSize: '11px', fontWeight: 'bold' }}>{s.label}</span>
                          )}
                          {s.id.startsWith('custom_') && (
                            <button type="button" onClick={() => removeCustomStream(s.id)} style={{ background: 'none', border: 'none', color: '#dc2626', cursor: 'pointer' }}><Trash2 size={15} /></button>
                          )}
                        </div>
                        <input type="number" value={s.userWeight} onChange={(e) => updateStreamConfig(s.id, 'userWeight', e.target.value)} style={{ ...inputStyle, marginTop: '4px' }} />
                      </div>
                    ))}
                  </div>
                  <button type="button" onClick={addCustomStream} style={{ marginTop: '10px', padding: '6px 12px', background: '#e2e8f0', color: '#334155', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '4px' }}>
                    <Plus size={14} /> Add Custom Fraction
                  </button>
                </>
              )}
            </div>
          )}

          <div style={{ marginBottom: '14px' }}>
            <strong style={{ fontSize: '13px' }}>{lang === 'hi' ? `महीने चुनें (${pricing.count} चयनित — ₹${pricing.total}):` : `Select Months (${pricing.count} Selected — ₹${pricing.total}):`}</strong>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(75px, 1fr))', gap: '6px', marginTop: '6px' }}>
              {MONTHS.map((m) => (
                <button key={m.id} type="button" onClick={() => toggleMonth(m.id)} style={{ padding: '6px 2px', borderRadius: '4px', border: selectedMonths.includes(m.id) ? '2px solid #0f172a' : '1px solid #cbd5e1', background: selectedMonths.includes(m.id) ? '#f1f5f9' : '#fff', fontWeight: selectedMonths.includes(m.id) ? 'bold' : 'normal', cursor: 'pointer', fontSize: '12px' }}>
                  {lang === 'hi' ? m.shortHi : m.shortEn}
                </button>
              ))}
            </div>
          </div>

          <button type="submit" disabled={generateDisabled} style={{ width: '100%', padding: '14px', background: generateDisabled ? '#94a3b8' : '#0f172a', color: '#fff', border: 'none', borderRadius: '6px', cursor: generateDisabled ? 'not-allowed' : 'pointer', fontWeight: 'bold', fontSize: '15px' }}>
            {lang === 'hi' ? `लॉग-बुक जनरेट करें (₹${pricing.total}) →` : `Generate Logbook Dataset (₹${pricing.total}) →`}
          </button>
        </form>

        {/* PREVIEW CONTAINER */}
        {generatedMonthlyData && (
          <div ref={resultsRef} style={{ background: '#fff', padding: '15px', borderRadius: '8px', border: '1px solid #cbd5e1' }}>
            
            <div style={{ display: 'flex', gap: '6px', marginBottom: '12px', overflowX: 'auto', borderBottom: '1px solid #e2e8f0', paddingBottom: '6px' }}>
              {selectedMonths.map(mId => (
                <button key={mId} type="button" onClick={() => setActiveTabMonth(mId)} style={{ padding: '6px 12px', borderRadius: '4px', border: activeTabMonth === mId ? '2px solid #0f172a' : '1px solid #cbd5e1', background: activeTabMonth === mId ? '#0f172a' : '#f8fafc', color: activeTabMonth === mId ? '#fff' : '#334155', fontWeight: 'bold', cursor: 'pointer', fontSize: '12px' }}>
                  {MONTHS.find(m => m.id === mId)?.fullEn}
                </button>
              ))}
            </div>

            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
              <strong>{name} — Sheet Preview</strong>
              {isPaid ? (
                <button onClick={downloadExcel} style={{ padding: '6px 12px', background: '#0f172a', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '4px' }}>
                  <FileSpreadsheet size={14} /> Download Excel Workbook
                </button>
              ) : (
                <span style={{ fontSize: '12px', color: '#dc2626', fontWeight: 'bold' }}>🔒 Locked Preview (Days 1–5 Only)</span>
              )}
            </div>

            <div onContextMenu={(e) => !isPaid && e.preventDefault()} style={{ overflowX: 'auto', border: '1px solid #cbd5e1', borderRadius: '4px', userSelect: isPaid ? 'text' : 'none' }}>
              <table cellPadding="8" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', minWidth: '600px' }}>
                <thead>
                  {facilityType === 'ULB' ? (
                    <tr style={{ background: '#f1f5f9', borderBottom: '1px solid #cbd5e1', textAlign: 'left' }}>
                      <th>Date</th><th>Day</th>
                      <th style={{ textAlign: 'right' }}>Total Collection (Tons)</th>
                      <th style={{ textAlign: 'right' }}>Domestic MSW (Tons)</th>
                      <th style={{ textAlign: 'right' }}>Commercial (Tons)</th>
                      <th style={{ textAlign: 'right' }}>C&D Waste (Tons)</th>
                      <th style={{ textAlign: 'right' }}>Drain Silt (Tons)</th>
                    </tr>
                  ) : (
                    <tr style={{ background: '#f1f5f9', borderBottom: '1px solid #cbd5e1', textAlign: 'left' }}>
                      <th>Date</th><th>Day</th>
                      <th style={{ textAlign: 'right' }}>Intake (Tons)</th>
                      {currentStreamConfig.map(frac => (
                        <th key={frac.id} style={{ textAlign: 'right' }}>{frac.label} ({frac.userWeight}%)</th>
                      ))}
                    </tr>
                  )}
                </thead>
                <tbody>
                  {visibleRows.map((r, idx) => (
                    <tr key={idx} style={{ borderBottom: '1px solid #e2e8f0' }}>
                      <td>{r.date}</td><td>{r.dayName}</td>
                      <td style={{ textAlign: 'right' }}><strong>{formatVal(r.intake)}</strong></td>
                      
                      {facilityType === 'ULB' ? (
                        <>
                          <td style={{ textAlign: 'right' }}>{formatVal(r.domestic)}</td>
                          <td style={{ textAlign: 'right' }}>{formatVal(r.commercial)}</td>
                          <td style={{ textAlign: 'right' }}>{formatVal(r.cnd)}</td>
                          <td style={{ textAlign: 'right' }}>{formatVal(r.drain)}</td>
                        </>
                      ) : (
                        <>
                          {currentStreamConfig.map(frac => (
                            <td key={frac.id} style={{ textAlign: 'right' }}>{formatVal(r.fractions[frac.id])}</td>
                          ))}
                        </>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {!isPaid && (
              <div style={{ border: '2px dashed #0f172a', background: '#f1f5f9', padding: '15px', textAlign: 'center', marginTop: '12px', borderRadius: '6px' }}>
                <Lock style={{ color: '#0f172a' }} size={18} />
                <h4 style={{ margin: '4px 0', color: '#334155' }}>Preview Locked (Days 1–5 Only)</h4>
                <button onClick={handlePayment} disabled={isProcessing} style={{ padding: '10px 20px', background: '#0f172a', color: '#fff', border: 'none', borderRadius: '6px', fontSize: '14px', fontWeight: 'bold', cursor: 'pointer', marginTop: '8px' }}>
                  {isProcessing ? 'Connecting...' : `Pay ₹${pricing.total} & Download File`}
                </button>
              </div>
            )}
          </div>
        )}

        {/* COMPLIANCE FOOTER */}
        <footer style={{ marginTop: '40px', paddingTop: '20px', borderTop: '1px solid #cbd5e1', textAlign: 'center', fontSize: '12px', color: '#64748b' }}>
          <div style={{ display: 'flex', justifyContent: 'center', gap: '20px', flexWrap: 'wrap', marginBottom: '10px' }}>
            <button type="button" onClick={() => setActivePolicyModal('contact')} style={{ background: 'none', border: 'none', color: '#0f172a', cursor: 'pointer', textDecoration: 'underline', fontWeight: '600' }}>Contact Us</button> |
            <button type="button" onClick={() => setActivePolicyModal('terms')} style={{ background: 'none', border: 'none', color: '#0f172a', cursor: 'pointer', textDecoration: 'underline', fontWeight: '600' }}>Terms & Conditions</button> |
            <button type="button" onClick={() => setActivePolicyModal('refunds')} style={{ background: 'none', border: 'none', color: '#0f172a', cursor: 'pointer', textDecoration: 'underline', fontWeight: '600' }}>Refunds & Cancellations</button> |
            <button type="button" onClick={() => setActivePolicyModal('pricing')} style={{ background: 'none', border: 'none', color: '#0f172a', cursor: 'pointer', textDecoration: 'underline', fontWeight: '600' }}>Services & Pricing (INR)</button>
          </div>
          <p style={{ margin: 0 }}>© {new Date().getFullYear()} Consilience Research Foundation / SWM Suite. All Rights Reserved.</p>
        </footer>

        {/* POLICY MODAL */}
        {activePolicyModal && (
          <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(0,0,0,0.6)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000, padding: '15px' }}>
            <div style={{ background: '#fff', padding: '24px', borderRadius: '8px', maxWidth: '600px', width: '100%', maxHeight: '80vh', overflowY: 'auto' }}>
              {activePolicyModal === 'contact' && (
                <div>
                  <h2 style={{ fontSize: '18px', fontWeight: 'bold' }}>Contact Us</h2>
                  <p style={{ fontSize: '13px' }}><strong>Organisation:</strong> Consilience Research Foundation</p>
                  <p style={{ fontSize: '13px' }}><strong>Address:</strong> Arjunganj, Lucknow, Uttar Pradesh, India</p>
                  <p style={{ fontSize: '13px' }}><strong>Email:</strong> support@consilience.res.in</p>
                </div>
              )}
              {activePolicyModal === 'terms' && (
                <div>
                  <h2 style={{ fontSize: '18px', fontWeight: 'bold' }}>Terms & Conditions</h2>
                  <p style={{ fontSize: '12px' }}>This tool provides engineered single-asset estimations for solid waste management facilities.</p>
                </div>
              )}
              {activePolicyModal === 'refunds' && (
                <div>
                  <h2 style={{ fontSize: '18px', fontWeight: 'bold' }}>Refunds & Cancellations</h2>
                  <p style={{ fontSize: '12px' }}>Digital Excel files are unlocked instantly upon payment confirmation. Failed unlocks after debit are refunded in 5-7 business days.</p>
                </div>
              )}
              {activePolicyModal === 'pricing' && (
                <div>
                  <h2 style={{ fontSize: '18px', fontWeight: 'bold' }}>Services & Pricing (INR)</h2>
                  <ul style={{ fontSize: '12px', lineHeight: '1.8' }}>
                    <li>ULB Collection Dataset: ₹100 / Month</li>
                    <li>MRF Processing Dataset: ₹150 / Month</li>
                    <li>Mixed Waste Processing Dataset: ₹200 / Month</li>
                  </ul>
                </div>
              )}
              <button type="button" onClick={() => setActivePolicyModal(null)} style={{ marginTop: '15px', padding: '8px 16px', background: '#0f172a', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>Close</button>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}
