import React, { useState, useRef, useEffect } from 'react';
import { Building2, Download, Lock, Globe, Check, Info, ShieldCheck, MapPin, AlertCircle, Phone, Plus, Trash2 } from 'lucide-react';
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
  const [facilityType, setFacilityType] = useState('INTEGRATED_3IN1');
  const [name, setName] = useState('Nagar Palika Parishad');
  const [phone, setPhone] = useState('');
  
  // Intake and Segregation
  const [ulbCalculationMode, setUlbCalculationMode] = useState('population');
  const [population, setPopulation] = useState(50000);
  const [perCapitaOption, setPerCapitaOption] = useState('450');
  const [customPerCapita, setCustomPerCapita] = useState('');
  const [actualAverageTpd, setActualAverageTpd] = useState(10);
  const [segregationRate, setSegregationRate] = useState(80);

  // Dynamic Infrastructure Asset Lists
  const [compostUnits, setCompostUnits] = useState([
    { id: 'c1', label: 'Windrow Pad Alpha', type: 'Windrow Composting', capacity: 10 }
  ]);
  const [mrfUnits, setMrfUnits] = useState([
    { id: 'm1', label: 'MRF Shed 1', type: 'Manual Sorting', capacity: 5 }
  ]);

  const [startYear, setStartYear] = useState(2026);
  const [selectedMonths, setSelectedMonths] = useState([1]);
  const [displayUnit, setDisplayUnit] = useState('Tons');
  
  const [generatedMonthlyData, setGeneratedMonthlyData] = useState(null);
  const [activeTabMonth, setActiveTabMonth] = useState(null);
  const [isPaid, setIsPaid] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);

  const resultsRef = useRef(null);
  const currentStateObj = STATES_LIST.find(s => s.nameEn === selectedState) || STATES_LIST[33];

  const parsedPerCapita = perCapitaOption === 'custom' ? Number(customPerCapita) : Number(perCapitaOption);
  const estimatedDailyWaste = ((Number(population) * parsedPerCapita) / 1000000).toFixed(2);

  // Dynamic Array Handlers
  const addCompostUnit = () => {
    const newId = `c${compostUnits.length + 1}`;
    setCompostUnits([...compostUnits, { id: newId, label: `Compost Unit ${compostUnits.length + 1}`, type: 'Vermicomposting', capacity: 5 }]);
  };

  const removeCompostUnit = (id) => {
    if (compostUnits.length > 1) {
      setCompostUnits(compostUnits.filter(u => u.id !== id));
    }
  };

  const updateCompostUnit = (id, field, value) => {
    setCompostUnits(compostUnits.map(u => u.id === id ? { ...u, [field]: value } : u));
  };

  const addMrfUnit = () => {
    const newId = `m${mrfUnits.length + 1}`;
    setMrfUnits([...mrfUnits, { id: newId, label: `MRF Shed ${mrfUnits.length + 1}`, type: 'Dry Sorting', capacity: 5 }]);
  };

  const removeMrfUnit = (id) => {
    if (mrfUnits.length > 1) {
      setMrfUnits(mrfUnits.filter(u => u.id !== id));
    }
  };

  const updateMrfUnit = (id, field, value) => {
    setMrfUnits(mrfUnits.map(u => u.id === id ? { ...u, [field]: value } : u));
  };

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
    const baseRate = 500; // Flat ₹500/month for Integrated 3-in-1 Master Suite
    
    const baseTotal = billableMonths * baseRate;
    const effectiveFeeRate = 0.0236; 
    const finalTotalWithCharges = Math.round(baseTotal / (1 - effectiveFeeRate));
    const gatewayFee = finalTotalWithCharges - baseTotal;

    return { count, freeMonths, billableMonths, baseRate, baseTotal, gatewayFee, total: finalTotalWithCharges };
  };

  const pricing = getPricingDetails();

  const handleGenerate = (e) => {
    e.preventDefault();
    let monthlyDataMap = {};

    selectedMonths.forEach((m) => {
      const days = new Date(startYear, m, 0).getDate();
      let targetTons = ulbCalculationMode === 'population' ? (Number(population) * parsedPerCapita) / 1000000 : Number(actualAverageTpd);

      const seedString = `${facilityType}-${selectedState}-${name}-${startYear}-${m}-${ulbCalculationMode}-${targetTons}-${segregationRate}`;
      const random = mulberry32(cyrb128(seedString));

      let logs = [];

      for (let day = 1; day <= days; day++) {
        const dateStr = `${startYear}-${String(m).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const dateObj = new Date(startYear, m - 1, day);
        const dayName = dateObj.toLocaleDateString('en-US', { weekday: 'short' });

        let noise = 0.95 + random() * 0.10;
        const dailyTotal = targetTons * noise;

        const segFrac = segregationRate / 100;
        const unsegFrac = 1 - segFrac;

        // Gate Log
        const segregatedTotal = dailyTotal * segFrac;
        const unsegregatedMixed = Number((dailyTotal * unsegFrac).toFixed(3));
        const wetSeg = Number((segregatedTotal * 0.60).toFixed(3));
        const drySeg = Number((segregatedTotal * 0.32).toFixed(3));
        const hazSeg = Number((segregatedTotal * 0.03).toFixed(3));
        const sanSeg = Number((segregatedTotal * 0.05).toFixed(3));

        // Pre-Sort Log
        const organicFines = Number((unsegregatedMixed * 0.45).toFixed(3));
        const dryOversize = Number((unsegregatedMixed * 0.35).toFixed(3));
        const heavyInerts = Number((unsegregatedMixed * 0.20).toFixed(3));

        // Proportional Asset Distribution
        const totalCompostFeed = wetSeg + organicFines;
        const totalCompostCapacity = compostUnits.reduce((acc, u) => acc + Number(u.capacity || 1), 0);
        
        let compostUnitBreakdown = {};
        compostUnits.forEach(unit => {
          const unitShare = (Number(unit.capacity || 1) / totalCompostCapacity) * totalCompostFeed;
          compostUnitBreakdown[unit.id] = {
            feed: Number(unitShare.toFixed(3)),
            compostYield: Number((unitShare * 0.18).toFixed(3))
          };
        });

        const totalMrfFeed = drySeg + dryOversize;
        const totalMrfCapacity = mrfUnits.reduce((acc, u) => acc + Number(u.capacity || 1), 0);

        let mrfUnitBreakdown = {};
        mrfUnits.forEach(unit => {
          const unitShare = (Number(unit.capacity || 1) / totalMrfCapacity) * totalMrfFeed;
          mrfUnitBreakdown[unit.id] = {
            feed: Number(unitShare.toFixed(3)),
            recyclables: Number((unitShare * 0.65).toFixed(3)),
            rdf: Number((unitShare * 0.25).toFixed(3))
          };
        });

        logs.push({
          date: dateStr, dayName, totalIntake: Number(dailyTotal.toFixed(3)),
          wetSeg, drySeg, hazSeg, sanSeg, unsegregatedMixed,
          organicFines, dryOversize, heavyInerts,
          totalCompostFeed: Number(totalCompostFeed.toFixed(3)),
          totalMrfFeed: Number(totalMrfFeed.toFixed(3)),
          compostUnitBreakdown, mrfUnitBreakdown
        });
      }
      monthlyDataMap[m] = logs;
    });

    setGeneratedMonthlyData(monthlyDataMap);
    setActiveTabMonth(selectedMonths[0]);
    setTimeout(() => resultsRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 100);
  };

  const formatVal = (v) => displayUnit === 'kg' ? Math.round(v * 1000) : Number(v || 0).toFixed(3);

  const downloadMultiSheetExcel = () => {
    if (!generatedMonthlyData) return;
    const u = displayUnit === 'kg' ? 'kg' : 'Tons';
    const wb = XLSX.utils.book_new();

    selectedMonths.forEach((mId) => {
      const monthName = MONTHS.find(m => m.id === mId)?.fullEn;

      // Gate Log
      const gateHeaders = ["Date", "Day", `Total Gate Intake (${u})`, `Segregated Wet (${u})`, `Segregated Dry (${u})`, `Domestic Hazardous (${u})`, `Domestic Sanitary (${u})`, `Unsegregated Mixed (${u})`];
      const gateRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.totalIntake), formatVal(r.wetSeg), formatVal(r.drySeg), formatVal(r.hazSeg), formatVal(r.sanSeg), formatVal(r.unsegregatedMixed)]);
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([gateHeaders, ...gateRows]), `${monthName}_Gate`);

      // Pre-Sorting Log
      const preHeaders = ["Date", "Day", `Mixed Intake (${u})`, `Fine Screen Fraction (${u})`, `Coarse Screen Fraction (${u})`, `Heavy Inerts (${u})`];
      const preRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.unsegregatedMixed), formatVal(r.organicFines), formatVal(r.dryOversize), formatVal(r.heavyInerts)]);
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([preHeaders, ...preRows]), `${monthName}_PreSort`);

      // Dynamic Compost Unit Tabs
      compostUnits.forEach(unit => {
        const cHeaders = ["Date", "Day", `Unit Feed (${u})`, `Compost Yield (${u})`];
        const cRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.compostUnitBreakdown[unit.id]?.feed), formatVal(r.compostUnitBreakdown[unit.id]?.compostYield)]);
        const safeTabName = `${monthName}_${unit.label.substring(0, 10)}`;
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([cHeaders, ...cRows]), safeTabName);
      });

      // Dynamic MRF Unit Tabs
      mrfUnits.forEach(unit => {
        const mHeaders = ["Date", "Day", `Unit Feed (${u})`, `Sorted Recyclables (${u})`, `RDF Dispatched (${u})`];
        const mRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.mrfUnitBreakdown[unit.id]?.feed), formatVal(r.mrfUnitBreakdown[unit.id]?.recyclables), formatVal(r.mrfUnitBreakdown[unit.id]?.rdf)]);
        const safeTabName = `${monthName}_${unit.label.substring(0, 10)}`;
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([mHeaders, ...mRows]), safeTabName);
      });
    });

    XLSX.writeFile(wb, `Integrated_Master_Logbook_${name.replace(/\s+/g, '_')}.xlsx`);
  };

  const activeRows = generatedMonthlyData?.[activeTabMonth] || [];
  const visibleRows = isPaid ? activeRows : activeRows.slice(0, 5);

  return (
    <div style={{ fontFamily: 'sans-serif', background: '#f8fafc', minHeight: '100vh', padding: '15px' }}>
      <div style={{ maxWidth: '1000px', margin: '0 auto' }}>
        
        {/* HEADER BANNER */}
        <div style={{ background: 'linear-gradient(135deg, #064e3b 0%, #047857 100%)', color: '#fff', padding: '20px', borderRadius: '8px', marginBottom: '20px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap' }}>
            <div>
              <span style={{ background: 'rgba(255,255,255,0.2)', padding: '3px 8px', borderRadius: '12px', fontSize: '11px', fontWeight: 'bold' }}>
                <ShieldCheck size={12} style={{ verticalAlign: 'middle' }} /> MULTI-ASSET SWM ESTIMATION ENGINE
              </span>
              <h1 style={{ fontSize: '22px', margin: '6px 0 2px 0', fontWeight: '800' }}>
                <Building2 size={22} style={{ verticalAlign: 'middle', marginRight: '6px' }} />
                {lang === 'hi' ? 'एकीकृत 3-इन-1 मल्टी-यूनिट लॉग-बुक जनरेटर' : 'Integrated 3-in-1 Multi-Unit Logbook Suite'}
              </h1>
            </div>
            <button type="button" onClick={() => setLang(lang === 'hi' ? 'en' : 'hi')} style={{ padding: '6px 12px', background: '#fff', color: '#047857', border: 'none', borderRadius: '5px', fontWeight: 'bold', cursor: 'pointer' }}>
              <Globe size={15} style={{ verticalAlign: 'middle' }} /> {lang === 'hi' ? 'English' : 'हिंदी'}
            </button>
          </div>
        </div>

        {/* INPUT FORM */}
        <form onSubmit={handleGenerate} style={{ background: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #cbd5e1', marginBottom: '20px' }}>
          
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '12px', marginBottom: '16px' }}>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'राज्य चुनें' : 'Select State'}</label>
              <select style={inputStyle} value={selectedState} onChange={(e) => setSelectedState(e.target.value)}>
                {STATES_LIST.map((s) => <option key={s.nameEn} value={s.nameEn}>{lang === 'hi' ? s.nameHi : s.nameEn}</option>)}
              </select>
            </div>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'निकाय / प्लांट का नाम' : 'ULB / Facility Name'}</label>
              <input style={inputStyle} type="text" required value={name} onChange={(e) => setName(e.target.value)} />
            </div>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'मोबाइल नंबर' : 'Mobile Number'}</label>
              <input style={inputStyle} type="tel" maxLength={10} required value={phone} onChange={(e) => setPhone(e.target.value.replace(/\D/g, ''))} />
            </div>
          </div>

          {/* INTAKE & SEGREGATION CONFIG */}
          <div style={{ background: '#f8fafc', padding: '12px', borderRadius: '6px', border: '1px solid #e2e8f0', marginBottom: '16px' }}>
            <strong style={{ fontSize: '13px', color: '#0f172a' }}>{lang === 'hi' ? 'कचरा आवक एवं पृथक्करण दक्षता' : 'Gate Intake & Segregation Efficiency'}</strong>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '12px', marginTop: '8px' }}>
              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold' }}>Population (Approx.)</label>
                <input style={inputStyle} type="number" value={population} onChange={(e) => setPopulation(e.target.value)} />
              </div>
              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold' }}>Source Segregation Rate: {segregationRate}%</label>
                <input type="range" min="20" max="95" step="5" value={segregationRate} onChange={(e) => setSegregationRate(Number(e.target.value))} style={{ width: '100%', marginTop: '8px' }} />
              </div>
            </div>
          </div>

          {/* DYNAMIC COMPOST UNITS */}
          <div style={{ background: '#ecfdf5', padding: '14px', borderRadius: '6px', border: '1px solid #a7f3d0', marginBottom: '16px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
              <strong style={{ fontSize: '13px', color: '#065f46' }}>Composting Assets (Wet Line)</strong>
              <button type="button" onClick={addCompostUnit} style={{ padding: '4px 8px', background: '#047857', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '11px', display: 'flex', alignItems: 'center', gap: '4px' }}>
                <Plus size={12} /> Add Compost Unit
              </button>
            </div>
            {compostUnits.map((unit) => (
              <div key={unit.id} style={{ display: 'grid', gridTemplateColumns: '2fr 2fr 1fr auto', gap: '8px', marginBottom: '8px', alignItems: 'center' }}>
                <input type="text" value={unit.label} onChange={(e) => updateCompostUnit(unit.id, 'label', e.target.value)} style={{ ...inputStyle, marginTop: 0 }} placeholder="Unit Label" />
                <select value={unit.type} onChange={(e) => updateCompostUnit(unit.id, 'type', e.target.value)} style={{ ...inputStyle, marginTop: 0 }}>
                  <option value="Windrow Composting">Windrow Pad</option>
                  <option value="Vermicomposting">Vermicompost Pit</option>
                  <option value="Bio-Methanation">OWC / Bio-Methanation</option>
                </select>
                <input type="number" value={unit.capacity} onChange={(e) => updateCompostUnit(unit.id, 'capacity', e.target.value)} style={{ ...inputStyle, marginTop: 0 }} placeholder="TPD" />
                {compostUnits.length > 1 && (
                  <button type="button" onClick={() => removeCompostUnit(unit.id)} style={{ color: '#dc2626', background: 'none', border: 'none', cursor: 'pointer' }}><Trash2 size={16} /></button>
                )}
              </div>
            ))}
          </div>

          {/* DYNAMIC MRF UNITS */}
          <div style={{ background: '#f0f9ff', padding: '14px', borderRadius: '6px', border: '1px solid #bae6fd', marginBottom: '16px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
              <strong style={{ fontSize: '13px', color: '#0369a1' }}>MRF / Sorting Shed Assets (Dry Line)</strong>
              <button type="button" onClick={addMrfUnit} style={{ padding: '4px 8px', background: '#0284c7', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '11px', display: 'flex', alignItems: 'center', gap: '4px' }}>
                <Plus size={12} /> Add MRF Shed
              </button>
            </div>
            {mrfUnits.map((unit) => (
              <div key={unit.id} style={{ display: 'grid', gridTemplateColumns: '2fr 2fr 1fr auto', gap: '8px', marginBottom: '8px', alignItems: 'center' }}>
                <input type="text" value={unit.label} onChange={(e) => updateMrfUnit(unit.id, 'label', e.target.value)} style={{ ...inputStyle, marginTop: 0 }} placeholder="Unit Label" />
                <select value={unit.type} onChange={(e) => updateMrfUnit(unit.id, 'type', e.target.value)} style={{ ...inputStyle, marginTop: 0 }}>
                  <option value="Manual Sorting">Manual Sorting Shed</option>
                  <option value="Semi-Automated MRF">Semi-Automated Line</option>
                  <option value="RDF Baling Shed">RDF Baling Storage</option>
                </select>
                <input type="number" value={unit.capacity} onChange={(e) => updateMrfUnit(unit.id, 'capacity', e.target.value)} style={{ ...inputStyle, marginTop: 0 }} placeholder="TPD" />
                {mrfUnits.length > 1 && (
                  <button type="button" onClick={() => removeMrfUnit(unit.id)} style={{ color: '#dc2626', background: 'none', border: 'none', cursor: 'pointer' }}><Trash2 size={16} /></button>
                )}
              </div>
            ))}
          </div>

          <button type="submit" style={{ width: '100%', padding: '12px', background: '#059669', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '15px' }}>
            Generate Master Dataset (₹500/mo Suite) →
          </button>
        </form>

        {/* DATASET TABLE PREVIEW */}
        {generatedMonthlyData && (
          <div ref={resultsRef} style={{ background: '#fff', padding: '15px', borderRadius: '8px', border: '1px solid #cbd5e1' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
              <strong style={{ fontSize: '14px' }}>{name} — Master Dataset Preview</strong>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button onClick={() => setDisplayUnit(displayUnit === 'Tons' ? 'kg' : 'Tons')} style={{ padding: '6px 10px', fontSize: '12px', borderRadius: '4px', border: '1px solid #cbd5e1' }}>Unit: <strong>{displayUnit}</strong></button>
                {isPaid ? (
                  <button onClick={downloadMultiSheetExcel} style={{ padding: '6px 12px', background: '#0f172a', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold' }}>
                    <Download size={13} /> Export Master Excel (.xlsx)
                  </button>
                ) : (
                  <button onClick={() => setIsPaid(true)} style={{ padding: '6px 12px', background: '#059669', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold' }}>
                    Pay ₹{pricing.total} to Unlock Full File
                  </button>
                )}
              </div>
            </div>

            <div style={{ overflowX: 'auto', border: '1px solid #cbd5e1', borderRadius: '4px' }}>
              <table cellPadding="8" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
                <thead>
                  <tr style={{ background: '#f1f5f9', borderBottom: '1px solid #cbd5e1' }}>
                    <th>Date</th><th>Day</th><th>Total Gate Intake</th><th>Seg. Wet Waste</th><th>Seg. Dry Waste</th><th>Unsegregated Mixed</th>
                  </tr>
                </thead>
                <tbody>
                  {visibleRows.map((r, idx) => (
                    <tr key={idx} style={{ borderBottom: '1px solid #e2e8f0' }}>
                      <td>{r.date}</td><td>{r.dayName}</td>
                      <td><strong>{formatVal(r.totalIntake)}</strong></td>
                      <td>{formatVal(r.wetSeg)}</td>
                      <td>{formatVal(r.drySeg)}</td>
                      <td>{formatVal(r.unsegregatedMixed)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}
