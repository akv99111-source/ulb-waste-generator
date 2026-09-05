import React, { useState, useRef, useEffect } from 'react';
import { Building2, Download, Lock, Globe, Check, Info, ShieldCheck, MapPin, AlertCircle, Phone, Plus, Trash2, ArrowLeft, Layers, Sparkles, BookOpen } from 'lucide-react';
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

const DEFAULT_MRF_STREAMS = [
  { id: 'pet', label: 'PET', defaultWeight: 15, min: 0, max: 90, isDefault: true },
  { id: 'hdpe', label: 'HDPE', defaultWeight: 10, min: 0, max: 90, isDefault: true },
  { id: 'paper', label: 'Paper/Cardboard', defaultWeight: 25, min: 0, max: 90, isDefault: true },
  { id: 'rdf', label: 'RDF/SCF', defaultWeight: 20, min: 0, max: 100, isDefault: true },
  { id: 'glass_metal', label: 'Glass & Metal', defaultWeight: 10, min: 0, max: 90, isDefault: true },
  { id: 'rejects', label: 'Rejects', defaultWeight: 20, min: 0, max: 100, isDefault: true },
  { id: 'other_mrf_fraction', label: 'Other/Custom Fraction', defaultWeight: 0, min: 0, max: 50, isDefault: false },
];

const DEFAULT_MIXED_STREAMS = [
  { id: 'fine_screen_undersize', label: 'Fine Screen Fraction (Organics)', defaultWeight: 45, min: 10, max: 80, isDefault: true },
  { id: 'coarse_screen_oversize', label: 'Coarse Screen Fraction (RDF)', defaultWeight: 25, min: 5, max: 60, isDefault: true },
  { id: 'recovered_recyclables', label: 'Sorted Recyclables', defaultWeight: 10, min: 1, max: 30, isDefault: true },
  { id: 'inerts_stones', label: 'Inerts, Silt & Stones', defaultWeight: 12, min: 2, max: 40, isDefault: true },
  { id: 'process_rejects', label: 'Process / Landfill Rejects', defaultWeight: 8, min: 1, max: 30, isDefault: true },
  { id: 'other_mixed_fraction', label: 'Other/Custom Fraction', defaultWeight: 0, min: 0, max: 50, isDefault: false },
];

const getSeasonalFractionsULB = (m, regionKey) => {
  const profile = REGION_PROFILES[regionKey] || REGION_PROFILES.north_plains;
  if ([5, 6, 7].includes(m)) return [profile.wetBase + 0.05, profile.dryBase - 0.02, 0.04, 0.02, 0.05, 0.12];
  if ([8, 9].includes(m)) return [profile.wetBase + 0.03, profile.dryBase - 0.01, 0.04, 0.02, 0.05, 0.13];
  return [profile.wetBase, profile.dryBase, 0.04, 0.02, 0.05, 0.15];
};

const cyrb128 = (str) => {
  let h1 = 1779033703, h2 = 3144134277, h3 = 1013904242, h4 = 2773480762;
  for (let i = 0; i < str.length; i++) { let k = str.charCodeAt(i); h1 = h2 ^ Math.imul(h1 ^ k, 597399067); h2 = h3 ^ Math.imul(h2 ^ k, 2869860233); h3 = h4 ^ Math.imul(h3 ^ k, 951274213); h4 = h1 ^ Math.imul(h4 ^ k, 2716044179); }
  return (Math.imul(h3 ^ (h1 >>> 18), 597399067) ^ Math.imul(h4 ^ (h2 >>> 22), 2869860233) ^ Math.imul(h1 ^ (h3 >>> 17), 951274213) ^ Math.imul(h2 ^ (h4 >>> 19), 2716044179)) >>> 0;
};

const mulberry32 = (a) => {
  return function() { var t = a += 0x6D2B79F5; t = Math.imul(t ^ t >>> 15, t | 1); t ^= t + Math.imul(t ^ t >>> 7, t | 61); return ((t ^ t >>> 14) >>> 0) / 4294967296; }
};

const inputStyle = { width: '100%', padding: '9px', borderRadius: '6px', border: '1px solid #cbd5e1', marginTop: '4px', boxSizing: 'border-box' };

export default function App() {
  const [lang, setLang] = useState('hi');
  const [appMode, setAppMode] = useState('STANDALONE'); 
  const [facilityType, setFacilityType] = useState('ULB'); 
  const [selectedState, setSelectedState] = useState('Uttar Pradesh');
  const [name, setName] = useState('Nagar Palika Parishad');
  const [phone, setPhone] = useState('');
  const [isAdvancedMode, setIsAdvancedMode] = useState(false);
  
  const [ulbCalculationMode, setUlbCalculationMode] = useState('population');
  const [population, setPopulation] = useState(50000);
  const [perCapitaOption, setPerCapitaOption] = useState('450');
  const [customPerCapita, setCustomPerCapita] = useState('');
  const [actualAverageTpd, setActualAverageTpd] = useState(10);
  const [segregationRate, setSegregationRate] = useState(80);

  const [mrfDailyDryTons, setMrfDailyDryTons] = useState(15);
  const [mrfMaxCapacityTons, setMrfMaxCapacityTons] = useState(25);
  const [mrfStreamsConfig, setMrfStreamsConfig] = useState(
    DEFAULT_MRF_STREAMS.map(s => ({ ...s, active: s.isDefault, userWeight: s.defaultWeight }))
  );
  const [mixedStreamsConfig, setMixedStreamsConfig] = useState(
    DEFAULT_MIXED_STREAMS.map(s => ({ ...s, active: s.isDefault, userWeight: s.defaultWeight }))
  );

  const [compostUnits, setCompostUnits] = useState([
    { id: 'c1', label: 'Windrow Pad Alpha', type: 'Windrow Pad', capacity: 10 }
  ]);
  const [mrfUnits, setMrfUnits] = useState([
    { id: 'm1', label: 'MRF Shed 1', type: 'Manual Sorting Shed', capacity: 5 }
  ]);
  
  const [startYear, setStartYear] = useState(2026);
  const [selectedMonths, setSelectedMonths] = useState([1]);
  const [displayUnit, setDisplayUnit] = useState('Tons');
  
  const [generatedMonthlyData, setGeneratedMonthlyData] = useState(null);
  const [generatedConfig, setGeneratedConfig] = useState(null);
  const [activeTabMonth, setActiveTabMonth] = useState(null);
  const [isPaid, setIsPaid] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);

  const resultsRef = useRef(null);
  const currentStateObj = STATES_LIST.find(s => s.nameEn === selectedState) || STATES_LIST[33];
  const currentRegionKey = currentStateObj.region;
  const parsedPerCapita = perCapitaOption === 'custom' ? Number(customPerCapita) : Number(perCapitaOption);
  
  const currentStreamConfig = facilityType === 'MIXED_PLANT' ? mixedStreamsConfig : mrfStreamsConfig;
  const activeStreams = isAdvancedMode ? currentStreamConfig.filter(s => s.active) : currentStreamConfig.filter(s => s.isDefault);
  const totalPercentage = activeStreams.reduce((acc, s) => acc + Number(s.userWeight || 0), 0);
  const isValidTotal = totalPercentage === 100;
  const generateDisabled = appMode === 'STANDALONE' && (facilityType === 'MRF' || facilityType === 'MIXED_PLANT') && isAdvancedMode && !isValidTotal;

  // FIX: Safe mode and facility switchers to prevent React UI crash from stale table data
  const switchAppMode = (mode) => {
    setAppMode(mode);
    setGeneratedMonthlyData(null);
    setGeneratedConfig(null);
  };

  const switchFacility = (type) => {
    setFacilityType(type);
    setGeneratedMonthlyData(null);
    setGeneratedConfig(null);
  };

  const getSessionKey = () => {
    const cleanName = name.trim().toLowerCase().replace(/\s+/g, '_');
    return `crf_paid_${appMode}_${facilityType}_${cleanName}_${selectedMonths.length}M`;
  };

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
  }, [name, appMode, facilityType, selectedMonths.length]);

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
    if (appMode === 'INTEGRATED_3IN1') baseRate = 500;
    else if (facilityType === 'MIXED_PLANT') baseRate = 200;
    else if (isAdvancedMode) baseRate = 150;
    
    const baseTotal = billableMonths * baseRate;
    const effectiveFeeRate = 0.0236; 
    const finalTotalWithCharges = Math.round(baseTotal / (1 - effectiveFeeRate));
    return { count, freeMonths, billableMonths, baseRate, baseTotal, total: finalTotalWithCharges };
  };

  const pricing = getPricingDetails();

  const handleGenerate = (e) => {
    e.preventDefault();
    if (generateDisabled) return;

    let monthlyDataMap = {};
    selectedMonths.forEach((m) => {
      const days = new Date(startYear, m, 0).getDate();
      let targetTons = (appMode === 'INTEGRATED_3IN1' || facilityType === 'ULB')
        ? (ulbCalculationMode === 'population' ? (Number(population) * parsedPerCapita) / 1000000 : Number(actualAverageTpd))
        : Number(mrfDailyDryTons);

      const seedString = `${appMode}-${facilityType}-${selectedState}-${name}-${startYear}-${m}-${ulbCalculationMode}-${targetTons}-${segregationRate}`;
      const random = mulberry32(cyrb128(seedString));
      let logs = [];

      for (let day = 1; day <= days; day++) {
        const dateStr = `${startYear}-${String(m).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const dateObj = new Date(startYear, m - 1, day);
        const dayName = dateObj.toLocaleDateString('en-US', { weekday: 'short' });
        let noise = 0.95 + random() * 0.10;
        const dailyTotal = targetTons * noise;

        if (appMode === 'INTEGRATED_3IN1') {
          const segFrac = segregationRate / 100;
          const unsegFrac = 1 - segFrac;
          const segregatedTotal = dailyTotal * segFrac;
          const unsegregatedMixed = Number((dailyTotal * unsegFrac).toFixed(3));
          
          const wetSeg = Number((segregatedTotal * 0.60).toFixed(3));
          const drySeg = Number((segregatedTotal * 0.32).toFixed(3));
          const hazSeg = Number((segregatedTotal * 0.03).toFixed(3));
          const sanSeg = Number((segregatedTotal * 0.05).toFixed(3));

          const organicFines = Number((unsegregatedMixed * 0.45).toFixed(3));
          const dryOversize = Number((unsegregatedMixed * 0.35).toFixed(3));
          const heavyInerts = Number((unsegregatedMixed * 0.20).toFixed(3));

          logs.push({
            date: dateStr, dayName, totalIntake: Number(dailyTotal.toFixed(3)),
            wetSeg, drySeg, hazSeg, sanSeg, unsegregatedMixed,
            organicFines, dryOversize, heavyInerts
          });
        } else if (facilityType === 'ULB') {
          const baseFractions = getSeasonalFractionsULB(m, currentRegionKey);
          let raw = baseFractions.map(r => r * (0.88 + random() * 0.24));
          let sum = raw.reduce((a, b) => a + b, 0);
          let norm = raw.map(r => r / sum);
          logs.push({
            date: dateStr, dayName,
            c1: Number((dailyTotal * norm[0]).toFixed(3)), c2: Number((dailyTotal * norm[1]).toFixed(3)),
            c3: Number((dailyTotal * norm[2]).toFixed(3)), c4: Number((dailyTotal * norm[3]).toFixed(3)),
            c5: Number((dailyTotal * norm[4]).toFixed(3)), c6: Number((dailyTotal * norm[5]).toFixed(3)),
            total: Number(dailyTotal.toFixed(3))
          });
        } else {
          const baseNorm = activeStreams.map(s => Number(s.userWeight || 0) / 100);
          let raw = baseNorm.map(r => r * (0.88 + random() * 0.24));
          let dynamicSum = raw.reduce((a, b) => a + b, 0);
          let dynamicNorm = raw.map(r => r / dynamicSum);
          let rowStreams = {}; let exactTotal = 0;
          activeStreams.forEach((stream, idx) => {
            let val = Number((dailyTotal * dynamicNorm[idx]).toFixed(3));
            rowStreams[stream.id] = val; exactTotal += val;
          });
          logs.push({ date: dateStr, dayName, streams: rowStreams, total: Number(exactTotal.toFixed(3)) });
        }
      }
      monthlyDataMap[m] = logs;
    });

    setGeneratedConfig({ appMode, type: facilityType, streams: activeStreams });
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
      const res = await fetch('/api/create-order', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ amount: pricing.total, customerName: name, customerPhone: phone })
      });

      const order = await res.json();
      if (!order.payment_session_id) throw new Error(order.message || 'Failed to initialize payment session.');

      const cashfree = window.Cashfree({ mode: import.meta.env.VITE_CASHFREE_MODE || 'sandbox' });

      cashfree.checkout({
        paymentSessionId: order.payment_session_id,
        redirectTarget: '_modal'
      }).then((result) => {
        if (result.error) {
          alert('Payment Failed: ' + result.error.message);
        } else if (result.paymentDetails) {
          setIsPaid(true);
          localStorage.setItem(getSessionKey(), JSON.stringify({ paid: true, timestamp: Date.now() }));
        }
        setIsProcessing(false);
      });
    } catch (err) {
      alert('Payment Error: ' + err.message);
      setIsProcessing(false);
    }
  };

  const formatVal = (v) => displayUnit === 'kg' ? Math.round(v * 1000) : Number(v || 0).toFixed(3);

  const downloadMultiSheetExcel = () => {
    if (!generatedMonthlyData) return;
    const u = displayUnit === 'kg' ? 'kg' : 'Tons';
    const wb = XLSX.utils.book_new();

    selectedMonths.forEach((mId) => {
      const monthName = MONTHS.find(m => m.id === mId)?.fullEn;

      if (appMode === 'INTEGRATED_3IN1') {
        const gateHeaders = ["Date", "Day", `Total Gate Intake (${u})`, `Segregated Wet (${u})`, `Segregated Dry (${u})`, `Domestic Hazardous (${u})`, `Domestic Sanitary (${u})`, `Unsegregated Mixed (${u})`];
        const gateRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.totalIntake), formatVal(r.wetSeg), formatVal(r.drySeg), formatVal(r.hazSeg), formatVal(r.sanSeg), formatVal(r.unsegregatedMixed)]);
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([gateHeaders, ...gateRows]), `${monthName}_Gate`);
      } else {
        let headers = facilityType === 'ULB' 
          ? ["Date", "Day", `Wet (${u})`, `Dry (${u})`, `Sanitary (${u})`, `Hazardous (${u})`, `C&D (${u})`, `Inerts (${u})`, `Total (${u})`]
          : ["Date", "Day", ...generatedConfig.streams.map(s => `${s.label} (${u})`), `Total (${u})`];

        const sheetData = [headers, ...generatedMonthlyData[mId].map(r => {
          if (facilityType === 'ULB') {
            return [r.date, r.dayName, formatVal(r.c1), formatVal(r.c2), formatVal(r.c3), formatVal(r.c4), formatVal(r.c5), formatVal(r.c6), formatVal(r.total)];
          } else {
            const row = [r.date, r.dayName];
            generatedConfig.streams.forEach(s => row.push(formatVal(r.streams[s.id])));
            row.push(formatVal(r.total));
            return row;
          }
        })];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sheetData), monthName);
      }
    });
    XLSX.writeFile(wb, `${appMode}_${facilityType}_${name.replace(/\s+/g, '_')}.xlsx`);
  };

  const activeRows = generatedMonthlyData?.[activeTabMonth] || [];
  const visibleRows = isPaid ? activeRows : activeRows.slice(0, 5);

  return (
    <div style={{ fontFamily: 'sans-serif', background: '#f8fafc', minHeight: '100vh', padding: '15px' }}>
      <div style={{ maxWidth: '1000px', margin: '0 auto' }}>
        
        {/* HEADER */}
        <div style={{ background: 'linear-gradient(135deg, #064e3b 0%, #047857 100%)', color: '#fff', padding: '20px', borderRadius: '8px', marginBottom: '16px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap' }}>
            <div>
              <span style={{ background: 'rgba(255,255,255,0.2)', padding: '3px 8px', borderRadius: '12px', fontSize: '11px', fontWeight: 'bold' }}>
                <ShieldCheck size={12} style={{ verticalAlign: 'middle' }} /> {lang === 'hi' ? 'SWM एस्टिमेशन इंजन' : 'SWM ESTIMATION ENGINE'}
              </span>
              <h1 style={{ fontSize: '22px', margin: '6px 0 2px 0', fontWeight: '800' }}>
                <Building2 size={22} style={{ verticalAlign: 'middle', marginRight: '6px' }} />
                {appMode === 'INTEGRATED_3IN1' 
                  ? (lang === 'hi' ? 'एकीकृत 3-इन-1 मल्टी-यूनिट सुइट' : 'Integrated 3-in-1 Multi-Unit Suite') 
                  : (lang === 'hi' ? 'ULB, MRF और मिक्स्ड कचरा लॉग-बुक जनरेटर' : 'ULB, MRF & Mixed Waste Logbook Generator')}
              </h1>
            </div>
            <button type="button" onClick={() => setLang(lang === 'hi' ? 'en' : 'hi')} style={{ padding: '6px 12px', background: '#fff', color: '#047857', border: 'none', borderRadius: '5px', fontWeight: 'bold', cursor: 'pointer' }}>
              <Globe size={15} style={{ verticalAlign: 'middle' }} /> {lang === 'hi' ? 'English' : 'हिंदी'}
            </button>
          </div>
        </div>

        {/* HELP TEXT */}
        <div style={{ background: '#f0f9ff', border: '1px solid #7dd3fc', padding: '12px 16px', borderRadius: '8px', marginBottom: '16px', display: 'flex', gap: '12px', alignItems: 'flex-start' }}>
          <BookOpen style={{ color: '#0284c7', marginTop: '2px' }} size={20} />
          <div>
            <strong style={{ color: '#0369a1', fontSize: '14px', display: 'block', marginBottom: '4px' }}>
              {lang === 'hi' ? 'इस टूल का उपयोग कैसे करें?' : 'How to use this tool?'}
            </strong>
            <p style={{ margin: 0, fontSize: '13px', color: '#0c4a6e', lineHeight: '1.5' }}>
              {lang === 'hi' 
                ? '1. अपने प्लांट का प्रकार चुनें। 2. शहर का नाम और कचरे की क्षमता दर्ज करें। 3. आवश्यक महीने चुनें। 4. डेटा जेनरेट करें और सुरक्षित भुगतान करके एक्सेल फाइल डाउनलोड करें।' 
                : '1. Select your facility type. 2. Enter your city details and capacity. 3. Select the required months. 4. Generate and pay to download your full Excel files.'}
            </p>
          </div>
        </div>

        {/* SWITCHER BANNER */}
        {appMode === 'STANDALONE' ? (
          <div style={{ background: 'linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%)', border: '2px solid #22c55e', padding: '14px 18px', borderRadius: '8px', marginBottom: '16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
            <div>
              <span style={{ background: '#15803d', color: '#fff', padding: '2px 8px', borderRadius: '4px', fontSize: '11px', fontWeight: 'bold', display: 'inline-flex', alignItems: 'center', gap: '4px' }}>
                <Sparkles size={12} /> {lang === 'hi' ? 'नया फीचर' : 'NEW FEATURE'}
              </span>
              <h3 style={{ margin: '4px 0 2px 0', color: '#166534', fontSize: '15px', fontWeight: 'bold' }}>
                {lang === 'hi' ? 'क्या आपको एक साथ जुड़ी हुई लॉग-बुक चाहिए?' : 'Need Interconnected Logbooks?'}
              </h3>
              <p style={{ margin: 0, color: '#15803d', fontSize: '12px' }}>
                {lang === 'hi' ? 'संपूर्ण RRC प्लांट के लिए एक्सेल लॉग-बुक (₹500/माह)।' : 'Synchronized Excel logbooks for complete facilities (₹500/mo).'}
              </p>
            </div>
            <button type="button" onClick={() => switchAppMode('INTEGRATED_3IN1')} style={{ padding: '8px 16px', background: '#166534', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px' }}>
              <Layers size={14} /> {lang === 'hi' ? 'एकीकृत 3-इन-1 सुइट खोलें' : 'Open Integrated 3-in-1 Suite'} →
            </button>
          </div>
        ) : (
          <button type="button" onClick={() => switchAppMode('STANDALONE')} style={{ marginBottom: '16px', padding: '6px 12px', background: '#e2e8f0', color: '#334155', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold', display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
            <ArrowLeft size={14} /> {lang === 'hi' ? 'सिंगल-फैसिलिटी लॉग-बुक पर वापस जाएँ' : 'Back to Single-Facility Logbooks'}
          </button>
        )}

        {/* FORM SECTION */}
        <form onSubmit={handleGenerate} style={{ background: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #cbd5e1', marginBottom: '20px' }}>
          
          {appMode === 'STANDALONE' && (
            <div style={{ marginBottom: '14px', display: 'flex', gap: '15px', alignItems: 'center', fontSize: '14px', flexWrap: 'wrap', background: '#f8fafc', padding: '10px', borderRadius: '6px', border: '1px solid #e2e8f0' }}>
              <strong>{lang === 'hi' ? 'सिंगल-फैसिलिटी चुनें:' : 'Select Standalone Logbook:'}</strong>
              <label style={{ cursor: 'pointer' }}><input type="radio" value="ULB" checked={facilityType === 'ULB'} onChange={() => switchFacility('ULB')} /> {lang === 'hi' ? 'निकाय (ULB) (₹100)' : 'ULB Collection (₹100)'}</label>
              <label style={{ cursor: 'pointer' }}><input type="radio" value="MRF" checked={facilityType === 'MRF'} onChange={() => switchFacility('MRF')} /> {lang === 'hi' ? 'एमआरएफ (MRF) (₹100–150)' : 'MRF Centre (₹100–150)'}</label>
              <label style={{ cursor: 'pointer' }}><input type="radio" value="MIXED_PLANT" checked={facilityType === 'MIXED_PLANT'} onChange={() => switchFacility('MIXED_PLANT')} /> {lang === 'hi' ? 'मिश्रित कचरा (₹200)' : 'Mixed Waste Plant (₹200)'}</label>
            </div>
          )}

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '12px', marginBottom: '14px' }}>
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

            {(appMode === 'INTEGRATED_3IN1' || facilityType === 'ULB') && (
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
                      <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'अनुमानित जनसंख्या' : 'Population'}</label>
                      <input style={inputStyle} type="number" value={population} onChange={(e) => setPopulation(e.target.value)} />
                    </div>
                    <div>
                      <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'प्रति व्यक्ति दर' : 'Per Capita Rate'}</label>
                      <select style={inputStyle} value={perCapitaOption} onChange={(e) => setPerCapitaOption(e.target.value)}>
                        <option value="300">300 {lang === 'hi' ? 'ग्राम/दिन' : 'g/day'}</option>
                        <option value="450">450 {lang === 'hi' ? 'ग्राम/दिन' : 'g/day'}</option>
                        <option value="500">500 {lang === 'hi' ? 'ग्राम/दिन' : 'g/day'}</option>
                      </select>
                    </div>
                  </>
                ) : (
                  <div>
                    <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'वास्तविक कचरा (TPD)' : 'Observed Waste (TPD)'}</label>
                    <input style={inputStyle} type="number" value={actualAverageTpd} onChange={(e) => setActualAverageTpd(e.target.value)} />
                  </div>
                )}
              </>
            )}

            {appMode === 'STANDALONE' && (facilityType === 'MRF' || facilityType === 'MIXED_PLANT') && (
              <>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'दैनिक आवक (TPD)' : 'Daily Intake (TPD)'}</label>
                  <input style={inputStyle} type="number" value={mrfDailyDryTons} onChange={(e) => setMrfDailyDryTons(e.target.value)} />
                </div>
              </>
            )}
          </div>

          {appMode === 'INTEGRATED_3IN1' && (
            <div style={{ background: '#ecfdf5', padding: '12px', borderRadius: '6px', border: '1px solid #a7f3d0', marginBottom: '14px' }}>
              <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#065f46' }}>
                {lang === 'hi' ? `कचरा पृथक्करण दर (%): ${segregationRate}%` : `Source Segregation Rate (%): ${segregationRate}%`}
              </label>
              <input type="range" min="20" max="95" step="5" value={segregationRate} onChange={(e) => setSegregationRate(Number(e.target.value))} style={{ width: '100%', marginTop: '6px' }} />
            </div>
          )}

          <div style={{ marginBottom: '14px' }}>
            <strong style={{ fontSize: '13px', display: 'flex', alignItems: 'center', gap: '10px' }}>
              {lang === 'hi' ? `महीने चुनें (${pricing.count} चयनित — ₹${pricing.total}):` : `Select Months (${pricing.count} Selected — ₹${pricing.total}):`}
              {pricing.freeMonths > 0 && (
                <span style={{ color: '#dc2626', fontSize: '11px', background: '#fee2e2', padding: '3px 8px', borderRadius: '12px' }}>
                  {lang === 'hi' ? `🎉 ${pricing.freeMonths} महीना मुफ़्त!` : `🎉 ${pricing.freeMonths} Month Free!`}
                </span>
              )}
            </strong>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(75px, 1fr))', gap: '6px', marginTop: '6px' }}>
              {MONTHS.map((m) => {
                const active = selectedMonths.includes(m.id);
                return (
                  <button key={m.id} type="button" onClick={() => toggleMonth(m.id)} style={{ padding: '6px 2px', borderRadius: '4px', border: active ? '2px solid #059669' : '1px solid #cbd5e1', background: active ? '#ecfdf5' : '#fff', fontWeight: active ? 'bold' : 'normal', cursor: 'pointer', fontSize: '12px' }}>
                    {lang === 'hi' ? m.shortHi : m.shortEn}
                  </button>
                );
              })}
            </div>
          </div>

          <button type="submit" disabled={generateDisabled} style={{ width: '100%', padding: '12px', background: generateDisabled ? '#94a3b8' : '#059669', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '15px' }}>
            {lang === 'hi' 
              ? `लॉग-बुक जनरेट करें (${appMode === 'INTEGRATED_3IN1' ? '₹500/माह' : `₹${pricing.total}`}) →` 
              : `Generate Logbook Dataset (${appMode === 'INTEGRATED_3IN1' ? '₹500/mo' : `₹${pricing.total}`}) →`}
          </button>
        </form>

        {/* RESULTS PREVIEW */}
        {generatedMonthlyData && (
          <div ref={resultsRef} style={{ background: '#fff', padding: '15px', borderRadius: '8px', border: '1px solid #cbd5e1' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px', flexWrap: 'wrap', gap: '10px' }}>
              <strong style={{ fontSize: '14px' }}>{name} — {lang === 'hi' ? 'लॉग-बुक डेटासेट' : 'Dataset Preview'}</strong>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button onClick={() => setDisplayUnit(displayUnit === 'Tons' ? 'kg' : 'Tons')} style={{ padding: '6px 10px', fontSize: '12px', borderRadius: '4px', border: '1px solid #cbd5e1' }}>
                  {lang === 'hi' ? 'यूनिट:' : 'Unit:'} <strong>{displayUnit}</strong>
                </button>
                {isPaid && (
                  <button onClick={downloadMultiSheetExcel} style={{ padding: '6px 12px', background: '#0f172a', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold' }}>
                    <Download size={13} /> Export Excel (.xlsx)
                  </button>
                )}
              </div>
            </div>

            <div onContextMenu={(e) => !isPaid && e.preventDefault()} style={{ overflowX: 'auto', border: '1px solid #cbd5e1', borderRadius: '4px', userSelect: isPaid ? 'text' : 'none' }}>
              <table cellPadding="8" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', minWidth: '600px' }}>
                <thead>
                  <tr style={{ background: '#f1f5f9', borderBottom: '1px solid #cbd5e1' }}>
                    <th>Date</th><th>Day</th>
                    {appMode === 'INTEGRATED_3IN1' ? (
                      <><th>Gate Intake</th><th>Wet Seg</th><th>Dry Seg</th><th>Mixed</th></>
                    ) : (
                      <>
                        {facilityType === 'ULB' ? (
                          <><th>Wet</th><th>Dry</th><th>Sanitary</th><th>Hazardous</th><th>C&D</th><th>Inerts</th><th>Total</th></>
                        ) : (
                          <>{generatedConfig?.streams.map(s => <th key={s.id}>{s.label}</th>)}<th>Total</th></>
                        )}
                      </>
                    )}
                  </tr>
                </thead>
                <tbody>
                  {visibleRows.map((r, idx) => (
                    <tr key={idx} style={{ borderBottom: '1px solid #e2e8f0' }}>
                      <td>{r.date}</td><td>{r.dayName}</td>
                      {appMode === 'INTEGRATED_3IN1' ? (
                        <>
                          <td><strong>{formatVal(r.totalIntake)}</strong></td>
                          <td>{formatVal(r.wetSeg)}</td>
                          <td>{formatVal(r.drySeg)}</td>
                          <td>{formatVal(r.unsegregatedMixed)}</td>
                        </>
                      ) : (
                        <>
                          {facilityType === 'ULB' ? (
                            <>
                              <td>{formatVal(r.c1)}</td><td>{formatVal(r.c2)}</td><td>{formatVal(r.c3)}</td>
                              <td>{formatVal(r.c4)}</td><td>{formatVal(r.c5)}</td><td>{formatVal(r.c6)}</td>
                              <td><strong>{formatVal(r.total)}</strong></td>
                            </>
                          ) : (
                            <>
                              {generatedConfig?.streams.map(s => <td key={s.id}>{formatVal(r.streams?.[s.id])}</td>)}
                              <td><strong>{formatVal(r.total)}</strong></td>
                            </>
                          )}
                        </>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {!isPaid && (
              <div style={{ border: '2px dashed #059669', background: '#ecfdf5', padding: '15px', textAlign: 'center', marginTop: '12px', borderRadius: '6px' }}>
                <Lock style={{ color: '#059669' }} size={18} />
                <h4 style={{ margin: '4px 0', color: '#065f46', fontSize: '15px' }}>
                  {lang === 'hi' ? 'प्रीव्यू लॉक है (केवल 1-5 दिन दिख रहे हैं)' : 'Preview Locked (Days 1–5 Only)'}
                </h4>
                <p style={{ margin: '4px 0 10px 0', color: '#047857', fontSize: '13px' }}>
                  {lang === 'hi' 
                    ? `पूरी एक्सेल फाइल डाउनलोड करने के लिए ₹${pricing.total} का सुरक्षित भुगतान करें।` 
                    : `Pay ₹${pricing.total} to unlock and download your complete Master Excel.`}
                </p>
                <button onClick={handlePayment} disabled={isProcessing} style={{ padding: '10px 20px', background: '#059669', color: '#fff', border: 'none', borderRadius: '6px', fontSize: '14px', fontWeight: 'bold', cursor: 'pointer' }}>
                  {isProcessing 
                    ? (lang === 'hi' ? 'प्रक्रिया जारी है...' : 'Connecting...') 
                    : (lang === 'hi' ? `₹${pricing.total} सुरक्षित भुगतान करें` : `Pay ₹${pricing.total} Securely`)}
                </button>
              </div>
            )}
          </div>
        )}

      </div>
    </div>
  );
}
