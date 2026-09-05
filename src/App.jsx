import React, { useState, useRef, useEffect } from 'react';
import { Building2, Download, Lock, Globe, Check, Info, ShieldCheck, MapPin, AlertCircle, Phone, Sliders, Layers } from 'lucide-react';
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
  { id: 'fine_screen_undersize', label: 'Fine Screen Fraction (Organics/Undersize)', defaultWeight: 45, min: 10, max: 80, isDefault: true },
  { id: 'coarse_screen_oversize', label: 'Coarse Screen Fraction (RDF/Oversize)', defaultWeight: 25, min: 5, max: 60, isDefault: true },
  { id: 'recovered_recyclables', label: 'Sorted Recyclables (Plastics/Metals)', defaultWeight: 10, min: 1, max: 30, isDefault: true },
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
  const [isAdvancedMode, setIsAdvancedMode] = useState(false);
  
  // ULB & Integrated State Variables
  const [ulbCalculationMode, setUlbCalculationMode] = useState('population');
  const [population, setPopulation] = useState(50000);
  const [perCapitaOption, setPerCapitaOption] = useState('450');
  const [customPerCapita, setCustomPerCapita] = useState('');
  const [actualAverageTpd, setActualAverageTpd] = useState(10);
  const [referencePeriod, setReferencePeriod] = useState(30);
  const [segregationRate, setSegregationRate] = useState(80);

  // MRF & Mixed Plant State Variables
  const [mrfDailyDryTons, setMrfDailyDryTons] = useState(15);
  const [mrfMaxCapacityTons, setMrfMaxCapacityTons] = useState(25);
  const [mrfStreamsConfig, setMrfStreamsConfig] = useState(
    DEFAULT_MRF_STREAMS.map(s => ({ ...s, active: s.isDefault, userWeight: s.defaultWeight }))
  );
  const [mixedStreamsConfig, setMixedStreamsConfig] = useState(
    DEFAULT_MIXED_STREAMS.map(s => ({ ...s, active: s.isDefault, userWeight: s.defaultWeight }))
  );
  
  const [startYear, setStartYear] = useState(2026);
  const [selectedMonths, setSelectedMonths] = useState([1]);
  const [displayUnit, setDisplayUnit] = useState('Tons');
  
  const [generatedMonthlyData, setGeneratedMonthlyData] = useState(null);
  const [generatedConfig, setGeneratedConfig] = useState(null);
  const [activeTabMonth, setActiveTabMonth] = useState(null);
  const [isPaid, setIsPaid] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [showStateInfo, setShowStateInfo] = useState(false);
  const [activePolicy, setActivePolicy] = useState(null);

  // Recovery State Variables
  const [showRestoreModal, setShowRestoreModal] = useState(false);
  const [restoreOrderId, setRestoreOrderId] = useState('');
  const [isRestoring, setIsRestoring] = useState(false);

  const resultsRef = useRef(null);
  const currentStateObj = STATES_LIST.find(s => s.nameEn === selectedState) || STATES_LIST[33];
  const currentRegionKey = currentStateObj.region;
  const currentRegionObj = REGION_PROFILES[currentRegionKey];

  const parsedPerCapita = perCapitaOption === 'custom' ? Number(customPerCapita) : Number(perCapitaOption);
  const estimatedDailyWaste = ((Number(population) * parsedPerCapita) / 1000000).toFixed(2);
  const mrfCapacityUtilization = Number(mrfMaxCapacityTons) > 0 ? (Number(mrfDailyDryTons) / Number(mrfMaxCapacityTons)) * 100 : 0;

  // Stream Selection & Mass Balance Logic
  const currentStreamConfig = facilityType === 'MIXED_PLANT' ? mixedStreamsConfig : mrfStreamsConfig;
  const activeStreams = isAdvancedMode ? currentStreamConfig.filter(s => s.active) : currentStreamConfig.filter(s => s.isDefault);
  const totalPercentage = activeStreams.reduce((acc, s) => acc + Number(s.userWeight || 0), 0);
  const isValidTotal = totalPercentage === 100;
  const generateDisabled = (facilityType === 'MRF' || facilityType === 'MIXED_PLANT') && isAdvancedMode && !isValidTotal;

  const getSessionKey = () => {
    const cleanName = name.trim().toLowerCase().replace(/\s+/g, '_');
    const cleanState = selectedState.trim().toLowerCase().replace(/\s+/g, '_');
    return `crf_paid_${facilityType}_${cleanState}_${cleanName}_${selectedMonths.length}M`;
  };

  useEffect(() => {
    const rawData = localStorage.getItem(getSessionKey());
    if (rawData) {
      try {
        const parsed = JSON.parse(rawData);
        const TwelveHours = 12 * 60 * 60 * 1000;
        if (parsed.paid && (Date.now() - parsed.timestamp < TwelveHours)) {
          setIsPaid(true);
          return;
        }
      } catch (e) {
        if (rawData === 'true') {
          setIsPaid(true);
          return;
        }
      }
    }
    setIsPaid(false);
  }, [name, facilityType, selectedState, selectedMonths.length]);

  const markSessionPaid = () => {
    const memoryData = { paid: true, timestamp: Date.now() };
    localStorage.setItem(getSessionKey(), JSON.stringify(memoryData));
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
    
    let baseRate = 100;
    if (facilityType === 'INTEGRATED_3IN1') {
      baseRate = 500;
    } else if (facilityType === 'MIXED_PLANT') {
      baseRate = 200;
    } else if (isAdvancedMode) {
      baseRate = 150;
    }
    
    const baseTotal = billableMonths * baseRate;
    const effectiveFeeRate = 0.0236; 
    
    const finalTotalWithCharges = Math.round(baseTotal / (1 - effectiveFeeRate));
    const gatewayFee = finalTotalWithCharges - baseTotal;

    return { count, freeMonths, billableMonths, baseRate, baseTotal, gatewayFee, total: finalTotalWithCharges };
  };

  const pricing = getPricingDetails();

  const handleFacilityChange = (type) => {
    setFacilityType(type);
    if (type === 'ULB' || type === 'INTEGRATED_3IN1') setIsAdvancedMode(false);
  };

  const updateStreamConfig = (id, field, value) => {
    const setConfig = facilityType === 'MIXED_PLANT' ? setMixedStreamsConfig : setMrfStreamsConfig;
    setConfig(prev => prev.map(s => {
      if (s.id === id) {
        if (field === 'userWeight') {
          let num = value === '' ? '' : Number(value);
          if (num !== '') {
            if (num < s.min) num = s.min;
            if (num > s.max) num = s.max;
          }
          return { ...s, [field]: num };
        }
        return { ...s, [field]: value };
      }
      return s;
    }));
  };

  const handleGenerate = (e) => {
    e.preventDefault();
    if (generateDisabled) return;

    let monthlyDataMap = {};

    selectedMonths.forEach((m) => {
      const days = new Date(startYear, m, 0).getDate();
      
      let targetTons = (facilityType === 'ULB' || facilityType === 'INTEGRATED_3IN1')
        ? (ulbCalculationMode === 'population' ? (Number(population) * parsedPerCapita) / 1000000 : Number(actualAverageTpd))
        : Number(mrfDailyDryTons);

      const seedString = `${facilityType}-${selectedState}-${name}-${startYear}-${m}-${ulbCalculationMode}-${targetTons}-${segregationRate}`;
      const random = mulberry32(cyrb128(seedString));

      let logs = [];

      for (let day = 1; day <= days; day++) {
        const dateStr = `${startYear}-${String(m).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const dateObj = new Date(startYear, m - 1, day);
        const dayName = dateObj.toLocaleDateString('en-US', { weekday: 'short' });
        const isWeekend = [0, 6].includes(dateObj.getDay());

        let noise = 0.95 + random() * 0.10;
        if (isWeekend) noise *= 1.05;
        const dailyTotal = targetTons * noise;

        if (facilityType === 'INTEGRATED_3IN1') {
          const segFrac = segregationRate / 100;
          const unsegFrac = 1 - segFrac;

          // Gate Log Splits
          const segregatedTotal = dailyTotal * segFrac;
          const unsegregatedMixed = Number((dailyTotal * unsegFrac).toFixed(3));

          const wetSeg = Number((segregatedTotal * 0.60).toFixed(3));
          const drySeg = Number((segregatedTotal * 0.32).toFixed(3));
          const hazSeg = Number((segregatedTotal * 0.03).toFixed(3));
          const sanSeg = Number((segregatedTotal * 0.05).toFixed(3));

          // Pre-Sorting Log Splits
          const organicFines = Number((unsegregatedMixed * 0.45).toFixed(3));
          const dryOversize = Number((unsegregatedMixed * 0.35).toFixed(3));
          const heavyInerts = Number((unsegregatedMixed * 0.20).toFixed(3));

          // Compost Log Splits
          const totalCompostInput = Number((wetSeg + organicFines).toFixed(3));
          const moistureLoss = Number((totalCompostInput * 0.65).toFixed(3));
          const compostYield = Number((totalCompostInput * 0.18).toFixed(3));
          const compostRejects = Number((totalCompostInput - moistureLoss - compostYield).toFixed(3));

          // MRF Log Splits
          const totalMrfInput = Number((drySeg + dryOversize).toFixed(3));
          const pet = Number((totalMrfInput * 0.18).toFixed(3));
          const hdpe = Number((totalMrfInput * 0.12).toFixed(3));
          const paper = Number((totalMrfInput * 0.25).toFixed(3));
          const glassMetal = Number((totalMrfInput * 0.10).toFixed(3));
          const rdf = Number((totalMrfInput * 0.22).toFixed(3));
          const mrfRejects = Number((totalMrfInput - pet - hdpe - paper - glassMetal - rdf).toFixed(3));

          logs.push({
            date: dateStr, dayName, totalIntake: Number(dailyTotal.toFixed(3)),
            wetSeg, drySeg, hazSeg, sanSeg, unsegregatedMixed,
            organicFines, dryOversize, heavyInerts,
            totalCompostInput, moistureLoss, compostYield, compostRejects,
            totalMrfInput, pet, hdpe, paper, glassMetal, rdf, mrfRejects
          });
        } else if (facilityType === 'ULB') {
          const baseFractions = getSeasonalFractionsULB(m, currentRegionKey);
          let raw = baseFractions.map(r => r * (0.88 + random() * 0.24));
          let sum = raw.reduce((a, b) => a + b, 0);
          let norm = raw.map(r => r / sum);

          let c1 = Number((dailyTotal * norm[0]).toFixed(3));
          let c2 = Number((dailyTotal * norm[1]).toFixed(3));
          let c3 = Number((dailyTotal * norm[2]).toFixed(3));
          let c4 = Number((dailyTotal * norm[3]).toFixed(3));
          let c5 = Number((dailyTotal * norm[4]).toFixed(3));
          let c6 = Number((dailyTotal * norm[5]).toFixed(3));
          let exactTotal = Number((c1 + c2 + c3 + c4 + c5 + c6).toFixed(3));

          logs.push({ date: dateStr, dayName, c1, c2, c3, c4, c5, c6, total: exactTotal });
        } else {
          const baseNorm = activeStreams.map(s => Number(s.userWeight || 0) / 100);
          let raw = baseNorm.map(r => r * (0.88 + random() * 0.24));
          let dynamicSum = raw.reduce((a, b) => a + b, 0);
          let dynamicNorm = raw.map(r => r / dynamicSum);

          let rowStreams = {};
          let exactTotal = 0;

          activeStreams.forEach((stream, idx) => {
            let val = Number((dailyTotal * dynamicNorm[idx]).toFixed(3));
            rowStreams[stream.id] = val;
            exactTotal += val;
          });

          logs.push({ date: dateStr, dayName, streams: rowStreams, total: Number(exactTotal.toFixed(3)) });
        }
      }
      monthlyDataMap[m] = logs;
    });

    setGeneratedConfig({ type: facilityType, streams: activeStreams });
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

      const cashfree = window.Cashfree({ mode: import.meta.env.VITE_CASHFREE_MODE || 'production' });

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
          markSessionPaid();
          downloadMultiSheetExcel();
        }
      });
    } catch (err) {
      alert('Payment Error: ' + err.message);
      setIsProcessing(false);
    }
  };

  const handleRestoreAccess = async () => {
    if (!restoreOrderId.trim()) return alert("Please enter your Order ID");
    setIsRestoring(true);

    try {
      const res = await fetch('/api/verify-order', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ order_id: restoreOrderId.trim() })
      });

      const data = await res.json();

      if (data.success) {
        setIsPaid(true);
        setShowRestoreModal(false);
        markSessionPaid();
        alert("Payment Verified! Dataset unlocked.");
      } else {
        alert("Verification Failed: " + data.message);
      }
    } catch (err) {
      alert("Error contacting verification server.");
    }
    setIsRestoring(false);
  };

  const formatVal = (v) => displayUnit === 'kg' ? Math.round(v * 1000) : v.toFixed(3);

  const downloadMultiSheetExcel = () => {
    try {
      if (!generatedMonthlyData || !generatedConfig) {
        alert("Wait for the data to finish calculating.");
        return;
      }
      
      const u = displayUnit === 'kg' ? 'kg' : 'Tons';
      const wb = XLSX.utils.book_new();

      selectedMonths.forEach((mId) => {
        const monthName = MONTHS.find(m => m.id === mId)?.fullEn;

        if (generatedConfig.type === 'INTEGRATED_3IN1') {
          // Sheet 1: Gate Log
          const gateHeaders = ["Date", "Day", `Total Intake (${u})`, `Segregated Wet (${u})`, `Segregated Dry (${u})`, `Domestic Hazardous (${u})`, `Domestic Sanitary (${u})`, `Unsegregated Mixed (${u})`];
          const gateRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.totalIntake), formatVal(r.wetSeg), formatVal(r.drySeg), formatVal(r.hazSeg), formatVal(r.sanSeg), formatVal(r.unsegregatedMixed)]);
          XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([gateHeaders, ...gateRows]), `${monthName}_Gate`);

          // Sheet 2: Pre-Sorting Log
          const preHeaders = ["Date", "Day", `Mixed Intake (${u})`, `Fine Screen Fraction (${u})`, `Coarse Screen Fraction (${u})`, `Heavy Inerts (${u})`];
          const preRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.unsegregatedMixed), formatVal(r.organicFines), formatVal(r.dryOversize), formatVal(r.heavyInerts)]);
          XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([preHeaders, ...preRows]), `${monthName}_PreSort`);

          // Sheet 3: Compost Log
          const compostHeaders = ["Date", "Day", `Total Organic Input (${u})`, `Moisture Loss (${u})`, `Compost Yield (${u})`, `Rejects (${u})`];
          const compostRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.totalCompostInput), formatVal(r.moistureLoss), formatVal(r.compostYield), formatVal(r.compostRejects)]);
          XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([compostHeaders, ...compostRows]), `${monthName}_Compost`);

          // Sheet 4: MRF Log
          const mrfHeaders = ["Date", "Day", `Total Dry Input (${u})`, `PET (${u})`, `HDPE (${u})`, `Paper/Cardboard (${u})`, `Glass & Metal (${u})`, `RDF (${u})`, `Rejects (${u})`];
          const mrfRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.totalMrfInput), formatVal(r.pet), formatVal(r.hdpe), formatVal(r.paper), formatVal(r.glassMetal), formatVal(r.rdf), formatVal(r.mrfRejects)]);
          XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([mrfHeaders, ...mrfRows]), `${monthName}_MRF`);

        } else {
          let headers = [];
          if (generatedConfig.type === 'ULB') {
            headers = ["Date", "Day", `Wet (${u})`, `Dry (${u})`, `Sanitary (${u})`, `Special Care (${u})`, `C&D (${u})`, `Inerts (${u})`, `Total (${u})`];
          } else {
            headers = ["Date", "Day", ...generatedConfig.streams.map(s => `${s.label} (${u})`), `Total Processing (${u})`];
          }

          const sheetData = [headers, ...generatedMonthlyData[mId].map(r => {
            if (generatedConfig.type === 'ULB') {
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

      const firstMonthObj = MONTHS.find(m => m.id === selectedMonths[0]);
      const monthStr = firstMonthObj ? firstMonthObj.shortEn : 'Jan';
      const yearStr = String(startYear).slice(-2);
      const safeName = name.trim().replace(/[^a-zA-Z0-9]/g, '_').replace(/_+/g, '_');
      
      const customFileName = `${facilityType}_${safeName}_${monthStr}_${yearStr}.xlsx`;
      XLSX.writeFile(wb, customFileName);

    } catch (error) {
      console.error(error);
      alert("Download blocked: " + error.message);
    }
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
                {lang === 'hi' ? 'एकीकृत 3-इन-1 यूएलबी, एमआरएफ एवं कम्पोस्ट लॉग-बुक जनरेटर' : 'Integrated 3-in-1 ULB, MRF & Compost Logbook Generator'}
              </h1>
              <p style={{ fontSize: '13px', margin: 0, color: '#a7f3d0' }}>
                {lang === 'hi' ? 'गेट एंट्री, प्री-सॉर्टिंग, खाद एवं एमआरएफ मास्टर लॉग-बुक इंजन || **केवल एजुकेशन व् एस्टीमेशन हेतु' : 'Master Gate, Pre-Sorting, Compost & MRF Logbook Engine || ** For Educational & Estimation Purpose'}
              </p>
            </div>
            <button type="button" onClick={() => setLang(lang === 'hi' ? 'en' : 'hi')} style={{ padding: '6px 12px', background: '#fff', color: '#047857', border: 'none', borderRadius: '5px', fontWeight: 'bold', cursor: 'pointer', fontSize: '13px' }}>
              <Globe size={15} style={{ verticalAlign: 'middle', marginRight: '4px' }} /> {lang === 'hi' ? 'English' : 'हिंदी'}
            </button>
          </div>
        </div>

        {/* FORM SECTION */}
        <div style={{ padding: '15px', maxWidth: '1000px', margin: '0 auto' }}>
          <form onSubmit={handleGenerate} style={{ background: '#fff', padding: '16px', borderRadius: '8px', border: '1px solid #cbd5e1', marginBottom: '20px' }}>
            
            <div style={{ marginBottom: '14px', display: 'flex', gap: '15px', alignItems: 'center', fontSize: '14px', flexWrap: 'wrap' }}>
              <strong>{lang === 'hi' ? 'सुविधा प्रकार:' : 'Facility Type:'}</strong>
              <label style={{ cursor: 'pointer', color: '#047857', fontWeight: 'bold' }}>
                <input type="radio" value="INTEGRATED_3IN1" checked={facilityType === 'INTEGRATED_3IN1'} onChange={() => handleFacilityChange('INTEGRATED_3IN1')} /> Integrated 3-in-1 Master Suite (₹500/mo)
              </label>
              <label style={{ cursor: 'pointer' }}><input type="radio" value="ULB" checked={facilityType === 'ULB'} onChange={() => handleFacilityChange('ULB')} /> ULB Only</label>
              <label style={{ cursor: 'pointer' }}><input type="radio" value="MRF" checked={facilityType === 'MRF'} onChange={() => handleFacilityChange('MRF')} /> MRF Centre Only</label>
              <label style={{ cursor: 'pointer' }}><input type="radio" value="MIXED_PLANT" checked={facilityType === 'MIXED_PLANT'} onChange={() => handleFacilityChange('MIXED_PLANT')} /> Mixed Waste Plant (Screening)</label>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '12px', marginBottom: '14px' }}>
              <div>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                  <label style={{ fontSize: '12px', fontWeight: '600', color: '#047857', display: 'flex', alignItems: 'center', gap: '4px' }}>
                    <MapPin size={13} /> {lang === 'hi' ? 'राज्य चुनें (State)' : 'Select State'}
                  </label>
                  <button type="button" onClick={() => setShowStateInfo(!showStateInfo)} style={{ background: 'none', border: 'none', color: '#0284c7', cursor: 'pointer', padding: '0 2px', display: 'flex', alignItems: 'center', gap: '2px', fontSize: '12px', fontWeight: 'bold' }}>
                    <Info size={14} />
                  </button>
                </div>
                <select style={{ ...inputStyle, border: '1px solid #059669', background: '#f0fdf4' }} value={selectedState} onChange={(e) => setSelectedState(e.target.value)}>
                  {STATES_LIST.map((s) => <option key={s.nameEn} value={s.nameEn}>{lang === 'hi' ? s.nameHi : s.nameEn}</option>)}
                </select>
              </div>

              <div>
                <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'संयंत्र / निकाय का नाम' : 'Plant / ULB Name'}</label>
                <input style={inputStyle} type="text" required value={name} onChange={(e) => setName(e.target.value)} />
              </div>

              <div>
                <label style={{ fontSize: '12px', fontWeight: '600', display: 'flex', alignItems: 'center', gap: '4px' }}>
                  <Phone size={12} /> {lang === 'hi' ? 'मोबाइल नंबर (रसीद के लिए)' : 'Mobile Number (for Receipt)'}
                </label>
                <input style={inputStyle} type="tel" maxLength={10} placeholder="9876543210" required value={phone} onChange={(e) => setPhone(e.target.value.replace(/\D/g, ''))} />
              </div>

              {(facilityType === 'ULB' || facilityType === 'INTEGRATED_3IN1') ? (
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

                  {facilityType === 'INTEGRATED_3IN1' && (
                    <div style={{ gridColumn: '1 / -1', background: '#ecfdf5', padding: '12px', borderRadius: '6px', border: '1px solid #a7f3d0', marginTop: '6px' }}>
                      <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#065f46' }}>
                        {lang === 'hi' ? 'स्रोत पृथक्करण दर (%)' : 'Source Segregation Efficiency (%)'}: {segregationRate}%
                      </label>
                      <input type="range" min="20" max="95" step="5" value={segregationRate} onChange={(e) => setSegregationRate(Number(e.target.value))} style={{ width: '100%', marginTop: '6px' }} />
                      <div style={{ fontSize: '11px', color: '#047857', marginTop: '4px' }}>
                        {segregationRate}% Segregated (Sent to Compost & MRF) | {100 - segregationRate}% Mixed Waste (Pre-screened at site front-end)
                      </div>
                    </div>
                  )}
                </>
              ) : (
                <>
                  <div>
                    <label style={{ fontSize: '12px', fontWeight: '600' }}>
                      {facilityType === 'MIXED_PLANT' ? (lang === 'hi' ? 'दैनिक मिश्रित कचरा आवक (TPD)' : 'Daily Mixed Waste Intake (TPD)') : (lang === 'hi' ? 'दैनिक सूखा कचरा आवक (TPD)' : 'Daily Dry Waste Input (TPD)')}
                    </label>
                    <input style={inputStyle} type="number" step="0.01" required value={mrfDailyDryTons} onChange={(e) => setMrfDailyDryTons(e.target.value)} />
                  </div>
                  <div>
                    <label style={{ fontSize: '12px', fontWeight: '600' }}>{lang === 'hi' ? 'संयंत्र क्षमता (TPD)' : 'Plant Capacity (TPD)'}</label>
                    <input style={inputStyle} type="number" step="0.01" required value={mrfMaxCapacityTons} onChange={(e) => setMrfMaxCapacityTons(e.target.value)} />
                  </div>
                  <div style={{ gridColumn: '1 / -1', fontSize: '13px', fontWeight: 'bold', display: 'flex', alignItems: 'center', flexWrap: 'wrap' }}>
                    {lang === 'hi' ? 'क्षमता उपयोग (Capacity Utilization):' : 'Capacity Utilization:'} {mrfCapacityUtilization.toFixed(1)}%
                    {mrfCapacityUtilization > 100 && (
                      <span style={{ color: '#dc2626', marginLeft: '6px', fontSize: '12px' }}>(Warning: Daily intake exceeds plant design capacity.)</span>
                    )}
                  </div>

                  {/* ADVANCED CONFIGURATION WITH OTHER/CUSTOM FRACTION BOX */}
                  <div style={{ gridColumn: '1 / -1', marginTop: '10px', background: isAdvancedMode ? '#fffbeb' : '#f8fafc', padding: '12px', borderRadius: '6px', border: '1px solid #cbd5e1' }}>
                    <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontWeight: 'bold', color: '#0f172a', cursor: 'pointer', fontSize: '13px' }}>
                      <input type="checkbox" checked={isAdvancedMode} onChange={(e) => setIsAdvancedMode(e.target.checked)} style={{ width: '16px', height: '16px', cursor: 'pointer' }} />
                      {lang === 'hi' 
                        ? `!! ADVANCE Log-Book !! एडवांस्ड ${facilityType === 'MIXED_PLANT' ? 'स्क्रीनिंग (₹200/माह)' : 'एमआरएफ (₹150/माह)'} कॉन्फ़िगरेशन सक्षम करें` 
                        : `Enable Advanced ${facilityType === 'MIXED_PLANT' ? 'Mixed Waste Screening (₹200/mo)' : 'MRF (₹150/mo)'} Configuration`}
                    </label>

                    {isAdvancedMode && (
                      <div style={{ marginTop: '14px', borderTop: '1px solid #e2e8f0', paddingTop: '12px' }}>
                        <p style={{ fontSize: '12px', color: '#475569', margin: '0 0 10px 0' }}>
                          {facilityType === 'MIXED_PLANT'
                            ? (lang === 'hi' ? 'फ़ाइन स्क्रीन एवं कोर्स स्क्रीन पृथक्करण प्रतिशत दर्ज करें।' : 'Set mass-balance splits based on fine screen and coarse screen fractions.')
                            : (lang === 'hi' ? 'विशिष्ट पृथक्कृत स्ट्रीम चुनें और सांकेतिक प्रतिशत दर्ज करें।' : 'Select sorted dry streams and input composition percentages.')}
                        </p>
                        
                        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', gap: '10px' }}>
                          {currentStreamConfig.map(stream => (
                            <div key={stream.id} style={{ display: 'flex', flexDirection: 'column', background: '#fff', padding: '8px', border: '1px solid #e2e8f0', borderRadius: '6px', opacity: stream.active ? 1 : 0.6 }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <input type="checkbox" checked={stream.active} onChange={(e) => updateStreamConfig(stream.id, 'active', e.target.checked)} style={{ cursor: 'pointer' }} />
                                {stream.id.startsWith('other_') ? (
                                  <input 
                                    type="text" 
                                    value={stream.label} 
                                    onChange={(e) => updateStreamConfig(stream.id, 'label', e.target.value)}
                                    placeholder="Custom Fraction Name"
                                    style={{ fontSize: '11px', fontWeight: 'bold', padding: '2px 4px', border: '1px solid #cbd5e1', borderRadius: '3px', width: '100%' }}
                                  />
                                ) : (
                                  <span style={{ fontSize: '11px', fontWeight: 'bold', color: '#1e293b' }}>{stream.label}</span>
                                )}
                              </div>
                              <div style={{ display: 'flex', alignItems: 'center', gap: '4px', marginTop: '6px' }}>
                                <input type="number" min={stream.min} max={stream.max} value={stream.userWeight} onChange={(e) => updateStreamConfig(stream.id, 'userWeight', e.target.value)} disabled={!stream.active} style={{ width: '100%', fontSize: '12px', padding: '4px', border: '1px solid #cbd5e1', borderRadius: '4px' }} />
                                <span style={{ fontSize: '11px', color: '#64748b' }}>%</span>
                              </div>
                            </div>
                          ))}
                        </div>

                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '15px', padding: '10px 15px', borderRadius: '6px', background: isValidTotal ? '#dcfce7' : '#fee2e2', border: isValidTotal ? '1px solid #22c55e' : '1px solid #ef4444' }}>
                           <span style={{ fontWeight: 'bold', color: isValidTotal ? '#166534' : '#991b1b', fontSize: '14px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                              {!isValidTotal && <AlertCircle size={16} />} 
                              {lang === 'hi' ? 'कुल द्रव्यमान संतुलन:' : 'Total Mass Balance:'} {totalPercentage}%
                           </span>
                           {!isValidTotal && (
                             <span style={{ fontSize: '12px', color: '#991b1b', fontWeight: 'bold' }}>
                               {lang === 'hi' ? 'आगे बढ़ने के लिए कुल योग सटीक 100% होना चाहिए।' : 'Total mass balance must equal exactly 100% to proceed.'}
                             </span>
                           )}
                        </div>
                      </div>
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
                </ul>
              </div>
            )}

            <div style={{ marginBottom: '14px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px', flexWrap: 'wrap' }}>
                <strong style={{ fontSize: '13px' }}>{lang === 'hi' ? 'माह चुनें:' : 'Select Months:'}</strong>
                <span style={{ color: '#059669', fontWeight: 'bold', fontSize: '13px', background: '#ecfdf5', padding: '4px 10px', borderRadius: '4px', border: '1px solid #a7f3d0' }}>
                  {pricing.count} {lang === 'hi' ? 'माह' : 'Month/s'} — ₹{pricing.total} 
                  <span style={{ fontSize: '10px', color: '#047857', marginLeft: '4px' }}>
                    (₹{pricing.baseTotal} + ₹{pricing.gatewayFee} PG Fee)
                  </span>
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

            <button type="submit" disabled={generateDisabled} style={{ width: '100%', padding: '12px', background: generateDisabled ? '#94a3b8' : '#059669', color: '#fff', border: 'none', borderRadius: '6px', cursor: generateDisabled ? 'not-allowed' : 'pointer', fontWeight: 'bold', fontSize: '15px' }}>
              {lang === 'hi' ? '3-इन-1 मास्टर लॉग-बुक जनरेट करें →' : 'Generate Master Logbook Dataset →'}
            </button>
          </form>

          {generatedMonthlyData && generatedConfig && (
            <div ref={resultsRef} style={{ background: '#fff', padding: '15px', borderRadius: '8px', border: '1px solid #cbd5e1', scrollMarginTop: '15px' }}>
              
              <div style={{ background: '#eff6ff', border: '1px solid #bfdbfe', padding: '10px', borderRadius: '6px', marginBottom: '12px', fontSize: '12px', color: '#1e40af' }}>
                <strong><Info size={14} style={{ verticalAlign: 'middle' }} /> Modelling configuration applied ({currentStateObj.nameEn}):</strong> 
                {generatedConfig.type === 'INTEGRATED_3IN1' ? ' Integrated 4-Tab Master Suite (Gate, Pre-Sort, Compost, MRF) generated with fine & coarse screen mass-balance.' : ' Single facility model applied.'}
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
                <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                  <button onClick={() => setDisplayUnit(displayUnit === 'Tons' ? 'kg' : 'Tons')} style={{ padding: '6px 10px', fontSize: '12px', borderRadius: '4px', border: '1px solid #cbd5e1' }}>
                    Unit: <strong>{displayUnit}</strong>
                  </button>
                  {isPaid ? (
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      <button onClick={downloadMultiSheetExcel} style={{ padding: '6px 12px', background: '#0f172a', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold' }}>
                        <Download size={13} /> Export Master Excel (.xlsx)
                      </button>
                    </div>
                  ) : (
                    <button onClick={handlePayment} disabled={isProcessing} style={{ padding: '6px 12px', background: '#059669', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold' }}>
                      {isProcessing ? 'Wait...' : `Pay ₹${pricing.total} to Unlock`}
                    </button>
                  )}
                </div>
              </div>

              {/* TABLE PREVIEW */}
              <div onContextMenu={(e) => e.preventDefault()} style={{ overflowX: 'auto', border: '1px solid #cbd5e1', borderRadius: '4px', userSelect: isPaid ? 'text' : 'none' }}>
                <table cellPadding="8" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', minWidth: '700px' }}>
                  <thead>
                    <tr style={{ background: '#f1f5f9', borderBottom: '1px solid #cbd5e1' }}>
                      <th>Date</th><th>Day</th>
                      {generatedConfig.type === 'INTEGRATED_3IN1' ? (
                        <>
                          <th>Total Gate Intake</th>
                          <th style={{ background: '#ecfdf5' }}>Seg. Wet Waste</th>
                          <th>Seg. Dry Waste</th>
                          <th>Mixed Waste (Unsegregated)</th>
                          <th>Compost Yield</th>
                          <th>MRF Recyclables</th>
                        </>
                      ) : (
                        <>
                          {generatedConfig.streams?.map(s => <th key={s.id}>{s.label} ({displayUnit})</th>)}
                          <th>Total ({displayUnit})</th>
                        </>
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {visibleRows.map((r, idx) => (
                      <tr key={idx} style={{ borderBottom: '1px solid #e2e8f0' }}>
                        <td>{r.date}</td><td>{r.dayName}</td>
                        {generatedConfig.type === 'INTEGRATED_3IN1' ? (
                          <>
                            <td><strong>{formatVal(r.totalIntake)}</strong></td>
                            <td style={{ background: '#ecfdf5' }}>{formatVal(r.wetSeg)}</td>
                            <td>{formatVal(r.drySeg)}</td>
                            <td>{formatVal(r.unsegregatedMixed)}</td>
                            <td>{formatVal(r.compostYield)}</td>
                            <td>{formatVal(r.pet + r.hdpe + r.paper)}</td>
                          </>
                        ) : (
                          generatedConfig.streams?.map(s => <td key={s.id}>{formatVal(r.streams[s.id])}</td>)
                        )}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {!isPaid && (
                <div style={{ border: '2px dashed #059669', background: '#ecfdf5', padding: '15px', textAlign: 'center', marginTop: '12px', borderRadius: '6px' }}>
                  <Lock style={{ color: '#059669' }} size={18} />
                  <h4 style={{ margin: '4px 0', color: '#065f46', fontSize: '15px' }}>Preview Locked (Days 1–5 Only)</h4>
                  <p style={{ margin: '4px 0 10px 0', color: '#047857', fontSize: '13px' }}>Pay ₹{pricing.total} to unlock complete 4-tab Master Excel workbook.</p>
                  
                  <button onClick={handlePayment} disabled={isProcessing} style={{ padding: '10px 20px', background: '#059669', color: '#fff', border: 'none', borderRadius: '6px', fontSize: '14px', fontWeight: 'bold', cursor: 'pointer' }}>
                    {isProcessing ? 'Connecting...' : `Pay ₹${pricing.total} & Download Complete Master File`}
                  </button>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
