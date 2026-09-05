import React, { useState, useRef, useEffect } from 'react';
import { Building2, Download, Lock, Globe, Check, Info, ShieldCheck, MapPin, AlertCircle, Phone, Plus, Trash2, ArrowLeft, Layers, Sparkles } from 'lucide-react';
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
  const [appMode, setAppMode] = useState('STANDALONE'); // 'STANDALONE' or 'INTEGRATED_3IN1'
  const [facilityType, setFacilityType] = useState('ULB'); // Options for Standalone: 'ULB', 'MRF', 'MIXED_PLANT'
  const [selectedState, setSelectedState] = useState('Uttar Pradesh');
  const [name, setName] = useState('Nagar Palika Parishad');
  const [phone, setPhone] = useState('');
  const [isAdvancedMode, setIsAdvancedMode] = useState(false);
  
  // Standalone ULB & Integrated State
  const [ulbCalculationMode, setUlbCalculationMode] = useState('population');
  const [population, setPopulation] = useState(50000);
  const [perCapitaOption, setPerCapitaOption] = useState('450');
  const [customPerCapita, setCustomPerCapita] = useState('');
  const [actualAverageTpd, setActualAverageTpd] = useState(10);
  const [referencePeriod, setReferencePeriod] = useState(30);
  const [segregationRate, setSegregationRate] = useState(80);

  // Standalone MRF & Mixed Plant State
  const [mrfDailyDryTons, setMrfDailyDryTons] = useState(15);
  const [mrfMaxCapacityTons, setMrfMaxCapacityTons] = useState(25);
  const [mrfStreamsConfig, setMrfStreamsConfig] = useState(
    DEFAULT_MRF_STREAMS.map(s => ({ ...s, active: s.isDefault, userWeight: s.defaultWeight }))
  );
  const [mixedStreamsConfig, setMixedStreamsConfig] = useState(
    DEFAULT_MIXED_STREAMS.map(s => ({ ...s, active: s.isDefault, userWeight: s.defaultWeight }))
  );

  // Dynamic Assets for Integrated 3-in-1 Mode
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
  const estimatedDailyWaste = ((Number(population) * parsedPerCapita) / 1000000).toFixed(2);
  const mrfCapacityUtilization = Number(mrfMaxCapacityTons) > 0 ? (Number(mrfDailyDryTons) / Number(mrfMaxCapacityTons)) * 100 : 0;

  // Stream Selection & Mass Balance Logic
  const currentStreamConfig = facilityType === 'MIXED_PLANT' ? mixedStreamsConfig : mrfStreamsConfig;
  const activeStreams = isAdvancedMode ? currentStreamConfig.filter(s => s.active) : currentStreamConfig.filter(s => s.isDefault);
  const totalPercentage = activeStreams.reduce((acc, s) => acc + Number(s.userWeight || 0), 0);
  const isValidTotal = totalPercentage === 100;
  const generateDisabled = appMode === 'STANDALONE' && (facilityType === 'MRF' || facilityType === 'MIXED_PLANT') && isAdvancedMode && !isValidTotal;

  // Dynamic Array Handlers for Integrated Mode
  const addCompostUnit = () => {
    const newId = `c${compostUnits.length + 1}`;
    setCompostUnits([...compostUnits, { id: newId, label: `Compost Unit ${compostUnits.length + 1}`, type: 'Vermicompost Pit', capacity: 5 }]);
  };

  const removeCompostUnit = (id) => {
    if (compostUnits.length > 1) setCompostUnits(compostUnits.filter(u => u.id !== id));
  };

  const updateCompostUnit = (id, field, value) => {
    setCompostUnits(compostUnits.map(u => u.id === id ? { ...u, [field]: value } : u));
  };

  const addMrfUnit = () => {
    const newId = `m${mrfUnits.length + 1}`;
    setMrfUnits([...mrfUnits, { id: newId, label: `MRF Shed ${mrfUnits.length + 1}`, type: 'Manual Sorting Shed', capacity: 5 }]);
  };

  const removeMrfUnit = (id) => {
    if (mrfUnits.length > 1) setMrfUnits(mrfUnits.filter(u => u.id !== id));
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
    
    let baseRate = 100;
    if (appMode === 'INTEGRATED_3IN1') {
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

          // Proportional Multi-Unit Distribution
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

          logs.push({ date: dateStr, dayName, c1, c2, c3, c4, c5, c6, total: Number((c1 + c2 + c3 + c4 + c5 + c6).toFixed(3)) });
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

    setGeneratedConfig({ appMode, type: facilityType, streams: activeStreams });
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

      if (appMode === 'INTEGRATED_3IN1') {
        const gateHeaders = ["Date", "Day", `Total Gate Intake (${u})`, `Segregated Wet (${u})`, `Segregated Dry (${u})`, `Domestic Hazardous (${u})`, `Domestic Sanitary (${u})`, `Unsegregated Mixed (${u})`];
        const gateRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.totalIntake), formatVal(r.wetSeg), formatVal(r.drySeg), formatVal(r.hazSeg), formatVal(r.sanSeg), formatVal(r.unsegregatedMixed)]);
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([gateHeaders, ...gateRows]), `${monthName}_Gate`);

        const preHeaders = ["Date", "Day", `Mixed Intake (${u})`, `Fine Screen Fraction (${u})`, `Coarse Screen Fraction (${u})`, `Heavy Inerts (${u})`];
        const preRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.unsegregatedMixed), formatVal(r.organicFines), formatVal(r.dryOversize), formatVal(r.heavyInerts)]);
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([preHeaders, ...preRows]), `${monthName}_PreSort`);

        compostUnits.forEach(unit => {
          const cHeaders = ["Date", "Day", `Unit Feed (${u})`, `Compost Yield (${u})`];
          const cRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.compostUnitBreakdown[unit.id]?.feed), formatVal(r.compostUnitBreakdown[unit.id]?.compostYield)]);
          XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([cHeaders, ...cRows]), `${monthName}_${unit.label.substring(0, 10)}`);
        });

        mrfUnits.forEach(unit => {
          const mHeaders = ["Date", "Day", `Unit Feed (${u})`, `Sorted Recyclables (${u})`, `RDF Dispatched (${u})`];
          const mRows = generatedMonthlyData[mId].map(r => [r.date, r.dayName, formatVal(r.mrfUnitBreakdown[unit.id]?.feed), formatVal(r.mrfUnitBreakdown[unit.id]?.recyclables), formatVal(r.mrfUnitBreakdown[unit.id]?.rdf)]);
          XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([mHeaders, ...mRows]), `${monthName}_${unit.label.substring(0, 10)}`);
        });
      } else {
        let headers = facilityType === 'ULB' 
          ? ["Date", "Day", `Wet (${u})`, `Dry (${u})`, `Sanitary (${u})`, `Special Care (${u})`, `C&D (${u})`, `Inerts (${u})`, `Total (${u})`]
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
        
        {/* HEADER BANNER */}
        <div style={{ background: 'linear-gradient(135deg, #064e3b 0%, #047857 100%)', color: '#fff', padding: '20px', borderRadius: '8px', marginBottom: '16px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap' }}>
            <div>
              <span style={{ background: 'rgba(255,255,255,0.2)', padding: '3px 8px', borderRadius: '12px', fontSize: '11px', fontWeight: 'bold' }}>
                <ShieldCheck size={12} style={{ verticalAlign: 'middle' }} /> SWM ESTIMATION ENGINE
              </span>
              <h1 style={{ fontSize: '22px', margin: '6px 0 2px 0', fontWeight: '800' }}>
                <Building2 size={22} style={{ verticalAlign: 'middle', marginRight: '6px' }} />
                {appMode === 'INTEGRATED_3IN1' ? 'Integrated 3-in-1 Multi-Unit Logbook Suite' : 'ULB, MRF & Mixed Waste Logbook Generator'}
              </h1>
            </div>
            <button type="button" onClick={() => setLang(lang === 'hi' ? 'en' : 'hi')} style={{ padding: '6px 12px', background: '#fff', color: '#047857', border: 'none', borderRadius: '5px', fontWeight: 'bold', cursor: 'pointer' }}>
              <Globe size={15} style={{ verticalAlign: 'middle' }} /> {lang === 'hi' ? 'English' : 'हिंदी'}
            </button>
          </div>
        </div>

        {/* SWITCHER BANNER TO INTEGRATED SUITE */}
        {appMode === 'STANDALONE' ? (
          <div style={{ background: 'linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%)', border: '2px solid #22c55e', padding: '14px 18px', borderRadius: '8px', marginBottom: '16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
            <div>
              <span style={{ background: '#15803d', color: '#fff', padding: '2px 8px', borderRadius: '4px', fontSize: '11px', fontWeight: 'bold', display: 'inline-flex', alignItems: 'center', gap: '4px' }}>
                <Sparkles size={12} /> NEW FEATURE
              </span>
              <h3 style={{ margin: '4px 0 2px 0', color: '#166534', fontSize: '15px', fontWeight: 'bold' }}>Need Interconnected Logbooks for Gate, Compost & MRF?</h3>
              <p style={{ margin: 0, color: '#15803d', fontSize: '12px' }}>Generate synchronized 4-tab Excel logbooks for complete resource recovery facilities (₹500/mo).</p>
            </div>
            <button type="button" onClick={() => setAppMode('INTEGRATED_3IN1')} style={{ padding: '8px 16px', background: '#166534', color: '#fff', border: 'none', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px' }}>
              <Layers size={14} /> Open Integrated 3-in-1 Suite →
            </button>
          </div>
        ) : (
          <button type="button" onClick={() => setAppMode('STANDALONE')} style={{ marginBottom: '16px', padding: '6px 12px', background: '#e2e8f0', color: '#334155', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold', display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
            <ArrowLeft size={14} /> Back to Single-Facility Logbooks
          </button>
        )}

        {/* FORM SECTION */}
        <form onSubmit={handleGenerate} style={{ background: '#fff', padding: '20px', borderRadius: '8px', border: '1px solid #cbd5e1', marginBottom: '20px' }}>
          
          {/* FACILITY SELECTION FOR STANDALONE MODE */}
          {appMode === 'STANDALONE' && (
            <div style={{ marginBottom: '14px', display: 'flex', gap: '15px', alignItems: 'center', fontSize: '14px', flexWrap: 'wrap', background: '#f8fafc', padding: '10px', borderRadius: '6px', border: '1px solid #e2e8f0' }}>
              <strong>Select Standalone Logbook:</strong>
              <label style={{ cursor: 'pointer' }}><input type="radio" value="ULB" checked={facilityType === 'ULB'} onChange={() => setFacilityType('ULB')} /> ULB Collection (₹100/mo)</label>
              <label style={{ cursor: 'pointer' }}><input type="radio" value="MRF" checked={facilityType === 'MRF'} onChange={() => setFacilityType('MRF')} /> MRF Centre (₹100–150/mo)</label>
              <label style={{ cursor: 'pointer' }}><input type="radio" value="MIXED_PLANT" checked={facilityType === 'MIXED_PLANT'} onChange={() => setFacilityType('MIXED_PLANT')} /> Mixed Waste Plant (₹200/mo)</label>
            </div>
          )}

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '12px', marginBottom: '14px' }}>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>Select State</label>
              <select style={inputStyle} value={selectedState} onChange={(e) => setSelectedState(e.target.value)}>
                {STATES_LIST.map((s) => <option key={s.nameEn} value={s.nameEn}>{lang === 'hi' ? s.nameHi : s.nameEn}</option>)}
              </select>
            </div>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>ULB / Facility Name</label>
              <input style={inputStyle} type="text" required value={name} onChange={(e) => setName(e.target.value)} />
            </div>
            <div>
              <label style={{ fontSize: '12px', fontWeight: '600' }}>Mobile Number</label>
              <input style={inputStyle} type="tel" maxLength={10} placeholder="9876543210" required value={phone} onChange={(e) => setPhone(e.target.value.replace(/\D/g, ''))} />
            </div>

            {(appMode === 'INTEGRATED_3IN1' || facilityType === 'ULB') && (
              <>
                <div style={{ gridColumn: '1 / -1', background: '#f1f5f9', padding: '10px', borderRadius: '6px', border: '1px solid #e2e8f0' }}>
                  <strong style={{ fontSize: '13px' }}>Waste Generation Estimation Basis</strong>
                  <div style={{ display: 'flex', gap: '15px', marginTop: '6px', fontSize: '13px' }}>
                    <label style={{ cursor: 'pointer' }}><input type="radio" checked={ulbCalculationMode === 'population'} onChange={() => setUlbCalculationMode('population')} /> Population Based</label>
                    <label style={{ cursor: 'pointer' }}><input type="radio" checked={ulbCalculationMode === 'actual'} onChange={() => setUlbCalculationMode('actual')} /> Actual TPD</label>
                  </div>
                </div>

                {ulbCalculationMode === 'population' ? (
                  <>
                    <div>
                      <label style={{ fontSize: '12px', fontWeight: '600' }}>Population</label>
                      <input style={inputStyle} type="number" value={population} onChange={(e) => setPopulation(e.target.value)} />
                    </div>
                    <div>
                      <label style={{ fontSize: '12px', fontWeight: '600' }}>Per Capita Rate (g/day)</label>
                      <select style={inputStyle} value={perCapitaOption} onChange={(e) => setPerCapitaOption(e.target.value)}>
                        <option value="300">300 g/person/day</option>
                        <option value="450">450 g/person/day</option>
                        <option value="500">500 g/person/day</option>
                      </select>
                    </div>
                  </>
                ) : (
                  <div>
                    <label style={{ fontSize: '12px', fontWeight: '600' }}>Observed Waste Generation (TPD)</label>
                    <input style={inputStyle} type="number" value={actualAverageTpd} onChange={(e) => setActualAverageTpd(e.target.value)} />
                  </div>
                )}
              </>
            )}

            {appMode === 'STANDALONE' && (facilityType === 'MRF' || facilityType === 'MIXED_PLANT') && (
              <>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: '600' }}>Daily Intake (TPD)</label>
                  <input style={inputStyle} type="number" value={mrfDailyDryTons} onChange={(e) => setMrfDailyDryTons(e.target.value)} />
                </div>
                <div>
                  <label style={{ fontSize: '12px', fontWeight: '600' }}>Plant Capacity (TPD)</label>
                  <input style={inputStyle} type="number" value={mrfMaxCapacityTons} onChange={(e) => setMrfMaxCapacityTons(e.target.value)} />
                </div>
              </>
            )}
          </div>

          {/* INTEGRATED MODE SPECIFIC MULTI-ASSET INPUTS */}
          {appMode === 'INTEGRATED_3IN1' && (
            <>
              <div style={{ background: '#ecfdf5', padding: '12px', borderRadius: '6px', border: '1px solid #a7f3d0', marginBottom: '14px' }}>
                <label style={{ fontSize: '12px', fontWeight: 'bold', color: '#065f46' }}>Source Segregation Rate (%): {segregationRate}%</label>
                <input type="range" min="20" max="95" step="5" value={segregationRate} onChange={(e) => setSegregationRate(Number(e.target.value))} style={{ width: '100%', marginTop: '6px' }} />
              </div>

              <div style={{ background: '#f0fdf4', padding: '12px', borderRadius: '6px', border: '1px solid #bbf7d0', marginBottom: '14px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                  <strong style={{ fontSize: '12px', color: '#166534' }}>Composting Assets (Wet Waste Line)</strong>
                  <button type="button" onClick={addCompostUnit} style={{ padding: '4px 8px', background: '#166534', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '11px', display: 'flex', alignItems: 'center', gap: '4px' }}><Plus size={12} /> Add Unit</button>
                </div>
                {compostUnits.map((u) => (
                  <div key={u.id} style={{ display: 'grid', gridTemplateColumns: '2fr 2fr 1fr auto', gap: '6px', marginBottom: '6px' }}>
                    <input type="text" value={u.label} onChange={(e) => updateCompostUnit(u.id, 'label', e.target.value)} style={{ ...inputStyle, marginTop: 0 }} />
                    <select value={u.type} onChange={(e) => updateCompostUnit(u.id, 'type', e.target.value)} style={{ ...inputStyle, marginTop: 0 }}>
                      <option value="Windrow Pad">Windrow Pad</option>
                      <option value="Vermicompost Pit">Vermicompost Pit</option>
                    </select>
                    <input type="number" value={u.capacity} onChange={(e) => updateCompostUnit(u.id, 'capacity', e.target.value)} style={{ ...inputStyle, marginTop: 0 }} placeholder="TPD" />
                    {compostUnits.length > 1 && <button type="button" onClick={() => removeCompostUnit(u.id)} style={{ color: '#dc2626', background: 'none', border: 'none', cursor: 'pointer' }}><Trash2 size={14} /></button>}
                  </div>
                ))}
              </div>

              <div style={{ background: '#f0f9ff', padding: '12px', borderRadius: '6px', border: '1px solid #bae6fd', marginBottom: '14px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                  <strong style={{ fontSize: '12px', color: '#0369a1' }}>MRF / Sorting Shed Assets (Dry Waste Line)</strong>
                  <button type="button" onClick={addMrfUnit} style={{ padding: '4px 8px', background: '#0284c7', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '11px', display: 'flex', alignItems: 'center', gap: '4px' }}><Plus size={12} /> Add Shed</button>
                </div>
                {mrfUnits.map((u) => (
                  <div key={u.id} style={{ display: 'grid', gridTemplateColumns: '2fr 2fr 1fr auto', gap: '6px', marginBottom: '6px' }}>
                    <input type="text" value={u.label} onChange={(e) => updateMrfUnit(u.id, 'label', e.target.value)} style={{ ...inputStyle, marginTop: 0 }} />
                    <select value={u.type} onChange={(e) => updateMrfUnit(u.id, 'type', e.target.value)} style={{ ...inputStyle, marginTop: 0 }}>
                      <option value="Manual Sorting Shed">Manual Sorting Shed</option>
                      <option value="Semi-Automated Line">Semi-Automated Line</option>
                    </select>
                    <input type="number" value={u.capacity} onChange={(e) => updateMrfUnit(u.id, 'capacity', e.target.value)} style={{ ...inputStyle, marginTop: 0 }} placeholder="TPD" />
                    {mrfUnits.length > 1 && <button type="button" onClick={() => removeMrfUnit(u.id)} style={{ color: '#dc2626', background: 'none', border: 'none', cursor: 'pointer' }}><Trash2 size={14} /></button>}
                  </div>
                ))}
              </div>
            </>
          )}

          {/* ADVANCED CUSTOM FRACTIONS TOGGLE FOR STANDALONE MRF/MIXED */}
          {appMode === 'STANDALONE' && (facilityType === 'MRF' || facilityType === 'MIXED_PLANT') && (
            <div style={{ marginTop: '10px', background: isAdvancedMode ? '#fffbeb' : '#f8fafc', padding: '12px', borderRadius: '6px', border: '1px solid #cbd5e1', marginBottom: '14px' }}>
              <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontWeight: 'bold', fontSize: '13px', cursor: 'pointer' }}>
                <input type="checkbox" checked={isAdvancedMode} onChange={(e) => setIsAdvancedMode(e.target.checked)} />
                Enable Advanced Stream Configuration & Custom Fractions
              </label>

              {isAdvancedMode && (
                <div style={{ marginTop: '10px', display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', gap: '8px' }}>
                  {currentStreamConfig.map(s => (
                    <div key={s.id} style={{ background: '#fff', padding: '6px', border: '1px solid #e2e8f0', borderRadius: '4px' }}>
                      {s.id.startsWith('other_') ? (
                        <input type="text" value={s.label} onChange={(e) => updateStreamConfig(s.id, 'label', e.target.value)} style={{ fontSize: '11px', fontWeight: 'bold', width: '100%' }} />
                      ) : (
                        <span style={{ fontSize: '11px', fontWeight: 'bold' }}>{s.label}</span>
                      )}
                      <input type="number" value={s.userWeight} onChange={(e) => updateStreamConfig(s.id, 'userWeight', e.target.value)} style={{ ...inputStyle, marginTop: '2px', padding: '4px' }} />
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {/* MONTH SELECTION */}
          <div style={{ marginBottom: '14px' }}>
            <strong style={{ fontSize: '13px' }}>Select Months ({pricing.count} Selected — ₹{pricing.total}):</strong>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(75px, 1fr))', gap: '6px', marginTop: '6px' }}>
              {MONTHS.map((m) => {
                const active = selectedMonths.includes(m.id);
                return (
                  <button key={m.id} type="button" onClick={() => toggleMonth(m.id)} style={{ padding: '6px 2px', borderRadius: '4px', border: active ? '2px solid #059669' : '1px solid #cbd5e1', background: active ? '#ecfdf5' : '#fff', fontWeight: active ? 'bold' : 'normal', cursor: 'pointer', fontSize: '12px' }}>
                    {m.shortEn}
                  </button>
                );
              })}
            </div>
          </div>

          <button type="submit" disabled={generateDisabled} style={{ width: '100%', padding: '12px', background: generateDisabled ? '#94a3b8' : '#059669', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '15px' }}>
            Generate Logbook Dataset ({appMode === 'INTEGRATED_3IN1' ? '₹500/mo Suite' : `₹${pricing.total}`}) →
          </button>
        </form>

        {/* RESULTS PREVIEW */}
        {generatedMonthlyData && (
          <div ref={resultsRef} style={{ background: '#fff', padding: '15px', borderRadius: '8px', border: '1px solid #cbd5e1' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
              <strong style={{ fontSize: '14px' }}>{name} — Dataset Preview</strong>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button onClick={() => setDisplayUnit(displayUnit === 'Tons' ? 'kg' : 'Tons')} style={{ padding: '6px 10px', fontSize: '12px', borderRadius: '4px', border: '1px solid #cbd5e1' }}>Unit: <strong>{displayUnit}</strong></button>
                {isPaid ? (
                  <button onClick={downloadMultiSheetExcel} style={{ padding: '6px 12px', background: '#0f172a', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold' }}>
                    <Download size={13} /> Export Excel (.xlsx)
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
                    <th>Date</th><th>Day</th>
                    {appMode === 'INTEGRATED_3IN1' ? (
                      <><th>Gate Intake</th><th>Seg. Wet Waste</th><th>Seg. Dry Waste</th><th>Mixed Waste</th></>
                    ) : (
                      <><th>Wet Waste</th><th>Dry Waste</th><th>Sanitary</th><th>Special Care</th><th>C&D</th><th>Inerts</th><th>Total</th></>
                    )}
                  </tr>
                </thead>
                <tbody>
                  {visibleRows.map((r, idx) => (
                    <tr key={idx} style={{ borderBottom: '1px solid #e2e8f0' }}>
                      <td>{r.date}</td><td>{r.dayName}</td>
                      {appMode === 'INTEGRATED_3IN1' ? (
                        <><td><strong>{formatVal(r.totalIntake)}</strong></td><td>{formatVal(r.wetSeg)}</td><td>{formatVal(r.drySeg)}</td><td>{formatVal(r.unsegregatedMixed)}</td></>
                      ) : (
                        <><td>{formatVal(r.c1)}</td><td>{formatVal(r.c2)}</td><td>{formatVal(r.c3)}</td><td>{formatVal(r.c4)}</td><td>{formatVal(r.c5)}</td><td>{formatVal(r.c6)}</td><td><strong>{formatVal(r.total)}</strong></td></>
                      )}
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
