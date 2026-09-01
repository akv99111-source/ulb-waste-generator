from pathlib import Path

code = r'''import React, { useMemo, useRef, useState } from 'react';
import {
  Building2,
  Download,
  Lock,
  Globe,
  Check,
  Info,
  ShieldCheck,
  MapPin
} from 'lucide-react';
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

// Indicative modelling profiles. These are NOT official regulatory composition standards.
const REGION_PROFILES = {
  north_plains: {
    wetBase: 0.54,
    dryBase: 0.20,
    nameEn: 'North Plains',
    nameHi: 'उत्तरी मैदानी क्षेत्र'
  },
  coastal_south: {
    wetBase: 0.62,
    dryBase: 0.16,
    nameEn: 'South & Coastal',
    nameHi: 'दक्षिण एवं तटीय क्षेत्र'
  },
  western_central: {
    wetBase: 0.48,
    dryBase: 0.26,
    nameEn: 'West & Central',
    nameHi: 'पश्चिम एवं मध्य भारत'
  },
  eastern_states: {
    wetBase: 0.56,
    dryBase: 0.18,
    nameEn: 'East India',
    nameHi: 'पूर्वी भारत'
  },
  hilly_ne: {
    wetBase: 0.45,
    dryBase: 0.28,
    nameEn: 'Hilly & North-East',
    nameHi: 'पहाड़ी व पूर्वोत्तर क्षेत्र'
  },
  national_avg: {
    wetBase: 0.52,
    dryBase: 0.22,
    nameEn: 'Pan-India Standard',
    nameHi: 'राष्ट्रीय औसत'
  }
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

const inputStyle = {
  width: '100%',
  padding: '9px',
  borderRadius: '6px',
  border: '1px solid #cbd5e1',
  marginTop: '4px',
  boxSizing: 'border-box'
};

function getSeasonalFractions(month, type, regionKey) {
  const profile = REGION_PROFILES[regionKey] || REGION_PROFILES.north_plains;

  if (type === 'ULB') {
    if ([5, 6, 7].includes(month)) {
      return [
        profile.wetBase + 0.05,
        profile.dryBase - 0.02,
        0.04,
        0.02,
        0.05,
        0.12
      ];
    }

    if ([8, 9].includes(month)) {
      return [
        profile.wetBase + 0.03,
        profile.dryBase - 0.01,
        0.04,
        0.02,
        0.05,
        0.13
      ];
    }

    return [
      profile.wetBase,
      profile.dryBase,
      0.04,
      0.02,
      0.05,
      0.15
    ];
  }

  // Indicative MRF distribution model.
  return [5, 6, 7].includes(month)
    ? [0.25, 0.15, 0.20, 0.22, 0.08, 0.10]
    : [0.20, 0.15, 0.25, 0.20, 0.10, 0.10];
}

// Small deterministic PRNG so identical inputs generate identical outputs.
function hashString(str) {
  let h = 2166136261;
  for (let i = 0; i < str.length; i += 1) {
    h ^= str.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return h >>> 0;
}

function mulberry32(seed) {
  return function random() {
    let t = (seed += 0x6D2B79F5);
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function makeLocalDateString(year, month, day) {
  return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

export default function App() {
  const [lang, setLang] = useState('hi');
  const [selectedState, setSelectedState] = useState('Uttar Pradesh');
  const [facilityType, setFacilityType] = useState('ULB');
  const [name, setName] = useState('Nagar Palika Parishad');

  // ULB mode: population OR actual TPD, never both.
  const [ulbCalculationMode, setUlbCalculationMode] = useState('population');
  const [population, setPopulation] = useState(150000);
  const [perCapita, setPerCapita] = useState('400');
  const [customPerCapita, setCustomPerCapita] = useState('');
  const [ulbActualTons, setUlbActualTons] = useState('');
  const [actualReferenceDays, setActualReferenceDays] = useState(30);

  // MRF
  const [mrfDailyDryTons, setMrfDailyDryTons] = useState(15);
  const [mrfMaxCapacityTons, setMrfMaxCapacityTons] = useState(25);

  const [startYear, setStartYear] = useState(new Date().getFullYear());
  const [selectedMonths, setSelectedMonths] = useState([1, 12]);
  const [displayUnit, setDisplayUnit] = useState('Tons');
  const [generatedMonthlyData, setGeneratedMonthlyData] = useState(null);
  const [activeTabMonth, setActiveTabMonth] = useState(null);
  const [isPaid, setIsPaid] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [showStateInfo, setShowStateInfo] = useState(false);

  const resultsRef = useRef(null);

  const currentStateObj = useMemo(
    () =>
      STATES_LIST.find((s) => s.nameEn === selectedState) ||
      STATES_LIST.find((s) => s.nameEn === 'Uttar Pradesh'),
    [selectedState]
  );

  const currentRegionKey = currentStateObj.region;
  const currentRegionObj = REGION_PROFILES[currentRegionKey];

  const effectivePerCapita =
    perCapita === 'custom'
      ? parseFloat(customPerCapita)
      : parseFloat(perCapita);

  const estimatedUlbTpd =
    Number(population) > 0 && Number(effectivePerCapita) > 0
      ? (Number(population) * Number(effectivePerCapita)) / 1_000_000
      : 0;

  const mrfUtilization =
    Number(mrfMaxCapacityTons) > 0
      ? (Number(mrfDailyDryTons) / Number(mrfMaxCapacityTons)) * 100
      : 0;

  const toggleMonth = (mId) => {
    if (selectedMonths.includes(mId)) {
      if (selectedMonths.length > 1) {
        setSelectedMonths(selectedMonths.filter((m) => m !== mId));
      }
      return;
    }

    if (selectedMonths.length >= 3) {
      alert(
        lang === 'hi'
          ? 'अधिकतम 3 माह चुन सकते हैं।'
          : 'Max 3 months allowed.'
      );
      return;
    }

    setSelectedMonths([...selectedMonths, mId].sort((a, b) => a - b));
  };

  const getPrice = () => {
    const prices = {
      1: 50,
      2: 100,
      3: 125
    };
    return prices[selectedMonths.length] ?? 50;
  };

  const validateInputs = () => {
    if (!name.trim()) {
      alert(lang === 'hi' ? 'नाम दर्ज करें।' : 'Please enter the ULB / MRF name.');
      return false;
    }

    if (!Number.isInteger(Number(startYear)) || startYear < 2000 || startYear > 2100) {
      alert(lang === 'hi' ? 'मान्य वर्ष दर्ज करें।' : 'Please enter a valid year.');
      return false;
    }

    if (facilityType === 'ULB') {
      if (ulbCalculationMode === 'population') {
        if (!(Number(population) > 0)) {
          alert(lang === 'hi' ? 'वर्तमान जनसंख्या दर्ज करें।' : 'Please enter current population.');
          return false;
        }

        if (!(Number(effectivePerCapita) > 0)) {
          alert(
            lang === 'hi'
              ? 'प्रतिव्यक्ति अपशिष्ट उत्पादन दर दर्ज करें।'
              : 'Please enter per-capita waste generation.'
          );
          return false;
        }

        if (effectivePerCapita < 250 || effectivePerCapita > 750) {
          alert(
            lang === 'hi'
              ? 'कृपया प्रतिव्यक्ति दर की पुनः जांच करें। सामान्य मॉडलिंग सीमा 250–750 ग्राम/व्यक्ति/दिन रखी गई है।'
              : 'Please re-check the per-capita rate. The model allows a broad validation range of 250–750 g/person/day.'
          );
          return false;
        }
      } else {
        if (!(Number(ulbActualTons) > 0)) {
          alert(
            lang === 'hi'
              ? 'औसत वास्तविक दैनिक अपशिष्ट (TPD) दर्ज करें।'
              : 'Please enter observed average daily waste generation (TPD).'
          );
          return false;
        }

        if (!(Number(actualReferenceDays) > 0)) {
          alert(
            lang === 'hi'
              ? 'संदर्भ अवधि (दिन) दर्ज करें।'
              : 'Please enter the reference period in days.'
          );
          return false;
        }
      }
    }

    if (facilityType === 'MRF') {
      if (!(Number(mrfDailyDryTons) > 0) || !(Number(mrfMaxCapacityTons) > 0)) {
        alert(
          lang === 'hi'
            ? 'MRF दैनिक आवक और क्षमता दोनों दर्ज करें।'
            : 'Please enter both MRF daily input and MRF capacity.'
        );
        return false;
      }

      if (Number(mrfDailyDryTons) > Number(mrfMaxCapacityTons)) {
        const ok = window.confirm(
          lang === 'hi'
            ? `दैनिक आवक (${mrfDailyDryTons} TPD) MRF क्षमता (${mrfMaxCapacityTons} TPD) से अधिक है। उपयोग ${mrfUtilization.toFixed(1)}% है। क्या फिर भी डेटा जनरेट करना है?`
            : `Daily input (${mrfDailyDryTons} TPD) exceeds MRF capacity (${mrfMaxCapacityTons} TPD). Utilization is ${mrfUtilization.toFixed(1)}%. Generate anyway?`
        );
        if (!ok) return false;
      }
    }

    return true;
  };

  const getTargetTons = () => {
    if (facilityType === 'MRF') {
      return Number(mrfDailyDryTons);
    }

    if (ulbCalculationMode === 'actual') {
      return Number(ulbActualTons);
    }

    // population × g/capita/day ÷ 1,000,000 = tonnes/day
    return (Number(population) * Number(effectivePerCapita)) / 1_000_000;
  };

  const handleGenerate = (e) => {
    e.preventDefault();

    if (!validateInputs()) return;

    setIsPaid(false);

    const targetTons = getTargetTons();
    const monthlyDataMap = {};

    selectedMonths.forEach((month) => {
      // Correct leap-year handling.
      const days = new Date(Number(startYear), month, 0).getDate();
      const baseFractions = getSeasonalFractions(
        month,
        facilityType,
        currentRegionKey
      );

      // Same inputs => same monthly data.
      const seedText = [
        facilityType,
        selectedState,
        name.trim().toLowerCase(),
        startYear,
        month,
        ulbCalculationMode,
        population,
        effectivePerCapita,
        ulbActualTons,
        mrfDailyDryTons
      ].join('|');

      const random = mulberry32(hashString(seedText));
      const logs = [];

      for (let day = 1; day <= days; day += 1) {
        const dateObj = new Date(Number(startYear), month - 1, day);
        const weekendMultiplier = [0, 6].includes(dateObj.getDay()) ? 1.05 : 1;

        // Approx. ±5% day-to-day total variation.
        const dailyTotal =
          targetTons * (0.95 + random() * 0.10) * weekendMultiplier;

        // Mild deterministic category variation.
        const raw = baseFractions.map(
          (fraction) => fraction * (0.88 + random() * 0.24)
        );

        const sum = raw.reduce((a, b) => a + b, 0);
        const norm = raw.map((value) => value / sum);

        const parts = norm.map((fraction) =>
          Number((dailyTotal * fraction).toFixed(2))
        );

        // Recalculate displayed total from categories so row totals match
        // the rounded category values exactly.
        const roundedTotal = Number(
          parts.reduce((a, b) => a + b, 0).toFixed(2)
        );

        logs.push({
          date: makeLocalDateString(startYear, month, day),
          dayName: dateObj.toLocaleDateString('en-US', { weekday: 'short' }),
          c1: parts[0],
          c2: parts[1],
          c3: parts[2],
          c4: parts[3],
          c5: parts[4],
          c6: parts[5],
          total: roundedTotal
        });
      }

      monthlyDataMap[month] = logs;
    });

    setGeneratedMonthlyData(monthlyDataMap);
    setActiveTabMonth(selectedMonths[0]);

    setTimeout(() => {
      resultsRef.current?.scrollIntoView({
        behavior: 'smooth',
        block: 'start'
      });
    }, 100);
  };

  const loadRazorpay = async () => {
    if (window.Razorpay) return true;

    return new Promise((resolve, reject) => {
      const existing = document.querySelector(
        'script[src*="checkout.razorpay.com"]'
      );

      if (existing) {
        existing.addEventListener('load', () => resolve(true), { once: true });
        existing.addEventListener(
          'error',
          () => reject(new Error('Unable to load Razorpay checkout.')),
          { once: true }
        );
        return;
      }

      const script = document.createElement('script');
      script.src = 'https://checkout.razorpay.com/v1/checkout.js';
      script.async = true;
      script.onload = () => resolve(true);
      script.onerror = () =>
        reject(new Error('Unable to load Razorpay checkout.'));
      document.body.appendChild(script);
    });
  };

  const handlePayment = async () => {
    if (!generatedMonthlyData || isProcessing) return;

    setIsProcessing(true);

    try {
      await loadRazorpay();

      const orderRes = await fetch('/api/create-order', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          amount: getPrice() * 100,
          months: selectedMonths.length,
          facilityType
        })
      });

      if (!orderRes.ok) {
        throw new Error('Unable to create payment order.');
      }

      const order = await orderRes.json();

      if (!order?.id || !order?.amount || !order?.currency) {
        throw new Error('Invalid order response from server.');
      }

      const options = {
        key: import.meta.env.VITE_RAZORPAY_KEY_ID,
        amount: order.amount,
        currency: order.currency,
        name: 'SWM Logbook Estimation Tool',
        description: `${facilityType} (${selectedMonths.length} Month/s) - ${name}`,
        order_id: order.id,

        handler: async (response) => {
          try {
            // IMPORTANT:
            // /api/verify-payment must verify the Razorpay signature on your server.
            const verifyRes = await fetch('/api/verify-payment', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                razorpay_order_id: response.razorpay_order_id,
                razorpay_payment_id: response.razorpay_payment_id,
                razorpay_signature: response.razorpay_signature
              })
            });

            const verification = await verifyRes.json();

            if (!verifyRes.ok || !verification?.verified) {
              throw new Error('Payment could not be verified.');
            }

            setIsPaid(true);
            setIsProcessing(false);
            downloadMultiSheetExcel();
          } catch (error) {
            setIsProcessing(false);
            alert(
              lang === 'hi'
                ? `भुगतान सत्यापन त्रुटि: ${error.message}`
                : `Payment verification error: ${error.message}`
            );
          }
        },

        modal: {
          ondismiss: () => setIsProcessing(false)
        },

        theme: { color: '#059669' }
      };

      const razorpay = new window.Razorpay(options);

      razorpay.on('payment.failed', (response) => {
        setIsProcessing(false);
        alert(
          lang === 'hi'
            ? `भुगतान विफल: ${response?.error?.description || 'Unknown error'}`
            : `Payment failed: ${response?.error?.description || 'Unknown error'}`
        );
      });

      razorpay.open();
    } catch (error) {
      setIsProcessing(false);
      alert(
        lang === 'hi'
          ? `भुगतान त्रुटि: ${error.message}`
          : `Payment Error: ${error.message}`
      );
    }
  };

  const formatVal = (value) =>
    displayUnit === 'kg'
      ? Math.round(Number(value) * 1000)
      : Number(value).toFixed(2);

  const downloadMultiSheetExcel = () => {
    if (!generatedMonthlyData) return;

    const unit = displayUnit === 'kg' ? 'kg' : 'Tons';

    const headers =
      facilityType === 'ULB'
        ? [
            'Date',
            'Day',
            `Wet (${unit})`,
            `Dry (${unit})`,
            `Sanitary (${unit})`,
            `Special Care/Hazardous (${unit})`,
            `C&D (${unit})`,
            `Inerts (${unit})`,
            `Total (${unit})`
          ]
        : [
            'Date',
            'Day',
            `PET (${unit})`,
            `HDPE (${unit})`,
            `Paper/Cardboard (${unit})`,
            `RDF/SCF (${unit})`,
            `Glass & Metal (${unit})`,
            `Rejects (${unit})`,
            `Total Dry (${unit})`
          ];

    const wb = XLSX.utils.book_new();

    selectedMonths.forEach((mId) => {
      const monthRows = generatedMonthlyData[mId] || [];

      const sheetData = [
        headers,
        ...monthRows.map((row) => [
          row.date,
          row.dayName,
          formatVal(row.c1),
          formatVal(row.c2),
          formatVal(row.c3),
          formatVal(row.c4),
          formatVal(row.c5),
          formatVal(row.c6),
          formatVal(row.total)
        ])
      ];

      const ws = XLSX.utils.aoa_to_sheet(sheetData);

      // Helpful widths.
      ws['!cols'] = [
        { wch: 12 },
        { wch: 8 },
        { wch: 18 },
        { wch: 18 },
        { wch: 18 },
        { wch: 26 },
        { wch: 18 },
        { wch: 18 },
        { wch: 18 }
      ];

      XLSX.utils.book_append_sheet(
        wb,
        ws,
        MONTHS.find((m) => m.id === mId)?.fullEn || `Month-${mId}`
      );
    });

    XLSX.writeFile(
      wb,
      `${name.trim().replace(/\s+/g, '_')}_SWM_Logbook_${selectedMonths.length}M.xlsx`
    );
  };

  const activeRows = generatedMonthlyData?.[activeTabMonth] || [];
  const visibleRows = isPaid ? activeRows : activeRows.slice(0, 5);
  const activeMonthObj = MONTHS.find((m) => m.id === activeTabMonth);

  return (
    <div
      style={{
        fontFamily: 'sans-serif',
        background: '#f8fafc',
        minHeight: '100vh',
        display: 'flex',
        flexDirection: 'column',
        justifyContent: 'space-between'
      }}
    >
      <div>
        {/* HEADER */}
        <div
          style={{
            background: 'linear-gradient(135deg, #064e3b 0%, #047857 100%)',
            color: '#fff',
            padding: '20px 15px'
          }}
        >
          <div
            style={{
              maxWidth: '1000px',
              margin: '0 auto',
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
              flexWrap: 'wrap',
              gap: '10px'
            }}
          >
            <div>
              <span
                style={{
                  background: 'rgba(255,255,255,0.2)',
                  padding: '3px 8px',
                  borderRadius: '12px',
                  fontSize: '11px',
                  fontWeight: 'bold'
                }}
              >
                <ShieldCheck size={12} style={{ verticalAlign: 'middle' }} />{' '}
                SWM-ALIGNED ESTIMATION TOOL
              </span>

              <h1
                style={{
                  fontSize: '22px',
                  margin: '6px 0 2px 0',
                  fontWeight: '800'
                }}
              >
                <Building2
                  size={22}
                  style={{ verticalAlign: 'middle', marginRight: '6px' }}
                />
                {lang === 'hi'
                  ? 'यूएलबी एवं एमआरएफ लोगबुक जनरेटर'
                  : 'ULB & MRF Waste Logbook Generator'}
              </h1>

              <p
                style={{
                  fontSize: '13px',
                  margin: 0,
                  color: '#a7f3d0'
                }}
              >
                {lang === 'hi'
                  ? 'मॉडल-आधारित अपशिष्ट श्रेणी वितरण एवं क्षेत्रीय अनुमान'
                  : 'Model-based waste category distribution with regional estimation adjustments'}
              </p>
            </div>

            <button
              onClick={() => setLang(lang === 'hi' ? 'en' : 'hi')}
              style={{
                padding: '6px 12px',
                background: '#fff',
                color: '#047857',
                border: 'none',
                borderRadius: '5px',
                fontWeight: 'bold',
                cursor: 'pointer',
                fontSize: '13px'
              }}
            >
              <Globe
                size={15}
                style={{ verticalAlign: 'middle', marginRight: '4px' }}
              />
              {lang === 'hi' ? 'English' : 'हिंदी'}
            </button>
          </div>
        </div>

        {/* FORM */}
        <div style={{ padding: '15px', maxWidth: '1000px', margin: '0 auto' }}>
          <form
            onSubmit={handleGenerate}
            style={{
              background: '#fff',
              padding: '16px',
              borderRadius: '8px',
              border: '1px solid #cbd5e1',
              marginBottom: '20px'
            }}
          >
            <div
              style={{
                marginBottom: '14px',
                display: 'flex',
                gap: '15px',
                alignItems: 'center',
                fontSize: '14px',
                flexWrap: 'wrap'
              }}
            >
              <strong>{lang === 'hi' ? 'सुविधा प्रकार:' : 'Facility:'}</strong>

              <label>
                <input
                  type="radio"
                  checked={facilityType === 'ULB'}
                  onChange={() => {
                    setFacilityType('ULB');
                    setGeneratedMonthlyData(null);
                  }}
                />{' '}
                ULB
              </label>

              <label>
                <input
                  type="radio"
                  checked={facilityType === 'MRF'}
                  onChange={() => {
                    setFacilityType('MRF');
                    setGeneratedMonthlyData(null);
                  }}
                />{' '}
                MRF Centre
              </label>
            </div>

            <div
              style={{
                display: 'grid',
                gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
                gap: '12px',
                marginBottom: '14px'
              }}
            >
              <div>
                <div
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between'
                  }}
                >
                  <label
                    style={{
                      fontSize: '12px',
                      fontWeight: '600',
                      color: '#047857',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '4px'
                    }}
                  >
                    <MapPin size={13} />
                    {lang === 'hi' ? 'राज्य चुनें' : 'Select State'}
                  </label>

                  <button
                    type="button"
                    onClick={() => setShowStateInfo(!showStateInfo)}
                    style={{
                      background: 'none',
                      border: 'none',
                      color: '#0284c7',
                      cursor: 'pointer',
                      padding: '0 2px',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '2px',
                      fontSize: '12px',
                      fontWeight: 'bold'
                    }}
                  >
                    <Info size={14} /> (i)
                  </button>
                </div>

                <select
                  style={{
                    ...inputStyle,
                    border: '1px solid #059669',
                    background: '#f0fdf4'
                  }}
                  value={selectedState}
                  onChange={(e) => setSelectedState(e.target.value)}
                >
                  {STATES_LIST.map((state) => (
                    <option key={state.nameEn} value={state.nameEn}>
                      {lang === 'hi' ? state.nameHi : state.nameEn}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label style={{ fontSize: '12px', fontWeight: '600' }}>
                  {lang === 'hi' ? 'निकाय / एमआरएफ का नाम' : 'ULB / MRF Name'}
                </label>

                <input
                  style={inputStyle}
                  type="text"
                  required
                  value={name}
                  onChange={(e) => setName(e.target.value)}
                />
              </div>

              <div>
                <label style={{ fontSize: '12px', fontWeight: '600' }}>
                  {lang === 'hi' ? 'वर्ष' : 'Year'}
                </label>

                <input
                  style={inputStyle}
                  type="number"
                  min="2000"
                  max="2100"
                  value={startYear}
                  onChange={(e) => setStartYear(Number(e.target.value))}
                />
              </div>
            </div>

            {showStateInfo && (
              <div
                style={{
                  background: '#f0f9ff',
                  border: '1px solid #7dd3fc',
                  padding: '10px 14px',
                  borderRadius: '6px',
                  marginBottom: '14px',
                  fontSize: '12px',
                  color: '#0369a1',
                  lineHeight: '1.4'
                }}
              >
                <div
                  style={{
                    fontWeight: 'bold',
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    marginBottom: '4px'
                  }}
                >
                  <span>
                    ℹ️{' '}
                    {lang === 'hi'
                      ? 'राज्य चयन का महत्व:'
                      : 'Relevance of Selecting Your State:'}
                  </span>

                  <button
                    type="button"
                    onClick={() => setShowStateInfo(false)}
                    style={{
                      background: 'none',
                      border: 'none',
                      cursor: 'pointer',
                      color: '#0369a1',
                      fontWeight: 'bold'
                    }}
                  >
                    ✕
                  </button>
                </div>

                <p style={{ margin: '0 0 4px 0' }}>
                  {lang === 'hi'
                    ? `चयनित राज्य (${currentStateObj.nameHi}) को ${currentRegionObj.nameHi} क्षेत्र में रखा गया है। क्षेत्रीय कारक केवल मॉडल-आधारित अनुमान हैं।`
                    : `Selected state (${currentStateObj.nameEn}) is mapped to the ${currentRegionObj.nameEn} zone. Regional factors are indicative model assumptions only.`}
                </p>
              </div>
            )}

            {/* ULB INPUTS */}
            {facilityType === 'ULB' && (
              <div
                style={{
                  border: '1px solid #d1fae5',
                  background: '#f0fdf4',
                  padding: '14px',
                  borderRadius: '8px',
                  marginBottom: '14px'
                }}
              >
                <div
                  style={{
                    fontSize: '13px',
                    fontWeight: 'bold',
                    marginBottom: '10px',
                    color: '#065f46'
                  }}
                >
                  {lang === 'hi'
                    ? 'यूएलबी अपशिष्ट गणना का आधार — केवल एक विकल्प चुनें'
                    : 'ULB Waste Calculation Basis — select only one'}
                </div>

                <div
                  style={{
                    display: 'flex',
                    gap: '18px',
                    flexWrap: 'wrap',
                    marginBottom: '12px',
                    fontSize: '13px'
                  }}
                >
                  <label>
                    <input
                      type="radio"
                      checked={ulbCalculationMode === 'population'}
                      onChange={() => setUlbCalculationMode('population')}
                    />{' '}
                    {lang === 'hi'
                      ? 'वर्तमान जनसंख्या आधारित अनुमान'
                      : 'Current population-based estimation'}
                  </label>

                  <label>
                    <input
                      type="radio"
                      checked={ulbCalculationMode === 'actual'}
                      onChange={() => setUlbCalculationMode('actual')}
                    />{' '}
                    {lang === 'hi'
                      ? 'वास्तविक / प्रेक्षित औसत TPD'
                      : 'Actual / observed average TPD'}
                  </label>
                </div>

                {ulbCalculationMode === 'population' ? (
                  <>
                    <div
                      style={{
                        display: 'grid',
                        gridTemplateColumns:
                          'repeat(auto-fit, minmax(220px, 1fr))',
                        gap: '12px'
                      }}
                    >
                      <div>
                        <label
                          style={{ fontSize: '12px', fontWeight: '600' }}
                        >
                          {lang === 'hi'
                            ? 'वर्तमान अनुमानित जनसंख्या'
                            : 'Current Estimated Population'}
                        </label>

                        <input
                          style={inputStyle}
                          type="number"
                          min="1"
                          value={population}
                          onChange={(e) =>
                            setPopulation(Number(e.target.value))
                          }
                        />
                      </div>

                      <div>
                        <label
                          style={{ fontSize: '12px', fontWeight: '600' }}
                        >
                          {lang === 'hi'
                            ? 'प्रतिव्यक्ति अपशिष्ट उत्पादन'
                            : 'Per-capita Waste Generation'}
                        </label>

                        <select
                          style={inputStyle}
                          value={perCapita}
                          onChange={(e) => setPerCapita(e.target.value)}
                        >
                          <option value="300">300 g/person/day</option>
                          <option value="350">350 g/person/day</option>
                          <option value="400">400 g/person/day</option>
                          <option value="450">450 g/person/day</option>
                          <option value="500">500 g/person/day</option>
                          <option value="custom">
                            {lang === 'hi'
                              ? 'कस्टम दर'
                              : 'Custom rate'}
                          </option>
                        </select>
                      </div>

                      {perCapita === 'custom' && (
                        <div>
                          <label
                            style={{ fontSize: '12px', fontWeight: '600' }}
                          >
                            {lang === 'hi'
                              ? 'कस्टम दर (ग्राम/व्यक्ति/दिन)'
                              : 'Custom rate (g/person/day)'}
                          </label>

                          <input
                            style={inputStyle}
                            type="number"
                            min="250"
                            max="750"
                            step="1"
                            value={customPerCapita}
                            onChange={(e) =>
                              setCustomPerCapita(e.target.value)
                            }
                          />
                        </div>
                      )}
                    </div>

                    <div
                      style={{
                        marginTop: '10px',
                        fontSize: '12px',
                        color: '#065f46'
                      }}
                    >
                      <strong>
                        {lang === 'hi'
                          ? 'अनुमानित दैनिक अपशिष्ट:'
                          : 'Estimated Daily Waste:'}
                      </strong>{' '}
                      {estimatedUlbTpd.toFixed(2)} TPD
                      <div style={{ marginTop: '3px', color: '#475569' }}>
                        {lang === 'hi'
                          ? 'सूत्र: वर्तमान जनसंख्या × ग्राम/व्यक्ति/दिन ÷ 10,00,000'
                          : 'Formula: Current population × g/person/day ÷ 1,000,000'}
                      </div>
                    </div>
                  </>
                ) : (
                  <div
                    style={{
                      display: 'grid',
                      gridTemplateColumns:
                        'repeat(auto-fit, minmax(220px, 1fr))',
                      gap: '12px'
                    }}
                  >
                    <div>
                      <label
                        style={{ fontSize: '12px', fontWeight: '600' }}
                      >
                        {lang === 'hi'
                          ? 'औसत वास्तविक दैनिक अपशिष्ट (TPD)'
                          : 'Observed Average Daily Waste (TPD)'}
                      </label>

                      <input
                        style={inputStyle}
                        type="number"
                        min="0.01"
                        step="0.01"
                        value={ulbActualTons}
                        onChange={(e) => setUlbActualTons(e.target.value)}
                      />
                    </div>

                    <div>
                      <label
                        style={{ fontSize: '12px', fontWeight: '600' }}
                      >
                        {lang === 'hi'
                          ? 'संदर्भ अवधि (दिन)'
                          : 'Reference Period (days)'}
                      </label>

                      <input
                        style={inputStyle}
                        type="number"
                        min="1"
                        value={actualReferenceDays}
                        onChange={(e) =>
                          setActualReferenceDays(Number(e.target.value))
                        }
                      />
                    </div>

                    <div
                      style={{
                        fontSize: '12px',
                        color: '#475569',
                        alignSelf: 'end',
                        paddingBottom: '8px'
                      }}
                    >
                      {lang === 'hi'
                        ? 'जहाँ विश्वसनीय स्थानीय रिकॉर्ड उपलब्ध हों, वहाँ वास्तविक औसत TPD को प्राथमिकता दें।'
                        : 'Where reliable local records are available, prefer observed average TPD over population-based estimation.'}
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* MRF INPUTS */}
            {facilityType === 'MRF' && (
              <div
                style={{
                  border: '1px solid #dbeafe',
                  background: '#eff6ff',
                  padding: '14px',
                  borderRadius: '8px',
                  marginBottom: '14px'
                }}
              >
                <div
                  style={{
                    display: 'grid',
                    gridTemplateColumns:
                      'repeat(auto-fit, minmax(220px, 1fr))',
                    gap: '12px'
                  }}
                >
                  <div>
                    <label style={{ fontSize: '12px', fontWeight: '600' }}>
                      {lang === 'hi'
                        ? 'सूखा कचरा आवक (टन/दिन)'
                        : 'Daily Dry Input (Tons/day)'}
                    </label>

                    <input
                      style={inputStyle}
                      type="number"
                      min="0.01"
                      step="0.01"
                      value={mrfDailyDryTons}
                      onChange={(e) => setMrfDailyDryTons(e.target.value)}
                    />
                  </div>

                  <div>
                    <label style={{ fontSize: '12px', fontWeight: '600' }}>
                      {lang === 'hi'
                        ? 'MRF क्षमता (टन/दिन)'
                        : 'MRF Capacity (Tons/day)'}
                    </label>

                    <input
                      style={inputStyle}
                      type="number"
                      min="0.01"
                      step="0.01"
                      value={mrfMaxCapacityTons}
                      onChange={(e) => setMrfMaxCapacityTons(e.target.value)}
                    />
                  </div>

                  <div
                    style={{
                      alignSelf: 'end',
                      paddingBottom: '8px',
                      fontSize: '12px',
                      color:
                        mrfUtilization > 100 ? '#b91c1c' : '#1e40af',
                      fontWeight: 'bold'
                    }}
                  >
                    {lang === 'hi'
                      ? 'क्षमता उपयोग'
                      : 'Capacity Utilization'}
                    : {Number.isFinite(mrfUtilization) ? mrfUtilization.toFixed(1) : '0.0'}%
                  </div>
                </div>
              </div>
            )}

            {/* MONTHS */}
            <div style={{ marginBottom: '14px' }}>
              <div
                style={{
                  display: 'flex',
                  justifyContent: 'space-between',
                  marginBottom: '6px',
                  fontSize: '13px'
                }}
              >
                <strong>
                  {lang === 'hi'
                    ? 'माह चुनें (अधिकतम 3):'
                    : 'Select Months (Max 3):'}
                </strong>

                <span style={{ color: '#059669', fontWeight: 'bold' }}>
                  {selectedMonths.length}{' '}
                  {lang === 'hi' ? 'माह' : 'Month/s'} — ₹{getPrice()}
                </span>
              </div>

              <div
                style={{
                  display: 'grid',
                  gridTemplateColumns:
                    'repeat(auto-fill, minmax(75px, 1fr))',
                  gap: '6px'
                }}
              >
                {MONTHS.map((month) => {
                  const active = selectedMonths.includes(month.id);

                  return (
                    <button
                      key={month.id}
                      type="button"
                      onClick={() => toggleMonth(month.id)}
                      style={{
                        padding: '8px 2px',
                        borderRadius: '5px',
                        border: active
                          ? '2px solid #059669'
                          : '1px solid #cbd5e1',
                        background: active ? '#ecfdf5' : '#fff',
                        color: active ? '#065f46' : '#334155',
                        fontWeight: active ? 'bold' : 'normal',
                        cursor: 'pointer',
                        fontSize: '12px'
                      }}
                    >
                      {active && <Check size={11} />}{' '}
                      {lang === 'hi' ? month.shortHi : month.shortEn}
                    </button>
                  );
                })}
              </div>
            </div>

            <button
              type="submit"
              style={{
                width: '100%',
                padding: '12px',
                background: '#059669',
                color: '#fff',
                border: 'none',
                borderRadius: '6px',
                cursor: 'pointer',
                fontWeight: 'bold',
                fontSize: '15px'
              }}
            >
              {lang === 'hi'
                ? 'लोगबुक डेटा जनरेट करें →'
                : 'Generate Logbook Dataset →'}
            </button>
          </form>

          {/* RESULTS */}
          {generatedMonthlyData && (
            <div
              ref={resultsRef}
              style={{
                background: '#fff',
                padding: '15px',
                borderRadius: '8px',
                border: '1px solid #cbd5e1',
                scrollMarginTop: '15px'
              }}
            >
              <div
                style={{
                  background: '#eff6ff',
                  border: '1px solid #bfdbfe',
                  padding: '10px',
                  borderRadius: '6px',
                  marginBottom: '12px',
                  fontSize: '12px',
                  color: '#1e40af'
                }}
              >
                <strong>
                  <Info
                    size={14}
                    style={{ verticalAlign: 'middle' }}
                  />{' '}
                  Model Guidance ({currentStateObj.nameEn} /{' '}
                  {currentRegionObj.nameEn}):
                </strong>{' '}
                Regional composition factors and day-to-day variations are
                indicative modelling assumptions. Use actual weighbridge /
                collection records where available.
              </div>

              <div
                style={{
                  display: 'flex',
                  borderBottom: '2px solid #e2e8f0',
                  marginBottom: '12px',
                  overflowX: 'auto'
                }}
              >
                {selectedMonths.map((mId) => (
                  <button
                    key={mId}
                    onClick={() => setActiveTabMonth(mId)}
                    style={{
                      padding: '8px 16px',
                      border: 'none',
                      borderBottom:
                        activeTabMonth === mId
                          ? '3px solid #059669'
                          : 'none',
                      background:
                        activeTabMonth === mId
                          ? '#ecfdf5'
                          : 'transparent',
                      fontWeight:
                        activeTabMonth === mId ? 'bold' : 'normal',
                      cursor: 'pointer',
                      fontSize: '13px'
                    }}
                  >
                    {
                      MONTHS.find((m) => m.id === mId)?.[
                        lang === 'hi' ? 'shortHi' : 'fullEn'
                      ]
                    }
                  </button>
                ))}
              </div>

              <div
                style={{
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'center',
                  marginBottom: '10px',
                  flexWrap: 'wrap',
                  gap: '8px'
                }}
              >
                <span style={{ fontSize: '13px', fontWeight: 'bold' }}>
                  {name} ({currentStateObj.nameEn}) —{' '}
                  {activeMonthObj?.fullEn} {startYear}
                </span>

                <div style={{ display: 'flex', gap: '8px' }}>
                  <button
                    onClick={() =>
                      setDisplayUnit(
                        displayUnit === 'Tons' ? 'kg' : 'Tons'
                      )
                    }
                    style={{
                      padding: '6px 10px',
                      fontSize: '12px',
                      borderRadius: '4px',
                      border: '1px solid #cbd5e1'
                    }}
                  >
                    Unit: <strong>{displayUnit}</strong>
                  </button>

                  {isPaid ? (
                    <button
                      onClick={downloadMultiSheetExcel}
                      style={{
                        padding: '6px 12px',
                        background: '#0f172a',
                        color: '#fff',
                        border: 'none',
                        borderRadius: '4px',
                        cursor: 'pointer',
                        fontSize: '12px',
                        fontWeight: 'bold'
                      }}
                    >
                      <Download size={13} /> Export Excel (.xlsx)
                    </button>
                  ) : (
                    <button
                      onClick={handlePayment}
                      disabled={isProcessing}
                      style={{
                        padding: '6px 12px',
                        background: '#059669',
                        color: '#fff',
                        border: 'none',
                        borderRadius: '4px',
                        cursor: 'pointer',
                        fontSize: '12px',
                        fontWeight: 'bold',
                        opacity: isProcessing ? 0.7 : 1
                      }}
                    >
                      {isProcessing
                        ? 'Wait...'
                        : `Pay ₹${getPrice()} to Unlock`}
                    </button>
                  )}
                </div>
              </div>

              <div
                style={{
                  overflowX: 'auto',
                  border: '1px solid #cbd5e1',
                  borderRadius: '4px'
                }}
              >
                <table
                  cellPadding="8"
                  style={{
                    width: '100%',
                    borderCollapse: 'collapse',
                    fontSize: '12px',
                    minWidth: '700px'
                  }}
                >
                  <thead>
                    <tr
                      style={{
                        background: '#f1f5f9',
                        borderBottom: '1px solid #cbd5e1'
                      }}
                    >
                      <th>Date</th>
                      <th>Day</th>

                      {facilityType === 'ULB' ? (
                        <>
                          <th style={{ background: '#ecfdf5' }}>
                            Wet ({displayUnit})
                          </th>
                          <th>Dry ({displayUnit})</th>
                          <th>Sanitary ({displayUnit})</th>
                          <th>Special Care ({displayUnit})</th>
                          <th>C&D ({displayUnit})</th>
                          <th>Inerts ({displayUnit})</th>
                          <th>Total ({displayUnit})</th>
                        </>
                      ) : (
                        <>
                          <th>PET ({displayUnit})</th>
                          <th>HDPE ({displayUnit})</th>
                          <th>Paper ({displayUnit})</th>
                          <th>RDF ({displayUnit})</th>
                          <th>Glass/Metal ({displayUnit})</th>
                          <th>Rejects ({displayUnit})</th>
                          <th>Total Dry ({displayUnit})</th>
                        </>
                      )}
                    </tr>
                  </thead>

                  <tbody>
                    {visibleRows.map((row) => (
                      <tr
                        key={row.date}
                        style={{
                          borderBottom: '1px solid #e2e8f0'
                        }}
                      >
                        <td>{row.date}</td>
                        <td>{row.dayName}</td>
                        <td
                          style={{
                            background:
                              facilityType === 'ULB'
                                ? '#ecfdf5'
                                : '#fff'
                          }}
                        >
                          {formatVal(row.c1)}
                        </td>
                        <td>{formatVal(row.c2)}</td>
                        <td>{formatVal(row.c3)}</td>
                        <td>{formatVal(row.c4)}</td>
                        <td>{formatVal(row.c5)}</td>
                        <td>{formatVal(row.c6)}</td>
                        <td>
                          <strong>{formatVal(row.total)}</strong>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {!isPaid && (
                <div
                  style={{
                    border: '2px dashed #059669',
                    background: '#ecfdf5',
                    padding: '15px',
                    textAlign: 'center',
                    marginTop: '12px',
                    borderRadius: '6px'
                  }}
                >
                  <Lock style={{ color: '#059669' }} size={18} />

                  <h4
                    style={{
                      margin: '4px 0',
                      color: '#065f46',
                      fontSize: '15px'
                    }}
                  >
                    Preview Locked (Days 1–5 Only)
                  </h4>

                  <p
                    style={{
                      margin: '4px 0 10px 0',
                      color: '#047857',
                      fontSize: '13px'
                    }}
                  >
                    Pay ₹{getPrice()} to unlock the complete dataset and
                    multi-tab Excel (.xlsx) file.
                  </p>

                  <button
                    onClick={handlePayment}
                    disabled={isProcessing}
                    style={{
                      padding: '10px 20px',
                      background: '#059669',
                      color: '#fff',
                      border: 'none',
                      borderRadius: '6px',
                      fontSize: '14px',
                      fontWeight: 'bold',
                      cursor: 'pointer',
                      opacity: isProcessing ? 0.7 : 1
                    }}
                  >
                    {isProcessing
                      ? 'Connecting...'
                      : `Pay ₹${getPrice()} & Download Complete File`}
                  </button>
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      <div
        style={{
          marginTop: '30px',
          padding: '15px 10px',
          borderTop: '1px solid #cbd5e1',
          textAlign: 'center',
          fontSize: '11px',
          color: '#64748b',
          lineHeight: '1.5',
          background: '#ffffff'
        }}
      >
        <p style={{ margin: '0 0 6px 0' }}>
          <strong>Disclaimer:</strong> This tool is developed for educational,
          research, planning and estimation purposes. Generated datasets are
          model-based estimates and are not a substitute for actual statutory
          records, weighbridge records or field measurements.
        </p>

        <p style={{ margin: 0 }}>
          Copyright © 2026 CRF | Engineered & Maintained by{' '}
          <strong>Team CRF</strong> —{' '}
          <a
            href="https://www.consilienceresearch.in/"
            target="_blank"
            rel="noopener noreferrer"
            style={{
              color: '#059669',
              fontWeight: 'bold',
              textDecoration: 'none'
            }}
          >
            Consilience Research Foundation
          </a>
          , an Urban & Infrastructure Research Consultancy Institute.
        </p>
      </div>
    </div>
  );
}
'''

path = Path('/mnt/data/App_corrected.jsx')
path.write_text(code, encoding='utf-8')
print(f"Created {path} with {len(code.splitlines())} lines.")
