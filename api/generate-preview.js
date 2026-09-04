// /api/generate-preview.js
import { cyrb128, mulberry32, getSeasonalFractionsULB, REGION_PROFILES } from '../../lib/swm-engine';

export default function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ message: 'Method not allowed' });

  const { facilityType, selectedState, name, startYear, selectedMonths, ulbCalculationMode, targetTons, isAdvancedMode, activeMrfStreams } = req.body;

  let truncatedMonthlyData = {};

  selectedMonths.forEach((m) => {
    // Generate days using your seeded deterministic algorithm
    const seedString = `${facilityType}-${selectedState}-${name}-${startYear}-${m}-${ulbCalculationMode}-${targetTons}-${isAdvancedMode}`;
    const random = mulberry32(cyrb128(seedString));
    const totalDaysInMonth = new Date(startYear, m, 0).getDate();

    let logs = [];
    for (let day = 1; day <= totalDaysInMonth; day++) {
      // Calculation logic...
      // (Your existing day-by-day math goes here)
      logs.push({ day, ...calculatedData });
    }

    // CRITICAL SECURITY STEP: Slice ONLY the first 5 days for preview
    truncatedMonthlyData[m] = logs.slice(0, 5);
  });

  return res.status(200).json({
    success: true,
    isPreview: true,
    data: truncatedMonthlyData
  });
}
