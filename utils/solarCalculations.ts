import { DailyData, AggregatedData, RawRow, SimulationResult, LossFactors, SystemSpecs } from '../types';
import { parse, isValid, compareAsc, subDays, startOfMonth, format, isSameMonth, isAfter, startOfYear, min } from 'date-fns';

/**
 * Normalizes keys from Excel to standard internal keys
 * NOTE: Strict separation of PV vs Inverter/BESS capacity to ensure Yield calc only uses PV Capacity.
 */
const normalizeKey = (key: string): string => {
  const k = key.toLowerCase().trim();
  
  // 1. Capacity (High Priority)
  // Distinguish between System (PV), Inverter, BESS
  if (k.includes('bess') || k.includes('battery')) {
      if (k.includes('capacity') || k.includes('size') || k.includes('kwh') || k.includes('mwh')) return 'bessCapacity';
      if (k.includes('mode') || k.includes('operation')) return 'bessOperationMode';
  }

  if (k.includes('inverter') && (k.includes('capacity') || k.includes('size') || k.includes('power'))) {
      return 'invCapacity';
  }

  // Fallback for PV Capacity if not caught above
  // Includes 'module power', 'array size', 'dc power', 'installed power'
  if (k.includes('capacity') || 
      (k.includes('system') && k.includes('size')) || 
      k.includes('dc power') || 
      k.includes('installed power') || 
      k.includes('pv capacity') || 
      k.includes('module power') || 
      k.includes('array size')) {
      
      // Ensure it's not specific yield or just 'power' which might be actual power
      if (!k.includes('specific') && !k.includes('actual') && !k.includes('output')) return 'systemCapacity';
  }

  // 2. Date
  // COD check
  if (k === 'cod' || (k.includes('commercial') && k.includes('operation'))) return 'cod';
  if (k === 'date' || k === 'datetime' || k === 'timestamp' || k.includes('date')) return 'date';

  // 3. Site Name
  if (k.includes('site') || k.includes('plant') || k.includes('project')) return 'siteName';

  // 4. Specific Yield (Ignore for calculation, prevent it from overriding kwhActual)
  if (k.includes('sy') || k.includes('specific yield') || k.includes('/kwp')) {
     return 'ignore_sy'; 
  }

  // 5a. POA (Plane of Array) — explicit; matched BEFORE GHI to prevent capture
  if (k.includes('poa') || (k.includes('plane') && k.includes('array')) || k === 'gpoa' || k.includes('tilted')) {
      if (k.includes('budget') || k.includes('target') || k.includes('expected')) return 'poaBudget';
      if (k.includes('actual') || k.includes('measured')) return 'poaActual';
      if (!k.includes('budget') && !k.includes('forecast')) return 'poaActual';
  }

  // 5b. GHI (Global Horizontal Irradiance) — fallback irradiance
  if (k.includes('ghi') || (k.includes('global') && k.includes('horizontal')) ||
      k.includes('irradiation') || k.includes('irradiance')) {
      if (k.includes('budget') || k.includes('target') || k.includes('expected')) return 'ghiBudget';
      if (k.includes('actual') || k.includes('measured')) return 'ghiActual';
      if (!k.includes('budget') && !k.includes('forecast')) return 'ghiActual';
  }
  
  // 6. Availability (New)
  if (k.includes('availability')) {
      if (k.includes('guaranteed') || k.includes('target') || k.includes('budget')) return 'guaranteedAvailability';
      if (k.includes('actual') || k.includes('measured')) return 'actualAvailability';
      // Default to actual if just "availability" or "plant availability"
      return 'actualAvailability';
  }

  // 7. Energy
  if (k.includes('kwh') || k.includes('energy') || k.includes('yield') || k.includes('production') || k.includes('loss')) {
      if (k.includes('estimated') || k.includes('loss') || k.includes('outage') || k.includes('curtailment')) {
           if (k.includes('curtailment') || k.includes('grid') || k.includes('constraint')) return 'curtailment';
           return 'estimatedLoss';
      }
      
      // PAE Energy Check
      if (k.includes('pae')) return 'paeEnergy';

      if (k.includes('budget') || k.includes('target') || k.includes('expected')) return 'kwhBudget';
      if (k.includes('actual') || k.includes('measured') || k.includes('inverter')) return 'kwhActual';
      // Fallback: if just "Energy" or "Production", treat as actual
      if (!k.includes('budget') && !k.includes('forecast')) return 'kwhActual';
  }
  
  // Catch PAE if it doesn't have "energy" or "kwh"
  if (k.includes('pae')) return 'paeEnergy';

  // 8. O&M Contractor PR Target (Distinct from Corrected PR)
  if ((k.includes('contractor') || k.includes('o&m')) && k.includes('pr')) {
      return 'contractorPrTarget';
  }

  // 9. Corrected PR Budget
  if (k.includes('corrected') || k.includes('weather adjusted') || k.includes('guaranteed')) {
      if (k.includes('pr') || k.includes('performance')) {
          if (k.includes('budget') || k.includes('target') || k.includes('expected')) return 'correctedPrBudget';
      }
  }

  // 10. PR
  if (k.includes('pr') || k.includes('performance ratio') || k.includes('efficiency')) {
      if (k.includes('budget') || k.includes('target') || k.includes('expected')) return 'prBudget';
      if (k.includes('actual') || k.includes('measured')) return 'prActual';
  }

  // 11. Module Temperature
  if (k.includes('temp') || k.includes('tmod') || k.includes('t_mod') || k.includes('module')) {
      if (k.includes('module') || k.includes('cell') || k.includes('panel')) return 'moduleTemp';
  }
  
  // 12. Curtailment
  if (k.includes('curtail') || k.includes('grid loss') || k.includes('limitation')) return 'curtailment';

  // 13. Notes
  if (k.includes('note') || k.includes('issue') || k.includes('comment') || k.includes('remark') || k.includes('reason')) {
      return 'notes';
  }
  
  return key;
};

/**
 * Parses raw excel rows into typed DailyData
 */
export const parseRawData = (rows: RawRow[]): DailyData[] => {
  const parseVal = (val: any) => {
      if (typeof val === 'number') return val;
      if (typeof val === 'string') {
          const cleaned = val.replace(/,/g, '').replace(/[^\d.-]/g, '');
          const float = parseFloat(cleaned);
          return isNaN(float) ? 0 : float;
      }
      return 0;
  };

  const parseDate = (val: any): Date | null => {
     if (!val) return null;
     if (typeof val === 'number') {
         return new Date(Math.round((val - 25569) * 86400 * 1000));
     }
     if (val instanceof Date) return val;
     const d = new Date(val);
     return isValid(d) ? d : null;
  };

  // Sanitize daily irradiance values (kWh/m²). Reject impossible, auto-convert obvious unit errors.
  const sanitizeIrradiance = (v: number | null): number | null => {
    if (v === null || v === undefined || isNaN(v)) return null;
    if (v <= 0) return null;
    if (v > 50) return v / 1000;   // assume Wh/m² → kWh/m²
    if (v > 15) return null;       // unrealistic for daily
    return v;
  };

  const parsed = rows.map(row => {
    const newRow: any = {};
    Object.keys(row).forEach(key => {
      newRow[normalizeKey(key)] = row[key];
    });

    let dateObj = parseDate(newRow.date);
    if (!dateObj) dateObj = new Date(); // Fallback but filtering later

    // Normalize PR
    let prB = parseFloat(newRow.prBudget) || 0;
    if (prB > 1.5) prB = prB / 100;

    let corrPrB = newRow.correctedPrBudget !== undefined ? parseFloat(newRow.correctedPrBudget) : null;
    if (corrPrB !== null && corrPrB > 1.5) corrPrB = corrPrB / 100;

    let contPrB = newRow.contractorPrTarget !== undefined ? parseFloat(newRow.contractorPrTarget) : null;
    if (contPrB !== null && contPrB > 1.5) contPrB = contPrB / 100;

    let prA = newRow.prActual !== undefined && newRow.prActual !== null ? parseFloat(newRow.prActual) : null;
    if (prA !== null && prA > 1.5) prA = prA / 100;

    // Normalize Availability
    let gAvail = newRow.guaranteedAvailability !== undefined ? parseFloat(newRow.guaranteedAvailability) : null;
    if (gAvail !== null && gAvail > 1.5) gAvail = gAvail / 100;

    let aAvail = newRow.actualAvailability !== undefined ? parseFloat(newRow.actualAvailability) : null;
    if (aAvail !== null && aAvail > 1.5) aAvail = aAvail / 100;

    let siteName = newRow.siteName ? String(newRow.siteName).trim() : 'Unknown Site';
    if (siteName === '') siteName = 'Unknown Site';

    const poaActualRaw = newRow.poaActual !== undefined && newRow.poaActual !== null ? parseVal(newRow.poaActual) : null;
    const poaBudgetRaw = newRow.poaBudget !== undefined && newRow.poaBudget !== null ? parseVal(newRow.poaBudget) : null;

    // moduleTemp: preserve null (do NOT coerce to 0) — important so lossFactor uses STC fallback
    const modTempRaw = newRow.moduleTemp;
    const modTemp = (modTempRaw !== undefined && modTempRaw !== null && modTempRaw !== '')
        ? parseVal(modTempRaw) : null;

    return {
      date: dateObj,
      siteName: siteName,
      ghiBudget: parseVal(newRow.ghiBudget),
      ghiActual: newRow.ghiActual !== undefined && newRow.ghiActual !== null ? parseVal(newRow.ghiActual) : null,
      poaBudget: sanitizeIrradiance(poaBudgetRaw),
      poaActual: sanitizeIrradiance(poaActualRaw),
      kwhBudget: parseVal(newRow.kwhBudget),
      prBudget: prB,
      correctedPrBudget: corrPrB,
      contractorPrTarget: contPrB,
      kwhActual: newRow.kwhActual !== undefined && newRow.kwhActual !== null ? parseVal(newRow.kwhActual) : null,
      prActual: prA,
      systemCapacity: newRow.systemCapacity !== undefined ? parseVal(newRow.systemCapacity) : null,
      invCapacity: newRow.invCapacity !== undefined ? parseVal(newRow.invCapacity) : null,
      bessCapacity: newRow.bessCapacity !== undefined ? parseVal(newRow.bessCapacity) : null,
      bessOperationMode: newRow.bessOperationMode ? String(newRow.bessOperationMode) : null,
      cod: parseDate(newRow.cod),
      kwhForecast: 0,
      ghiForecast: 0,
      isForecast: false,
      notes: newRow.notes ? String(newRow.notes) : undefined,
      estimatedLoss: newRow.estimatedLoss !== undefined ? parseVal(newRow.estimatedLoss) : undefined,
      moduleTemp: modTemp,
      curtailment: newRow.curtailment !== undefined ? parseVal(newRow.curtailment) : null,
      thermalVariance: 0,
      lossFactor: 1.0,                   // computed in runProjection
      theoPoaAct: 0,                     // computed in runProjection
      theoPoaBud: 0,                     // computed in runProjection
      paeEnergy: newRow.paeEnergy !== undefined && newRow.paeEnergy !== null ? parseVal(newRow.paeEnergy) : null,
      guaranteedAvailability: gAvail,
      actualAvailability: aAvail
    };
  }).filter(d => isValid(d.date)).sort((a, b) => compareAsc(a.date, b.date));

  // Backfill Static Data (Capacities, COD) per Site
  const siteStaticData = new Map<string, {
      pv: number, inv: number, bess: number, cod: Date | null, modes: Set<string> 
  }>();

  // First pass: gather non-nulls
  parsed.forEach(d => {
      if (!siteStaticData.has(d.siteName)) {
          siteStaticData.set(d.siteName, { pv: 0, inv: 0, bess: 0, cod: null, modes: new Set() });
      }
      const staticRec = siteStaticData.get(d.siteName)!;
      
      if (d.systemCapacity && d.systemCapacity > staticRec.pv) staticRec.pv = d.systemCapacity;
      if (d.invCapacity && d.invCapacity > staticRec.inv) staticRec.inv = d.invCapacity;
      if (d.bessCapacity && d.bessCapacity > staticRec.bess) staticRec.bess = d.bessCapacity;
      if (d.cod && !staticRec.cod) staticRec.cod = d.cod; 
      if (d.cod && staticRec.cod && d.cod < staticRec.cod) staticRec.cod = d.cod;
      
      if (d.bessOperationMode) staticRec.modes.add(d.bessOperationMode);
  });

    // Carry-forward only — no MAX fallback. Pre-first-valid rows leave capacity at 0 (= inactive/pre-COD).
    // Rationale: filling pre-COD rows with future MAX capacity makes PR denominator falsely large,
    // understating PR for periods when the plant didn't yet exist.
    const result: DailyData[] = [];
    const lastKnownCap = new Map<string, { pv: number, inv: number, bess: number }>();
    for (const d of parsed) {
        const prev = lastKnownCap.get(d.siteName) ?? { pv: 0, inv: 0, bess: 0 };
        const next = {
            pv:   (d.systemCapacity && d.systemCapacity > 0) ? d.systemCapacity : prev.pv,
            inv:  (d.invCapacity   && d.invCapacity   > 0) ? d.invCapacity   : prev.inv,
            bess: (d.bessCapacity  && d.bessCapacity  > 0) ? d.bessCapacity  : prev.bess,
        };
        lastKnownCap.set(d.siteName, next);

        const rec = siteStaticData.get(d.siteName)!;
        result.push({
            ...d,
            systemCapacity: (d.systemCapacity && d.systemCapacity > 0) ? d.systemCapacity : next.pv,
            invCapacity:    (d.invCapacity   && d.invCapacity   > 0) ? d.invCapacity   : next.inv,
            bessCapacity:   (d.bessCapacity  && d.bessCapacity  > 0) ? d.bessCapacity  : next.bess,
            cod: d.cod || rec.cod,
        });
    }
    return result;
};

/**
 * Returns the LATEST (most-recent-date) capacity per site.
 * Use for: SystemSpecs header, "current installed" displays, trendPr lookback.
 * NOT for historical period aggregation — use `d.systemCapacity` per-row instead,
 * because aggregating with a single capacity over a period when capacity changed
 * would produce wrong PR ratios for sub-periods.
 */
export const getSiteCapacityMap = (data: DailyData[]): Map<string, number> => {
  const map = new Map<string, number>();
  const uniqueSites = Array.from(new Set(data.map(d => d.siteName)));

  uniqueSites.forEach(site => {
      // data is already sorted ascending by date in parseRawData,
      // so the last non-zero entry is the most recent installed capacity.
      const siteRows = data.filter(d => d.siteName === site && d.systemCapacity && d.systemCapacity > 0);
      if (siteRows.length === 0) { map.set(site, 0); return; }
      const latest = siteRows[siteRows.length - 1].systemCapacity!;
      map.set(site, latest);
  });
  return map;
};

/**
 * Returns the MAX capacity ever recorded per site.
 * Use for: nameplate / contract limits / historical peak displays.
 */
export const getSiteMaxCapacityMap = (data: DailyData[]): Map<string, number> => {
  const map = new Map<string, number>();
  const uniqueSites = Array.from(new Set(data.map(d => d.siteName)));
  uniqueSites.forEach(site => {
      const caps = data.filter(d => d.siteName === site && d.systemCapacity && d.systemCapacity > 0)
                       .map(d => d.systemCapacity!);
      map.set(site, caps.length > 0 ? Math.max(...caps) : 0);
  });
  return map;
};

export const runProjection = (
    data: DailyData[], 
    overrideCapacityMap?: Map<string, number>,
    prLookbackDays: number = 30
): SimulationResult => {
  if (data.length === 0) throw new Error("No valid data rows found.");

  const actuals = data.filter(d => 
    d.kwhActual !== null && d.kwhActual !== undefined && !isNaN(d.kwhActual) &&
    (d.kwhActual > 0 || (d.ghiActual !== null && d.ghiActual > 0))
  );
  
  const lastActualDate = actuals.length > 0 ? actuals[actuals.length - 1].date : data[0].date;

  const uniqueSites = Array.from(new Set(data.map(d => d.siteName)));
  let totalPVCapacity = 0;
  let totalInvCapacity = 0;
  let totalBessCapacity = 0;
  const codList: Date[] = [];
  const modeSet = new Set<string>();
  
  // SystemSpecs reflects CURRENT INSTALLED state (latest non-zero per site), not first row.
  // First-row was buggy: if plant expanded mid-year, header showed pre-expansion capacity.
  const latestCapMap = getSiteCapacityMap(data);
  uniqueSites.forEach(site => {
      const siteRows = data.filter(d => d.siteName === site);
      if (siteRows.length === 0) return;

      totalPVCapacity += latestCapMap.get(site) || 0;

      // Inv & BESS: scan from the end for the last non-zero entry
      let latestInv = 0;
      let latestBess = 0;
      for (let i = siteRows.length - 1; i >= 0; i--) {
          const r = siteRows[i];
          if (latestInv === 0 && r.invCapacity && r.invCapacity > 0) latestInv = r.invCapacity;
          if (latestBess === 0 && r.bessCapacity && r.bessCapacity > 0) latestBess = r.bessCapacity;
          if (latestInv !== 0 && latestBess !== 0) break;
      }
      totalInvCapacity += latestInv;
      totalBessCapacity += latestBess;

      // COD: parseRawData already backfilled cod from the earliest valid value
      const lastRow = siteRows[siteRows.length - 1];
      if (lastRow.cod) codList.push(lastRow.cod);
  });

  data.forEach(d => {
      if (d.bessOperationMode) modeSet.add(d.bessOperationMode);
  });

  let codDisplay = 'N/A';
  if (codList.length > 0) {
      const minCod = codList.reduce((a, b) => a < b ? a : b);
      if (uniqueSites.length > 1 && codList.length > 1) {
          const allSame = codList.every(d => d.getTime() === codList[0].getTime());
          codDisplay = allSame ? format(minCod, 'dd MMM yyyy') : 'Various';
      } else {
          codDisplay = format(minCod, 'dd MMM yyyy');
      }
  }

  const systemSpecs: SystemSpecs = {
      pvCapacityKW: totalPVCapacity,
      invCapacityKW: totalInvCapacity,
      bessCapacityKWh: totalBessCapacity,
      bessOperationMode: Array.from(modeSet).join(', ') || 'N/A',
      cod: codDisplay
  };

  const siteCapacityMap = overrideCapacityMap || getSiteCapacityMap(data);
  const finalSystemCapacityMW = totalPVCapacity / 1000;

  const lookbackStart = subDays(lastActualDate, prLookbackDays);
  const recentData = data.filter(d => 
    d.date > lookbackStart && d.date <= lastActualDate && d.kwhActual !== null && d.kwhActual >= 0 && d.ghiActual !== null && d.ghiActual > 0
  );

  let sumEnergyActual = 0;
  let sumEnergyTheoretical = 0;
  let sumEnergyBudgetTheoretical = 0;
  let sumPrBudgetWeighted = 0;

  recentData.forEach(d => {
      const cap = d.systemCapacity || siteCapacityMap.get(d.siteName) || 0;
      const irr = (d.poaActual && d.poaActual > 0) ? d.poaActual : (d.ghiActual || 0);
      const theoretical = irr * cap;
      sumEnergyActual += (d.kwhActual || 0);
      sumEnergyTheoretical += theoretical;

      if (theoretical > 0) {
        sumPrBudgetWeighted += (d.prBudget * theoretical);
        sumEnergyBudgetTheoretical += theoretical;
      }
  });

  const trendPr = sumEnergyTheoretical > 0 ? sumEnergyActual / sumEnergyTheoretical : (data[0].prBudget || 0.8);
  const trendPrBudget = sumEnergyBudgetTheoretical > 0 ? sumPrBudgetWeighted / sumEnergyBudgetTheoretical : (data[0].prBudget || 0.8);
  const trendPrVariance = trendPr - trendPrBudget;

  const mtdStartDate = startOfMonth(lastActualDate);
  const ytdStartDate = startOfYear(lastActualDate);

  // Temperature coefficient (γ) — per IEC 61724-3 capacity evaluation method.
  // γ = -0.0029 /°C is conservative; configurable per-site is recommended (see refactor backlog).
  const GAMMA = -0.0029;
  const STC_TEMP = 25;
  const T_FALLBACK = 25;             // safest fallback when moduleTemp missing → lossFactor = 1.0
  const LOSS_FACTOR_MIN = 0.5;       // saturation guards against extreme/anomalous temperatures
  const LOSS_FACTOR_MAX = 1.2;

  let ytdKwhBudget = 0;
  let ytdKwhActual = 0;
  let ytdVarianceIrradiance = 0;
  let ytdVariancePr = 0;
  let ytdTheoreticalActual = 0;
  let ytdTheoreticalBudget = 0;
  let ytdCorrectedPrBudgetSum = 0;
  let ytdCurtailment = 0;
  let ytdTempNum = 0;
  let ytdTempDenom = 0;
  let ytdVarianceThermal = 0; 
  let ytdGhiBudget = 0;
  let ytdGhiActual = 0;

  let mtdKwhBudget = 0;
  let mtdKwhActual = 0;
  let mtdVarianceIrradiance = 0;
  let mtdVariancePr = 0;
  let mtdTheoreticalActual = 0;
  let mtdTheoreticalBudget = 0;
  let mtdCorrectedPrBudgetSum = 0;
  let mtdCurtailment = 0;
  let mtdTempNum = 0;
  let mtdTempDenom = 0;
  let mtdVarianceThermal = 0;
  let mtdGhiBudget = 0;
  let mtdGhiActual = 0;

  const processedData = data.map(d => {
    const isPast = d.date <= lastActualDate;
    let kwhForecast = 0;
    let ghiForecast = 0;

    let effectiveKwhBudget = d.kwhBudget;
    const effectiveCapacity = d.systemCapacity || 0;       // per-row (carry-forward in parseRawData)

    // ── Temperature loss factor (IEC 61724-3): lossFactor = 1 + γ·ΔT, γ = -0.0029 /°C ──
    // Fallback: moduleTemp null → assume STC → lossFactor = 1.0 (no correction).
    const moduleT = d.moduleTemp !== null ? d.moduleTemp : T_FALLBACK;
    const deltaT  = moduleT - STC_TEMP;
    let lossFactor = 1 + (GAMMA * deltaT);
    if (lossFactor < LOSS_FACTOR_MIN) lossFactor = LOSS_FACTOR_MIN;
    if (lossFactor > LOSS_FACTOR_MAX) lossFactor = LOSS_FACTOR_MAX;

    // ── Irradiance: prefer measured POA; fallback to GHI for legacy datasets ──
    const poaAct = (d.poaActual && d.poaActual > 0) ? d.poaActual
                 : (d.ghiActual && d.ghiActual > 0) ? d.ghiActual : 0;
    const poaBud = (d.poaBudget && d.poaBudget > 0) ? d.poaBudget : d.ghiBudget;

    // ── Theoretical energies ──
    // Theo_POA_Act = lossFactor × cap × POA_Actual  (per user spec)
    // Theo_POA_Bud = cap × POA_Budget               (budget assumed at STC, no temp correction)
    const theoPoaAct = lossFactor * effectiveCapacity * poaAct;
    const theoPoaBud = effectiveCapacity * poaBud;

    // ── Thermal variance (signed, kept negative for "loss" in waterfall UI) ──
    // Magnitude = ((1/lossFactor) - 1) × kWh_Actual  (per user spec; positive when hot).
    // Stored as NEGATIVE so existing waterfall/loss UI (red = loss) renders correctly.
    let dailyThermalVariance = 0;
    if (d.kwhActual && d.kwhActual > 0 && lossFactor > 0 && d.moduleTemp !== null) {
        const magnitude = ((1 / lossFactor) - 1) * d.kwhActual;
        dailyThermalVariance = -magnitude;
    }

    if (!isPast) {
      if (poaBud > 0 && effectiveCapacity > 0) {
          kwhForecast = poaBud * trendPr * effectiveCapacity;
      } else if (d.kwhBudget > 0 && d.prBudget > 0) {
          kwhForecast = d.kwhBudget * (trendPr / d.prBudget);
      } else {
          kwhForecast = d.kwhBudget;
      }
      ghiForecast = d.ghiBudget;
      if ((!effectiveKwhBudget || effectiveKwhBudget === 0) && poaBud > 0) {
          const prForBudget = d.prBudget > 0 ? d.prBudget : trendPr;
          effectiveKwhBudget = poaBud * prForBudget * effectiveCapacity;
      }
    } else {
      if (d.kwhActual !== null) {
        const actual = d.kwhActual;
        if ((!effectiveKwhBudget || effectiveKwhBudget === 0) && d.ghiBudget > 0 && d.prBudget > 0) {
             effectiveKwhBudget = d.ghiBudget * d.prBudget * effectiveCapacity;
        }
        const budget = effectiveKwhBudget;
        let weatherCorrected = budget;
        // Weather correction still uses GHI ratio (legacy contract definition);
        // can switch to POA ratio later when budgets are POA-defined.
        if (d.ghiBudget > 0 && d.ghiActual !== null && d.ghiActual > 0) {
          weatherCorrected = budget * (d.ghiActual / d.ghiBudget);
        }
        const irrVar = weatherCorrected - budget;
        const prVar  = actual - weatherCorrected;

        if (d.date >= ytdStartDate) {
            ytdKwhBudget += budget;
            ytdKwhActual += actual;
            ytdVarianceIrradiance += irrVar;
            ytdVariancePr += prVar;
            ytdTheoreticalActual += theoPoaAct;
            ytdTheoreticalBudget += theoPoaBud;
            if (d.correctedPrBudget) ytdCorrectedPrBudgetSum += (d.correctedPrBudget * theoPoaBud);
            ytdCurtailment += (d.curtailment || 0);
            if (d.moduleTemp !== null) { ytdTempNum += (d.moduleTemp * poaAct); ytdTempDenom += poaAct; }
            ytdVarianceThermal += dailyThermalVariance;
            ytdGhiBudget += d.ghiBudget;
            ytdGhiActual += (d.ghiActual || 0);
        }

        if (d.date >= mtdStartDate) {
            mtdKwhBudget += budget;
            mtdKwhActual += actual;
            mtdVarianceIrradiance += irrVar;
            mtdVariancePr += prVar;
            mtdTheoreticalActual += theoPoaAct;
            mtdTheoreticalBudget += theoPoaBud;
            if (d.correctedPrBudget) mtdCorrectedPrBudgetSum += (d.correctedPrBudget * theoPoaBud);
            mtdCurtailment += (d.curtailment || 0);
            if (d.moduleTemp !== null) { mtdTempNum += (d.moduleTemp * poaAct); mtdTempDenom += poaAct; }
            mtdVarianceThermal += dailyThermalVariance;
            mtdGhiBudget += d.ghiBudget;
            mtdGhiActual += (d.ghiActual || 0);
        }
      }
    }
    return {
      ...d,
      kwhBudget: effectiveKwhBudget,
      kwhActual: isPast ? (d.kwhActual || 0) : null,
      kwhForecast: isPast ? 0 : kwhForecast,
      ghiActual: isPast ? (d.ghiActual || 0) : null,
      ghiForecast: isPast ? 0 : ghiForecast,
      isForecast: !isPast,
      thermalVariance: dailyThermalVariance,
      lossFactor,
      theoPoaAct,
      theoPoaBud,
    };
  });

  const monthlyMap = new Map<string, any>();

  processedData.forEach(d => {
    const monthKey = format(d.date, 'yyyy-MM');
    const monthLabel = format(d.date, 'MMM yyyy');

    if (!monthlyMap.has(monthKey)) {
      monthlyMap.set(monthKey, {
        periodKey: monthKey,
        periodLabel: monthLabel,
        kwhBudget: 0, kwhActual: 0, kwhForecast: 0,
        ghiBudget: 0, ghiActual: 0, ghiForecast: 0,
        theoreticalKwhBudget: 0, theoreticalKwhActual: 0, theoreticalKwhForecast: 0,
        weightedPrBudgetSum: 0, 
        weightedCorrectedPrBudgetSum: 0,
        weightedContractorPrTargetSum: 0,
        tempWeightedSum: 0, irradianceSumForTemp: 0, curtailmentSum: 0, energyActualSum: 0,
        varianceThermalSum: 0,
        paeEnergySum: 0,
        // Availability Accumulators
        guaranteedAvailabilitySum: 0,
        guaranteedAvailabilityCount: 0,
        actualAvailabilitySum: 0,
        actualAvailabilityCount: 0
      });
    }

    const m = monthlyMap.get(monthKey)!;
    // PER-ROW capacity. Pre-COD rows (cap=0) skipped from aggregation entirely
    // so theoretical denominators don't get inflated by future capacity.
    const capacity = d.systemCapacity || 0;

    m.kwhBudget += d.kwhBudget;
    m.kwhActual += (d.kwhActual || 0);
    m.kwhForecast += d.kwhForecast || 0;

    m.ghiBudget += d.ghiBudget;
    m.ghiActual += (d.ghiActual || 0);
    m.ghiForecast += d.ghiForecast || 0;

    if (d.curtailment && d.curtailment > 0) m.curtailmentSum += d.curtailment;

    if (d.paeEnergy && d.paeEnergy > 0) {
        m.paeEnergySum += d.paeEnergy;
    }

    if (d.guaranteedAvailability !== null) {
        m.guaranteedAvailabilitySum += d.guaranteedAvailability;
        m.guaranteedAvailabilityCount++;
    }
    if (d.actualAvailability !== null) {
        m.actualAvailabilitySum += d.actualAvailability;
        m.actualAvailabilityCount++;
    }

    if (capacity <= 0) return;     // pre-COD / inactive — skip energy theoreticals only

    // Theoreticals computed in daily loop; here we just sum them.
    m.theoreticalKwhBudget += d.theoPoaBud;
    m.weightedPrBudgetSum += (d.prBudget * d.theoPoaBud);

    if (d.correctedPrBudget) m.weightedCorrectedPrBudgetSum += (d.correctedPrBudget * d.theoPoaBud);
    if (d.contractorPrTarget) m.weightedContractorPrTargetSum += (d.contractorPrTarget * d.theoPoaBud);

    if (!d.isForecast && d.kwhActual !== null) {
        // d.theoPoaAct already includes lossFactor (temp correction) — single source of truth.
        m.theoreticalKwhActual += d.theoPoaAct;
        m.energyActualSum += (d.kwhActual || 0);

        const poaAct = (d.poaActual && d.poaActual > 0) ? d.poaActual
                     : (d.ghiActual && d.ghiActual > 0) ? d.ghiActual : 0;
        if (d.moduleTemp !== null && d.moduleTemp > -50 && poaAct > 0) {
            m.tempWeightedSum += (d.moduleTemp * poaAct);
            m.irradianceSumForTemp += poaAct;
        }
        m.varianceThermalSum += d.thermalVariance;
    }
    if (d.isForecast) {
        const irr = d.ghiForecast > 0 ? d.ghiForecast : d.ghiBudget;
        m.theoreticalKwhForecast += (irr * capacity);
    }
  });

  const monthlyData = Array.from(monthlyMap.values()).map(m => {
    const totalProjected = m.kwhActual + m.kwhForecast;
    const variancePct = m.kwhBudget > 0 ? ((totalProjected - m.kwhBudget) / m.kwhBudget) * 100 : 0;
    
    const prBudget = m.theoreticalKwhBudget > 0 ? m.weightedPrBudgetSum / m.theoreticalKwhBudget : 0;
    const correctedPrBudget = m.theoreticalKwhBudget > 0 ? m.weightedCorrectedPrBudgetSum / m.theoreticalKwhBudget : null;
    const contractorPrTarget = m.theoreticalKwhBudget > 0 ? m.weightedContractorPrTargetSum / m.theoreticalKwhBudget : null;

    const prActual = m.theoreticalKwhActual > 0 ? m.kwhActual / m.theoreticalKwhActual : null;
    const prForecast = m.theoreticalKwhForecast > 0 ? m.kwhForecast / m.theoreticalKwhForecast : null;

    const curtailment = m.curtailmentSum;
    const moduleTemp = m.irradianceSumForTemp > 0 ? m.tempWeightedSum / m.irradianceSumForTemp : null;

    // Corrected PR (user spec, IEC 61724-3-style):
    //   Corr_PR = Σ kWh_Actual / Σ Theo_POA_Act
    // Temperature correction is already in the denominator (theoPoaAct = lossFactor × cap × POA).
    // Curtailment is NOT credited — kWh_Actual is post-curtailment per data convention.
    const correctedPr = m.theoreticalKwhActual > 0
        ? m.energyActualSum / m.theoreticalKwhActual
        : null;

    // Availability Averages
    const guaranteedAvailability = m.guaranteedAvailabilityCount > 0 ? m.guaranteedAvailabilitySum / m.guaranteedAvailabilityCount : null;
    const actualAvailability = m.actualAvailabilityCount > 0 ? m.actualAvailabilitySum / m.actualAvailabilityCount : null;

    return {
        periodKey: m.periodKey,
        periodLabel: m.periodLabel,
        kwhBudget: m.kwhBudget,
        kwhActual: m.kwhActual,
        kwhForecast: m.kwhForecast,
        totalProjected,
        variancePct,
        ghiBudget: m.ghiBudget,
        ghiActual: m.ghiActual,
        ghiForecast: m.ghiForecast,
        prBudget,
        correctedPrBudget,
        contractorPrTarget,
        prActual,
        prForecast,
        moduleTemp,
        curtailment,
        correctedPr,
        cumulativeKwhBudget: 0, cumulativeKwhProjected: 0, cumulativeKwhActual: null, cumulativeKwhForecast: null,
        varianceThermal: m.varianceThermalSum,
        paeEnergy: m.paeEnergySum,
        paeGuaranteed: m.paeEnergySum * 0.925,
        guaranteedAvailability,
        actualAvailability
    } as AggregatedData;
  }).sort((a, b) => a.periodKey.localeCompare(b.periodKey));

  let runningBudget = 0;
  let runningProjected = 0;
  const tempWithRunning = monthlyData.map((m) => {
      runningBudget += m.kwhBudget;
      runningProjected += m.totalProjected;
      return { ...m, cumulativeKwhBudget: runningBudget, cumulativeKwhProjected: runningProjected, isFullyActual: m.kwhForecast === 0 };
  });

  let lastActualIndex = -1;
  for (let i = 0; i < tempWithRunning.length; i++) {
      if (tempWithRunning[i].isFullyActual) lastActualIndex = i; else break; 
  }

  const monthlyDataWithCumulative = tempWithRunning.map((m, i) => {
      return {
          ...m,
          cumulativeKwhActual: i <= lastActualIndex ? m.cumulativeKwhProjected : null,
          cumulativeKwhForecast: i >= lastActualIndex ? m.cumulativeKwhProjected : null
      };
  });

  const totalBudget = monthlyData.reduce((sum, m) => sum + m.kwhBudget, 0);
  const totalProjected = monthlyData.reduce((sum, m) => sum + m.totalProjected, 0);
  const varianceYearEnd = totalBudget > 0 ? ((totalProjected - totalBudget) / totalBudget) * 100 : 0;

  // PR formulas (per user spec, IEC 61724-3 aligned):
  //   PR Actual    = Σ kWh_Actual / Σ Theo_POA_Act_uncorrected (here: theoAct already includes temp correction)
  //   PR Budget    = Σ kWh_Budget / Σ Theo_POA_Bud
  //   Corr PR      = Σ kWh_Actual / Σ Theo_POA_Act  (temp correction baked into denominator)
  //   Corr PR Bud  = weighted Σ correctedPrBudget × Theo_POA_Bud / Σ Theo_POA_Bud
  // Curtailment is NOT credited (kWh_Actual already includes curtailment per data convention).
  // _cur, _tNum, _tDen, _varianceThermal are kept in the signature for backward compatibility
  // but no longer participate in the math.
  const calcPrs = (
      kwhAct: number, kwhBud: number,
      theoAct: number, theoBud: number,
      _cur: number, _tNum: number, _tDen: number,
      corrPrBudSum: number, _varianceThermal: number
  ) => {
      const prBudget          = theoBud > 0 ? kwhBud / theoBud : 0;
      const prActual          = theoAct > 0 ? kwhAct / theoAct : 0;
      const correctedPrBudget = theoBud > 0 ? corrPrBudSum / theoBud : 0;
      const correctedPr       = theoAct > 0 ? kwhAct / theoAct : 0;
      return { prBudget, prActual, correctedPr, correctedPrBudget };
  };

  const ytdPrs = calcPrs(ytdKwhActual, ytdKwhBudget, ytdTheoreticalActual, ytdTheoreticalBudget, ytdCurtailment, ytdTempNum, ytdTempDenom, ytdCorrectedPrBudgetSum, ytdVarianceThermal);
  const mtdPrs = calcPrs(mtdKwhActual, mtdKwhBudget, mtdTheoreticalActual, mtdTheoreticalBudget, mtdCurtailment, mtdTempNum, mtdTempDenom, mtdCorrectedPrBudgetSum, mtdVarianceThermal);

  const ytdThermalVar = ytdVarianceThermal;
  const mtdThermalVar = mtdVarianceThermal;

  const ytdCurtailmentVar = -ytdCurtailment;
  const mtdCurtailmentVar = -mtdCurtailment;

  const ytdOtherPrVar = ytdVariancePr - ytdThermalVar - ytdCurtailmentVar;
  const mtdOtherPrVar = mtdVariancePr - mtdThermalVar - mtdCurtailmentVar;

  const lossAnalysisYTD: LossFactors = {
    periodLabel: 'YTD',
    kwhBudget: ytdKwhBudget, kwhActual: ytdKwhActual,
    varianceTotal: ytdKwhActual - ytdKwhBudget,
    varianceIrradiance: ytdVarianceIrradiance, 
    variancePr: ytdVariancePr,
    varianceThermal: ytdThermalVar,
    varianceCurtailment: ytdCurtailmentVar,
    varianceOtherPr: ytdOtherPrVar,
    prBudget: ytdPrs.prBudget, correctedPrBudget: ytdPrs.correctedPrBudget,
    prActual: ytdPrs.prActual, correctedPr: ytdPrs.correctedPr,
    ghiBudget: ytdGhiBudget, ghiActual: ytdGhiActual
  };

  const lossAnalysisMTD: LossFactors = {
    periodLabel: format(mtdStartDate, 'MMM yyyy'),
    kwhBudget: mtdKwhBudget, kwhActual: mtdKwhActual,
    varianceTotal: mtdKwhActual - mtdKwhBudget,
    varianceIrradiance: mtdVarianceIrradiance, 
    variancePr: mtdVariancePr,
    varianceThermal: mtdThermalVar,
    varianceCurtailment: mtdCurtailmentVar,
    varianceOtherPr: mtdOtherPrVar,
    prBudget: mtdPrs.prBudget, correctedPrBudget: mtdPrs.correctedPrBudget,
    prActual: mtdPrs.prActual, correctedPr: mtdPrs.correctedPr,
    ghiBudget: mtdGhiBudget, ghiActual: mtdGhiActual
  };

  return {
    dailyData: processedData,
    monthlyData: monthlyDataWithCumulative,
    metrics: {
      lastDataDate: lastActualDate,
      trendPr, trendPrVariance,
      impliedCapacity: finalSystemCapacityMW, 
      totalBudget, totalProjected, varianceYearEnd,
      systemSpecs 
    },
    lossAnalysisYTD, lossAnalysisMTD
  };
};
