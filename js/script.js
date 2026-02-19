/* Shared helpers */
let draggingRow = null;
let outlaysTable = null;
let toggleSailing = null;
let densityComfortable = null;
let densityDense = null;
let printRestoreDensity = null;
let outlaysCurrency = null;
let roundPdaPrices = null;
const STORAGE_KEYS = {
  vesselName: 'pda_vessel_name',
  gt: 'pda_gt',
  quantity: 'pda_quantity',
  indexState: 'pda_index_state',
  towageTotal: 'pda_towage_total',
  towageArrivalCount: 'pda_towage_arrival_count',
  towageDepartureCount: 'pda_towage_departure_count',
  tugsState: 'pda_tugs_state',
  towageTotalSailing: 'pda_towage_total_sailing',
  towageArrivalCountSailing: 'pda_towage_arrival_count_sailing',
  towageDepartureCountSailing: 'pda_towage_departure_count_sailing',
  tugsStateSailing: 'pda_tugs_state_sailing',
  lightDuesState: 'pda_light_dues_state',
  lightDuesStateSailing: 'pda_light_dues_state_sailing',
  lightDuesAmountPda: 'pda_light_dues_amount_pda',
  lightDuesTariffPda: 'pda_light_dues_tariff_pda',
  lightDuesAmountSailing: 'pda_light_dues_amount_sailing',
  lightDuesTariffSailing: 'pda_light_dues_tariff_sailing',
  portDuesState: 'pda_port_dues_state',
  portDuesStateSailing: 'pda_port_dues_state_sailing',
  portDuesAmountPda: 'pda_port_dues_amount_pda',
  portDuesAmountSailing: 'pda_port_dues_amount_sailing',
  portDuesCargoAmountPda: 'pda_port_dues_cargo_amount_pda',
  portDuesCargoAmountSailing: 'pda_port_dues_cargo_amount_sailing',
  portDuesBunkeringAmountPda: 'pda_port_dues_bunkering_amount_pda',
  portDuesBunkeringAmountSailing: 'pda_port_dues_bunkering_amount_sailing',
  mooringState: 'pda_mooring_state',
  mooringStateSailing: 'pda_mooring_state_sailing',
  mooringAmountPda: 'pda_mooring_amount_pda',
  mooringAmountSailing: 'pda_mooring_amount_sailing'
};

const INDEX_FIELD_IDS = [
  'logoLeftNote',
  'titleNote',
  'dateInput',
  'vesselNameIndex',
  'grossTonnage',
  'lengthOverall',
  'bowThrusterFitted',
  'portInput',
  'berthTerminal',
  'operationsInput',
  'cargoInput',
  'quantityInput',
  'toggleSailing',
  'roundPdaPrices',
  'outlaysCurrency',
  'bankRate'
];

const TUG_STORAGE = {
  standard: {
    towageTotal: STORAGE_KEYS.towageTotal,
    towageArrivalCount: STORAGE_KEYS.towageArrivalCount,
    towageDepartureCount: STORAGE_KEYS.towageDepartureCount,
    tugsState: STORAGE_KEYS.tugsState
  },
  sailing: {
    towageTotal: STORAGE_KEYS.towageTotalSailing,
    towageArrivalCount: STORAGE_KEYS.towageArrivalCountSailing,
    towageDepartureCount: STORAGE_KEYS.towageDepartureCountSailing,
    tugsState: STORAGE_KEYS.tugsStateSailing
  }
};

const LIGHT_DUES_TARIFFS = {
  cargo: {
    rate30: 0.5088,
    rate12: 1.696,
    label30: '0,5088',
    label12: '1,696'
  },
  tanker: {
    rate30: 0.579768,
    rate12: 1.940448,
    label30: '0,579768',
    label12: '1,940448'
  },
  roroCargo: {
    rate30: 0.2,
    rate12: 0.664,
    label30: '0,2',
    label12: '0,664'
  },
  passengerRoroFerry: {
    rate30: 0.2014,
    rate12: 0.6784,
    label30: '0,2014',
    label12: '0,6784'
  },
  supply: {
    rate30: 0.5088,
    rate12: 1.696,
    label30: '0,5088',
    label12: '1,696'
  },
  tugboatPusher: {
    rate30: 0.1325,
    rate12: 0.4399,
    label30: '0,1325',
    label12: '0,4399'
  },
  fishing: {
    rate30: 0.2014,
    rate12: 0.6784,
    label30: '0,2014',
    label12: '0,6784'
  },
  technicalCraft: {
    rate30: 0.2014,
    rate12: 0.6784,
    label30: '0,2014',
    label12: '0,6784'
  },
  nonSelfPropelled: {
    rate30: 0.2014,
    rate12: 0.6784,
    label30: '0,2014',
    label12: '0,6784'
  },
  otherUnidentified: {
    rate30: 0.5088,
    rate12: 1.696,
    label30: '0,5088',
    label12: '1,696'
  },
  bulk: {
    le30000: {
      rate30: 0.45792,
      rate12: 1.5264,
      label30: '0,45792',
      label12: '1,5264'
    },
    '30001to50000': {
      rate30: 0.4028,
      rate12: 1.378,
      label30: '0,4028',
      label12: '1,378'
    },
    gt50000: {
      rate30: 0.2968,
      rate12: 0.8904,
      label30: '0,2968',
      label12: '0,8904'
    }
  },
  container: {
    le40000: {
      rate30: 0.2332,
      rate12: 1.06,
      label30: '0,2332',
      label12: '1,06'
    },
    '40001to65000': {
      rate30: 0.10388,
      rate12: 0.34238,
      label30: '0,10388',
      label12: '0,34238'
    },
    '65001to100000': {
      rate30: 0.0742,
      rate12: 0.2862,
      label30: '0,0742',
      label12: '0,2862'
    },
    gt100000: {
      rate30: 0.053,
      rate12: 0.1908,
      label30: '0,053',
      label12: '0,1908'
    }
  },
  cruise: {
    le20000: {
      rate30: 0.25,
      rate12: 0.875,
      label30: '0,25',
      label12: '0,875'
    },
    '20001to50000': {
      rate30: 0.20125,
      rate12: 0.6875,
      label30: '0,20125',
      label12: '0,6875'
    },
    '50001to80000': {
      rate30: 0.1725,
      rate12: 0.575,
      label30: '0,1725',
      label12: '0,575'
    },
    gt80000: {
      rate30: 0.17125,
      rate12: 0.5,
      label30: '0,17125',
      label12: '0,5'
    }
  }
};

const LIGHT_DUES_TIER_OPTIONS = {
  bulk: [
    { value: 'le30000', label: 'if ≤ 30.000 GT', max: 30000 },
    { value: '30001to50000', label: 'if 30.001 - 50.000 GT', min: 30001, max: 50000 },
    { value: 'gt50000', label: 'if > 50.000 GT', min: 50001 }
  ],
  container: [
    { value: 'le40000', label: 'if ≤ 40.000 GT', max: 40000 },
    { value: '40001to65000', label: 'if 40.001 - 65.000 GT', min: 40001, max: 65000 },
    { value: '65001to100000', label: 'if 65.001 - 100.000 GT', min: 65001, max: 100000 },
    { value: 'gt100000', label: 'if > 100.000 GT', min: 100001 }
  ],
  cruise: [
    { value: 'le20000', label: 'if ≤ 20.000 GT', max: 20000 },
    { value: '20001to50000', label: 'if 20.001 - 50.000 GT', min: 20001, max: 50000 },
    { value: '50001to80000', label: 'if 50.001 - 80.000 GT', min: 50001, max: 80000 },
    { value: 'gt80000', label: 'if > 80.000 GT', min: 80001 }
  ]
};

const PORT_DUES_CARGO_TYPES = {
  bulkCargo: { rate: 0.6, label: '0,60' },
  cementGravelSawdustBulkStoneSlagPetroleumCokeClinkerCoal: { rate: 0.33, label: '0,33' },
  grains: { rate: 0.5, label: '0,50' },
  sulphur: { rate: 0.48, label: '0,48' },
  scrapIron: { rate: 0.33, label: '0,33' },
  copper: { rate: 0.4, label: '0,40' },
  liquidCargo: { rate: 1, label: '1,00' },
  generalCargo: { rate: 1, label: '1,00' },
  baggedCement: { rate: 0.8, label: '0,80' },
  explosiveCargo: { rate: 2.9, label: '2,90' },
  heavyCargo: { rate: 2, label: '2,00' },
  containerPerGt: { rate: 0.84, label: '0,84' }
};

const MOORING_GT_TARIFFS = [
  { min: 0, max: 250, amount: 8.4 },
  { min: 251, max: 500, amount: 14.7 },
  { min: 501, max: 1000, amount: 21.7 },
  { min: 1001, max: 2000, amount: 30.8 },
  { min: 2001, max: 4000, amount: 43.4 },
  { min: 4001, max: 7000, amount: 61.04 },
  { min: 7001, max: 11000, amount: 86.1 },
  { min: 11001, max: 15000, amount: 121.8 },
  { min: 15001, max: 20000, amount: 162.4 },
  { min: 20001, max: 25000, amount: 203 },
  { min: 25001, max: 30000, amount: 243.6 },
  { min: 30001, max: 35000, amount: 285.6 }
];
const BUNKER_ROW_DESCRIPTION = 'BUNKER (EUR 0,30 x loaded bunker / MT )';

function safeStorageGet(key) {
  try {
    return localStorage.getItem(key);
  } catch (error) {
    return null;
  }
}

function safeStorageSet(key, value) {
  try {
    localStorage.setItem(key, String(value));
  } catch (error) {
    // ignore storage errors
  }
}

function safeStorageRemove(key) {
  try {
    localStorage.removeItem(key);
  } catch (error) {
    // ignore storage errors
  }
}

function readFieldValue(field) {
  if (!field) return '';
  if (field.type === 'checkbox') return Boolean(field.checked);
  return field.value;
}

function applyFieldValue(field, value) {
  if (!field || value === undefined || value === null) return;
  if (field.type === 'checkbox') {
    field.checked = Boolean(value);
    return;
  }
  field.value = String(value);
}

function saveIndexState() {
  const outlaysBody = document.getElementById('outlaysBody');
  if (!outlaysBody) return;

  const fields = {};
  INDEX_FIELD_IDS.forEach((id) => {
    const field = document.getElementById(id);
    if (!field) return;
    fields[id] = readFieldValue(field);
  });

  const outlaysValues = Array.from(outlaysBody.querySelectorAll('tr')).map((row) => {
    return Array.from(row.querySelectorAll('textarea, input')).map((field) => field.value);
  });

  const state = {
    fields,
    outlaysHtml: outlaysBody.innerHTML,
    outlaysValues,
    density: getDensityMode()
  };
  safeStorageSet(STORAGE_KEYS.indexState, JSON.stringify(state));
}

function restoreIndexState() {
  const outlaysBody = document.getElementById('outlaysBody');
  if (!outlaysBody) return false;

  const raw = safeStorageGet(STORAGE_KEYS.indexState);
  if (!raw) return false;

  let state = null;
  try {
    state = JSON.parse(raw);
  } catch (error) {
    return false;
  }
  if (!state || typeof state !== 'object') return false;

  if (typeof state.outlaysHtml === 'string' && state.outlaysHtml.trim()) {
    outlaysBody.innerHTML = state.outlaysHtml;
  }

  if (Array.isArray(state.outlaysValues)) {
    const rows = outlaysBody.querySelectorAll('tr');
    state.outlaysValues.forEach((values, rowIndex) => {
      const row = rows[rowIndex];
      if (!row || !Array.isArray(values)) return;
      const fields = row.querySelectorAll('textarea, input');
      values.forEach((value, fieldIndex) => {
        const field = fields[fieldIndex];
        if (!field) return;
        field.value = value == null ? '' : String(value);
      });
    });
  }

  if (state.fields && typeof state.fields === 'object') {
    INDEX_FIELD_IDS.forEach((id) => {
      const field = document.getElementById(id);
      if (!field) return;
      applyFieldValue(field, state.fields[id]);
    });
  }

  if (state.density === 'comfortable' || state.density === 'dense') {
    setDensity(state.density);
  }
  return true;
}

let isRestoringTugs = false;

function isSailingTugsPage() {
  return document.body && document.body.classList.contains('page-tugs-sailing');
}

function getTugStorageKeys(isSailing) {
  return isSailing ? TUG_STORAGE.sailing : TUG_STORAGE.standard;
}

function getTugsState(isSailing) {
  const keys = getTugStorageKeys(isSailing);
  const raw = safeStorageGet(keys.tugsState);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (error) {
    return null;
  }
}

function saveTugsState() {
  if (isRestoringTugs) return;
  const tugCards = document.getElementById('tugCards');
  if (!tugCards) return;

  const keys = getTugStorageKeys(isSailingTugsPage());
  const tugs = Array.from(tugCards.querySelectorAll('.card')).map((card) => {
    const id = card.id.split('_')[1];
    return {
      op: document.getElementById(`op_${id}`)?.value || 'arrival',
      voyage: document.getElementById(`voyage_${id}`)?.value || '1',
      assist: document.getElementById(`assist_${id}`)?.value || '0.5',
      imo: document.getElementById(`imo_${id}`)?.checked || false,
      lines: document.getElementById(`lines_${id}`)?.checked || false,
      kw: document.getElementById(`kw_${id}`)?.checked || false,
      voy_ot: document.getElementById(`voy_ot_${id}`)?.value || '0',
      assist_ot: document.getElementById(`assist_ot_${id}`)?.value || '0'
    };
  });

  const state = {
    vesselName: document.getElementById('vesselName')?.value || '',
    gt: document.getElementById('gt')?.value || '',
    imoMaster: document.getElementById('imoMaster')?.checked || false,
    linesMaster: document.getElementById('linesMaster')?.checked || false,
    tugs
  };

  safeStorageSet(keys.tugsState, JSON.stringify(state));
}

function restoreTugsState() {
  const tugCards = document.getElementById('tugCards');
  if (!tugCards) return false;

  const state = getTugsState(isSailingTugsPage());
  if (!state || !Array.isArray(state.tugs)) return false;

  isRestoringTugs = true;
  tugCards.innerHTML = '';
  tugCount = 0;

  state.tugs.forEach((tug) => {
    addTug();
    const id = tugCount;
    const op = document.getElementById(`op_${id}`);
    const voyage = document.getElementById(`voyage_${id}`);
    const assist = document.getElementById(`assist_${id}`);
    const imo = document.getElementById(`imo_${id}`);
    const lines = document.getElementById(`lines_${id}`);
    const kw = document.getElementById(`kw_${id}`);
    const voyOt = document.getElementById(`voy_ot_${id}`);
    const assistOt = document.getElementById(`assist_ot_${id}`);

    if (op) op.value = tug.op || 'arrival';
    if (voyage) voyage.value = tug.voyage || '1';
    if (assist) assist.value = tug.assist || '0.5';
    if (imo) imo.checked = Boolean(tug.imo);
    if (lines) lines.checked = Boolean(tug.lines);
    if (kw) kw.checked = Boolean(tug.kw);
    if (voyOt) voyOt.value = tug.voy_ot || '0';
    if (assistOt) assistOt.value = tug.assist_ot || '0';
  });

  const vesselName = document.getElementById('vesselName');
  if (vesselName) {
    vesselName.value = state.vesselName || '';
    updateVesselNameFromStorage(vesselName);
  }

  const imoMasterInput = document.getElementById('imoMaster');
  if (imoMasterInput) {
    imoMasterInput.checked = Boolean(state.imoMaster);
  }
  const linesMasterInput = document.getElementById('linesMaster');
  if (linesMasterInput) {
    linesMasterInput.checked = Boolean(state.linesMaster);
  }

  const gtInput = document.getElementById('gt');
  if (gtInput) {
    if (state.gt) gtInput.value = state.gt;
    const gtValue = Number(gtInput.value);
    const tariff = getTariffFromGT(gtValue);
    const tariffInput = document.getElementById('tariff');
    if (tariffInput) tariffInput.value = tariff || '';
  }

  isRestoringTugs = false;
  updateTugTitles();
  syncImoMaster();
  syncLinesMaster();
  calculate();
  saveTugsState();
  return true;
}

function addOutlayRow(values = {}) {
  const tbody = document.getElementById('outlaysBody');
  const template = document.getElementById('outlayRowTemplate');
  if (!tbody || !template) return;

  const row = template.content.firstElementChild.cloneNode(true);
  const desc = row.querySelector('[data-field="desc"]');
  const pda = row.querySelector('[data-field="pda"]');
  const sailing = row.querySelector('[data-field="sailing"]');

  if (desc) desc.value = values.desc || '';
  if (pda) pda.value = values.pda || '';
  if (sailing) sailing.value = values.sailing || '';

  const bankRow = tbody.querySelector('tr[data-row="bank-charges"]');
  if (bankRow) {
    tbody.insertBefore(row, bankRow);
  } else {
    tbody.appendChild(row);
  }
  if (desc && desc.tagName === 'TEXTAREA') autoResizeTextarea(desc);
  wrapMoneyFields();
  decorateMoneyEditCells();
  recalcOutlayTotals();
  saveIndexState();
}

function clearOutlayRow(row) {
  row.querySelectorAll('input, textarea').forEach((field) => {
    field.value = '';
    if (field.tagName === 'TEXTAREA') autoResizeTextarea(field);
  });
  recalcOutlayTotals();
  saveIndexState();
}

function decorateOutlayRows() {
  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return;

  tbody.querySelectorAll('tr').forEach((row) => {
    const descCell = row.querySelector('td.desc');
    if (!descCell) return;
    if (row.dataset.row === 'bank-charges') return;
    if (descCell.querySelector('.row-item')) return;

    const input = descCell.querySelector('input, textarea');
    if (!input) return;

    const wrapper = document.createElement('div');
    wrapper.className = 'row-item';

    descCell.removeChild(input);
    const handle = document.createElement('button');
    handle.type = 'button';
    handle.className = 'row-handle';
    handle.setAttribute('draggable', 'true');
    handle.setAttribute('aria-label', 'Move row');
    const handleImg = document.createElement('img');
    handleImg.src = 'assets/icons/move_grabber_48×48.png';
    handleImg.alt = '';
    handle.appendChild(handleImg);
    wrapper.appendChild(handle);

    wrapper.appendChild(input);

    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'row-remove';
    removeBtn.setAttribute('aria-label', 'Remove row');
    const removeImg = document.createElement('img');
    removeImg.src = 'assets/icons/remove_48×48.png';
    removeImg.alt = '';
    removeBtn.appendChild(removeImg);
    wrapper.appendChild(removeBtn);

    descCell.appendChild(wrapper);
  });
}

function autoResizeTextarea(textarea) {
  if (!textarea) return;
  textarea.style.height = 'auto';
  textarea.style.height = `${textarea.scrollHeight}px`;
}

function stripTowageCounts(text) {
  return String(text || '').replace(/\s*\(Arrival tugs:.*?Departure tugs:.*?\)\s*$/i, '').trim();
}

function updateTowageFromStorage() {
  const towageRow = document.querySelector('tr[data-row="towage"]');
  if (!towageRow) return;

  const descInput = towageRow.querySelector('textarea.cell-input.text');
  if (descInput && !descInput.dataset.baseText) {
    descInput.dataset.baseText = 'TOWAGE';
  }

  const pdaKeys = getTugStorageKeys(false);
  const sailingKeys = getTugStorageKeys(true);

  const totalRaw = safeStorageGet(pdaKeys.towageTotal);
  const sailingTotalRaw = safeStorageGet(sailingKeys.towageTotal);
  const arrivalCountRaw = safeStorageGet(pdaKeys.towageArrivalCount);
  const departureCountRaw = safeStorageGet(pdaKeys.towageDepartureCount);

  const arrivalCount = Number(arrivalCountRaw);
  const departureCount = Number(departureCountRaw);
  if (descInput) {
    descInput.dataset.baseText = 'TOWAGE';
    const safeArrival = Number.isFinite(arrivalCount) ? arrivalCount : 0;
    const safeDeparture = Number.isFinite(departureCount) ? departureCount : 0;
    descInput.value = `TOWAGE (Arrival tugs: ${safeArrival}, Departure tugs: ${safeDeparture})`;
    autoResizeTextarea(descInput);
  }

  const totalValue = Number(totalRaw);
  const pdaInput = towageRow.querySelector('td:nth-child(2) input');
  const sailingInput = towageRow.querySelector('td:nth-child(3) input');

  if (Number.isFinite(totalValue) && pdaInput) {
    const formattedTowage = formatMoneyValue(totalValue);
    pdaInput.value = formattedTowage;
    if (isPdaRoundingEnabled()) {
      pdaInput.dataset.rawValue = formattedTowage;
    } else {
      delete pdaInput.dataset.rawValue;
    }
  }
  const sailingTotalValue = Number(sailingTotalRaw);
  if (Number.isFinite(sailingTotalValue) && sailingInput) {
    sailingInput.value = formatMoneyValue(sailingTotalValue);
  }
  if (Number.isFinite(totalValue) || Number.isFinite(sailingTotalValue)) {
    recalcOutlayTotals();
    saveIndexState();
  }
}

function findLightDuesRow() {
  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return null;
  return Array.from(tbody.querySelectorAll('tr')).find((row) => {
    const descField = row.querySelector('td.desc textarea, td.desc input');
    const desc = String(descField?.value || '').trim().toUpperCase();
    return desc.startsWith('LIGHT DUES');
  }) || null;
}

function getCurrentGtValue() {
  const gtInput = document.getElementById('grossTonnage');
  const raw = gtInput ? gtInput.value : safeStorageGet(STORAGE_KEYS.gt);
  return parseMoneyValue(raw);
}

function getCurrentQuantityValue() {
  const quantityInput = document.getElementById('quantityInput');
  const raw = quantityInput ? quantityInput.value : safeStorageGet(STORAGE_KEYS.quantity);
  return parseMoneyValue(raw);
}

function parseStoredJson(raw) {
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch (error) {
    return null;
  }
}

function getSelectedPeriodFromState(state) {
  if (!state || typeof state !== 'object') return '30';
  if (state.period12 === true) return '12';
  return '30';
}

function getLightDuesCalculationFromState(state, gt) {
  if (!state || typeof state !== 'object' || !Number.isFinite(gt) || gt <= 0) return null;
  const type = typeof state.type === 'string' ? state.type : 'cargo';
  const tierBand = state.tierBand || state.bulkBand || '';
  const tariff = getLightDuesTariff(type, tierBand, gt);
  const selectedPeriod = getSelectedPeriodFromState(state);
  if (!tariff) return null;
  if (selectedPeriod === '12') {
    return {
      tariffLabel: tariff.label12,
      amount: gt * tariff.rate12
    };
  }
  return {
    tariffLabel: tariff.label30,
    amount: gt * tariff.rate30
  };
}

function getCoefficientFromDescription(descValue) {
  const match = String(descValue || '').match(/EUR\s*([0-9.,]+)\s*x/i);
  if (!match || !match[1]) return null;
  const coefficient = parseMoneyValue(match[1]);
  if (!Number.isFinite(coefficient) || coefficient <= 0) return null;
  return {
    value: coefficient,
    label: match[1].replace(/\s+/g, '')
  };
}

function updateLightDuesFromStorage() {
  const lightDuesRow = findLightDuesRow();
  if (!lightDuesRow) return;

  let changed = false;
  const gt = getCurrentGtValue();
  const descInput = lightDuesRow.querySelector('td.desc textarea, td.desc input');
  const pdaState = parseStoredJson(safeStorageGet(STORAGE_KEYS.lightDuesState));
  const sailingState = parseStoredJson(safeStorageGet(STORAGE_KEYS.lightDuesStateSailing));
  const pdaCalc = getLightDuesCalculationFromState(pdaState, gt);
  const sailingCalc = getLightDuesCalculationFromState(sailingState, gt);
  const descCoefficient = getCoefficientFromDescription(descInput ? descInput.value : '');

  const pdaInput = lightDuesRow.querySelector('td:nth-child(2) input.cell-input.money');
  let pdaTariffLabel = safeStorageGet(STORAGE_KEYS.lightDuesTariffPda);
  let pdaAmount = Number(safeStorageGet(STORAGE_KEYS.lightDuesAmountPda));
  if (pdaCalc) {
    pdaTariffLabel = pdaCalc.tariffLabel;
    pdaAmount = pdaCalc.amount;
    safeStorageSet(STORAGE_KEYS.lightDuesTariffPda, pdaTariffLabel);
    safeStorageSet(STORAGE_KEYS.lightDuesAmountPda, pdaAmount.toFixed(2));
  } else if (Number.isFinite(gt) && gt > 0 && descCoefficient) {
    pdaTariffLabel = descCoefficient.label;
    pdaAmount = gt * descCoefficient.value;
    safeStorageSet(STORAGE_KEYS.lightDuesTariffPda, pdaTariffLabel);
    safeStorageSet(STORAGE_KEYS.lightDuesAmountPda, pdaAmount.toFixed(2));
  }

  if (pdaTariffLabel && descInput) {
    const nextDesc = `LIGHT DUES (EUR ${pdaTariffLabel} x vsl's GRT)`;
    if (descInput.value !== nextDesc) {
      descInput.value = nextDesc;
      if (descInput.tagName === 'TEXTAREA') autoResizeTextarea(descInput);
      changed = true;
    }
  }

  if (pdaInput && Number.isFinite(pdaAmount)) {
    const formatted = formatMoneyValue(pdaAmount);
    if (pdaInput.value !== formatted) {
      pdaInput.value = formatted;
      if (isPdaRoundingEnabled()) pdaInput.dataset.rawValue = formatted;
      else delete pdaInput.dataset.rawValue;
      changed = true;
    }
  }

  const sailingInput = lightDuesRow.querySelector('td:nth-child(3) input.cell-input.money');
  let sailingAmount = Number(safeStorageGet(STORAGE_KEYS.lightDuesAmountSailing));
  if (sailingCalc) {
    sailingAmount = sailingCalc.amount;
    safeStorageSet(STORAGE_KEYS.lightDuesTariffSailing, sailingCalc.tariffLabel);
    safeStorageSet(STORAGE_KEYS.lightDuesAmountSailing, sailingAmount.toFixed(2));
  } else if (Number.isFinite(gt) && gt > 0) {
    if (sailingState && sailingCalc === null) {
      // keep existing value if sailing state is invalid.
    } else if (pdaCalc) {
      sailingAmount = pdaCalc.amount;
      safeStorageSet(STORAGE_KEYS.lightDuesTariffSailing, pdaCalc.tariffLabel);
      safeStorageSet(STORAGE_KEYS.lightDuesAmountSailing, sailingAmount.toFixed(2));
    } else if (descCoefficient) {
      sailingAmount = gt * descCoefficient.value;
      safeStorageSet(STORAGE_KEYS.lightDuesAmountSailing, sailingAmount.toFixed(2));
    }
  }

  if (sailingInput && Number.isFinite(sailingAmount)) {
    const formatted = formatMoneyValue(sailingAmount);
    if (sailingInput.value !== formatted) {
      sailingInput.value = formatted;
      changed = true;
    }
  }

  if (changed) {
    recalcOutlayTotals();
    saveIndexState();
  }
}

function findPortDuesRow() {
  const byDataRow = document.querySelector('tr[data-row="port-dues"]');
  if (byDataRow) return byDataRow;
  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return null;
  return Array.from(tbody.querySelectorAll('tr')).find((row) => {
    const descField = row.querySelector('td.desc textarea, td.desc input');
    const desc = String(descField?.value || '').trim().toUpperCase();
    return desc.startsWith('PORT DUES');
  }) || null;
}

function findBunkerRow() {
  return document.querySelector('tr[data-row="bunker"]');
}

function createBunkerRow() {
  const row = document.createElement('tr');
  row.dataset.row = 'bunker';
  row.innerHTML = `
    <td class="desc"><textarea class="cell-input text" rows="1">${BUNKER_ROW_DESCRIPTION}</textarea></td>
    <td><input class="cell-input money" value="" /></td>
    <td><input class="cell-input money" value="" /></td>
  `;
  return row;
}

function ensureBunkerRowAfterPortDues() {
  const portDuesRow = findPortDuesRow();
  if (!portDuesRow || !portDuesRow.parentElement) return null;
  const tbody = portDuesRow.parentElement;

  let bunkerRow = findBunkerRow();
  if (!bunkerRow) {
    bunkerRow = createBunkerRow();
  }

  const targetNext = portDuesRow.nextElementSibling;
  if (targetNext !== bunkerRow) {
    tbody.insertBefore(bunkerRow, targetNext);
  }

  decorateOutlayRows();
  wrapMoneyFields();
  decorateMoneyEditCells();
  const descInput = bunkerRow.querySelector('textarea.cell-input.text');
  if (descInput) autoResizeTextarea(descInput);
  return bunkerRow;
}

function updateBunkerRowFromStorage() {
  const pdaBunkeringRaw = Number(safeStorageGet(STORAGE_KEYS.portDuesBunkeringAmountPda));
  const sailingBunkeringRaw = Number(safeStorageGet(STORAGE_KEYS.portDuesBunkeringAmountSailing));

  const pdaBunkering = Number.isFinite(pdaBunkeringRaw) ? pdaBunkeringRaw : 0;
  const sailingBunkering = Number.isFinite(sailingBunkeringRaw) ? sailingBunkeringRaw : 0;
  const hasBunkering = pdaBunkering > 0 || sailingBunkering > 0;

  const existingBunkerRow = findBunkerRow();
  if (!hasBunkering) {
    if (existingBunkerRow) {
      existingBunkerRow.remove();
      recalcOutlayTotals();
      saveIndexState();
    }
    return;
  }

  const bunkerRow = ensureBunkerRowAfterPortDues();
  if (!bunkerRow) return;

  let changed = false;
  const descInput = bunkerRow.querySelector('td.desc textarea, td.desc input');
  if (descInput && descInput.value !== BUNKER_ROW_DESCRIPTION) {
    descInput.value = BUNKER_ROW_DESCRIPTION;
    if (descInput.tagName === 'TEXTAREA') autoResizeTextarea(descInput);
    changed = true;
  }

  const pdaInput = bunkerRow.querySelector('td:nth-child(2) input.cell-input.money');
  if (pdaInput) {
    const formatted = formatMoneyValue(pdaBunkering);
    if (pdaInput.value !== formatted) {
      pdaInput.value = formatted;
      if (isPdaRoundingEnabled()) pdaInput.dataset.rawValue = formatted;
      else delete pdaInput.dataset.rawValue;
      changed = true;
    }
  }

  const sailingInput = bunkerRow.querySelector('td:nth-child(3) input.cell-input.money');
  if (sailingInput) {
    const formatted = formatMoneyValue(sailingBunkering);
    if (sailingInput.value !== formatted) {
      sailingInput.value = formatted;
      changed = true;
    }
  }

  if (changed) {
    recalcOutlayTotals();
    saveIndexState();
  }
}

function getPortDuesCalculationFromState(state, quantityValue) {
  if (!state || typeof state !== 'object') return null;
  const cargoType = getPortDuesCargoTypeConfig(state.cargoType);
  const cargoQuantityFromState = parseMoneyValue(state.cargoQuantity);
  const cargoQuantity = quantityValue > 0 ? quantityValue : cargoQuantityFromState;
  const cargoAmount = cargoQuantity > 0 ? cargoQuantity * cargoType.rate : 0;

  const bunkeringEnabled = Boolean(state.bunkeringEnabled);
  const bunkeringQuantity = parseMoneyValue(state.bunkeringQuantity);
  const bunkeringAmount = bunkeringEnabled && bunkeringQuantity > 0 ? bunkeringQuantity * 0.3 : 0;

  return {
    coefficientLabel: cargoType.label,
    cargoAmount,
    bunkeringAmount,
    totalAmount: cargoAmount + bunkeringAmount
  };
}

function updatePortDuesFromStorage() {
  const portDuesRow = findPortDuesRow();
  if (!portDuesRow) return;

  let changed = false;
  const quantity = getCurrentQuantityValue();
  const pdaState = parseStoredJson(safeStorageGet(STORAGE_KEYS.portDuesState));
  const sailingState = parseStoredJson(safeStorageGet(STORAGE_KEYS.portDuesStateSailing));
  const pdaCalc = getPortDuesCalculationFromState(pdaState, quantity);
  const sailingCalc = getPortDuesCalculationFromState(sailingState, quantity);

  const descInput = portDuesRow.querySelector('td.desc textarea, td.desc input');
  const descCoefficient = getCoefficientFromDescription(descInput ? descInput.value : '');
  if (descInput) {
    let coefficientLabel = '';
    if (pdaCalc) {
      coefficientLabel = pdaCalc.coefficientLabel;
    } else if (pdaState && typeof pdaState.cargoType === 'string') {
      coefficientLabel = getPortDuesCargoTypeConfig(pdaState.cargoType).label;
    } else if (descCoefficient) {
      coefficientLabel = descCoefficient.label;
    }
    if (coefficientLabel) {
      const nextDesc = `PORT DUES (EUR ${coefficientLabel} x cargo / MT)`;
      if (descInput.value !== nextDesc) {
        descInput.value = nextDesc;
        if (descInput.tagName === 'TEXTAREA') autoResizeTextarea(descInput);
        changed = true;
      }
    }
  }

  const pdaInput = portDuesRow.querySelector('td:nth-child(2) input.cell-input.money');
  let pdaAmount = Number(safeStorageGet(STORAGE_KEYS.portDuesCargoAmountPda) ?? safeStorageGet(STORAGE_KEYS.portDuesAmountPda));
  if (pdaCalc) {
    pdaAmount = pdaCalc.totalAmount;
    safeStorageSet(STORAGE_KEYS.portDuesAmountPda, pdaCalc.totalAmount.toFixed(2));
    safeStorageSet(STORAGE_KEYS.portDuesCargoAmountPda, pdaCalc.cargoAmount.toFixed(2));
    if (pdaCalc.bunkeringAmount > 0) {
      safeStorageSet(STORAGE_KEYS.portDuesBunkeringAmountPda, pdaCalc.bunkeringAmount.toFixed(2));
    } else {
      safeStorageRemove(STORAGE_KEYS.portDuesBunkeringAmountPda);
    }
  } else if (quantity > 0 && descCoefficient) {
    pdaAmount = quantity * descCoefficient.value;
    safeStorageSet(STORAGE_KEYS.portDuesAmountPda, pdaAmount.toFixed(2));
    safeStorageSet(STORAGE_KEYS.portDuesCargoAmountPda, pdaAmount.toFixed(2));
  }

  if (pdaInput && Number.isFinite(pdaAmount)) {
    const formatted = formatMoneyValue(pdaAmount);
    if (pdaInput.value !== formatted) {
      pdaInput.value = formatted;
      if (isPdaRoundingEnabled()) pdaInput.dataset.rawValue = formatted;
      else delete pdaInput.dataset.rawValue;
      changed = true;
    }
  }

  const sailingInput = portDuesRow.querySelector('td:nth-child(3) input.cell-input.money');
  let sailingAmount = Number(safeStorageGet(STORAGE_KEYS.portDuesCargoAmountSailing) ?? safeStorageGet(STORAGE_KEYS.portDuesAmountSailing));
  if (sailingCalc) {
    sailingAmount = sailingCalc.totalAmount;
    safeStorageSet(STORAGE_KEYS.portDuesAmountSailing, sailingCalc.totalAmount.toFixed(2));
    safeStorageSet(STORAGE_KEYS.portDuesCargoAmountSailing, sailingCalc.cargoAmount.toFixed(2));
    if (sailingCalc.bunkeringAmount > 0) {
      safeStorageSet(STORAGE_KEYS.portDuesBunkeringAmountSailing, sailingCalc.bunkeringAmount.toFixed(2));
    } else {
      safeStorageRemove(STORAGE_KEYS.portDuesBunkeringAmountSailing);
    }
  } else if (quantity > 0) {
    if (sailingState && sailingCalc === null) {
      // keep existing value if sailing state exists but cannot be calculated.
    } else if (pdaCalc) {
      sailingAmount = pdaCalc.totalAmount;
      safeStorageSet(STORAGE_KEYS.portDuesAmountSailing, sailingAmount.toFixed(2));
      safeStorageSet(STORAGE_KEYS.portDuesCargoAmountSailing, sailingAmount.toFixed(2));
    } else if (descCoefficient) {
      sailingAmount = quantity * descCoefficient.value;
      safeStorageSet(STORAGE_KEYS.portDuesAmountSailing, sailingAmount.toFixed(2));
      safeStorageSet(STORAGE_KEYS.portDuesCargoAmountSailing, sailingAmount.toFixed(2));
    }
  }

  if (sailingInput && Number.isFinite(sailingAmount)) {
    const formatted = formatMoneyValue(sailingAmount);
    if (sailingInput.value !== formatted) {
      sailingInput.value = formatted;
      changed = true;
    }
  }

  if (changed) {
    recalcOutlayTotals();
    saveIndexState();
  }
  updateBunkerRowFromStorage();
}

function findMooringRow() {
  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return null;
  return Array.from(tbody.querySelectorAll('tr')).find((row) => {
    const descField = row.querySelector('td.desc textarea, td.desc input');
    const desc = String(descField?.value || '').trim().toUpperCase();
    return desc.startsWith('MOORING/UNMOORING');
  }) || null;
}

function updateMooringFromStorage() {
  const mooringRow = findMooringRow();
  if (!mooringRow) return;

  let changed = false;
  const gt = getCurrentGtValue();
  const pdaState = parseStoredJson(safeStorageGet(STORAGE_KEYS.mooringState));
  const sailingState = parseStoredJson(safeStorageGet(STORAGE_KEYS.mooringStateSailing));
  const pdaCalc = getMooringCalculationFromState(pdaState, gt);
  const sailingCalc = getMooringCalculationFromState(sailingState, gt);

  const pdaInput = mooringRow.querySelector('td:nth-child(2) input.cell-input.money');
  let pdaAmount = Number(safeStorageGet(STORAGE_KEYS.mooringAmountPda));
  if (pdaCalc) {
    pdaAmount = pdaCalc.totalAmount;
    safeStorageSet(STORAGE_KEYS.mooringAmountPda, pdaAmount.toFixed(2));
  }
  if (pdaInput && Number.isFinite(pdaAmount)) {
    const formatted = formatMoneyValue(pdaAmount);
    if (pdaInput.value !== formatted) {
      pdaInput.value = formatted;
      if (isPdaRoundingEnabled()) pdaInput.dataset.rawValue = formatted;
      else delete pdaInput.dataset.rawValue;
      changed = true;
    }
  }

  const sailingInput = mooringRow.querySelector('td:nth-child(3) input.cell-input.money');
  let sailingAmount = Number(safeStorageGet(STORAGE_KEYS.mooringAmountSailing));
  if (sailingCalc) {
    sailingAmount = sailingCalc.totalAmount;
    safeStorageSet(STORAGE_KEYS.mooringAmountSailing, sailingAmount.toFixed(2));
  } else if ((!Number.isFinite(sailingAmount) || sailingAmount <= 0) && pdaCalc) {
    sailingAmount = pdaCalc.totalAmount;
    safeStorageSet(STORAGE_KEYS.mooringAmountSailing, sailingAmount.toFixed(2));
  }
  if (sailingInput && Number.isFinite(sailingAmount)) {
    const formatted = formatMoneyValue(sailingAmount);
    if (sailingInput.value !== formatted) {
      sailingInput.value = formatted;
      changed = true;
    }
  }

  if (changed) {
    recalcOutlayTotals();
    saveIndexState();
  }
}

function updateIndexGtFromStorage() {
  const gtInputIndex = document.getElementById('grossTonnage');
  if (!gtInputIndex) return;
  const storedGt = safeStorageGet(STORAGE_KEYS.gt);
  if (!storedGt) return;
  if (gtInputIndex.value !== storedGt) {
    gtInputIndex.value = storedGt;
  }
}

function updateIndexQuantityFromStorage() {
  const quantityInput = document.getElementById('quantityInput');
  if (!quantityInput) return;
  const storedQuantity = safeStorageGet(STORAGE_KEYS.quantity);
  if (storedQuantity === null) return;
  if (quantityInput.value !== storedQuantity) {
    quantityInput.value = storedQuantity;
  }
}

function updateVesselNameFromStorage(targetInput) {
  if (!targetInput) return;
  const storedName = safeStorageGet(STORAGE_KEYS.vesselName);
  if (!storedName) return;
  if (targetInput.value !== storedName) {
    targetInput.value = storedName;
  }
}

function refreshOutlayLayout() {
  document.querySelectorAll('#outlaysBody textarea.cell-input.text').forEach(autoResizeTextarea);
}

function parseMoneyValue(raw) {
  if (!raw) return 0;
  const cleaned = String(raw).trim().replace(/[^0-9,.-]/g, '');
  if (!cleaned) return 0;

  const hasComma = cleaned.includes(',');
  const hasDot = cleaned.includes('.');

  let normalized = cleaned;
  if (hasComma && hasDot) {
    const lastComma = cleaned.lastIndexOf(',');
    const lastDot = cleaned.lastIndexOf('.');
    if (lastComma > lastDot) {
      normalized = cleaned.replace(/\./g, '').replace(/,/g, '.');
    } else {
      normalized = cleaned.replace(/,/g, '');
    }
  } else if (hasComma && !hasDot) {
    if (/^-?\d{1,3}(,\d{3})+$/.test(cleaned)) {
      normalized = cleaned.replace(/,/g, '');
    } else {
      normalized = cleaned.replace(/,/g, '.');
    }
  } else if (hasDot && !hasComma) {
    if (/^-?\d{1,3}(\.\d{3})+$/.test(cleaned)) {
      normalized = cleaned.replace(/\./g, '');
    }
  }

  const value = Number(normalized);
  return Number.isFinite(value) ? value : 0;
}

function parsePercentValue(raw) {
  return parseMoneyValue(raw);
}

function formatMoneyValue(value) {
  const fixed = Number.isFinite(value) ? value.toFixed(2) : '0.00';
  const [whole, decimal] = fixed.split('.');
  const withThousands = whole.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return `${withThousands},${decimal}`;
}

function getCurrencySymbol(code) {
  const map = {
    EUR: '€',
    USD: '$',
    GBP: '£',
    HRK: 'kn'
  };
  return map[code] || code;
}

function updateCurrencySymbol(code) {
  const symbol = getCurrencySymbol(code);
  document.querySelectorAll('[data-currency-symbol]').forEach((el) => {
    el.textContent = symbol;
  });
}

function wrapMoneyFields() {
  document.querySelectorAll('input.cell-input.money').forEach((input) => {
    if (input.closest('.money-field')) return;
    const wrapper = document.createElement('div');
    wrapper.className = 'money-field';
    const symbol = document.createElement('span');
    symbol.className = 'currency-symbol';
    symbol.setAttribute('data-currency-symbol', '');
    symbol.textContent = getCurrencySymbol(outlaysCurrency?.value || 'EUR');
    const parent = input.parentNode;
    parent.insertBefore(wrapper, input);
    wrapper.appendChild(symbol);
    wrapper.appendChild(input);
  });
}

function isPdaRoundingEnabled() {
  return Boolean(roundPdaPrices && roundPdaPrices.checked);
}

function isPdaMoneyInput(input) {
  if (!input) return false;
  const cell = input.closest('td');
  return Boolean(cell && cell.cellIndex === 1);
}

function getRawOrCurrentPdaValue(input) {
  if (!input) return 0;
  const source = input.dataset.rawValue !== undefined ? input.dataset.rawValue : input.value;
  return parseMoneyValue(source);
}

function focusMoneyFromEdit(button) {
  if (!button) return;
  const cell = button.closest('td');
  if (!cell) return;
  const input = cell.querySelector('input.cell-input.money');
  if (!input) return;
  input.focus();
  if (typeof input.select === 'function') input.select();
}

function getOutlayDescription(row) {
  const descField = row.querySelector('td.desc textarea, td.desc input');
  return String(descField?.value || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

function shouldShowInlineMoneyEdit(row) {
  const desc = getOutlayDescription(row);
  const editablePrefixes = [
    'LIGHT DUES',
    'PORT DUES',
    'PILOTAGE',
    'PILOT BOAT',
    'MOORING/UNMOORING'
  ];
  return editablePrefixes.some((prefix) => desc.startsWith(prefix));
}

function removeInlineMoneyEdit(cell) {
  const wrapper = cell.querySelector('.money-edit');
  if (!wrapper) return;
  const moneyField = wrapper.querySelector('.money-field');
  if (!moneyField) return;
  wrapper.replaceWith(moneyField);
}

function getInlineMoneyEditConfig(row, columnIndex) {
  const desc = getOutlayDescription(row);
  if (desc.startsWith('LIGHT DUES')) {
    if (columnIndex === 2) {
      return {
        onClick: 'openLightDuesPda()',
        ariaLabel: 'Edit light dues PDA'
      };
    }
    return {
      onClick: 'openLightDuesSailingPda()',
      ariaLabel: 'Edit light dues sailing PDA'
    };
  }
  if (desc.startsWith('PORT DUES')) {
    if (columnIndex === 2) {
      return {
        onClick: 'openPortDuesPda()',
        ariaLabel: 'Edit port dues PDA'
      };
    }
    return {
      onClick: 'openPortDuesSailingPda()',
      ariaLabel: 'Edit port dues sailing PDA'
    };
  }
  if (desc.startsWith('MOORING/UNMOORING')) {
    if (columnIndex === 2) {
      return {
        onClick: 'openMooringPda()',
        ariaLabel: 'Edit mooring PDA'
      };
    }
    return {
      onClick: 'openMooringSailingPda()',
      ariaLabel: 'Edit mooring sailing PDA'
    };
  }
  return {
    onClick: 'focusMoneyFromEdit(this)',
    ariaLabel: columnIndex === 2 ? 'Edit PDA amount' : 'Edit Sailing PDA amount'
  };
}

function decorateMoneyEditCells() {
  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return;

  tbody.querySelectorAll('tr').forEach((row) => {
    if (row.dataset.row === 'towage' || row.dataset.row === 'bank-charges') return;
    const shouldShowEdit = shouldShowInlineMoneyEdit(row);

    [2, 3].forEach((columnIndex) => {
      const cell = row.querySelector(`td:nth-child(${columnIndex})`);
      if (!cell) return;

      if (!shouldShowEdit) {
        removeInlineMoneyEdit(cell);
        return;
      }

      let button = cell.querySelector('.row-edit');
      if (button) {
        const config = getInlineMoneyEditConfig(row, columnIndex);
        button.setAttribute('aria-label', config.ariaLabel);
        button.setAttribute('onclick', config.onClick);
        return;
      }

      const moneyField = cell.querySelector('.money-field');
      if (!moneyField) return;

      const wrapper = document.createElement('div');
      wrapper.className = 'money-edit';

      button = document.createElement('button');
      button.type = 'button';
      button.className = 'row-edit money-edit-btn';
      const config = getInlineMoneyEditConfig(row, columnIndex);
      button.setAttribute('aria-label', config.ariaLabel);
      button.setAttribute('onclick', config.onClick);

      const icon = document.createElement('img');
      icon.src = 'assets/icons/edit_48x48.png';
      icon.alt = '';
      button.appendChild(icon);

      moneyField.remove();
      wrapper.appendChild(button);
      wrapper.appendChild(moneyField);
      cell.appendChild(wrapper);
    });
  });
}

function recalcOutlayTotals() {
  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return;
  const roundPda = isPdaRoundingEnabled();

  let totalPda = 0;
  let totalSailing = 0;
  let runningPda = 0;
  let runningSailing = 0;
  const bankRow = tbody.querySelector('tr[data-row="bank-charges"]');
  const bankRateInput = document.getElementById('bankRate');
  const bankRate = parsePercentValue(bankRateInput?.value) / 100;

  tbody.querySelectorAll('tr').forEach((row) => {
    if (row === bankRow) {
      const bankPdaRaw = runningPda * bankRate;
      const bankPda = roundPda ? Math.ceil(bankPdaRaw) : bankPdaRaw;
      const bankSailing = runningSailing * bankRate;
      const bankPdaInput = document.getElementById('bankPda');
      const bankSailingInput = document.getElementById('bankSailing');
      if (bankPdaInput) bankPdaInput.value = formatMoneyValue(bankPda);
      if (bankSailingInput) bankSailingInput.value = formatMoneyValue(bankSailing);
      totalPda += bankPda;
      totalSailing += bankSailing;
      return;
    }

    const inputs = row.querySelectorAll('input.cell-input.money');
    const pdaInput = inputs.length >= 1 ? inputs[0] : null;
    let pdaValue = pdaInput ? getRawOrCurrentPdaValue(pdaInput) : 0;
    if (roundPda) {
      if (pdaInput && !pdaInput.readOnly && pdaInput.dataset.rawValue === undefined) {
        pdaInput.dataset.rawValue = pdaInput.value;
      }
      pdaValue = Math.ceil(pdaValue);
      if (pdaInput && !pdaInput.readOnly) {
        pdaInput.value = formatMoneyValue(pdaValue);
      }
    } else if (pdaInput && pdaInput.dataset.rawValue !== undefined && !pdaInput.readOnly) {
      pdaInput.value = pdaInput.dataset.rawValue;
      delete pdaInput.dataset.rawValue;
      pdaValue = parseMoneyValue(pdaInput.value);
    }
    const sailingValue = inputs.length >= 2 ? parseMoneyValue(inputs[1].value) : 0;
    runningPda += pdaValue;
    runningSailing += sailingValue;
    totalPda += pdaValue;
    totalSailing += sailingValue;
  });

  const totalPdaInput = document.getElementById('totalPda');
  const totalSailingInput = document.getElementById('totalSailing');
  if (totalPdaInput) totalPdaInput.value = formatMoneyValue(totalPda);
  if (totalSailingInput) totalSailingInput.value = formatMoneyValue(totalSailing);
}

function removeOutlayRow(row) {
  const tbody = document.getElementById('outlaysBody');
  if (!tbody || !row) return;

  const descField = row.querySelector('td.desc textarea, td.desc input');
  const descText = String(descField?.value || '').trim();
  const label = descText ? `\n\nRow: ${descText}` : '';
  const confirmed = window.confirm(`Delete this row?${label}`);
  if (!confirmed) return;

  const rows = tbody.querySelectorAll('tr');
  if (rows.length <= 1) {
    clearOutlayRow(row);
    return;
  }

  row.remove();
  recalcOutlayTotals();
  saveIndexState();
}

function setSailingVisible(visible) {
  if (!outlaysTable) return;
  outlaysTable.classList.toggle('hide-sailing', !visible);
  requestAnimationFrame(() => {
    void outlaysTable.offsetWidth;
    requestAnimationFrame(refreshOutlayLayout);
  });
}

function autoSizeInput(input) {
  if (!input) return;
  const value = input.value || input.placeholder || '';
  const length = Math.max(value.length, 4);
  input.style.width = `${length + 1}ch`;
}

function setDensity(mode) {
  document.body.classList.toggle('density-comfortable', mode === 'comfortable');
  document.body.classList.toggle('density-dense', mode === 'dense');
  if (densityComfortable) densityComfortable.classList.toggle('active', mode === 'comfortable');
  if (densityDense) densityDense.classList.toggle('active', mode === 'dense');
}

function getDensityMode() {
  if (document.body.classList.contains('density-dense')) return 'dense';
  if (document.body.classList.contains('density-comfortable')) return 'comfortable';
  return 'none';
}

function applyPrintDensity() {
  const current = getDensityMode();
  if (current !== 'dense') {
    printRestoreDensity = current;
    setDensity('dense');
  }
}

function printFit() {
  applyPrintDensity();
  document.body.classList.add('print-fit');
  updatePrintHidden();
  setTimeout(() => {
    window.print();
  }, 50);
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function exportToExcel() {
  const outlaysBody = document.getElementById('outlaysBody');
  if (!outlaysBody) return;

  const dateValue = document.getElementById('dateInput')?.value || '';
  const noteValue = document.getElementById('titleNote')?.value || '';
  const currencyValue = outlaysCurrency ? outlaysCurrency.value : '';
  const includeSailing = outlaysTable ? !outlaysTable.classList.contains('hide-sailing') : true;

  const headerRows = [];
  headerRows.push(['Title', 'PRO-FORMA D/A']);
  if (noteValue.trim()) headerRows.push(['Note', noteValue]);
  headerRows.push(['Split', dateValue]);
  if (currencyValue) headerRows.push(['Outlays Currency', currencyValue]);

  const detailRows = [];
  document.querySelectorAll('.card-row .card .row > div').forEach((group) => {
    const label = group.querySelector('label')?.textContent?.trim() || '';
    const value = group.querySelector('input')?.value || '';
    if (label) detailRows.push([label, value]);
  });

  const outlayRows = [];
  outlaysBody.querySelectorAll('tr').forEach((row) => {
    const descInput = row.querySelector('textarea, input');
    let descValue = descInput ? descInput.value : '';
    if (row.dataset.row === 'bank-charges') {
      const rate = document.getElementById('bankRate')?.value || '';
      if (rate) descValue = `${descValue} (${rate}%)`;
    }
    const pdaValue = row.querySelector('td:nth-child(2) input')?.value || '';
    const sailingValue = row.querySelector('td:nth-child(3) input')?.value || '';
    outlayRows.push([descValue, pdaValue, sailingValue]);
  });

  const totalPda = document.getElementById('totalPda')?.value || '';
  const totalSailing = document.getElementById('totalSailing')?.value || '';

  const columnCount = includeSailing ? 3 : 2;
  const outlayHeaderCells = includeSailing
    ? `<th>Outlays</th><th>PDA (${escapeHtml(currencyValue)})</th><th>Sailing PDA (${escapeHtml(currencyValue)})</th>`
    : `<th>Outlays</th><th>PDA (${escapeHtml(currencyValue)})</th>`;

  const outlayBodyRows = outlayRows
    .map((row) => {
      const cells = includeSailing
        ? `<td>${escapeHtml(row[0])}</td><td>${escapeHtml(row[1])}</td><td>${escapeHtml(row[2])}</td>`
        : `<td>${escapeHtml(row[0])}</td><td>${escapeHtml(row[1])}</td>`;
      return `<tr>${cells}</tr>`;
    })
    .join('');

  const totalCells = includeSailing
    ? `<td>Total (${escapeHtml(currencyValue)})</td><td>${escapeHtml(totalPda)}</td><td>${escapeHtml(totalSailing)}</td>`
    : `<td>Total (${escapeHtml(currencyValue)})</td><td>${escapeHtml(totalPda)}</td>`;

  const headerTable = `
    <table>
      ${headerRows
        .map((row) => `<tr><th>${escapeHtml(row[0])}</th><td>${escapeHtml(row[1])}</td></tr>`)
        .join('')}
    </table>
  `;

  const detailTable = `
    <table>
      <tr><th colspan="2">Details</th></tr>
      ${detailRows
        .map((row) => `<tr><th>${escapeHtml(row[0])}</th><td>${escapeHtml(row[1])}</td></tr>`)
        .join('')}
    </table>
  `;

  const outlayTable = `
    <table>
      <tr><th colspan="${columnCount}">Outlays &amp; Charges Expressed In ${escapeHtml(currencyValue)}</th></tr>
      <tr>${outlayHeaderCells}</tr>
      ${outlayBodyRows}
      <tr>${totalCells}</tr>
    </table>
  `;

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        table { border-collapse: collapse; margin-bottom: 12px; }
        th, td { border: 1px solid #9ca3af; padding: 4px 6px; text-align: left; }
        th { background: #e5e7eb; font-weight: 700; }
      </style>
    </head>
    <body>
      ${headerTable}
      ${detailTable}
      ${outlayTable}
    </body>
    </html>
  `;

  const blob = new Blob([html], { type: 'application/vnd.ms-excel' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  const safeDate = dateValue.replace(/[^\d-]/g, '') || 'pro-forma';
  link.href = url;
  link.download = `pro-forma-da-${safeDate}.xls`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function openTugsCalculator() {
  saveIndexState();
  window.location.href = 'tugs-pda.html';
}

function openTugsCalculatorSailing() {
  saveIndexState();
  window.location.href = 'tugs-sailing-pda.html';
}

function openLightDuesPda() {
  saveIndexState();
  const gtInputIndex = document.getElementById('grossTonnage');
  if (gtInputIndex) {
    const gtValue = gtInputIndex.value.trim();
    if (gtValue) safeStorageSet(STORAGE_KEYS.gt, gtValue);
    else safeStorageRemove(STORAGE_KEYS.gt);
  }
  window.location.href = 'light-dues-pda.html';
}

function openLightDuesSailingPda() {
  saveIndexState();
  const gtInputIndex = document.getElementById('grossTonnage');
  if (gtInputIndex) {
    const gtValue = gtInputIndex.value.trim();
    if (gtValue) safeStorageSet(STORAGE_KEYS.gt, gtValue);
    else safeStorageRemove(STORAGE_KEYS.gt);
  }
  window.location.href = 'light-dues-sailing-pda.html';
}

function openPortDuesPda() {
  saveIndexState();
  const quantityInput = document.getElementById('quantityInput');
  if (quantityInput) {
    const quantityValue = quantityInput.value.trim();
    if (quantityValue) safeStorageSet(STORAGE_KEYS.quantity, quantityValue);
    else safeStorageRemove(STORAGE_KEYS.quantity);
  }
  window.location.href = 'port-dues-pda.html';
}

function openPortDuesSailingPda() {
  saveIndexState();
  const quantityInput = document.getElementById('quantityInput');
  if (quantityInput) {
    const quantityValue = quantityInput.value.trim();
    if (quantityValue) safeStorageSet(STORAGE_KEYS.quantity, quantityValue);
    else safeStorageRemove(STORAGE_KEYS.quantity);
  }
  window.location.href = 'port-dues-sailing-pda.html';
}

function openMooringPda() {
  saveIndexState();
  const gtInputIndex = document.getElementById('grossTonnage');
  if (gtInputIndex) {
    const gtValue = gtInputIndex.value.trim();
    if (gtValue) safeStorageSet(STORAGE_KEYS.gt, gtValue);
    else safeStorageRemove(STORAGE_KEYS.gt);
  }
  window.location.href = 'mooring-pda.html';
}

function openMooringSailingPda() {
  saveIndexState();
  const gtInputIndex = document.getElementById('grossTonnage');
  if (gtInputIndex) {
    const gtValue = gtInputIndex.value.trim();
    if (gtValue) safeStorageSet(STORAGE_KEYS.gt, gtValue);
    else safeStorageRemove(STORAGE_KEYS.gt);
  }
  window.location.href = 'mooring-sailing-pda.html';
}

function goHome() {
  saveTugsState();
  window.location.href = 'index.html';
}

function updatePrintHidden() {
  document.querySelectorAll('[data-print-hide-empty]').forEach((input) => {
    const empty = !(input.value || '').trim();
    input.classList.toggle('print-hide', empty);
    const group = input.closest('[data-print-hide-group]');
    if (group) group.classList.toggle('print-hide', empty);
  });
}

function clearPrintHidden() {
  document.querySelectorAll('.print-hide').forEach((el) => {
    el.classList.remove('print-hide');
  });
}

function getDragAfterElement(container, y) {
  const rows = [...container.querySelectorAll('tr:not(.dragging)')];
  let closest = { offset: Number.NEGATIVE_INFINITY, element: null };
  rows.forEach((row) => {
    const box = row.getBoundingClientRect();
    const offset = y - box.top - box.height / 2;
    if (offset < 0 && offset > closest.offset) {
      closest = { offset, element: row };
    }
  });
  return closest.element;
}

function initIndex() {
  const outlaysBody = document.getElementById('outlaysBody');
  if (!outlaysBody) return;

  document.body.classList.add('density-comfortable');
  restoreIndexState();
  decorateOutlayRows();
  recalcOutlayTotals();
  refreshOutlayLayout();

  outlaysBody.addEventListener('click', (event) => {
    const button = event.target.closest('.row-remove');
    if (!button) return;
    const row = button.closest('tr');
    removeOutlayRow(row);
  });

  outlaysBody.addEventListener('dragstart', (event) => {
    const handle = event.target.closest('.row-handle');
    if (!handle) return;
    const row = handle.closest('tr');
    if (!row) return;
    draggingRow = row;
    row.classList.add('dragging');
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', '');
  });

  outlaysBody.addEventListener('dragend', () => {
    if (draggingRow) draggingRow.classList.remove('dragging');
    draggingRow = null;
    saveIndexState();
  });

  outlaysBody.addEventListener('dragover', (event) => {
    if (!draggingRow) return;
    event.preventDefault();
    const afterElement = getDragAfterElement(outlaysBody, event.clientY);
    const bankRow = outlaysBody.querySelector('tr[data-row="bank-charges"]');
    if (!afterElement) {
      if (bankRow) {
        outlaysBody.insertBefore(draggingRow, bankRow);
      } else {
        outlaysBody.appendChild(draggingRow);
      }
    } else if (afterElement !== draggingRow) {
      if (bankRow && afterElement === bankRow) {
        outlaysBody.insertBefore(draggingRow, bankRow);
      } else {
        outlaysBody.insertBefore(draggingRow, afterElement);
      }
    }
  });

  outlaysBody.addEventListener('input', (event) => {
    const target = event.target;
    if (!target) return;
    if (target.tagName === 'TEXTAREA') {
      autoResizeTextarea(target);
      if (target.classList.contains('text')) {
        decorateMoneyEditCells();
      }
    }
    if (target.classList.contains('money') && isPdaRoundingEnabled() && isPdaMoneyInput(target) && !target.readOnly) {
      target.dataset.rawValue = target.value;
    }
    if (target.classList.contains('money') || target.classList.contains('percent-input')) {
      recalcOutlayTotals();
    }
    saveIndexState();
  });

  outlaysTable = document.querySelector('.outlays-table');
  toggleSailing = document.getElementById('toggleSailing');
  outlaysCurrency = document.getElementById('outlaysCurrency');
  roundPdaPrices = document.getElementById('roundPdaPrices');
  if (toggleSailing) {
    toggleSailing.addEventListener('change', () => {
      setSailingVisible(toggleSailing.checked);
      saveIndexState();
    });
    setSailingVisible(toggleSailing.checked);
  }
  if (roundPdaPrices) {
    roundPdaPrices.addEventListener('change', () => {
      recalcOutlayTotals();
      saveIndexState();
    });
  }
  wrapMoneyFields();
  decorateMoneyEditCells();
  recalcOutlayTotals();
  if (outlaysCurrency) {
    updateCurrencySymbol(outlaysCurrency.value);
    outlaysCurrency.addEventListener('change', () => {
      updateCurrencySymbol(outlaysCurrency.value);
      saveIndexState();
    });
  }

  const vesselNameIndex = document.getElementById('vesselNameIndex');
  if (vesselNameIndex) {
    updateVesselNameFromStorage(vesselNameIndex);
    vesselNameIndex.addEventListener('input', () => {
      const value = vesselNameIndex.value.trim();
      if (value) safeStorageSet(STORAGE_KEYS.vesselName, value);
      else safeStorageRemove(STORAGE_KEYS.vesselName);
      saveIndexState();
    });
  }

  const gtInputIndex = document.getElementById('grossTonnage');
  if (gtInputIndex) {
    updateIndexGtFromStorage();
    gtInputIndex.addEventListener('input', () => {
      const value = gtInputIndex.value.trim();
      if (value) safeStorageSet(STORAGE_KEYS.gt, value);
      else safeStorageRemove(STORAGE_KEYS.gt);
      updateLightDuesFromStorage();
      updateMooringFromStorage();
      saveIndexState();
    });
  }

  const quantityInputIndex = document.getElementById('quantityInput');
  if (quantityInputIndex) {
    const storedQuantity = safeStorageGet(STORAGE_KEYS.quantity);
    if (storedQuantity === null) {
      const currentQuantity = quantityInputIndex.value.trim();
      if (currentQuantity) safeStorageSet(STORAGE_KEYS.quantity, currentQuantity);
    }
    updateIndexQuantityFromStorage();
    quantityInputIndex.addEventListener('input', () => {
      const value = quantityInputIndex.value.trim();
      if (value) safeStorageSet(STORAGE_KEYS.quantity, value);
      else safeStorageRemove(STORAGE_KEYS.quantity);
      updatePortDuesFromStorage();
      saveIndexState();
    });
  }

  const towageDesc = document.querySelector('tr[data-row="towage"] textarea.cell-input.text');
  if (towageDesc) {
    towageDesc.dataset.baseText = 'TOWAGE';
    towageDesc.addEventListener('input', () => {
      towageDesc.dataset.baseText = 'TOWAGE';
    });
  }

  const titleNote = document.getElementById('titleNote');
  if (titleNote) {
    autoSizeInput(titleNote);
    titleNote.addEventListener('input', () => {
      autoSizeInput(titleNote);
      saveIndexState();
    });
  }

  const logoLeftNote = document.getElementById('logoLeftNote');
  if (logoLeftNote) {
    autoResizeTextarea(logoLeftNote);
    logoLeftNote.addEventListener('input', () => {
      autoResizeTextarea(logoLeftNote);
      saveIndexState();
    });
  }

  const dateInput = document.getElementById('dateInput');
  if (dateInput && !dateInput.value) {
    const now = new Date();
    const day = String(now.getDate());
    const month = String(now.getMonth() + 1);
    const year = String(now.getFullYear()).slice(-2);
    dateInput.value = `${day}/${month}/${year}`;
  }
  if (dateInput) {
    dateInput.addEventListener('input', saveIndexState);
  }

  densityComfortable = document.getElementById('densityComfortable');
  densityDense = document.getElementById('densityDense');
  if (densityComfortable) {
    densityComfortable.addEventListener('click', () => {
      setDensity('comfortable');
      saveIndexState();
    });
  }
  if (densityDense) {
    densityDense.addEventListener('click', () => {
      setDensity('dense');
      saveIndexState();
    });
  }

  updateTowageFromStorage();
  updateLightDuesFromStorage();
  updatePortDuesFromStorage();
  updateMooringFromStorage();
  updateIndexGtFromStorage();
  updateIndexQuantityFromStorage();
  window.addEventListener('pageshow', () => {
    updateTowageFromStorage();
    updateLightDuesFromStorage();
    updatePortDuesFromStorage();
    updateMooringFromStorage();
    updateIndexQuantityFromStorage();
  });
  window.addEventListener('storage', () => {
    updateTowageFromStorage();
    updateLightDuesFromStorage();
    updatePortDuesFromStorage();
    updateMooringFromStorage();
    updateIndexGtFromStorage();
    updateIndexQuantityFromStorage();
    updateVesselNameFromStorage(vesselNameIndex);
  });

  window.addEventListener('beforeunload', saveIndexState);
  window.addEventListener('pagehide', saveIndexState);

  window.addEventListener('afterprint', clearPrintHidden);
}

function isLightDuesPdaPage() {
  return Boolean(document.body && document.body.classList.contains('page-light-dues-pda'));
}

function isLightDuesSailingPage() {
  return Boolean(document.body && document.body.classList.contains('page-light-dues-sailing'));
}

function getLightDuesStateKey() {
  if (isLightDuesSailingPage()) return STORAGE_KEYS.lightDuesStateSailing;
  if (isLightDuesPdaPage()) return STORAGE_KEYS.lightDuesState;
  return null;
}

function getTierBandFromGt(type, gt) {
  const tierOptions = LIGHT_DUES_TIER_OPTIONS[type];
  if (!Array.isArray(tierOptions) || tierOptions.length === 0) return '';
  if (!Number.isFinite(gt) || gt <= 0) return tierOptions[0].value;

  const matching = tierOptions.find((option) => {
    const minOk = option.min === undefined || gt >= option.min;
    const maxOk = option.max === undefined || gt <= option.max;
    return minOk && maxOk;
  });
  return matching ? matching.value : tierOptions[tierOptions.length - 1].value;
}

function rebuildTierOptions(typeInput, tierBandInput, preferredBand, gt) {
  if (!typeInput || !tierBandInput) return;
  const options = LIGHT_DUES_TIER_OPTIONS[typeInput.value] || [];
  tierBandInput.innerHTML = '';
  options.forEach((option) => {
    const node = document.createElement('option');
    node.value = option.value;
    node.textContent = option.label;
    tierBandInput.appendChild(node);
  });
  if (options.length === 0) return;

  const hasPreferred = preferredBand && options.some((option) => option.value === preferredBand);
  tierBandInput.value = hasPreferred ? preferredBand : getTierBandFromGt(typeInput.value, gt);
}

function getLightDuesTariff(type, tierBand, gt) {
  const tierOptions = LIGHT_DUES_TIER_OPTIONS[type];
  if (Array.isArray(tierOptions) && tierOptions.length > 0) {
    const resolvedBand = tierBand || getTierBandFromGt(type, gt);
    return LIGHT_DUES_TARIFFS[type][resolvedBand] || LIGHT_DUES_TARIFFS[type][tierOptions[0].value];
  }
  return LIGHT_DUES_TARIFFS[type] || LIGHT_DUES_TARIFFS.cargo;
}

function setTierVisibility(typeInput, tierWrap) {
  if (!typeInput || !tierWrap) return;
  const hasTierOptions = Boolean(LIGHT_DUES_TIER_OPTIONS[typeInput.value]);
  tierWrap.style.display = hasTierOptions ? '' : 'none';
}

function saveLightDuesState() {
  const stateKey = getLightDuesStateKey();
  if (!stateKey) return;
  const typeInput = document.getElementById('lightDuesType');
  const tierBandInput = document.getElementById('lightDuesTierBand');
  const gtInput = document.getElementById('lightDuesGt');
  const period30Input = document.getElementById('lightDuesPeriod30');
  const period12Input = document.getElementById('lightDuesPeriod12');
  if (!typeInput || !tierBandInput || !gtInput || !period30Input || !period12Input) return;

  const state = {
    type: typeInput.value,
    tierBand: tierBandInput.value,
    gt: gtInput.value,
    period30: Boolean(period30Input.checked),
    period12: Boolean(period12Input.checked)
  };
  safeStorageSet(stateKey, JSON.stringify(state));
}

function restoreLightDuesState(typeInput, tierBandInput, gtInput, period30Input, period12Input) {
  const stateKey = getLightDuesStateKey();
  let preferredTierBand = '';

  if (stateKey) {
    const raw = safeStorageGet(stateKey);
    if (raw) {
      let state = null;
      try {
        state = JSON.parse(raw);
      } catch (error) {
        state = null;
      }
      if (state && typeof state === 'object') {
        if (typeof state.type === 'string' && state.type) typeInput.value = state.type;
        if (typeof state.gt === 'string') gtInput.value = state.gt;
        preferredTierBand = state.tierBand || state.bulkBand || '';
        if (typeof state.period30 === 'boolean') period30Input.checked = state.period30;
        if (typeof state.period12 === 'boolean') period12Input.checked = state.period12;
      }
    }
  }

  const sharedGt = safeStorageGet(STORAGE_KEYS.gt);
  if (sharedGt !== null) {
    gtInput.value = sharedGt;
  }

  if (!period30Input.checked && !period12Input.checked) {
    period30Input.checked = true;
  }
  if (period30Input.checked && period12Input.checked) {
    period12Input.checked = false;
  }
  rebuildTierOptions(typeInput, tierBandInput, preferredTierBand, Number(gtInput.value));
}

function enforceSingleLightDuesPeriod(changedInput, otherInput) {
  if (!changedInput || !otherInput) return;
  if (changedInput.checked) {
    otherInput.checked = false;
    return;
  }
  if (!otherInput.checked) {
    changedInput.checked = true;
  }
}

function calculateLightDues() {
  const typeInput = document.getElementById('lightDuesType');
  const tierBandInput = document.getElementById('lightDuesTierBand');
  const gtInput = document.getElementById('lightDuesGt');
  const period30Input = document.getElementById('lightDuesPeriod30');
  const period12Input = document.getElementById('lightDuesPeriod12');
  const tariff30 = document.getElementById('lightDuesTariff30');
  const tariff12 = document.getElementById('lightDuesTariff12');
  const amount30 = document.getElementById('lightDuesAmount30');
  const amount12 = document.getElementById('lightDuesAmount12');

  if (
    !typeInput || !tierBandInput || !gtInput || !period30Input || !period12Input ||
    !tariff30 || !tariff12 || !amount30 || !amount12
  ) return;

  if (!period30Input.checked && !period12Input.checked) {
    period30Input.checked = true;
  }
  if (period30Input.checked && period12Input.checked) {
    period12Input.checked = false;
  }

  const gt = Number(gtInput.value);
  const tariff = getLightDuesTariff(typeInput.value, tierBandInput.value, gt);
  tariff30.value = tariff.label30;
  tariff12.value = tariff.label12;

  const selectedPeriod = period12Input.checked ? '12' : '30';
  const selectedTariffLabel = selectedPeriod === '12' ? tariff.label12 : tariff.label30;
  const amountKey = isLightDuesSailingPage() ? STORAGE_KEYS.lightDuesAmountSailing : STORAGE_KEYS.lightDuesAmountPda;
  const tariffKey = isLightDuesSailingPage() ? STORAGE_KEYS.lightDuesTariffSailing : STORAGE_KEYS.lightDuesTariffPda;

  safeStorageSet(tariffKey, selectedTariffLabel);

  if (!Number.isFinite(gt) || gt <= 0) {
    amount30.value = '';
    amount12.value = '';
    safeStorageRemove(amountKey);
    saveLightDuesState();
    return;
  }

  const calculated30 = gt * tariff.rate30;
  const calculated12 = gt * tariff.rate12;
  amount30.value = formatMoneyValue(calculated30);
  amount12.value = formatMoneyValue(calculated12);

  const selectedAmount = selectedPeriod === '12' ? calculated12 : calculated30;
  safeStorageSet(amountKey, selectedAmount.toFixed(2));
  saveLightDuesState();
}

function initLightDues() {
  const typeInput = document.getElementById('lightDuesType');
  const tierBandInput = document.getElementById('lightDuesTierBand');
  const tierWrap = document.getElementById('lightDuesTierWrap');
  const gtInput = document.getElementById('lightDuesGt');
  const period30Input = document.getElementById('lightDuesPeriod30');
  const period12Input = document.getElementById('lightDuesPeriod12');
  if (!typeInput || !tierBandInput || !tierWrap || !gtInput || !period30Input || !period12Input) return;

  restoreLightDuesState(typeInput, tierBandInput, gtInput, period30Input, period12Input);
  setTierVisibility(typeInput, tierWrap);
  calculateLightDues();

  typeInput.addEventListener('change', () => {
    rebuildTierOptions(typeInput, tierBandInput, '', Number(gtInput.value));
    setTierVisibility(typeInput, tierWrap);
    calculateLightDues();
  });

  tierBandInput.addEventListener('change', calculateLightDues);

  period30Input.addEventListener('change', () => {
    enforceSingleLightDuesPeriod(period30Input, period12Input);
    calculateLightDues();
  });

  period12Input.addEventListener('change', () => {
    enforceSingleLightDuesPeriod(period12Input, period30Input);
    calculateLightDues();
  });

  gtInput.addEventListener('input', () => {
    const raw = gtInput.value.trim();
    if (raw) safeStorageSet(STORAGE_KEYS.gt, raw);
    else safeStorageRemove(STORAGE_KEYS.gt);

    if (LIGHT_DUES_TIER_OPTIONS[typeInput.value]) {
      rebuildTierOptions(typeInput, tierBandInput, '', Number(gtInput.value));
    }
    calculateLightDues();
  });

  window.addEventListener('storage', (event) => {
    if (event.key !== STORAGE_KEYS.gt) return;
    const newGt = safeStorageGet(STORAGE_KEYS.gt) || '';
    if (gtInput.value !== newGt) {
      gtInput.value = newGt;
      if (LIGHT_DUES_TIER_OPTIONS[typeInput.value]) {
        rebuildTierOptions(typeInput, tierBandInput, '', Number(gtInput.value));
      }
      calculateLightDues();
    }
  });
}

function isPortDuesPdaPage() {
  return Boolean(document.body && document.body.classList.contains('page-port-dues-pda'));
}

function isPortDuesSailingPage() {
  return Boolean(document.body && document.body.classList.contains('page-port-dues-sailing'));
}

function getPortDuesStateKey() {
  if (isPortDuesSailingPage()) return STORAGE_KEYS.portDuesStateSailing;
  if (isPortDuesPdaPage()) return STORAGE_KEYS.portDuesState;
  return null;
}

function getPortDuesAmountKey() {
  if (isPortDuesSailingPage()) return STORAGE_KEYS.portDuesAmountSailing;
  if (isPortDuesPdaPage()) return STORAGE_KEYS.portDuesAmountPda;
  return null;
}

function getPortDuesCargoAmountKey() {
  if (isPortDuesSailingPage()) return STORAGE_KEYS.portDuesCargoAmountSailing;
  if (isPortDuesPdaPage()) return STORAGE_KEYS.portDuesCargoAmountPda;
  return null;
}

function getPortDuesBunkeringAmountKey() {
  if (isPortDuesSailingPage()) return STORAGE_KEYS.portDuesBunkeringAmountSailing;
  if (isPortDuesPdaPage()) return STORAGE_KEYS.portDuesBunkeringAmountPda;
  return null;
}

function getPortDuesCargoTypeConfig(type) {
  return PORT_DUES_CARGO_TYPES[type] || PORT_DUES_CARGO_TYPES.bulkCargo;
}

function setPortDuesBunkeringState(toggleInput, quantityInput) {
  if (!toggleInput || !quantityInput) return;
  quantityInput.disabled = !toggleInput.checked;
}

function savePortDuesState() {
  const stateKey = getPortDuesStateKey();
  if (!stateKey) return;

  const cargoQuantityInput = document.getElementById('portDuesCargoQuantity');
  const cargoTypeInput = document.getElementById('portDuesCargoType');
  const bunkeringToggleInput = document.getElementById('portDuesBunkeringToggle');
  const bunkeringQuantityInput = document.getElementById('portDuesBunkeringQuantity');
  if (!cargoQuantityInput || !cargoTypeInput || !bunkeringToggleInput || !bunkeringQuantityInput) return;

  const state = {
    cargoQuantity: cargoQuantityInput.value,
    cargoType: cargoTypeInput.value,
    bunkeringEnabled: Boolean(bunkeringToggleInput.checked),
    bunkeringQuantity: bunkeringQuantityInput.value
  };
  safeStorageSet(stateKey, JSON.stringify(state));
}

function restorePortDuesState(cargoQuantityInput, cargoTypeInput, bunkeringToggleInput, bunkeringQuantityInput) {
  const stateKey = getPortDuesStateKey();
  if (stateKey) {
    const raw = safeStorageGet(stateKey);
    if (raw) {
      let state = null;
      try {
        state = JSON.parse(raw);
      } catch (error) {
        state = null;
      }
      if (state && typeof state === 'object') {
        if (typeof state.cargoQuantity === 'string') cargoQuantityInput.value = state.cargoQuantity;
        if (typeof state.cargoType === 'string' && PORT_DUES_CARGO_TYPES[state.cargoType]) {
          cargoTypeInput.value = state.cargoType;
        }
        if (typeof state.bunkeringEnabled === 'boolean') bunkeringToggleInput.checked = state.bunkeringEnabled;
        if (typeof state.bunkeringQuantity === 'string') bunkeringQuantityInput.value = state.bunkeringQuantity;
      }
    }
  }

  const sharedQuantity = safeStorageGet(STORAGE_KEYS.quantity);
  if (sharedQuantity !== null) {
    cargoQuantityInput.value = sharedQuantity;
  }

  if (!PORT_DUES_CARGO_TYPES[cargoTypeInput.value]) {
    cargoTypeInput.value = 'bulkCargo';
  }
  setPortDuesBunkeringState(bunkeringToggleInput, bunkeringQuantityInput);
}

function calculatePortDues() {
  const cargoQuantityInput = document.getElementById('portDuesCargoQuantity');
  const cargoTypeInput = document.getElementById('portDuesCargoType');
  const bunkeringToggleInput = document.getElementById('portDuesBunkeringToggle');
  const bunkeringQuantityInput = document.getElementById('portDuesBunkeringQuantity');
  const cargoTariffInput = document.getElementById('portDuesCargoTariff');
  const cargoAmountInput = document.getElementById('portDuesCargoAmount');
  const bunkeringTariffInput = document.getElementById('portDuesBunkeringTariff');
  const bunkeringAmountInput = document.getElementById('portDuesBunkeringAmount');
  const totalAmountInput = document.getElementById('portDuesTotalAmount');
  if (
    !cargoQuantityInput || !cargoTypeInput || !bunkeringToggleInput || !bunkeringQuantityInput ||
    !cargoTariffInput || !cargoAmountInput || !bunkeringTariffInput || !bunkeringAmountInput || !totalAmountInput
  ) return;

  const cargoType = getPortDuesCargoTypeConfig(cargoTypeInput.value);
  const cargoQuantity = parseMoneyValue(cargoQuantityInput.value);
  const cargoAmount = cargoQuantity > 0 ? cargoQuantity * cargoType.rate : 0;

  const bunkeringQuantity = parseMoneyValue(bunkeringQuantityInput.value);
  const bunkeringAmount = bunkeringToggleInput.checked && bunkeringQuantity > 0 ? bunkeringQuantity * 0.3 : 0;

  const totalAmount = cargoAmount + bunkeringAmount;
  const amountKey = getPortDuesAmountKey();
  const cargoAmountKey = getPortDuesCargoAmountKey();
  const bunkeringAmountKey = getPortDuesBunkeringAmountKey();
  if (amountKey) {
    safeStorageSet(amountKey, totalAmount.toFixed(2));
  }
  if (cargoAmountKey) {
    safeStorageSet(cargoAmountKey, cargoAmount.toFixed(2));
  }
  if (bunkeringAmountKey) {
    if (bunkeringAmount > 0) safeStorageSet(bunkeringAmountKey, bunkeringAmount.toFixed(2));
    else safeStorageRemove(bunkeringAmountKey);
  }

  const quantityRaw = cargoQuantityInput.value.trim();
  if (quantityRaw) safeStorageSet(STORAGE_KEYS.quantity, quantityRaw);
  else safeStorageRemove(STORAGE_KEYS.quantity);

  cargoTariffInput.value = cargoType.label;
  cargoAmountInput.value = formatMoneyValue(cargoAmount);
  bunkeringTariffInput.value = '0,30';
  bunkeringAmountInput.value = formatMoneyValue(bunkeringAmount);
  totalAmountInput.value = formatMoneyValue(totalAmount);
  savePortDuesState();
}

function initPortDues() {
  const cargoQuantityInput = document.getElementById('portDuesCargoQuantity');
  const cargoTypeInput = document.getElementById('portDuesCargoType');
  const bunkeringToggleInput = document.getElementById('portDuesBunkeringToggle');
  const bunkeringQuantityInput = document.getElementById('portDuesBunkeringQuantity');
  if (!cargoQuantityInput || !cargoTypeInput || !bunkeringToggleInput || !bunkeringQuantityInput) return;

  restorePortDuesState(cargoQuantityInput, cargoTypeInput, bunkeringToggleInput, bunkeringQuantityInput);
  calculatePortDues();

  cargoQuantityInput.addEventListener('input', calculatePortDues);
  cargoTypeInput.addEventListener('change', calculatePortDues);
  bunkeringToggleInput.addEventListener('change', () => {
    setPortDuesBunkeringState(bunkeringToggleInput, bunkeringQuantityInput);
    calculatePortDues();
  });
  bunkeringQuantityInput.addEventListener('input', calculatePortDues);

  window.addEventListener('storage', (event) => {
    if (event.key !== STORAGE_KEYS.quantity) return;
    const nextQuantity = safeStorageGet(STORAGE_KEYS.quantity) || '';
    if (cargoQuantityInput.value !== nextQuantity) {
      cargoQuantityInput.value = nextQuantity;
      calculatePortDues();
    }
  });
}

function isMooringPdaPage() {
  return Boolean(document.body && document.body.classList.contains('page-mooring-pda'));
}

function isMooringSailingPage() {
  return Boolean(document.body && document.body.classList.contains('page-mooring-sailing'));
}

function normalizeMooringOperation(value) {
  if (value === 'departure') return 'departure';
  if (value === 'shifting') return 'shifting';
  return 'arrival';
}

function getMooringStateKey() {
  if (isMooringSailingPage()) return STORAGE_KEYS.mooringStateSailing;
  if (isMooringPdaPage()) return STORAGE_KEYS.mooringState;
  return null;
}

function getMooringAmountKey() {
  if (isMooringSailingPage()) return STORAGE_KEYS.mooringAmountSailing;
  if (isMooringPdaPage()) return STORAGE_KEYS.mooringAmountPda;
  return null;
}

function getMooringBaseByGt(gt) {
  if (!Number.isFinite(gt) || gt <= 0) return 0;
  const matched = MOORING_GT_TARIFFS.find((item) => gt >= item.min && gt <= item.max);
  if (matched) return matched.amount;

  const baseAt35000 = 285.6;
  const additionalThousands = Math.ceil((gt - 35000) / 1000);
  return baseAt35000 + Math.max(0, additionalThousands) * 10;
}

function setMooringAddonState(enabled, fieldsWrap, ...inputs) {
  if (fieldsWrap) {
    fieldsWrap.style.display = enabled ? '' : 'none';
  }
  inputs.forEach((input) => {
    if (!input) return;
    input.disabled = !enabled;
  });
}

function getMooringLegacyCardFromState(state) {
  const hasLegacyFields =
    Object.prototype.hasOwnProperty.call(state, 'overtime25') ||
    Object.prototype.hasOwnProperty.call(state, 'overtime50') ||
    Object.prototype.hasOwnProperty.call(state, 'boatEnabled') ||
    Object.prototype.hasOwnProperty.call(state, 'boatHours') ||
    Object.prototype.hasOwnProperty.call(state, 'manEnabled') ||
    Object.prototype.hasOwnProperty.call(state, 'manHours') ||
    Object.prototype.hasOwnProperty.call(state, 'manPersons');
  if (!hasLegacyFields) return null;

  return {
    operation: normalizeMooringOperation(state.operation),
    overtime25: Boolean(state.overtime25),
    overtime50: Boolean(state.overtime50),
    boatEnabled: Boolean(state.boatEnabled),
    boatHours: state.boatHours == null ? '' : String(state.boatHours),
    manEnabled: Boolean(state.manEnabled),
    manHours: state.manHours == null ? '' : String(state.manHours),
    manPersons: state.manPersons == null ? '' : String(state.manPersons)
  };
}

function getMooringCardsFromState(state) {
  if (!state || typeof state !== 'object') return [];

  if (Array.isArray(state.cards)) {
    return state.cards.map((card) => ({
      operation: normalizeMooringOperation(card?.operation),
      overtime25: Boolean(card?.overtime25),
      overtime50: Boolean(card?.overtime50),
      boatEnabled: Boolean(card?.boatEnabled),
      boatHours: card?.boatHours == null ? '' : String(card.boatHours),
      manEnabled: Boolean(card?.manEnabled),
      manHours: card?.manHours == null ? '' : String(card.manHours),
      manPersons: card?.manPersons == null ? '' : String(card.manPersons)
    }));
  }

  const legacyCard = getMooringLegacyCardFromState(state);
  return legacyCard ? [legacyCard] : [];
}

function getMooringCardCalculation(cardState, gt) {
  const operation = normalizeMooringOperation(cardState?.operation);
  const overtime25 = Boolean(cardState?.overtime25);
  const overtime50 = Boolean(cardState?.overtime50);
  const boatEnabled = Boolean(cardState?.boatEnabled);
  const boatHours = parseMoneyValue(cardState?.boatHours);
  const manEnabled = Boolean(cardState?.manEnabled);
  const manHours = parseMoneyValue(cardState?.manHours);
  const manPersons = parseMoneyValue(cardState?.manPersons);

  const baseAmount = getMooringBaseByGt(gt);
  const overtimeFactor = 1 + (overtime25 ? 0.25 : 0) + (overtime50 ? 0.5 : 0);
  const baseWithOvertime = baseAmount * overtimeFactor;
  const boatAmount = boatEnabled ? 125 * boatHours : 0;
  const manAmount = manEnabled ? 25 * manHours * manPersons : 0;
  const totalAmount = baseWithOvertime + boatAmount + manAmount;

  return {
    operation,
    overtime25,
    overtime50,
    boatEnabled,
    boatHours,
    manEnabled,
    manHours,
    manPersons,
    baseAmount,
    overtimeFactor,
    baseWithOvertime,
    boatAmount,
    manAmount,
    totalAmount
  };
}

function getMooringCalculationFromState(state, gtValue) {
  if (!state || typeof state !== 'object') return null;
  const gt = Number.isFinite(gtValue) && gtValue > 0 ? gtValue : parseMoneyValue(state.gt);
  if (!Number.isFinite(gt) || gt <= 0) return null;

  const cardStates = getMooringCardsFromState(state);
  const cards = cardStates.map((cardState) => getMooringCardCalculation(cardState, gt));

  let arrivalTotal = 0;
  let departureTotal = 0;
  let shiftingTotal = 0;
  cards.forEach((card) => {
    if (card.operation === 'departure') {
      departureTotal += card.totalAmount;
    } else if (card.operation === 'shifting') {
      shiftingTotal += card.totalAmount;
    } else {
      arrivalTotal += card.totalAmount;
    }
  });
  const totalAmount = arrivalTotal + departureTotal + shiftingTotal;
  const firstCard = cards[0] || {
    baseAmount: 0,
    overtimeFactor: 1,
    baseWithOvertime: 0,
    boatAmount: 0,
    manAmount: 0
  };

  return {
    gt,
    cards,
    arrivalTotal,
    departureTotal,
    shiftingTotal,
    baseAmount: firstCard.baseAmount,
    overtimeFactor: firstCard.overtimeFactor,
    baseWithOvertime: firstCard.baseWithOvertime,
    boatAmount: firstCard.boatAmount,
    manAmount: firstCard.manAmount,
    totalAmount
  };
}

let mooringCardCount = 0;
let isRestoringMooring = false;

function getMooringCardState(id) {
  return {
    operation: normalizeMooringOperation(document.getElementById(`mooringOp_${id}`)?.value),
    overtime25: Boolean(document.getElementById(`mooringOt25_${id}`)?.checked),
    overtime50: Boolean(document.getElementById(`mooringOt50_${id}`)?.checked),
    boatEnabled: Boolean(document.getElementById(`mooringBoatEnabled_${id}`)?.checked),
    boatHours: document.getElementById(`mooringBoatHours_${id}`)?.value || '',
    manEnabled: Boolean(document.getElementById(`mooringManEnabled_${id}`)?.checked),
    manHours: document.getElementById(`mooringManHours_${id}`)?.value || '',
    manPersons: document.getElementById(`mooringManPersons_${id}`)?.value || ''
  };
}

function updateMooringCardTitles() {
  const cards = document.querySelectorAll('#mooringCards .card[data-mooring-card]');
  const operations = Array.from(cards).map((card) => {
    const id = card.dataset.cardId;
    return normalizeMooringOperation(document.getElementById(`mooringOp_${id}`)?.value);
  });
  const operationCounts = operations.reduce((acc, operation) => {
    acc[operation] = (acc[operation] || 0) + 1;
    return acc;
  }, {});

  let arrivalIndex = 0;
  let departureIndex = 0;
  let shiftingIndex = 0;

  cards.forEach((card) => {
    const id = card.dataset.cardId;
    const operation = normalizeMooringOperation(document.getElementById(`mooringOp_${id}`)?.value);
    const title = card.querySelector('.mooring-title');
    if (!title) return;

    if (operation === 'departure') {
      departureIndex += 1;
      const prefix = operationCounts.departure > 1 ? `${departureIndex}${getOrdinal(departureIndex)} ` : '';
      title.textContent = `${prefix}Unmooring on departure`;
      return;
    }

    if (operation === 'shifting') {
      shiftingIndex += 1;
      const prefix = operationCounts.shifting > 1 ? `${shiftingIndex}${getOrdinal(shiftingIndex)} ` : '';
      title.textContent = `${prefix}Shifting berth to berth`;
      return;
    }

    arrivalIndex += 1;
    const prefix = operationCounts.arrival > 1 ? `${arrivalIndex}${getOrdinal(arrivalIndex)} ` : '';
    title.textContent = `${prefix}Mooring on arrival`;
  });
}

function addMooringCard(initialState = {}) {
  const cardsWrap = document.getElementById('mooringCards');
  if (!cardsWrap) return;

  mooringCardCount += 1;
  const id = mooringCardCount;
  const operation = normalizeMooringOperation(initialState.operation);

  cardsWrap.insertAdjacentHTML('beforeend', `
    <div class="card" id="mooring_${id}" data-mooring-card="1" data-card-id="${id}">
      <div class="tug-header">
        <button class="icon-btn" onclick="removeMooringCard(${id})" aria-label="Remove mooring card">
          <img src="assets/icons/remove_48×48.png" alt="Remove" />
        </button>
        <h3 class="mooring-title">Mooring</h3>
        <span></span>
      </div>

      <label>Operation</label>
      <select id="mooringOp_${id}">
        <option value="arrival"${operation === 'arrival' ? ' selected' : ''}>Mooring on arrival</option>
        <option value="departure"${operation === 'departure' ? ' selected' : ''}>Unmooring on departure</option>
        <option value="shifting"${operation === 'shifting' ? ' selected' : ''}>Shifting berth to berth</option>
      </select>

      <div class="section-title">Overtime</div>
      <label class="checkbox"><input id="mooringOt25_${id}" type="checkbox" /> 25% overtime - 22:00 - 6:00 hrs</label>
      <label class="checkbox"><input id="mooringOt50_${id}" type="checkbox" /> 50% overtime - Holidays</label>

      <div class="section-title">Additions</div>
      <label class="checkbox"><input id="mooringBoatEnabled_${id}" type="checkbox" /> Use mooring boat (125 EUR / hour)</label>
      <div id="mooringBoatFields_${id}">
        <label>Boat Hours</label>
        <input id="mooringBoatHours_${id}" type="number" min="0" step="0.25" placeholder="0" />
      </div>

      <label class="checkbox"><input id="mooringManEnabled_${id}" type="checkbox" /> Use additional mooring man (25 EUR / hour)</label>
      <div id="mooringManFields_${id}">
        <div class="row">
          <div>
            <label>Man Hours</label>
            <input id="mooringManHours_${id}" type="number" min="0" step="0.25" placeholder="0" />
          </div>
          <div>
            <label>Persons</label>
            <input id="mooringManPersons_${id}" type="number" min="0" step="1" placeholder="0" />
          </div>
        </div>
      </div>

      <div class="tug-total" id="mooringCardTotal_${id}">Card total: €0.00</div>
    </div>
  `);

  if (Object.prototype.hasOwnProperty.call(initialState, 'overtime25')) {
    const input = document.getElementById(`mooringOt25_${id}`);
    if (input) input.checked = Boolean(initialState.overtime25);
  }
  if (Object.prototype.hasOwnProperty.call(initialState, 'overtime50')) {
    const input = document.getElementById(`mooringOt50_${id}`);
    if (input) input.checked = Boolean(initialState.overtime50);
  }
  if (Object.prototype.hasOwnProperty.call(initialState, 'boatEnabled')) {
    const input = document.getElementById(`mooringBoatEnabled_${id}`);
    if (input) input.checked = Boolean(initialState.boatEnabled);
  }
  if (typeof initialState.boatHours === 'string') {
    const input = document.getElementById(`mooringBoatHours_${id}`);
    if (input) input.value = initialState.boatHours;
  }
  if (Object.prototype.hasOwnProperty.call(initialState, 'manEnabled')) {
    const input = document.getElementById(`mooringManEnabled_${id}`);
    if (input) input.checked = Boolean(initialState.manEnabled);
  }
  if (typeof initialState.manHours === 'string') {
    const input = document.getElementById(`mooringManHours_${id}`);
    if (input) input.value = initialState.manHours;
  }
  if (typeof initialState.manPersons === 'string') {
    const input = document.getElementById(`mooringManPersons_${id}`);
    if (input) input.value = initialState.manPersons;
  }

  const boatEnabledInput = document.getElementById(`mooringBoatEnabled_${id}`);
  const boatFields = document.getElementById(`mooringBoatFields_${id}`);
  const boatHoursInput = document.getElementById(`mooringBoatHours_${id}`);
  const manEnabledInput = document.getElementById(`mooringManEnabled_${id}`);
  const manFields = document.getElementById(`mooringManFields_${id}`);
  const manHoursInput = document.getElementById(`mooringManHours_${id}`);
  const manPersonsInput = document.getElementById(`mooringManPersons_${id}`);
  setMooringAddonState(Boolean(boatEnabledInput?.checked), boatFields, boatHoursInput);
  setMooringAddonState(Boolean(manEnabledInput?.checked), manFields, manHoursInput, manPersonsInput);

  updateMooringCardTitles();
  if (!isRestoringMooring) calculateMooring();
}

function addMooringArrivalCard() {
  addMooringCard({ operation: 'arrival' });
}

function addMooringDepartureCard() {
  addMooringCard({ operation: 'departure' });
}

function removeMooringCard(id) {
  const card = document.getElementById(`mooring_${id}`);
  if (!card) return;
  card.remove();
  updateMooringCardTitles();
  calculateMooring();
}

function saveMooringState() {
  const stateKey = getMooringStateKey();
  if (!stateKey) return;

  const gtInput = document.getElementById('mooringGt');
  const cardsWrap = document.getElementById('mooringCards');
  if (!gtInput || !cardsWrap) return;

  const cards = Array.from(cardsWrap.querySelectorAll('.card[data-mooring-card]')).map((card) => {
    const id = card.dataset.cardId;
    return getMooringCardState(id);
  });

  const state = {
    gt: gtInput.value,
    cards
  };
  safeStorageSet(stateKey, JSON.stringify(state));
}

function restoreMooringState() {
  const gtInput = document.getElementById('mooringGt');
  const cardsWrap = document.getElementById('mooringCards');
  if (!gtInput || !cardsWrap) return;

  const stateKey = getMooringStateKey();
  let state = null;
  if (stateKey) {
    state = parseStoredJson(safeStorageGet(stateKey));
    if (state && typeof state.gt === 'string') gtInput.value = state.gt;
  }

  const sharedGt = safeStorageGet(STORAGE_KEYS.gt);
  if (sharedGt !== null) {
    gtInput.value = sharedGt;
  }

  const cards = getMooringCardsFromState(state || {});
  isRestoringMooring = true;
  cardsWrap.innerHTML = '';
  mooringCardCount = 0;
  if (cards.length === 0) {
    addMooringCard({ operation: 'arrival' });
    addMooringCard({ operation: 'departure' });
  } else {
    cards.forEach((cardState) => addMooringCard(cardState));
    const hasArrival = cards.some((cardState) => normalizeMooringOperation(cardState.operation) === 'arrival');
    const hasDeparture = cards.some((cardState) => normalizeMooringOperation(cardState.operation) === 'departure');
    if (!hasArrival) addMooringCard({ operation: 'arrival' });
    if (!hasDeparture) addMooringCard({ operation: 'departure' });
  }
  isRestoringMooring = false;
  updateMooringCardTitles();
}

function calculateMooring() {
  const gtInput = document.getElementById('mooringGt');
  const cardsWrap = document.getElementById('mooringCards');
  const final = document.getElementById('finalMooringTotal');
  if (!gtInput || !cardsWrap || !final) return;

  const gtRaw = gtInput.value.trim();
  if (gtRaw) safeStorageSet(STORAGE_KEYS.gt, gtRaw);
  else safeStorageRemove(STORAGE_KEYS.gt);
  const gt = parseMoneyValue(gtInput.value);
  const hasValidGt = Number.isFinite(gt) && gt > 0;

  let arrivalTotal = 0;
  let departureTotal = 0;
  let shiftingTotal = 0;
  const cardStates = [];
  const cards = Array.from(cardsWrap.querySelectorAll('.card[data-mooring-card]'));
  cards.forEach((card) => {
    const id = card.dataset.cardId;
    const cardState = getMooringCardState(id);
    cardStates.push(cardState);

    const cardTotal = document.getElementById(`mooringCardTotal_${id}`);
    if (!hasValidGt) {
      if (cardTotal) cardTotal.textContent = 'Card total: €0.00';
      return;
    }

    const calculation = getMooringCardCalculation(cardState, gt);
    if (calculation.operation === 'departure') {
      departureTotal += calculation.totalAmount;
    } else if (calculation.operation === 'shifting') {
      shiftingTotal += calculation.totalAmount;
    } else {
      arrivalTotal += calculation.totalAmount;
    }
    if (cardTotal) cardTotal.textContent = `Card total: €${calculation.totalAmount.toFixed(2)}`;
  });

  const grandTotal = arrivalTotal + departureTotal + shiftingTotal;
  final.style.display = 'block';
  final.innerHTML = `
    <div class="summary">
      <div><strong>Arrival total</strong><br>€${arrivalTotal.toFixed(2)}</div>
      <div><strong>Departure total</strong><br>€${departureTotal.toFixed(2)}</div>
      <div><strong>Shifting total</strong><br>€${shiftingTotal.toFixed(2)}</div>
      <div><strong>Grand total</strong><br>€${grandTotal.toFixed(2)}</div>
    </div>
  `;

  const amountKey = getMooringAmountKey();

  if (!hasValidGt) {
    if (amountKey) safeStorageRemove(amountKey);
  } else if (amountKey) {
    safeStorageSet(amountKey, grandTotal.toFixed(2));
  }

  const stateKey = getMooringStateKey();
  if (stateKey) {
    safeStorageSet(stateKey, JSON.stringify({ gt: gtInput.value, cards: cardStates }));
  } else {
    saveMooringState();
  }
}

function initMooring() {
  const gtInput = document.getElementById('mooringGt');
  const cardsWrap = document.getElementById('mooringCards');
  if (!gtInput || !cardsWrap) return;

  restoreMooringState();
  calculateMooring();

  gtInput.addEventListener('input', calculateMooring);

  cardsWrap.addEventListener('input', (event) => {
    const target = event.target;
    if (!target) return;
    if (target.id && target.id.startsWith('mooringOp_')) {
      updateMooringCardTitles();
    }
    calculateMooring();
  });

  cardsWrap.addEventListener('change', (event) => {
    const target = event.target;
    if (!target) return;
    if (target.id && target.id.startsWith('mooringOp_')) {
      updateMooringCardTitles();
    }
    if (target.id && target.id.startsWith('mooringBoatEnabled_')) {
      const id = target.id.replace('mooringBoatEnabled_', '');
      const boatFields = document.getElementById(`mooringBoatFields_${id}`);
      const boatHoursInput = document.getElementById(`mooringBoatHours_${id}`);
      setMooringAddonState(Boolean(target.checked), boatFields, boatHoursInput);
    }
    if (target.id && target.id.startsWith('mooringManEnabled_')) {
      const id = target.id.replace('mooringManEnabled_', '');
      const manFields = document.getElementById(`mooringManFields_${id}`);
      const manHoursInput = document.getElementById(`mooringManHours_${id}`);
      const manPersonsInput = document.getElementById(`mooringManPersons_${id}`);
      setMooringAddonState(Boolean(target.checked), manFields, manHoursInput, manPersonsInput);
    }
    calculateMooring();
  });

  window.addEventListener('storage', (event) => {
    if (event.key !== STORAGE_KEYS.gt) return;
    const nextGt = safeStorageGet(STORAGE_KEYS.gt) || '';
    if (gtInput.value !== nextGt) {
      gtInput.value = nextGt;
      calculateMooring();
    }
  });
}

// Tug calculator logic
let tugCount = 0;
const MIN_VOYAGE = 1;
const MIN_ASSIST = 0.5;
let imoMaster = null;
let linesMaster = null;

function getTariffFromGT(gt) {
  if (!gt || gt <= 0) return 0;

  if (gt <= 2000) return 698;
  if (gt <= 5000) return 908;
  if (gt <= 10000) return 1133;
  if (gt <= 15000) return 1294;
  if (gt <= 25000) return 1358;
  if (gt <= 30000) return 1646;
  if (gt <= 35000) return 1863;

  const extraThousands = Math.ceil((gt - 35000) / 1000);
  return 1863 + (extraThousands * 16);
}

function addTug() {
  const tugCards = document.getElementById('tugCards');
  if (!tugCards) return;
  tugCount++;
  tugCards.insertAdjacentHTML('beforeend', `
    <div class="card" id="tug_${tugCount}">
      <div class="tug-header">
        <button class="icon-btn" onclick="removeTug(${tugCount})" aria-label="Remove tugboat"><img src="assets/icons/remove_48×48.png" alt="Remove"></button>
        <h3 class="tug-title">Tugboat</h3>
        <button class="icon-btn" onclick="duplicateTug(${tugCount})" aria-label="Duplicate tugboat"><img src="assets/icons/duplicate_48x48.png" alt="Duplicate"></button>
      </div>

      <label>Operation</label>
      <select id="op_${tugCount}">
        <option value="arrival">Arrival (IN)</option>
        <option value="departure">Departure (OUT)</option>
      </select>

      <label>Voyage Time</label>
      <select id="voyage_${tugCount}">
        <option value="1">1 hour</option>
        <option value="1.5">1.5 hours</option>
      </select>

      <label>Assistance Time</label>
      <select id="assist_${tugCount}">
        <option value="0.5">Up to 30 min</option>
        <option value="1">Within 1 hour</option>
        <option value="1.5">Within 1.5 hours</option>
      </select>

      <div class="section-title">Assistance Surcharges</div>
      <label class="checkbox"><input type="checkbox" id="imo_${tugCount}" /> 20% IMO Code Classes I-III</label>
      <label class="checkbox"><input type="checkbox" id="lines_${tugCount}" /> 15% Usage of tug lines</label>
      <label class="checkbox"><input type="checkbox" id="kw_${tugCount}" /> 30% Tug above 2000 kW</label>

      <div class="section-title">Overtime</div>
      <label>Voyage Overtime</label>
      <select id="voy_ot_${tugCount}">
        <option value="0"></option>
        <option value="0.25">25% (Mon–Sat)</option>
        <option value="0.50">50% (Sun & Holidays)</option>
      </select>

      <label>Assistance Overtime</label>
      <select id="assist_ot_${tugCount}">
        <option value="0"></option>
        <option value="0.25">25% (Mon–Sat)</option>
        <option value="0.50">50% (Sun & Holidays)</option>
      </select>

      <div class="tug-total" id="tugTotal_${tugCount}">Tug total: €0.00</div>
    </div>
  `);

  if (isRestoringTugs) return;

  applyImoMaster();
  applyLinesMaster();
  syncImoMaster();
  syncLinesMaster();
  updateTugTitles();
  calculate();
}

function removeTug(id) {
  document.getElementById(`tug_${id}`)?.remove();
  syncImoMaster();
  syncLinesMaster();
  updateTugTitles();
  calculate();
}

function duplicateTug(id) {
  const source = document.getElementById(`tug_${id}`);
  if (!source) return;

  addTug();
  const newId = tugCount;

  const fields = [
    ['op', 'value'],
    ['voyage', 'value'],
    ['assist', 'value'],
    ['imo', 'checked'],
    ['lines', 'checked'],
    ['kw', 'checked'],
    ['voy_ot', 'value'],
    ['assist_ot', 'value']
  ];

  for (const [key, prop] of fields) {
    const from = document.getElementById(`${key}_${id}`);
    const to = document.getElementById(`${key}_${newId}`);
    if (!from || !to) continue;
    to[prop] = from[prop];
  }

  syncImoMaster();
  syncLinesMaster();
  updateTugTitles();
  calculate();
}

function duplicateLastTug() {
  const selectedId = getSelectedTugId();
  if (selectedId) {
    duplicateTug(selectedId);
    return;
  }
  const cards = document.querySelectorAll('#tugCards .card');
  if (cards.length === 0) {
    addTug();
    return;
  }
  const last = cards[cards.length - 1];
  const id = last.id.split('_')[1];
  duplicateTug(Number(id));
}

function setSelectedTug(id) {
  const cards = document.querySelectorAll('#tugCards .card');
  cards.forEach(card => card.classList.remove('selected'));
  const target = document.getElementById(`tug_${id}`);
  if (target) target.classList.add('selected');
}

function getSelectedTugId() {
  const selected = document.querySelector('#tugCards .card.selected');
  if (!selected) return null;
  const parts = selected.id.split('_');
  return parts.length > 1 ? Number(parts[1]) : null;
}

function applyImoMaster() {
  if (!imoMaster) return;
  const checked = imoMaster.checked;
  imoMaster.indeterminate = false;
  for (let i = 1; i <= tugCount; i++) {
    const cb = document.getElementById(`imo_${i}`);
    if (cb) cb.checked = checked;
  }
}

function applyLinesMaster() {
  if (!linesMaster) return;
  const checked = linesMaster.checked;
  linesMaster.indeterminate = false;
  for (let i = 1; i <= tugCount; i++) {
    const cb = document.getElementById(`lines_${i}`);
    if (cb) cb.checked = checked;
  }
}

function syncImoMaster() {
  if (!imoMaster) return;
  let total = 0;
  let checked = 0;
  for (let i = 1; i <= tugCount; i++) {
    const cb = document.getElementById(`imo_${i}`);
    if (!cb) continue;
    total++;
    if (cb.checked) checked++;
  }
  if (total === 0) {
    imoMaster.checked = false;
    imoMaster.indeterminate = false;
    return;
  }
  if (checked === 0) {
    imoMaster.checked = false;
    imoMaster.indeterminate = false;
  } else if (checked === total) {
    imoMaster.checked = true;
    imoMaster.indeterminate = false;
  } else {
    imoMaster.checked = false;
    imoMaster.indeterminate = true;
  }
}

function syncLinesMaster() {
  if (!linesMaster) return;
  let total = 0;
  let checked = 0;
  for (let i = 1; i <= tugCount; i++) {
    const cb = document.getElementById(`lines_${i}`);
    if (!cb) continue;
    total++;
    if (cb.checked) checked++;
  }
  if (total === 0) {
    linesMaster.checked = false;
    linesMaster.indeterminate = false;
    return;
  }
  if (checked === 0) {
    linesMaster.checked = false;
    linesMaster.indeterminate = false;
  } else if (checked === total) {
    linesMaster.checked = true;
    linesMaster.indeterminate = false;
  } else {
    linesMaster.checked = false;
    linesMaster.indeterminate = true;
  }
}

function updateTugTitles() {
  const cards = document.querySelectorAll('#tugCards .card');
  const operations = Array.from(cards).map((card) => {
    const id = card.id.split('_')[1];
    return document.getElementById(`op_${id}`)?.value === 'departure' ? 'departure' : 'arrival';
  });
  const operationCounts = operations.reduce((acc, operation) => {
    acc[operation] = (acc[operation] || 0) + 1;
    return acc;
  }, {});

  let arrivalIndex = 0;
  let departureIndex = 0;

  cards.forEach(card => {
    const id = card.id.split('_')[1];
    const op = document.getElementById(`op_${id}`)?.value;
    const title = card.querySelector('.tug-title');
    if (!title) return;

    if (op === 'arrival') {
      arrivalIndex++;
      const prefix = operationCounts.arrival > 1 ? `${arrivalIndex}${getOrdinal(arrivalIndex)} ` : '';
      title.innerText = `${prefix}Tugboat on arrival`;
    } else {
      departureIndex++;
      const prefix = operationCounts.departure > 1 ? `${departureIndex}${getOrdinal(departureIndex)} ` : '';
      title.innerText = `${prefix}Tugboat on departure`;
    }
  });
}

function getOrdinal(n) {
  if (n % 10 === 1 && n % 100 !== 11) return 'st';
  if (n % 10 === 2 && n % 100 !== 12) return 'nd';
  if (n % 10 === 3 && n % 100 !== 13) return 'rd';
  return 'th';
}

function calculate() {
  saveTugsState();
  const tariffInput = document.getElementById('tariff');
  const final = document.getElementById('finalTotal');
  if (!tariffInput || !final) return;

  const tariff = Number(tariffInput.value) || 0;
  if (!tariff) {
    final.style.display = 'none';
    const keys = getTugStorageKeys(isSailingTugsPage());
    safeStorageSet(keys.towageArrivalCount, 0);
    safeStorageSet(keys.towageDepartureCount, 0);
    return;
  }

  let arrivalTotal = 0;
  let departureTotal = 0;
  let arrivalCount = 0;
  let departureCount = 0;

  for (let i = 1; i <= tugCount; i++) {
    if (!document.getElementById(`tug_${i}`)) continue;

    let voyage = Number(document.getElementById(`voyage_${i}`).value);
    let assist = Number(document.getElementById(`assist_${i}`).value);
    const op = document.getElementById(`op_${i}`).value;

    if (op === 'arrival') arrivalCount += 1;
    else departureCount += 1;

    voyage = Math.max(voyage, MIN_VOYAGE);
    assist = Math.max(assist, MIN_ASSIST);

    const voyageOT = Number(document.getElementById(`voy_ot_${i}`).value);
    const assistOT = Number(document.getElementById(`assist_ot_${i}`).value);

    const voyageCost = tariff * voyage * (1 + voyageOT);

    let assistMultiplier = 1 + assistOT;
    if (document.getElementById(`imo_${i}`).checked) assistMultiplier += 0.20;
    if (document.getElementById(`lines_${i}`).checked) assistMultiplier += 0.15;
    if (document.getElementById(`kw_${i}`).checked) assistMultiplier += 0.30;

    const assistCost = tariff * assist * assistMultiplier;
    const total = voyageCost + assistCost;

    document.getElementById(`tugTotal_${i}`).innerText = `Tug total: €${total.toFixed(2)}`;

    if (op === 'arrival') arrivalTotal += total;
    else departureTotal += total;
  }

  if (arrivalTotal === 0 && departureTotal === 0) {
    final.style.display = 'none';
    const keys = getTugStorageKeys(isSailingTugsPage());
    safeStorageSet(keys.towageArrivalCount, arrivalCount);
    safeStorageSet(keys.towageDepartureCount, departureCount);
    return;
  }

  final.style.display = 'block';
  final.innerHTML = `
    <div class="summary">
      <div><strong>Arrival total</strong><br>€${arrivalTotal.toFixed(2)}</div>
      <div><strong>Departure total</strong><br>€${departureTotal.toFixed(2)}</div>
      <div><strong>Grand total</strong><br>€${(arrivalTotal + departureTotal).toFixed(2)}</div>
    </div>
  `;

  const grandTotal = arrivalTotal + departureTotal;
  if (Number.isFinite(grandTotal)) {
    const keys = getTugStorageKeys(isSailingTugsPage());
    safeStorageSet(keys.towageTotal, grandTotal.toFixed(2));
    safeStorageSet(keys.towageArrivalCount, arrivalCount);
    safeStorageSet(keys.towageDepartureCount, departureCount);
  }
}

function initTugs() {
  const tugCards = document.getElementById('tugCards');
  if (!tugCards) return;

  const vesselNameInput = document.getElementById('vesselName');
  if (vesselNameInput) {
    updateVesselNameFromStorage(vesselNameInput);
    window.addEventListener('storage', (event) => {
      if (event.key === STORAGE_KEYS.vesselName) {
        updateVesselNameFromStorage(vesselNameInput);
      }
    });
  }

  const gtInput = document.getElementById('gt');
  if (gtInput) {
    const syncGtAndTariff = () => {
      const storedGt = safeStorageGet(STORAGE_KEYS.gt);
      if (storedGt && gtInput.value !== storedGt) {
        gtInput.value = storedGt;
      }
      const gt = Number(gtInput.value);
      const tariff = getTariffFromGT(gt);
      const tariffInput = document.getElementById('tariff');
      if (tariffInput) tariffInput.value = tariff || '';
    };

    syncGtAndTariff();
    gtInput.addEventListener('input', () => {
      const gt = Number(gtInput.value);
      const tariff = getTariffFromGT(gt);
      const tariffInput = document.getElementById('tariff');
      if (tariffInput) tariffInput.value = tariff || '';
      const value = gtInput.value.trim();
      if (value) safeStorageSet(STORAGE_KEYS.gt, value);
      else safeStorageRemove(STORAGE_KEYS.gt);
      calculate();
    });

    window.addEventListener('storage', (event) => {
      if (event.key === STORAGE_KEYS.gt) {
        syncGtAndTariff();
        calculate();
      }
    });
  }

  document.addEventListener('input', calculate);
  document.addEventListener('change', () => {
    updateTugTitles();
    calculate();
  });

  tugCards.addEventListener('click', (event) => {
    const card = event.target.closest('.card');
    if (!card) return;
    const id = card.id.split('_')[1];
    setSelectedTug(id);
  });

  imoMaster = document.getElementById('imoMaster');
  linesMaster = document.getElementById('linesMaster');

  if (imoMaster) {
    imoMaster.addEventListener('change', () => {
      applyImoMaster();
      calculate();
    });
  }

  if (linesMaster) {
    linesMaster.addEventListener('change', () => {
      applyLinesMaster();
      calculate();
    });
  }

  document.addEventListener('change', (event) => {
    const target = event.target;
    if (target && target.id && target.id.startsWith('imo_')) {
      syncImoMaster();
    }
    if (target && target.id && target.id.startsWith('lines_')) {
      syncLinesMaster();
    }
  });

  restoreTugsState();
  window.addEventListener('beforeunload', saveTugsState);
  window.addEventListener('pagehide', saveTugsState);
}

window.addEventListener('afterprint', () => {
  document.body.classList.remove('print-fit');
  if (printRestoreDensity === 'comfortable') {
    setDensity('comfortable');
  } else if (printRestoreDensity === 'none') {
    document.body.classList.remove('density-comfortable', 'density-dense');
  }
  printRestoreDensity = null;
  clearPrintHidden();
});

window.addEventListener('beforeprint', () => {
  const logoLeftNote = document.getElementById('logoLeftNote');
  if (logoLeftNote) autoResizeTextarea(logoLeftNote);
  applyPrintDensity();
  updatePrintHidden();
});

window.addEventListener('DOMContentLoaded', () => {
  initIndex();
  initLightDues();
  initPortDues();
  initMooring();
  initTugs();
});
