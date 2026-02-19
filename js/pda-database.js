(() => {
  const FORM_PAGE = 'pda-split.html';
  const DASHBOARD_PAGE = 'index.html';

  const DB_STORAGE = {
    positions: 'pda_positions_database_v1',
    selected: 'pda_selected_position_id',
    indexState: 'pda_index_state',
    vesselName: 'pda_vessel_name',
    gt: 'pda_gt',
    quantity: 'pda_quantity'
  };

  let tableBody = null;
  let statusNode = null;

  function storageGet(key) {
    try {
      return localStorage.getItem(key);
    } catch (error) {
      return null;
    }
  }

  function storageSet(key, value) {
    try {
      localStorage.setItem(key, String(value));
    } catch (error) {
      // ignore storage failures
    }
  }

  function storageRemove(key) {
    try {
      localStorage.removeItem(key);
    } catch (error) {
      // ignore storage failures
    }
  }

  function readJson(raw, fallback) {
    if (!raw) return fallback;
    try {
      const parsed = JSON.parse(raw);
      return parsed == null ? fallback : parsed;
    } catch (error) {
      return fallback;
    }
  }

  function getPositions() {
    const parsed = readJson(storageGet(DB_STORAGE.positions), []);
    if (!Array.isArray(parsed)) return [];
    return parsed.filter((item) => item && typeof item === 'object' && typeof item.id === 'string' && item.id.trim());
  }

  function savePositions(positions) {
    storageSet(DB_STORAGE.positions, JSON.stringify(positions));
  }

  function makeId() {
    return `pda_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  }

  function readField(id) {
    const field = document.getElementById(id);
    if (!field) return '';
    return String(field.value || '').trim();
  }

  function writeField(id, value) {
    const field = document.getElementById(id);
    if (!field) return;
    field.value = value == null ? '' : String(value);
  }

  function formatSavedAt(savedAt) {
    if (!savedAt) return '-';
    const parsed = new Date(savedAt);
    if (Number.isNaN(parsed.getTime())) return '-';
    return parsed.toLocaleString();
  }

  function setStatus(message, isError) {
    if (!statusNode) return;
    statusNode.textContent = message || '';
    statusNode.classList.toggle('error', Boolean(isError));
  }

  function getUrl() {
    return new URL(window.location.href);
  }

  function getQueryValue(name) {
    return (getUrl().searchParams.get(name) || '').trim();
  }

  function isNewMode() {
    return getQueryValue('new') === '1';
  }

  function getCurrentPositionId() {
    const queryId = getQueryValue('pda');
    if (queryId) return queryId;
    return (storageGet(DB_STORAGE.selected) || '').trim();
  }

  function setCurrentPositionId(id) {
    const normalized = String(id || '').trim();
    const url = getUrl();

    if (normalized) {
      url.searchParams.set('pda', normalized);
      storageSet(DB_STORAGE.selected, normalized);
    } else {
      url.searchParams.delete('pda');
      storageRemove(DB_STORAGE.selected);
    }
    url.searchParams.delete('new');

    history.replaceState({}, '', `${url.pathname}${url.search}${url.hash}`);
  }

  function readMetadataFromForm() {
    return {
      date: readField('dateInput'),
      vesselName: readField('vesselNameIndex'),
      port: readField('portInput'),
      operation: readField('operationsInput'),
      agent: readField('agentInput')
    };
  }

  function snapshotIndexState() {
    if (typeof saveIndexState === 'function') {
      saveIndexState();
    }
    return readJson(storageGet(DB_STORAGE.indexState), null);
  }

  function syncSharedValuesFromForm() {
    const vesselName = readField('vesselNameIndex');
    const gt = readField('grossTonnage');
    const quantity = readField('quantityInput');

    if (vesselName) storageSet(DB_STORAGE.vesselName, vesselName);
    else storageRemove(DB_STORAGE.vesselName);

    if (gt) storageSet(DB_STORAGE.gt, gt);
    else storageRemove(DB_STORAGE.gt);

    if (quantity) storageSet(DB_STORAGE.quantity, quantity);
    else storageRemove(DB_STORAGE.quantity);
  }

  function buildRecord(existingId) {
    const metadata = readMetadataFromForm();
    return {
      id: existingId || makeId(),
      date: metadata.date,
      vesselName: metadata.vesselName,
      port: metadata.port,
      operation: metadata.operation,
      agent: metadata.agent,
      savedAt: new Date().toISOString(),
      indexState: snapshotIndexState()
    };
  }

  function renderPositionsTable() {
    if (!tableBody) return;

    const currentId = getCurrentPositionId();
    const positions = getPositions();
    tableBody.innerHTML = '';

    if (positions.length === 0) {
      const emptyRow = document.createElement('tr');
      const emptyCell = document.createElement('td');
      emptyCell.colSpan = 7;
      emptyCell.className = 'pda-db-empty';
      emptyCell.textContent = 'No PDA positions saved yet.';
      emptyRow.appendChild(emptyCell);
      tableBody.appendChild(emptyRow);
      return;
    }

    positions.forEach((position) => {
      const row = document.createElement('tr');
      if (currentId && position.id === currentId) {
        row.classList.add('pda-db-active');
      }

      const actionCell = document.createElement('td');
      const editButton = document.createElement('button');
      editButton.type = 'button';
      editButton.className = 'mini pda-db-edit-btn';
      editButton.textContent = 'Edit';
      editButton.dataset.action = 'edit';
      editButton.dataset.id = position.id;
      actionCell.appendChild(editButton);
      row.appendChild(actionCell);

      const dateCell = document.createElement('td');
      dateCell.textContent = position.date || '-';
      row.appendChild(dateCell);

      const vesselCell = document.createElement('td');
      const vesselName = document.createElement('div');
      vesselName.className = 'pda-db-vessel';
      vesselName.textContent = position.vesselName || 'Untitled PDA';
      vesselCell.appendChild(vesselName);

      const savedAt = document.createElement('div');
      savedAt.className = 'pda-db-saved';
      savedAt.textContent = `Saved: ${formatSavedAt(position.savedAt)}`;
      vesselCell.appendChild(savedAt);
      row.appendChild(vesselCell);

      const portCell = document.createElement('td');
      portCell.textContent = position.port || '-';
      row.appendChild(portCell);

      const operationCell = document.createElement('td');
      operationCell.textContent = position.operation || '-';
      row.appendChild(operationCell);

      const agentCell = document.createElement('td');
      agentCell.textContent = position.agent || '-';
      row.appendChild(agentCell);

      const deleteCell = document.createElement('td');
      const deleteButton = document.createElement('button');
      deleteButton.type = 'button';
      deleteButton.className = 'mini pda-db-delete-btn';
      deleteButton.textContent = 'Delete';
      deleteButton.dataset.action = 'delete';
      deleteButton.dataset.id = position.id;
      deleteCell.appendChild(deleteButton);
      row.appendChild(deleteCell);

      tableBody.appendChild(row);
    });
  }

  function setTodayDateIfPossible() {
    const dateInput = document.getElementById('dateInput');
    if (!dateInput) return;
    const now = new Date();
    const day = String(now.getDate());
    const month = String(now.getMonth() + 1);
    const year = String(now.getFullYear()).slice(-2);
    dateInput.value = `${day}/${month}/${year}`;
  }

  function clearFormForNewPda() {
    storageRemove(DB_STORAGE.indexState);
    storageRemove(DB_STORAGE.vesselName);
    storageRemove(DB_STORAGE.gt);
    storageRemove(DB_STORAGE.quantity);
    storageRemove(DB_STORAGE.selected);

    writeField('logoLeftNote', '');
    writeField('titleNote', '');
    writeField('vesselNameIndex', '');
    writeField('grossTonnage', '');
    writeField('lengthOverall', '');
    writeField('bowThrusterFitted', '');
    writeField('portInput', '');
    writeField('berthTerminal', '');
    writeField('operationsInput', '');
    writeField('cargoInput', '');
    writeField('quantityInput', '');
    writeField('agentInput', '');
    setTodayDateIfPossible();

    if (typeof saveIndexState === 'function') saveIndexState();
  }

  function focusVesselInput() {
    const vesselField = document.getElementById('vesselNameIndex');
    if (!vesselField) return;
    vesselField.focus();
    vesselField.select();
  }

  function restoreRecordToForm(record) {
    if (!record || typeof record !== 'object') return;

    if (record.indexState && typeof record.indexState === 'object') {
      storageSet(DB_STORAGE.indexState, JSON.stringify(record.indexState));
      if (typeof restoreIndexState === 'function') restoreIndexState();
    }

    writeField('dateInput', record.date || '');
    writeField('vesselNameIndex', record.vesselName || '');
    writeField('portInput', record.port || '');
    writeField('operationsInput', record.operation || '');
    writeField('agentInput', record.agent || '');

    syncSharedValuesFromForm();

    if (typeof decorateOutlayRows === 'function') decorateOutlayRows();
    if (typeof wrapMoneyFields === 'function') wrapMoneyFields();
    if (typeof decorateMoneyEditCells === 'function') decorateMoneyEditCells();
    if (typeof recalcOutlayTotals === 'function') recalcOutlayTotals();
    if (typeof refreshOutlayLayout === 'function') refreshOutlayLayout();

    const toggleSailing = document.getElementById('toggleSailing');
    if (toggleSailing && typeof setSailingVisible === 'function') {
      setSailingVisible(Boolean(toggleSailing.checked));
    }

    if (typeof updateTowageFromStorage === 'function') updateTowageFromStorage();
    if (typeof updateLightDuesFromStorage === 'function') updateLightDuesFromStorage();
    if (typeof updatePortDuesFromStorage === 'function') updatePortDuesFromStorage();
    if (typeof updatePilotageFromStorage === 'function') updatePilotageFromStorage();
    if (typeof updateMooringFromStorage === 'function') updateMooringFromStorage();

    if (typeof saveIndexState === 'function') saveIndexState();
  }

  function saveCurrentPosition() {
    const currentId = getCurrentPositionId();
    const positions = getPositions();
    const existingIndex = positions.findIndex((item) => item.id === currentId);
    const record = buildRecord(existingIndex >= 0 ? currentId : '');

    if (existingIndex >= 0) {
      positions.splice(existingIndex, 1);
    }
    positions.unshift(record);

    savePositions(positions);
    setCurrentPositionId(record.id);
    setStatus('PDA saved to database.', false);
  }

  function openDashboard() {
    window.location.href = DASHBOARD_PAGE;
  }

  function saveCurrentPositionAndOpenDashboard() {
    saveCurrentPosition();
    openDashboard();
  }

  function openNewDraftInForm() {
    storageRemove(DB_STORAGE.indexState);
    storageRemove(DB_STORAGE.vesselName);
    storageRemove(DB_STORAGE.gt);
    storageRemove(DB_STORAGE.quantity);
    storageRemove(DB_STORAGE.selected);
    window.location.href = `${FORM_PAGE}?new=1`;
  }

  function openRecordInForm(recordId) {
    const id = String(recordId || '').trim();
    if (!id) return;
    storageSet(DB_STORAGE.selected, id);
    window.location.href = `${FORM_PAGE}?pda=${encodeURIComponent(id)}`;
  }

  function deleteRecordFromDatabase(recordId) {
    const id = String(recordId || '').trim();
    if (!id) return;

    const positions = getPositions();
    const nextPositions = positions.filter((item) => item.id !== id);
    if (nextPositions.length === positions.length) return;

    savePositions(nextPositions);
    if (getCurrentPositionId() === id) {
      setCurrentPositionId('');
    }
    setStatus('PDA position deleted from database.', false);
    renderPositionsTable();
  }

  function handleTableClick(event) {
    const button = event.target.closest('button[data-action]');
    if (!button) return;
    const action = button.dataset.action;
    const recordId = button.dataset.id;

    if (action === 'edit') {
      openRecordInForm(recordId);
      return;
    }
    if (action === 'delete') {
      deleteRecordFromDatabase(recordId);
    }
  }

  function initDashboard() {
    if (!tableBody) return;

    const addBtn = document.getElementById('addPdaPositionBtn');
    if (addBtn) addBtn.addEventListener('click', openNewDraftInForm);

    tableBody.addEventListener('click', handleTableClick);
    renderPositionsTable();
  }

  function initFormPage() {
    const vesselField = document.getElementById('vesselNameIndex');
    if (!vesselField) return;

    const saveBtn = document.getElementById('savePdaPositionBtn');
    if (saveBtn) {
      saveBtn.addEventListener('click', saveCurrentPosition);
    }

    document.querySelectorAll('#openPdaDatabaseBtn, #openPdaDatabaseTopBtn').forEach((homeBtn) => {
      homeBtn.addEventListener('click', saveCurrentPositionAndOpenDashboard);
    });

    if (isNewMode()) {
      clearFormForNewPda();
      setStatus('New PDA ready. Enter Vessel Name.', false);
      focusVesselInput();
      return;
    }

    const currentId = getCurrentPositionId();
    if (!currentId) return;

    const positions = getPositions();
    const currentRecord = positions.find((item) => item.id === currentId);
    if (!currentRecord) {
      setStatus('Selected PDA record was not found.', true);
      return;
    }

    restoreRecordToForm(currentRecord);
    setCurrentPositionId(currentRecord.id);
    setStatus(`Opened PDA: ${currentRecord.vesselName || 'Untitled PDA'}.`, false);
  }

  function initPdaDatabase() {
    tableBody = document.getElementById('pdaDatabaseBody');
    statusNode = document.getElementById('pdaDatabaseStatus');

    initDashboard();
    initFormPage();
  }

  window.addEventListener('DOMContentLoaded', initPdaDatabase);
})();
