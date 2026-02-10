import * as XLSX from 'xlsx';

interface Contact {
  fecha: string;
  nombre: string;
  apellidos: string;
  telefono: string;
  bienvenida: boolean;
  comercial: string;
}

// State
let contacts: Contact[] = [];
let currentIndex = 0;
let copiedSet = new Set<number>();
let sentSet = new Set<number>();

// DOM Elements
const stepUpload = document.getElementById('step-upload')!;
const stepSheets = document.getElementById('step-sheets')!;
const stepMessaging = document.getElementById('step-messaging')!;
const dropZone = document.getElementById('drop-zone')!;
const fileInput = document.getElementById('file-input') as HTMLInputElement;
const uploadError = document.getElementById('upload-error')!;
const sheetsList = document.getElementById('sheets-list')!;
const btnReset = document.getElementById('btn-reset')!;
const btnPrev = document.getElementById('btn-prev') as HTMLButtonElement;
const btnNext = document.getElementById('btn-next') as HTMLButtonElement;
const progressText = document.getElementById('progress-text')!;
const progressBar = document.getElementById('progress-bar') as HTMLElement;
const skippedCount = document.getElementById('skipped-count')!;
const sentCount = document.getElementById('sent-count')!;
const contactName = document.getElementById('contact-name')!;
const contactPhone = document.getElementById('contact-phone')!;
const contactDate = document.getElementById('contact-date')!;
const contactCommercial = document.getElementById('contact-commercial')!;
const contactBadge = document.getElementById('contact-badge')!;
const contactList = document.getElementById('contact-list')!;
const messageTemplate = document.getElementById('message-template') as HTMLTextAreaElement;
const messagePreview = document.getElementById('message-preview')!;
const copyFeedback = document.getElementById('copy-feedback')!;
const allDone = document.getElementById('all-done')!;
const genderFemale = document.getElementById('gender-female') as HTMLButtonElement;
const genderMale = document.getElementById('gender-male') as HTMLButtonElement;
const btnToggleTemplate = document.getElementById('btn-toggle-template')!;
const templateEditor = document.getElementById('template-editor')!;
const templateChevron = document.getElementById('template-chevron')!;
const btnCopyPreview = document.getElementById('btn-copy-preview')!;
const toast = document.getElementById('toast')!;

let selectedGender: 'femenino' | 'masculino' = 'femenino';
let toastTimeout: ReturnType<typeof setTimeout> | null = null;
let workbook: XLSX.WorkBook | null = null;

// ==================== FILE HANDLING ====================

function showStep(step: 'upload' | 'sheets' | 'messaging') {
  stepUpload.classList.toggle('hidden', step !== 'upload');
  stepSheets.classList.toggle('hidden', step !== 'sheets');
  stepMessaging.classList.toggle('hidden', step !== 'messaging');
  btnReset.classList.toggle('hidden', step === 'upload');
  if (step !== 'upload') {
    btnReset.classList.add('flex');
  } else {
    btnReset.classList.remove('flex');
  }
}

function showError(msg: string) {
  uploadError.textContent = msg;
  uploadError.classList.remove('hidden');
}

function hideError() {
  uploadError.classList.add('hidden');
}

function handleFile(file: File) {
  hideError();

  if (!file.name.match(/\.xlsx?$/i)) {
    showError('Formato no valido. Sube un archivo .xlsx o .xls');
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target!.result as ArrayBuffer);
      workbook = XLSX.read(data, { type: 'array' });

      if (workbook.SheetNames.length === 0) {
        showError('El archivo no contiene hojas.');
        return;
      }

      if (workbook.SheetNames.length === 1) {
        loadSheet(workbook.SheetNames[0]);
      } else {
        showSheetSelector(workbook.SheetNames);
      }
    } catch {
      showError('Error al leer el archivo. Asegurate de que es un Excel valido.');
    }
  };
  reader.readAsArrayBuffer(file);
}

// Prevent browser from opening files dropped anywhere on the page
document.addEventListener('dragover', (e) => { e.preventDefault(); e.stopPropagation(); });
document.addEventListener('drop', (e) => { e.preventDefault(); e.stopPropagation(); });

// Click on drop zone opens file picker
dropZone.addEventListener('click', () => fileInput.click());

// Drag & drop on the drop zone
dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  e.stopPropagation();
  dropZone.classList.add('border-primary', 'bg-primary-light/30');
});

dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('border-primary', 'bg-primary-light/30');
});

dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  e.stopPropagation();
  dropZone.classList.remove('border-primary', 'bg-primary-light/30');
  const file = e.dataTransfer?.files[0];
  if (file) handleFile(file);
});

fileInput.addEventListener('change', () => {
  const file = fileInput.files?.[0];
  if (file) handleFile(file);
});

// ==================== SHEET SELECTOR ====================

function showSheetSelector(sheetNames: string[]) {
  showStep('sheets');
  sheetsList.innerHTML = '';

  sheetNames.forEach((name) => {
    const btn = document.createElement('button');
    btn.className =
      'w-full flex items-center gap-3 px-5 py-4 bg-white border border-slate-200 rounded-xl hover:border-blue-500 hover:bg-blue-50 transition-all text-left cursor-pointer group';
    btn.innerHTML = `
      <div class="w-10 h-10 bg-blue-100 rounded-lg flex items-center justify-center group-hover:bg-blue-200 transition-colors shrink-0">
        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="text-blue-600"><path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M10 13H8"/><path d="M16 17H8"/><path d="M16 13h-2"/></svg>
      </div>
      <div>
        <p class="font-semibold text-slate-800">${name}</p>
        <p class="text-xs text-slate-500">Hoja del Excel</p>
      </div>
    `;
    btn.addEventListener('click', () => loadSheet(name));
    sheetsList.appendChild(btn);
  });
}

// ==================== LOAD SHEET DATA ====================

function normalizeHeader(header: string): string {
  return header
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function loadSheet(sheetName: string) {
  if (!workbook) return;

  const sheet = workbook.Sheets[sheetName];
  const rawData = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' });

  if (rawData.length === 0) {
    showStep('upload');
    showError('La hoja seleccionada esta vacia.');
    return;
  }

  // Map columns flexibly
  const headers = Object.keys(rawData[0]);
  const headerMap: Record<string, string> = {};

  for (const h of headers) {
    const norm = normalizeHeader(h);
    if (norm.includes('fecha')) headerMap['fecha'] = h;
    else if (norm.includes('apellido')) headerMap['apellidos'] = h;
    else if (norm.includes('nombre')) headerMap['nombre'] = h;
    else if (norm.includes('telefono') || norm.includes('tel')) headerMap['telefono'] = h;
    else if (norm.includes('bienvenida') || norm.includes('mensaje')) headerMap['bienvenida'] = h;
    else if (norm.includes('comercial')) headerMap['comercial'] = h;
  }

  // Validate required columns
  const requiredColumns: { key: string; label: string }[] = [
    { key: 'nombre', label: 'Nombre' },
    { key: 'telefono', label: 'Telefono' },
  ];
  const missingColumns = requiredColumns.filter((col) => !headerMap[col.key]);

  if (missingColumns.length > 0) {
    const missingLabels = missingColumns.map((c) => c.label).join(', ');
    const foundLabels = Object.keys(headerMap).length > 0
      ? Object.keys(headerMap).map((k) => headerMap[k]).join(', ')
      : 'ninguna reconocida';
    showStep('upload');
    showError(
      `El archivo no tiene las columnas requeridas: ${missingLabels}. ` +
      `Columnas encontradas: ${headers.join(', ')}. ` +
      `Columnas reconocidas: ${foundLabels}. ` +
      `Asegurate de que el Excel contiene al menos: fecha, nombre, apellidos, telefono, mensaje bienvenida, comercial.`
    );
    return;
  }

  contacts = rawData.map((row) => {
    const bienvenidaVal = row[headerMap['bienvenida'] || ''];
    let bienvenida = false;
    if (typeof bienvenidaVal === 'boolean') {
      bienvenida = bienvenidaVal;
    } else if (typeof bienvenidaVal === 'string') {
      bienvenida =
        bienvenidaVal.toLowerCase() === 'si' ||
        bienvenidaVal.toLowerCase() === 'sí' ||
        bienvenidaVal === '1' ||
        bienvenidaVal.toLowerCase() === 'true' ||
        bienvenidaVal.toLowerCase() === 'x';
    } else if (typeof bienvenidaVal === 'number') {
      bienvenida = bienvenidaVal === 1;
    }

    const fechaRaw = row[headerMap['fecha'] || ''];
    let fecha = String(fechaRaw || '-');
    if (typeof fechaRaw === 'number') {
      try {
        const dateObj = XLSX.SSF.parse_date_code(fechaRaw);
        fecha = `${String(dateObj.d).padStart(2, '0')}/${String(dateObj.m).padStart(2, '0')}/${dateObj.y}`;
      } catch {
        fecha = String(fechaRaw);
      }
    }

    return {
      fecha,
      nombre: String(row[headerMap['nombre'] || ''] || '-'),
      apellidos: String(row[headerMap['apellidos'] || ''] || ''),
      telefono: String(row[headerMap['telefono'] || ''] || '-'),
      bienvenida,
      comercial: String(row[headerMap['comercial'] || ''] || '-'),
    };
  });

  if (contacts.length === 0) {
    showStep('upload');
    showError('No se encontraron contactos en la hoja.');
    return;
  }

  currentIndex = 0;
  copiedSet = new Set();
  sentSet = new Set();
  showStep('messaging');
  renderContactList();
  renderCurrent();
}

// ==================== MESSAGING INTERFACE ====================

function getPersonalizedMessage(contact: Contact): string {
  const template = messageTemplate.value;
  const fullName = [contact.nombre, contact.apellidos].filter(Boolean).join(' ');
  const bienvenidaText = selectedGender === 'masculino' ? 'Bienvenido' : 'Bienvenida';
  return template
    .replace(/\{bienvenida\}/gi, bienvenidaText)
    .replace(/\{nombre\}/gi, contact.nombre)
    .replace(/\{apellidos\}/gi, contact.apellidos)
    .replace(/\{nombre_completo\}/gi, fullName)
    .replace(/\{telefono\}/gi, contact.telefono)
    .replace(/\{fecha\}/gi, contact.fecha)
    .replace(/\{comercial\}/gi, contact.comercial);
}

function renderCurrent() {
  if (contacts.length === 0) return;

  const contact = contacts[currentIndex];

  // Contact info
  const fullName = [contact.nombre, contact.apellidos].filter(Boolean).join(' ');
  contactName.textContent = fullName || '-';
  contactPhone.textContent = contact.telefono;
  contactDate.textContent = contact.fecha;
  contactCommercial.textContent = contact.comercial;

  // Badge
  if (contact.bienvenida) {
    contactBadge.className =
      'mb-4 px-3 py-1.5 rounded-full text-xs font-semibold inline-flex items-center gap-1.5 bg-success-light text-green-800';
    contactBadge.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg> Bienvenida enviada`;
    contactBadge.classList.remove('hidden');
  } else if (sentSet.has(currentIndex)) {
    contactBadge.className =
      'mb-4 px-3 py-1.5 rounded-full text-xs font-semibold inline-flex items-center gap-1.5 bg-success-light text-green-800';
    contactBadge.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg> Enviado`;
    contactBadge.classList.remove('hidden');
  } else if (copiedSet.has(currentIndex)) {
    contactBadge.className =
      'mb-4 px-3 py-1.5 rounded-full text-xs font-semibold inline-flex items-center gap-1.5 bg-primary-light text-blue-800';
    contactBadge.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg> Copiado`;
    contactBadge.classList.remove('hidden');
  } else {
    contactBadge.classList.add('hidden');
  }

  // Message preview
  messagePreview.textContent = getPersonalizedMessage(contact);

  // Navigation buttons
  btnPrev.disabled = currentIndex === 0;
  btnNext.disabled = currentIndex === contacts.length - 1;

  // Progress
  const totalSent = contacts.filter((c, i) => c.bienvenida || sentSet.has(i)).length;
  const totalCopied = copiedSet.size;
  const totalPending = contacts.filter((_c, i) => !contacts[i].bienvenida && !sentSet.has(i) && !copiedSet.has(i)).length;

  progressText.textContent = `${currentIndex + 1} de ${contacts.length}`;
  progressBar.style.width = `${((currentIndex + 1) / contacts.length) * 100}%`;

  if (totalSent > 0) {
    sentCount.textContent = `Enviados: ${totalSent}/${contacts.length}`;
  } else {
    sentCount.textContent = '';
  }

  if (totalCopied > 0) {
    skippedCount.textContent = `${totalCopied} copiados en esta sesion`;
  } else {
    skippedCount.textContent = '';
  }

  // All done
  if (totalPending === 0 && contacts.length > 0) {
    allDone.classList.remove('hidden');
  } else {
    allDone.classList.add('hidden');
  }

  // Update list highlight
  updateListHighlight();

  // Reset copy button state
  resetCopyButton();
}

function renderContactList() {
  contactList.innerHTML = '';

  contacts.forEach((contact, index) => {
    const item = document.createElement('button');
    const fullName = [contact.nombre, contact.apellidos].filter(Boolean).join(' ');

    let statusDot = '';
    if (contact.bienvenida) {
      statusDot = '<span class="w-2 h-2 rounded-full bg-green-500 shrink-0"></span>';
    } else if (sentSet.has(index)) {
      statusDot = '<span class="w-2 h-2 rounded-full bg-purple-500 shrink-0"></span>';
    } else if (copiedSet.has(index)) {
      statusDot = '<span class="w-2 h-2 rounded-full bg-blue-500 shrink-0"></span>';
    } else {
      statusDot = '<span class="w-2 h-2 rounded-full bg-slate-300 shrink-0"></span>';
    }

    item.className =
      'w-full flex items-center gap-3 px-4 py-3 text-left hover:bg-slate-50 transition-colors cursor-pointer text-sm contact-item';
    item.dataset.index = String(index);
    item.innerHTML = `
      ${statusDot}
      <div class="min-w-0 flex-1">
        <p class="font-medium text-slate-800 truncate">${fullName || '-'}</p>
        <p class="text-xs text-slate-500">${contact.telefono}</p>
      </div>
    `;

    item.addEventListener('click', () => {
      currentIndex = index;
      renderCurrent();
    });

    contactList.appendChild(item);
  });
}

function updateListHighlight() {
  const items = contactList.querySelectorAll('.contact-item');
  items.forEach((item) => {
    const el = item as HTMLElement;
    const idx = parseInt(el.dataset.index || '0');
    if (idx === currentIndex) {
      el.classList.add('bg-primary-light/50');
      el.classList.remove('hover:bg-slate-50');
      el.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    } else {
      el.classList.remove('bg-primary-light/50');
      el.classList.add('hover:bg-slate-50');
    }

    // Update status dots
    const dot = el.querySelector('span');
    if (dot) {
      dot.className = 'w-2 h-2 rounded-full shrink-0';
      if (contacts[idx].bienvenida) {
        dot.classList.add('bg-green-500');
      } else if (sentSet.has(idx)) {
        dot.classList.add('bg-purple-500');
      } else if (copiedSet.has(idx)) {
        dot.classList.add('bg-blue-500');
      } else {
        dot.classList.add('bg-slate-300');
      }
    }
  });
}

// ==================== TOAST ====================

function showToast() {
  // Clear any existing timeout so rapid clicks don't leave stale timers
  if (toastTimeout) {
    clearTimeout(toastTimeout);
    toastTimeout = null;
  }

  // Reset animation state
  toast.classList.remove('toast-enter', 'toast-exit', 'hidden');
  toast.classList.add('flex');

  // Force reflow to restart animation if toast was already visible
  toast.offsetHeight;

  toast.classList.add('toast-enter');

  toastTimeout = setTimeout(() => {
    toast.classList.remove('toast-enter');
    toast.classList.add('toast-exit');

    toast.addEventListener('animationend', () => {
      toast.classList.remove('flex', 'toast-exit');
      toast.classList.add('hidden');
    }, { once: true });
  }, 2500);
}

// ==================== ACTIONS ====================

function resetCopyButton() {
  copyFeedback.classList.add('hidden');
  copyFeedback.classList.remove('flex');
}

async function copyMessage() {
  const contact = contacts[currentIndex];
  const message = getPersonalizedMessage(contact);

  try {
    await navigator.clipboard.writeText(message);
    copiedSet.add(currentIndex);

    // Show header feedback
    copyFeedback.classList.remove('hidden');
    copyFeedback.classList.add('flex');

    // Show toast notification
    showToast();

    // Update list
    updateListHighlight();
    renderCurrent();
  } catch {
    // Fallback for older browsers
    const textarea = document.createElement('textarea');
    textarea.value = message;
    textarea.style.position = 'fixed';
    textarea.style.opacity = '0';
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand('copy');
    document.body.removeChild(textarea);
    copiedSet.add(currentIndex);

    // Show toast notification
    showToast();

    updateListHighlight();
    renderCurrent();
  }
}

// Event listeners
btnPrev.addEventListener('click', () => {
  if (currentIndex > 0) {
    currentIndex--;
    renderCurrent();
  }
});

btnNext.addEventListener('click', () => {
  if (currentIndex < contacts.length - 1) {
    // Mark the current contact as sent before advancing
    sentSet.add(currentIndex);
    currentIndex++;
    renderCurrent();
  }
});

btnReset.addEventListener('click', () => {
  contacts = [];
  currentIndex = 0;
  copiedSet = new Set();
  sentSet = new Set();
  workbook = null;
  fileInput.value = '';
  showStep('upload');
});

// Update message preview when template changes
messageTemplate.addEventListener('input', () => {
  if (contacts.length > 0) {
    messagePreview.textContent = getPersonalizedMessage(contacts[currentIndex]);
  }
});

// ==================== GENDER TOGGLE ====================

function setGender(gender: 'femenino' | 'masculino') {
  selectedGender = gender;

  const activeClasses = ['bg-primary', 'text-white', 'shadow-sm'];
  const inactiveClasses = ['text-muted', 'hover:text-slate-700'];

  if (gender === 'femenino') {
    genderFemale.classList.add(...activeClasses);
    genderFemale.classList.remove(...inactiveClasses);
    genderFemale.setAttribute('aria-checked', 'true');
    genderMale.classList.remove(...activeClasses);
    genderMale.classList.add(...inactiveClasses);
    genderMale.setAttribute('aria-checked', 'false');
  } else {
    genderMale.classList.add(...activeClasses);
    genderMale.classList.remove(...inactiveClasses);
    genderMale.setAttribute('aria-checked', 'true');
    genderFemale.classList.remove(...activeClasses);
    genderFemale.classList.add(...inactiveClasses);
    genderFemale.setAttribute('aria-checked', 'false');
  }

  // Update preview if contacts are loaded
  if (contacts.length > 0) {
    messagePreview.textContent = getPersonalizedMessage(contacts[currentIndex]);
  }
}

genderFemale.addEventListener('click', () => setGender('femenino'));
genderMale.addEventListener('click', () => setGender('masculino'));

// ==================== TEMPLATE EDITOR TOGGLE ====================

btnToggleTemplate.addEventListener('click', () => {
  const isHidden = templateEditor.classList.contains('hidden');

  if (isHidden) {
    templateEditor.classList.remove('hidden');
    templateChevron.style.transform = 'rotate(180deg)';
    btnToggleTemplate.setAttribute('aria-expanded', 'true');
    // Remove rounded bottom from button when editor is open
    btnToggleTemplate.classList.remove('rounded-xl');
    btnToggleTemplate.classList.add('rounded-t-xl');
  } else {
    templateEditor.classList.add('hidden');
    templateChevron.style.transform = 'rotate(0deg)';
    btnToggleTemplate.setAttribute('aria-expanded', 'false');
    // Restore full rounding when editor is closed
    btnToggleTemplate.classList.remove('rounded-t-xl');
    btnToggleTemplate.classList.add('rounded-xl');
  }
});

// ==================== COPY PREVIEW BUTTON ====================

btnCopyPreview.addEventListener('click', copyMessage);

// Keyboard shortcuts
document.addEventListener('keydown', (e) => {
  if (stepMessaging.classList.contains('hidden')) return;
  if (document.activeElement === messageTemplate) return;

  if (e.key === 'ArrowLeft' && !btnPrev.disabled) {
    btnPrev.click();
  } else if (e.key === 'ArrowRight' && !btnNext.disabled) {
    btnNext.click();
  } else if (e.key === 'c' || e.key === 'C') {
    if (!e.ctrlKey && !e.metaKey) {
      copyMessage();
    }
  }
});
