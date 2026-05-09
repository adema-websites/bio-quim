/**
 * script.js
 * =========
 * Carga datos desde Google Sheets publicado (pubhtml) → lee configuración (brand, colores, placeholder), emojis de categoría,
 * parsea productos (columnas obligatorias: name, description, category, price, images),
 * convierte columnas adicionales en variantes dinámicas,
 * renderiza catálogo filtrable y buscable,
 * muestra modal de detalle con carousel de imágenes + variantes,
 * administra carrito dinámico (add/remove) → genera pedido formateado para WhatsApp.
 */

// URL base del Google Sheets publicado
const GOOGLE_SHEETS_BASE = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vT-gY5ZFYqOKI9jtxCHEkpG1279g4qQA8YvhLuSzNM5pcE_IQ8itaNPm_7yEMU_-R299ItDgoz79W-6';

/**
 * Extrae el texto de una celda de la tabla pubhtml de Google Sheets.
 * Algunas celdas usan <div class="softmerge-inner"> para el contenido real.
 * @param {HTMLElement} cell
 * @returns {string}
 */
function getCellText(cell) {
  const inner = cell.querySelector('.softmerge-inner');
  return (inner || cell).textContent.trim();
}

/**
 * Parsea la tabla <tbody> de un pubhtml/sheet de Google Sheets.
 * Estructura real: cada <tr> tiene un <th> (número de fila) seguido de <td> de datos.
 * La primera fila de datos es el encabezado de columnas.
 * @param {Document} doc - documento HTML ya parseado de la hoja individual
 * @returns {Array<Object>}
 */
function parseSheetDoc(doc) {
  const rows = Array.from(doc.querySelectorAll('tbody tr'));
  if (rows.length < 2) return [];

  // Primera fila → encabezados (saltar el primer <th> = número de fila)
  const headerCells = Array.from(rows[0].querySelectorAll('td'));
  const headers = headerCells.map(getCellText);

  // Resto de filas → datos
  return rows.slice(1).map(row => {
    const cells = Array.from(row.querySelectorAll('td'));
    const obj = {};
    headers.forEach((header, i) => {
      if (!header) return;
      const val = cells[i] ? getCellText(cells[i]) : '';
      if (val !== '') obj[header] = val;
    });
    return obj;
  }).filter(obj => Object.keys(obj).length > 0);
}

/**
 * Obtiene el Google Sheets publicado y construye un workbook compatible.
 * Paso 1: fetch del pubhtml principal → extrae nombres de hoja y GIDs vía regex en el JS inline.
 * Paso 2: fetch individual de cada hoja en pubhtml/sheet?headers=false&gid={gid}.
 * wb.SheetNames → array de nombres de hoja
 * wb.Sheets[nombre] → array de objetos de fila (ya parseados)
 * @returns {Object} workbook
 */
async function fetchGoogleSheets() {
  // 1. Obtener lista de hojas y sus GIDs desde el pubhtml principal
  const indexResp = await fetch(`${GOOGLE_SHEETS_BASE}/pubhtml`);
  if (!indexResp.ok) throw new Error(`Error al obtener el índice de hojas: ${indexResp.status}`);
  const indexHtml = await indexResp.text();

  // El JS inline contiene: items.push({name: "Nombre", ..., gid: "12345", ...})
  const sheetList = [];
  const itemRegex = /items\.push\(\{name:\s*"([^"]+)"[^}]+gid:\s*"([^"]+)"/g;
  let match;
  while ((match = itemRegex.exec(indexHtml)) !== null) {
    sheetList.push({ name: match[1], gid: match[2] });
  }

  if (sheetList.length === 0) throw new Error('No se encontraron hojas en el Google Sheets publicado.');

  // 2. Fetch y parseo de cada hoja individualmente
  const wb = { SheetNames: [], Sheets: {} };
  const parser = new DOMParser();

  await Promise.all(sheetList.map(async ({ name, gid }) => {
    const url = `${GOOGLE_SHEETS_BASE}/pubhtml/sheet?headers=false&gid=${gid}`;
    const resp = await fetch(url);
    if (!resp.ok) {
      console.warn(`No se pudo cargar la hoja "${name}" (gid=${gid}): ${resp.status}`);
      return;
    }
    const html = await resp.text();
    const doc = parser.parseFromString(html, 'text/html');
    wb.SheetNames.push(name);
    wb.Sheets[name] = parseSheetDoc(doc);
  }));

  return wb;
}

// Elementos DOM principales
const loading = document.getElementById('loading-indicator');
const resetFilterBtn = document.getElementById('reset-filter');
const productsContainer = document.getElementById('products-container');
const logoEl = document.querySelector('.logo');
const categoryFilters = document.getElementById('category-filters');
const sendBtn = document.getElementById('send-whatsapp');
sendBtn.onclick = sendOrder;

const openCartBtn = document.getElementById('open-cart');
const emptyCartMessage = document.getElementById('empty-cart-message');
const fallbackWhatsAppNumber = (
  document.getElementById('cart-send')?.dataset?.whatsapp ||
  sendBtn.dataset.whatsapp ||
  '5491123456789'
).replace(/\D/g, '');


// Variables globales
let products = [], 
    filtered = [], 
    emojis = {}, 
    cart = JSON.parse(localStorage.getItem('cart') || '[]'), 
    currentCategory = null;
let currentIndex = 0, 
    currentList = [];

let clients = [];

function normalizeWhatsAppPhone(value) {
  const raw = String(value || '').trim();
  if (!raw || /e\+?/i.test(raw)) return '';
  const digits = raw.replace(/\D/g, '');
  return digits.length >= 10 ? digits : '';
}

function getWhatsAppPhone() {
  return normalizeWhatsAppPhone(sendBtn.dataset.whatsapp) || fallbackWhatsAppNumber;
}

const b2bConsultFlows = {
  lavadero: {
    kicker: 'Flujo lavadero / laverrap',
    title: 'Abastecimiento mayorista para lavaderos y laverraps',
    image: 'images/laverrap.png',
    description: 'Trabajamos con lavaderos, laverraps y negocios que necesitan insumos constantes, precios por volumen y entrega directa. Te asesoramos para armar una compra práctica según tu consumo real.',
    points: [
      'Venta por mayor en bidones y presentaciones para uso profesional.',
      'Entregas directas a comercios, con rutas programadas en Zona Oeste y CABA.',
      'Fragancias, suavizantes, jabones, detergentes y productos de limpieza para el día a día del local.',
      'Atención comercial para consultar precios, reposición y frecuencia de entrega.'
    ],
    cta: 'Consultar precios para lavadero',
    whatsappMessage: 'Hola! Vengo desde la consulta para LAVADERO/LAVERRAP de la web. Quiero consultar precios mayoristas, entregas a negocio y productos recomendados para mi local.'
  },
  gastro: {
    kicker: 'Flujo gastronomía / comercio',
    title: 'Insumos por mayor para cocinas, bares y comercios',
    image: 'images/cocina.png',
    description: 'Proveemos productos de limpieza profesional para negocios gastronómicos y comercios que buscan comprar directo, resolver la reposición y mantener stock sin depender de compras sueltas.',
    points: [
      'Desengrasantes, detergentes, lavandina, sanitizantes y productos de limpieza de uso intensivo.',
      'Venta mayorista para bares, restaurantes, cocinas, locales y empresas.',
      'Entrega directa al negocio y asesoramiento según rubro, frecuencia y volumen.',
      'Presupuesto por WhatsApp con identificación de este flujo para responder más rápido.'
    ],
    cta: 'Pedir presupuesto mayorista',
    whatsappMessage: 'Hola! Vengo desde la consulta para GASTRONOMIA/COMERCIOS de la web. Quiero pedir un presupuesto mayorista y consultar entregas directas a mi negocio.'
  }
};
/**
 * Inicialización: enlaza eventos y carga el Excel
 */
document.addEventListener('DOMContentLoaded', init);

async function init() {
  // Configurar event listeners
  resetFilterBtn.onclick = () => applyFilter(null);
  initB2BConsults();
  
  const searchInput = document.getElementById('search-input');
  if (searchInput) {
    searchInput.oninput = searchProducts;
  }
  
  sendBtn.onclick = sendOrder;
  openCartBtn.onclick = () => {
    renderCart();
    document.getElementById('cart-modal').classList.remove('hidden');
  };
  
  // Mobile menu toggle
  const mobileMenuToggle = document.getElementById('mobile-menu-toggle');
  const navMenu = document.getElementById('nav-menu');
  
  mobileMenuToggle.addEventListener('click', () => {
    navMenu.classList.toggle('active');
    mobileMenuToggle.classList.toggle('active');
  });

  // Close mobile menu when clicking on nav links
  const navLinks = document.querySelectorAll('.nav-link');
  navLinks.forEach(link => {
    link.addEventListener('click', () => {
      navMenu.classList.remove('active');
      mobileMenuToggle.classList.remove('active');
    });
  });

  // Cart button in nav
  const cartNavBtn = document.getElementById('cart-nav-btn');
  cartNavBtn.onclick = () => {
    renderCart();
    document.getElementById('cart-modal').classList.remove('hidden');
  };

  // Cart send button
  const cartSendBtn = document.getElementById('cart-send');
  if (cartSendBtn) {
    cartSendBtn.onclick = () => sendOrder();
  }

  // Clear cart button
  const clearCartBtn = document.getElementById('clear-cart');
  if (clearCartBtn) {
    clearCartBtn.onclick = clearCart;
  }
  
  // Smooth scrolling for nav links
  navLinks.forEach(link => {
    link.addEventListener('click', (e) => {
      e.preventDefault();
      const targetId = link.getAttribute('href').substring(1);
      const targetElement = document.getElementById(targetId);
      if (targetElement) {
        targetElement.scrollIntoView({
          behavior: 'smooth',
          block: 'start'
        });
      }
    });
  });
  
  // Cerrar modales con Escape
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      document.querySelectorAll('.modal').forEach(modal => modal.classList.add('hidden'));
    }
  });
  
  // Cerrar modales al hacer clic en backdrop
  document.querySelectorAll('.modal-backdrop').forEach(el => {
    el.addEventListener('click', () => {
      el.closest('.modal').classList.add('hidden');
    });
  });
  
  // Cerrar modales con botones de cierre
  document.querySelectorAll('.modal-close').forEach(btn => {
    btn.addEventListener('click', () => {
      btn.closest('.modal').classList.add('hidden');
    });
  });
  
  // Cargar datos
  await loadExcel();
  loadPromociones();

  // Agregar event listener para PDF después de cargar productos
  document.getElementById('download-pdf').addEventListener('click', generatePriceListPDF);
}

function initB2BConsults() {
  document.querySelectorAll('[data-b2b-flow]').forEach(card => {
    card.addEventListener('click', () => openB2BConsultModal(card.dataset.b2bFlow));
  });
}

function openB2BConsultModal(flowKey) {
  const flow = b2bConsultFlows[flowKey];
  const modal = document.getElementById('b2b-modal');
  if (!flow || !modal) return;

  document.getElementById('b2b-modal-kicker').textContent = flow.kicker;
  document.getElementById('b2b-modal-title').textContent = flow.title;
  document.getElementById('b2b-modal-description').textContent = flow.description;

  const media = document.getElementById('b2b-modal-media');
  media.style.backgroundImage = `linear-gradient(180deg, rgba(8, 20, 16, 0.1), rgba(8, 20, 16, 0.36)), url('${flow.image}')`;

  const points = document.getElementById('b2b-modal-points');
  points.innerHTML = flow.points.map(point => `<li>${point}</li>`).join('');

  const phone = getWhatsAppPhone();
  const whatsapp = document.getElementById('b2b-modal-whatsapp');
  whatsapp.href = `https://wa.me/${phone}?text=${encodeURIComponent(flow.whatsappMessage)}`;
  whatsapp.innerHTML = `<i class="fab fa-whatsapp"></i> ${flow.cta}`;

  modal.classList.remove('hidden');
}
/**
 * Lee la hoja "promociones" del Google Sheets y renderiza dinámicamente
 * los combos en #promos-container. Si la hoja no existe o está vacía,
 * se mantienen los combos estáticos definidos en el HTML.
 * Columnas esperadas: nombre, items, precio, mensaje_wa
 */
async function loadPromociones() {
  try {
    const wb = await fetchGoogleSheets();
    const sheetName = wb.SheetNames.find(n => /promociones/i.test(n));
    if (!sheetName) return; // sin hoja → mantener HTML estático

    const rows = wb.Sheets[sheetName];
    if (!rows || rows.length === 0) return;

    const container = document.getElementById('promos-container');
    if (!container) return;

    // Limpiar combos estáticos
    container.innerHTML = '';

    const isPremium = (nombre) => /premium/i.test(nombre);

    rows.forEach((row, idx) => {
      const nombre  = row['nombre']    || row['Nombre']   || `Combo ${idx + 1}`;
      const items   = row['items']     || row['Items']     || '';
      const precio  = row['precio']    || row['Precio']    || '';
      const msgWa   = row['mensaje_wa']|| row['Mensaje WA']||
                      `Hola! Quiero pedir el ${nombre} de la web.`;
      const waNumber = getWhatsAppPhone();
      const waUrl   = `https://wa.me/${waNumber}?text=${encodeURIComponent(msgWa)}`;

      const premium = isPremium(nombre);
      const itemsList = items.split(',').map(i => i.trim()).filter(Boolean);

      const card = document.createElement('div');
      card.className = `promo-card${premium ? ' promo-card--premium' : ''}`;
      card.innerHTML = `
        <div class="promo-badge${premium ? ' promo-badge--premium' : ''}">${nombre}</div>
        <ul class="promo-items">
          ${itemsList.map(item => `<li><i class="fas fa-check-circle"></i> ${item}</li>`).join('')}
        </ul>
        <p class="promo-price">${precio}</p>
        <a class="btn-promo-wa${premium ? ' btn-promo-wa--premium' : ''}"
           href="${waUrl}" target="_blank" rel="noopener">
          <i class="fab fa-whatsapp"></i> Pedir por WhatsApp
        </a>`;
      container.appendChild(card);
    });
  } catch (err) {
    console.warn('No se pudieron cargar promociones dinámicas. Se usan los combos estáticos.', err);
  }
}

/**
 * Carga y parsea los datos desde Google Sheets
 */
async function loadExcel() {
  loading.style.display = 'flex';
  try {
    const wb = await fetchGoogleSheets();

    // Aplicar configuración y cargar datos
    applyConfig(wb);
    loadEmojis(wb);
    loadClients(wb);
    populateClientSelect();

    products = parseProducts(wb);
    filtered = [...products];

    // Renderizar interfaz
    renderCategories();
    renderProducts();
    categoryFilters.style.display = 'block';
    updateCartBadge();
  } catch (err) {
    console.error('Error al cargar Google Sheets:', err);
    alert('Error al cargar los productos: ' + err.message);
  } finally {
    loading.style.display = 'none';
  }
}


function applyConfig(wb) {

  const cfgSheet = wb.SheetNames.find(n => /configuracion/i.test(n));
  if (!cfgSheet) return;
  const cfg = (wb.Sheets[cfgSheet] || [])[0] || {};
  
  const configuredWhatsApp = normalizeWhatsAppPhone(cfg.WhatsAppNumber);
  if (configuredWhatsApp) {
    sendBtn.dataset.whatsapp = configuredWhatsApp;
  }
  
  // Brand name (texto) dentro de <h1 class="logo"><a>
  const brandLink = document.querySelector('.brand .logo a');
  if (cfg.BrandName) brandLink.textContent = cfg.BrandName;

  // El hero principal es contenido editorial/SEO y queda definido en el HTML.
  // No se pisa desde Google Sheets para evitar cambios accidentales de posicionamiento.
  
  // Configurar LogoURL - aplicar a la logo-icon y mantener texto
  if (cfg.LogoURL) {
    const logoIcon = document.querySelector('.logo-icon');
    if (logoIcon) {
      logoIcon.innerHTML = `<img src="${cfg.LogoURL}" alt="${cfg.BrandName || 'Logo'}" />`;
    }
  }
  
  const searchInput = document.getElementById('search-input');
  if (cfg.SearchPlaceholder && searchInput) searchInput.placeholder = cfg.SearchPlaceholder;
  
  // Configurar colores
  const root = document.documentElement.style;
  if (cfg.PrimaryColor) root.setProperty('--primary', cfg.PrimaryColor);
  if (cfg.PrimaryDarkColor) root.setProperty('--primary-dark', cfg.PrimaryDarkColor);
  if (cfg.SecondaryColor) root.setProperty('--secondary', cfg.SecondaryColor);
  if (cfg.AccentColor) root.setProperty('--accent', cfg.AccentColor);
}

/**
 * Carga emojis de categorías desde hoja "categoryemojis"
 * @param {Object} wb - Workbook de Excel
 */
function loadEmojis(wb) {
  const sheet = wb.SheetNames.find(n => /categoryemojis/i.test(n));
  if (!sheet) return;
  
  (wb.Sheets[sheet] || []).forEach(r => {
    if (r.Category && r.Emoji) emojis[r.Category] = r.Emoji;
  });
}

/**
 * Parsea hoja "productos" a array de objetos con variantes dinámicas
 * @param {Object} wb - Workbook de Excel
 * @returns {Array} Array de productos
 */
function parseProducts(wb) {
  const sheetName = wb.SheetNames.find(n => /productos/i.test(n));
  if (!sheetName) throw new Error('Hoja "Productos" no encontrada.');
  
  const rows = wb.Sheets[sheetName] || [];
  const reserved = ['id','Id','name','Name','description','Description','category','Category','price','Price','images','Images'];

  return rows.map(r => {
    // Procesar imágenes
    const images = String(r.images || r.Images || '')
      .split(',')
      .map(i => i.trim())
      .filter(Boolean);
    
    // Si no hay imágenes, usar placeholder
    if (images.length === 0) {
      images.push('images/placeholder.svg');
    }

    // Procesar variantes dinámicas
    const variants = {};
    Object.entries(r).forEach(([key, val]) => {
      if (!reserved.includes(key) && val) {
        variants[key] = String(val).split(',').map(x => x.trim());
      }
    });

    // Retornar objeto de producto normalizado
    return {
      id: String(r.id || r.Id || Date.now()),
      name: r.name || r.Name || 'Producto sin nombre',
      description: r.description || r.Description || '',
      category: r.category || r.Category || 'Sin categoría',
      price: parseFloat(r.price || r.Price) || 0,
      images,
      variants
    };
  });
}

/**
 * Carga clientes desde hoja "clientes"
 * @param {Object} wb - Workbook de Excel
 */
function loadClients(wb) {
  const sheet = wb.SheetNames.find(n => /clientes/i.test(n));
  if (!sheet) return;
  clients = wb.Sheets[sheet] || [];
}

/**
 * Renderiza dropdown de categoría
 */
function renderCategories() {
  const categorySelect = document.getElementById('category-select');
  categorySelect.innerHTML = '<option value="">Todas las categorías</option>';

  // Obtener categorías únicas
  const categories = [...new Set(products.map(p => p.category))];

  // Crear opciones para cada categoría
  categories.forEach(cat => {
    const option = document.createElement('option');
    option.value = cat;
    option.textContent = `${emojis[cat] || '🏷️'} ${cat}`;
    option.selected = cat === currentCategory;
    categorySelect.appendChild(option);
  });

  // Configurar event listener para el select
  categorySelect.onchange = () => applyFilter(categorySelect.value || null);
}

/**
 * Aplica filtro por categoría y re-renderiza productos
 * @param {string|null} cat - Categoría a filtrar, null para mostrar todas
 */
function applyFilter(cat) {
  currentCategory = cat;
  filtered = cat ? products.filter(p => p.category === cat) : [...products];
  resetFilterBtn.style.display = cat ? 'inline-flex' : 'none';
  renderProducts();
}

/**
 * Filtrado en vivo por búsqueda
 */
function searchProducts() {
  const searchInput = document.getElementById('search-input');
  const q = searchInput ? searchInput.value.trim().toLowerCase() : '';
  
  // Filtrar productos por búsqueda y categoría actual
  filtered = products.filter(p =>
    (!currentCategory || p.category === currentCategory) &&
    (p.name.toLowerCase().includes(q) || p.description.toLowerCase().includes(q))
  );
  
  renderProducts();
}

/**
 * Renderiza grid de productos
 */
function renderProducts() {
  productsContainer.innerHTML = '';
  
  // Crear tarjeta para cada producto
  filtered.forEach((p, i) => {
    const card = document.createElement('div');
    card.className = 'product-card';
  const quantityText = getQuantity(p) || '';
    card.innerHTML = `
      <div class="product-image-container">
        <img src="${p.images[0]}" alt="${p.name}" class="product-image" loading="lazy"/>
      </div>
      <div class="product-info-container">
        <h3 class="product-name">${p.name}</h3>
    <p class="product-meta"><span class="product-code">Cod: ${p.id}</span>${quantityText ? ` • <span class="product-qty">${quantityText}</span>` : ''}</p>
    <p class="product-price">$${p.price.toFixed(2)}</p>
      </div>`;
    
    // Abrir modal al hacer clic
    card.onclick = () => showProductDetail(filtered[i]);
    productsContainer.appendChild(card);
  });
}

/**
 * Muestra modal de detalle con carousel + variantes
 * @param {Object} product - Producto a mostrar
 */
function showProductDetail(product) {
  // Configurar carousel
  currentList = product.images; 
  currentIndex = 0;
  updateModalImage();
  
  // Llenar información del producto
  document.getElementById('modal-name').textContent = product.name;
  document.getElementById('modal-desc').textContent = product.description;
  document.getElementById('modal-price').textContent = `$${product.price.toFixed(2)}`;
  document.getElementById('modal-category').textContent = product.category;
  
  // Agregar código y cantidad en la descripción si falta
  const qty = getQuantity(product);
  if (qty) {
    const descEl = document.getElementById('modal-desc');
    if (!/Cantidad:/i.test(descEl.textContent)) {
      descEl.textContent += `\nCantidad: ${qty}`;
    }
  }
  
  // Construir miniaturas y variantes
  buildThumbnails();
  buildVariants(product.variants);
  
  // Resetear cantidad
  document.getElementById('modal-quantity').value = 1;
  
  // Configurar botón de agregar al carrito
  document.getElementById('add-to-cart').onclick = () => addToCart(product);
  
  // Mostrar modal
  document.getElementById('product-modal').classList.remove('hidden');
}

/**
 * Construye miniaturas para el carousel
 */
function buildThumbnails() {
  const cont = document.querySelector('.modal-thumbnails');
  cont.innerHTML = '';
  
  // Crear miniatura para cada imagen
  currentList.forEach((src, i) => {
    const img = document.createElement('img');
    img.src = src;
    img.alt = 'Miniatura';
    img.classList.toggle('active', i === currentIndex);
    img.onclick = () => { 
      currentIndex = i; 
      updateModalImage(); 
    };
    cont.appendChild(img);
  });
}

/**
 * Actualiza imagen principal del carousel
 */
function updateModalImage() {
  document.getElementById('modal-image').src = currentList[currentIndex];
  document.getElementById('modal-image').alt = `Imagen ${currentIndex + 1}`;
  
  // Actualizar estado activo de miniaturas
  document.querySelectorAll('.modal-thumbnails img').forEach((t, i) =>
    t.classList.toggle('active', i === currentIndex)
  );
}

/**
 * Construye select o span para variantes
 * @param {Object} vars - Objeto de variantes
 */
function buildVariants(vars) {
  const cont = document.getElementById('modal-variants');
  cont.innerHTML = '';

  // Ejemplo de mapa para SVGs (opcional)
  const svgMap = {
    'Rojo': '<svg width="16" height="16" viewBox="0 0 16 16"><!-- SVG Rojo --></svg>',
    'Verde': '<svg width="16" height="16" viewBox="0 0 16 16"><!-- SVG Verde --></svg>',
    'Azul': '<svg width="16" height="16" viewBox="0 0 16 16"><!-- SVG Azul --></svg>'
    // Agrega más opciones según necesites
  };

  Object.entries(vars).forEach(([name, opts]) => {
    const wrapper = document.createElement('div');
    wrapper.className = 'variant-wrapper';

    const label = document.createElement('label');
    label.textContent = name;
    wrapper.appendChild(label);

    // Si hay múltiples opciones, crear select
    if (opts.length > 1) {
      const select = document.createElement('select');
      select.id = `variant-${name}`;
      select.name = name;

      opts.forEach(o => {
        const option = document.createElement('option');
        option.value = o;
        // Agrega SVG si existe en el mapa
        option.innerHTML = (svgMap[o] || '') + ' ' + o;
        select.appendChild(option);
      });

      wrapper.appendChild(select);
    } else {
      // Solo hay una opción: analizar si es un rango usando el guion "-" como separador
      const value = opts[0];
      // Expresión regular para detectar un rango en formato "min-max" (por ejemplo, "0-100")
      const rangePattern = /^\s*(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)\s*$/;
      const rangeMatch = value.match(rangePattern);

      if (rangeMatch) {
        const min = rangeMatch[1];
        const max = rangeMatch[2];
        const rangeInput = document.createElement('input');
        rangeInput.type = 'range';
        rangeInput.min = min;
        rangeInput.max = max;
        rangeInput.value = min;

        // Mostrar el valor actual al lado del range
        const rangeValue = document.createElement('span');
        rangeValue.textContent = min;
        rangeInput.addEventListener('input', () => {
          rangeValue.textContent = rangeInput.value;
        });

        wrapper.appendChild(rangeInput);
        wrapper.appendChild(rangeValue);
      } else {
        // Si no es un rango, se muestra en un campo readonly
        const span = document.createElement('span');
        span.className = 'variant-single';
        span.textContent = opts[0];
        wrapper.appendChild(span);
      }
    }

    cont.appendChild(wrapper);
  });
}

/**
 * Navega carousel
 * @param {number} step - Dirección de navegación (-1 o 1)
 */
function changeImage(step) {
  currentIndex = (currentIndex + step + currentList.length) % currentList.length;
  updateModalImage();
}

/**
 * Cierra modal detalle
 */
function closeModal() {
  document.getElementById('product-modal').classList.add('hidden');
}

/**
 * Ajusta cantidad en el selector de cantidad
 * @param {number} change - Cantidad a ajustar (-1 o 1)
 */
function adjustQuantity(change) {
  const input = document.getElementById('modal-quantity');
  const newValue = Math.max(1, parseInt(input.value) + change);
  input.value = newValue;
}

/**
 * Añade item al carrito
 * @param {Object} product - Producto a añadir
 */
function addToCart(product) {
  // Recopilar variantes seleccionadas
  const selected = {};
  Object.keys(product.variants).forEach(k => {
    const variantEl = document.getElementById(`variant-${k}`);
    selected[k] = variantEl ? variantEl.value : product.variants[k][0];
  });
  
  // Obtener cantidad
  const qty = Math.max(1, parseInt(document.getElementById('modal-quantity').value, 10));
  
  // Añadir al carrito
  cart.push({ 
    id: product.id,
    name: product.name, 
    variants: selected, 
    qty, 
    price: product.price,
    image: product.images[0]
  });
  
  // Guardar en localStorage
  localStorage.setItem('cart', JSON.stringify(cart));
  
  // Notificar al usuario
  const notification = document.createElement('div');
  notification.className = 'notification';
  notification.innerHTML = `
    <div class="notification-content">
      <i class="fas fa-check-circle"></i>
      <span>Producto agregado al carrito</span>
    </div>
  `;
  document.body.appendChild(notification);
  
  // Eliminar notificación después de 3 segundos
  setTimeout(() => {
    notification.classList.add('fade-out');
    setTimeout(() => notification.remove(), 300);
  }, 2000);
  
  // Cerrar modal y actualizar carrito
  closeModal();
  renderCart();
  updateCartBadge();
}

/**
 * Renderiza modal carrito
 */
function renderCart() {
  const list = document.getElementById('cart-items');
  list.innerHTML = '';
  let total = 0;

  const clearCartBtn = document.getElementById('clear-cart');
  
  // Mostrar/ocultar mensaje de carrito vacío
  if (cart.length === 0) {
    emptyCartMessage.style.display = 'flex';
    document.querySelector('.cart-container').style.display = 'none';
  } else {
    emptyCartMessage.style.display = 'none';
    document.querySelector('.cart-container').style.display = 'block';
    
    // Mostrar botón de vaciar carrito solo si hay items
    if (clearCartBtn) {
      clearCartBtn.style.display = 'inline-flex';
    }
    
    // Crear elemento para cada item del carrito
    cart.forEach((item, i) => {
      const line = item.price * item.qty;
      total += line;

      const li = document.createElement('li');
      li.className = 'cart-item';
      
      // Información del producto
      const infoDiv = document.createElement('div');
      infoDiv.className = 'cart-item-info';
      infoDiv.innerHTML = `
        <strong>${item.name}</strong>
        ${Object.entries(item.variants).map(([k,v]) => `<small>${k}: ${v}</small>`).join(' • ')}
      `;
      
      // Acciones (cantidad, precio, eliminar)
      const actionsDiv = document.createElement('div');
      actionsDiv.className = 'cart-item-actions';
      
      const qtyInput = document.createElement('input');
      qtyInput.type = 'number';
      qtyInput.min = '1';
      qtyInput.value = item.qty;
      qtyInput.className = 'cart-qty-input';
      qtyInput.addEventListener('change', () => updateCartItem(i, qtyInput.value));
      
      const priceSpan = document.createElement('span');
      priceSpan.className = 'cart-line-price';
      priceSpan.textContent = `$${line.toFixed(2)}`;
      
      const removeBtn = document.createElement('button');
      removeBtn.className = 'cart-remove';
      removeBtn.innerHTML = '<i class="fas fa-times"></i>';
      removeBtn.addEventListener('click', () => removeFromCart(i));
      
      actionsDiv.appendChild(qtyInput);
      actionsDiv.appendChild(priceSpan);
      actionsDiv.appendChild(removeBtn);
      
      li.appendChild(infoDiv);
      li.appendChild(actionsDiv);
      list.appendChild(li);
    });
  }

  // Actualizar total
  document.getElementById('cart-total').textContent = `$${total.toFixed(2)}`;
}

/**
 * Vacía completamente el carrito
 */
function clearCart() {
  if (cart.length === 0) {
    alert('El carrito ya está vacío');
    return;
  }
  
  if (confirm('¿Estás seguro de que quieres vaciar el carrito?')) {
    cart = [];
    localStorage.setItem('cart', JSON.stringify(cart));
    renderCart();
    updateCartBadge();
    
    // Mostrar notificación
    const notification = document.createElement('div');
    notification.className = 'notification';
    notification.innerHTML = `
      <div class="notification-content">
        <i class="fas fa-trash"></i>
        <span>Carrito vaciado</span>
      </div>
    `;
    document.body.appendChild(notification);
    
    // Eliminar notificación después de 2 segundos
    setTimeout(() => {
      notification.classList.add('fade-out');
      setTimeout(() => notification.remove(), 300);
    }, 1500);
  }
}

/**
 * Elimina item del carrito
 * @param {number} idx - Índice del item a eliminar
 */
function removeFromCart(idx) {
  cart.splice(idx, 1);
  localStorage.setItem('cart', JSON.stringify(cart));
  renderCart();
  updateCartBadge();
}

/**
 * Actualiza cantidad de un item del carrito
 * @param {number} index - Índice del item
 * @param {number|string} newQty - Nueva cantidad
 */
function updateCartItem(index, newQty) {
  cart[index].qty = Math.max(1, parseInt(newQty, 10));
  localStorage.setItem('cart', JSON.stringify(cart));
  renderCart();
  updateCartBadge();
}

/**
 * Actualiza badge del carrito
 */
function updateCartBadge() {
  const cartCount = document.querySelector('.cart-count');
  const cartNavCount = document.querySelector('.cart-nav-count');
  
  cartCount.textContent = cart.length;
  cartNavCount.textContent = cart.length;
  
  // Mostrar/ocultar botón de WhatsApp
  sendBtn.classList.toggle('hidden', cart.length === 0);
}

/**
 * Arma texto WhatsApp y abre chat
 */
function sendOrder() {
  if (!cart.length) {
    alert('El carrito está vacío');
    return;
  }
  
  const name = prompt('¿Cómo te llamas?') || 'Cliente';
  let total = 0;
  let text = `Hola, soy ${name}. Te paso mi pedido:%0A%0A`;
  
  // Construir mensaje con items del carrito
  cart.forEach((item, i) => {
    const lineTotal = item.price * item.qty;
    total += lineTotal;
    
    text += `*${i+1}.* ${item.name} (${item.qty}x $${item.price.toFixed(2)})%0A`;
    Object.entries(item.variants).forEach(([k,v]) => text += `• ${k}: ${v}%0A`);
    text += `• Subtotal: $${lineTotal.toFixed(2)}%0A%0A`;
  });
  
  // Añadir total
  text += `*TOTAL: $${total.toFixed(2)}*%0A%0A`;
  text += `Gracias!`;
  
  // Abrir WhatsApp
  const phone = getWhatsAppPhone();
  window.open(`https://wa.me/${phone}?text=${text}`, '_blank');
  
}

// Configurar event listeners para navegación del carousel
document.getElementById('prev-img').onclick = () => changeImage(-1);
document.getElementById('next-img').onclick = () => changeImage(1);

// Añadir estilos dinámicos para notificaciones
const style = document.createElement('style');
style.textContent = `
  .notification {
    position: fixed;
    bottom: 20px;
    left: 50%;
    transform: translateX(-50%);
    background: var(--primary);
    color: white;
    padding: 12px 20px;
    border-radius: 50px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    z-index: 1000;
    animation: slideUp 0.3s ease;
  }
  
  .notification-content {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  
  .notification.fade-out {
    opacity: 0;
    transform: translate(-50%, 10px);
    transition: all 0.3s ease;
  }
  
  @keyframes slideUp {
    from { opacity: 0; transform: translate(-50%, 10px); }
    to { opacity: 1; transform: translate(-50%, 0); }
  }
`;
document.head.appendChild(style);


function loadClients(wb) {
  const sheet = wb.SheetNames.find(n => /clientes/i.test(n));
  if (!sheet) return;
  clients = wb.Sheets[sheet] || [];
}

function populateClientSelect() {
  // Esta función ya no es necesaria ya que removimos el client selector
  // Mantener vacía por compatibilidad
}

// Función para generar PDF de lista de precios
async function generatePriceListPDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();
  const primaryColor = getComputedStyle(document.documentElement).getPropertyValue('--primary').trim() || '#2980b9';

  // Helper to convert hex to RGB
  function hexToRgb(hex) {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? [parseInt(result[1], 16), parseInt(result[2], 16), parseInt(result[3], 16)] : [41, 128, 185];
  }
  const primaryRgb = hexToRgb(primaryColor);

  // Configurar fuente
  doc.setFont('helvetica', 'normal');

  // --- HEADER ---
  // Logo
  let logoData = null;
  try {
    const response = await fetch('images/logo.png');
    if (response.ok) {
        const blob = await response.blob();
        logoData = await new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.readAsDataURL(blob);
        });
    }
  } catch (e) {
    console.log('No se pudo cargar el logo para el PDF');
  }

  if (logoData) {
    doc.addImage(logoData, 'PNG', 15, 15, 30, 30);
  }

  // Título
  doc.setFontSize(22);
  doc.setFont('helvetica', 'bold');
  doc.text('Lista de Precios', doc.internal.pageSize.getWidth() / 2, 25, { align: 'center' });
  doc.setFontSize(12);
  doc.setFont('helvetica', 'normal');
  doc.text(new Date().toLocaleDateString('es-AR'), doc.internal.pageSize.getWidth() / 2, 32, { align: 'center' });


  // Datos de la empresa
  doc.setFontSize(10);
  const companyInfoX = doc.internal.pageSize.getWidth() - 15;
  doc.text('Bio-Quim', companyInfoX, 20, { align: 'right' });
  doc.text('Buenos Aires, Argentina', companyInfoX, 25, { align: 'right' });
  doc.text('administracion@bio-quim.com.ar', companyInfoX, 30, { align: 'right' });
  doc.text('Vendedor: Valeria', companyInfoX, 40, { align: 'right' });
  doc.text('Tel: +54 9 11 2744 4019', companyInfoX, 45, { align: 'right' });


  // --- BODY (TABLA DE PRODUCTOS) ---
  let startY = 60;

  const categories = [...new Set(products.map(p => p.category))].sort();

  categories.forEach((category, index) => {
    const categoryProducts = products.filter(p => p.category === category);
    if (categoryProducts.length === 0) return;

    // Evitar que el título de la categoría quede solo al final de la página
    if (startY > 250) {
        doc.addPage();
        startY = 20;
    }

    // Título de la categoría
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(...primaryRgb);
    doc.text(category, 14, startY);
    doc.setTextColor(0, 0, 0); // Reset text color
    startY += 8;

    const rows = categoryProducts.map(p => [
      p.name,
      getQuantity(p) || '',
      `$${p.price.toFixed(2)}`
    ]);

    doc.autoTable({
      startY: startY,
      head: [['Producto', 'Presentación', 'Precio']],
      body: rows,
      theme: 'striped', // 'striped', 'grid', 'plain'
      styles: {
        fontSize: 9,
        cellPadding: 2,
        valign: 'middle'
      },
      headStyles: {
        fillColor: primaryRgb,
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        halign: 'center'
      },
      columnStyles: {
        0: { cellWidth: 'auto' },
        1: { cellWidth: 30, halign: 'center' },
        2: { cellWidth: 25, halign: 'right' }
      },
      didDrawPage: function (data) {
        // --- FOOTER ---
        const pageCount = doc.internal.getNumberOfPages();
        doc.setFontSize(8);
        doc.setTextColor(150);
        doc.text(
          `Página ${data.pageNumber} de ${pageCount}`,
          doc.internal.pageSize.getWidth() / 2,
          doc.internal.pageSize.getHeight() - 10,
          { align: 'center' }
        );
      },
      margin: { top: 10, bottom: 20 } // Margen para footer
    });
    
    startY = doc.autoTable.previous.finalY + 12;
  });


  // Guardar el PDF
  doc.save(`lista_precios_bio_quim_${new Date().toISOString().split('T')[0]}.pdf`);
}

// Event listener para el botón de descarga
// document.getElementById('download-pdf').addEventListener('click', generatePriceListPDF);

// =====================
// Helpers extra
// =====================
function getQuantity(product) {
  // Se pidió usar directamente la columna "description" como cantidad.
  if (product.description) {
    return String(product.description).trim();
  }
  // Fallback (por si algún producto no trae description) usando lógica anterior resumida
  const name = product.name || '';
  const match = name.match(/(\d+[\.,]?\d*)\s*(LT|L|KG|KGS|G|GR|GRS|ML|CC|UNID(?:ADES)?|U|UND|UNIDADES)/i);
  if (match) return match[1].replace(',', '.') + ' ' + match[2].toUpperCase();
  return '';
}

