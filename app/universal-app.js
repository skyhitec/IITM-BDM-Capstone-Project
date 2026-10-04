// Trackify Pro - Universal Business Operating System JavaScript Engine

// State & Repositories
let currentLanguage = 'hi';

const businessCatalogues = {
  kirana: [
    { sku: 'KIR-001', name: 'Fortune Mustard Oil 1L', category: 'Grocery', mrp: 180, purchaseRate: 145, stock: 45, unit: 'Ltr' },
    { sku: 'KIR-002', name: 'Aashirvaad Atta 10Kg', category: 'Grocery', mrp: 420, purchaseRate: 360, stock: 28, unit: 'Box' },
    { sku: 'KIR-003', name: 'Tata Salt 1Kg', category: 'Grocery', mrp: 28, purchaseRate: 21, stock: 120, unit: 'Pcs' },
    { sku: 'KIR-004', name: 'Amul Butter 500g', category: 'Dairy', mrp: 275, purchaseRate: 240, stock: 15, unit: 'Pcs' },
    { sku: 'KIR-005', name: 'Surf Excel Quick Wash 1Kg', category: 'Detergents', mrp: 210, purchaseRate: 175, stock: 34, unit: 'Pcs' }
  ],
  garments: [
    { sku: 'GAR-101', name: 'Levi Slim Fit Denim Jeans 32', category: 'Menswear', mrp: 1899, purchaseRate: 1100, stock: 18, unit: 'Pcs' },
    { sku: 'GAR-102', name: 'Cotton Casual Shirt XL', category: 'Menswear', mrp: 999, purchaseRate: 550, stock: 25, unit: 'Pcs' },
    { sku: 'GAR-103', name: 'Designer Silk Saree Red', category: 'Womenswear', mrp: 3499, purchaseRate: 2100, stock: 8, unit: 'Pcs' },
    { sku: 'GAR-104', name: 'Kids Kurta Pyjama Set 6Y', category: 'Kidswear', mrp: 799, purchaseRate: 400, stock: 14, unit: 'Pcs' }
  ],
  electronics: [
    { sku: 'ELE-201', name: 'Samsung 25W Fast Charger Type-C', category: 'Accessories', mrp: 1299, purchaseRate: 750, stock: 40, unit: 'Pcs' },
    { sku: 'ELE-202', name: 'boAt Airdopes 141 TWS', category: 'Audio', mrp: 1499, purchaseRate: 890, stock: 19, unit: 'Pcs' },
    { sku: 'ELE-203', name: 'Realme 10,000mAh Powerbank', category: 'Gadgets', mrp: 1199, purchaseRate: 780, stock: 12, unit: 'Pcs' }
  ],
  hardware: [
    { sku: 'HDW-301', name: 'Asian Paints Apex Emulsion 20L', category: 'Paints', mrp: 4800, purchaseRate: 3950, stock: 6, unit: 'Box' },
    { sku: 'HDW-302', name: 'Finolex 2.5sqmm Wire Roll 90m', category: 'Electricals', mrp: 2450, purchaseRate: 1850, stock: 15, unit: 'Pcs' },
    { sku: 'HDW-303', name: 'Taparia 8-inch Adjustable Spanner', category: 'Tools', mrp: 380, purchaseRate: 260, stock: 22, unit: 'Pcs' }
  ],
  pharmacy: [
    { sku: 'PHA-401', name: 'Paracetamol 650mg Dolo (Strip 15)', category: 'Analgesics', mrp: 34, purchaseRate: 22, stock: 150, unit: 'Pcs' },
    { sku: 'PHA-402', name: 'Azithromycin 500mg (Strip 5)', category: 'Antibiotics', mrp: 118, purchaseRate: 78, stock: 60, unit: 'Pcs' },
    { sku: 'PHA-403', name: 'Volini Pain Relief Spray 100g', category: 'OTC', mrp: 295, purchaseRate: 210, stock: 24, unit: 'Pcs' }
  ],
  restaurant: [
    { sku: 'RES-501', name: 'Special Paneer Butter Masala', category: 'Main Course', mrp: 260, purchaseRate: 110, stock: 99, unit: 'Pcs' },
    { sku: 'RES-502', name: 'Butter Naan Basket (4 Pcs)', category: 'Breads', mrp: 140, purchaseRate: 40, stock: 99, unit: 'Pcs' },
    { sku: 'RES-503', name: 'Veg Dum Biryani Family Pack', category: 'Biryani', mrp: 320, purchaseRate: 135, stock: 99, unit: 'Pcs' }
  ],
  services: [
    { sku: 'SRV-601', name: 'Smartphone Screen Replacement', category: 'Mobile Repair', mrp: 1800, purchaseRate: 850, stock: 50, unit: 'Pcs' },
    { sku: 'SRV-602', name: 'Split AC Deep Foam Cleaning', category: 'Appliance Repair', mrp: 699, purchaseRate: 200, stock: 50, unit: 'Pcs' }
  ]
};

let currentInventory = [...businessCatalogues.kirana];
let cart = [];

let khataList = [
  { name: 'Ramesh Kumar Singh', phone: '+91 98350 12345', credit: 5400, paid: 2000, pending: 3400, lastTx: '2026-10-02' },
  { name: 'Sunita Devi', phone: '+91 94310 67890', credit: 2800, paid: 1000, pending: 1800, lastTx: '2026-10-03' },
  { name: 'Mohd. Imran Khan', phone: '+91 99340 54321', credit: 11600, paid: 0, pending: 11600, lastTx: '2026-09-28' }
];

let weeklyChart = null;
let paymentChart = null;

// INIT
document.addEventListener('DOMContentLoaded', () => {
  setupNavigation();
  changeBusinessType();
  renderKhataTable();
  initCharts();
});

// NAVIGATION
function setupNavigation() {
  const navItems = document.querySelectorAll('.nav-item');
  navItems.forEach(item => {
    item.addEventListener('click', () => {
      const tabId = item.getAttribute('data-tab');
      switchTab(tabId);
    });
  });
}

function switchTab(tabId) {
  document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));

  const activeNav = document.querySelector(`.nav-item[data-tab="${tabId}"]`);
  const activeTab = document.getElementById(tabId);

  if (activeNav) activeNav.classList.add('active');
  if (activeTab) activeTab.classList.add('active');
}

// BUSINESS TYPE SWITCHING
function changeBusinessType() {
  const select = document.getElementById('biz-type-select');
  const selectedType = select.value;
  
  if (businessCatalogues[selectedType]) {
    currentInventory = [...businessCatalogues[selectedType]];
  }

  // Update POS select options
  const posSelect = document.getElementById('pos-select');
  posSelect.innerHTML = '';
  
  currentInventory.forEach((item, index) => {
    const opt = document.createElement('option');
    opt.value = index;
    opt.textContent = `${item.name} - ₹${item.mrp} (Stock: ${item.stock} ${item.unit})`;
    posSelect.appendChild(opt);
  });

  updatePosDetailPill();
  renderInventoryTable();
  updateKpiMetrics();
}

function updatePosDetailPill() {
  const select = document.getElementById('pos-select');
  const index = select.value;
  const strip = document.getElementById('pos-detail-strip');
  
  if (index !== undefined && currentInventory[index]) {
    const item = currentInventory[index];
    document.getElementById('pos-cat-val').textContent = item.category;
    document.getElementById('pos-mrp-val').textContent = `₹${item.mrp.toFixed(2)}`;
    document.getElementById('pos-stock-val').textContent = `${item.stock} ${item.unit}`;
    strip.style.display = 'flex';
  } else {
    strip.style.display = 'none';
  }
}

function handlePosSearch() {
  const query = document.getElementById('pos-search').value.toLowerCase();
  const select = document.getElementById('pos-select');
  select.innerHTML = '';
  
  currentInventory.forEach((item, index) => {
    if (item.name.toLowerCase().includes(query) || item.sku.toLowerCase().includes(query) || item.category.toLowerCase().includes(query)) {
      const opt = document.createElement('option');
      opt.value = index;
      opt.textContent = `${item.name} - ₹${item.mrp} (Stock: ${item.stock} ${item.unit})`;
      select.appendChild(opt);
    }
  });

  updatePosDetailPill();
}

// POS CART MANAGEMENT
function addPosCartItem() {
  const select = document.getElementById('pos-select');
  const index = select.value;
  const qtyInput = document.getElementById('pos-qty');
  const discInput = document.getElementById('pos-disc');

  if (index === undefined || !currentInventory[index]) {
    alert('Please select a valid item first.');
    return;
  }

  const item = currentInventory[index];
  const qty = parseInt(qtyInput.value) || 1;
  const discountPercent = parseFloat(discInput.value) || 0;

  const discountedPrice = item.mrp * (1 - discountPercent / 100);
  const total = discountedPrice * qty;

  // Check if item exists in cart
  const existing = cart.find(c => c.sku === item.sku);
  if (existing) {
    existing.qty += qty;
    existing.total += total;
  } else {
    cart.push({
      sku: item.sku,
      name: item.name,
      mrp: item.mrp,
      price: discountedPrice,
      qty: qty,
      total: total
    });
  }

  renderCart();
  qtyInput.value = 1;
  discInput.value = 0;
}

function removeCartItem(index) {
  cart.splice(index, 1);
  renderCart();
}

function renderCart() {
  const tbody = document.getElementById('cart-tbody');
  tbody.innerHTML = '';

  if (cart.length === 0) {
    tbody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: var(--muted);">Cart is empty. Select item above.</td></tr>`;
    updateCartSummary(0, 0, 0);
    return;
  }

  let subtotal = 0;
  cart.forEach((c, index) => {
    subtotal += c.total;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${c.name}</td>
      <td>${c.qty}</td>
      <td>₹${c.price.toFixed(2)}</td>
      <td>₹${c.total.toFixed(2)}</td>
      <td><button class="btn btn-sm btn-outline" style="border-color: var(--rose); color: var(--rose);" onclick="removeCartItem(${index})">&times;</button></td>
    `;
    tbody.appendChild(tr);
  });

  const gst = subtotal * 0.12;
  const grand = subtotal + gst;
  updateCartSummary(subtotal, gst, grand);
}

function updateCartSummary(subtotal, gst, grand) {
  document.getElementById('cs-subtotal').textContent = `₹${subtotal.toFixed(2)}`;
  document.getElementById('cs-gst').textContent = `₹${gst.toFixed(2)}`;
  document.getElementById('cs-grand').textContent = `₹${grand.toFixed(2)}`;
}

function printReceipt(type) {
  if (cart.length === 0) {
    alert('Cart is empty. Add items to print receipt.');
    return;
  }
  const custName = document.getElementById('pos-cust-name').value || 'Walk-in Customer';
  const grandTotal = document.getElementById('cs-grand').textContent;
  alert(`🖨️ Printing ${type.toUpperCase()} Receipt for ${custName}!\nTotal Amount: ${grandTotal}\nThank you for shopping at Rajesh Super Mart!`);
}

function sendWhatsAppInvoice() {
  if (cart.length === 0) {
    alert('Cart is empty.');
    return;
  }
  const phone = document.getElementById('pos-cust-phone').value;
  if (!phone) {
    alert('Please enter customer WhatsApp phone number.');
    return;
  }
  const grand = document.getElementById('cs-grand').textContent;
  const text = encodeURIComponent(`Hello! Your bill at Rajesh Super Mart total is ${grand}. Thank you for shopping with us!`);
  window.open(`https://wa.me/${phone.replace(/[^0-9]/g, '')}?text=${text}`, '_blank');
}

// INVENTORY TABLE
function renderInventoryTable() {
  const tbody = document.getElementById('inv-tbody');
  const filter = (document.getElementById('inv-search-input')?.value || '').toLowerCase();
  tbody.innerHTML = '';

  currentInventory.forEach(item => {
    if (filter && !item.name.toLowerCase().includes(filter) && !item.category.toLowerCase().includes(filter) && !item.sku.toLowerCase().includes(filter)) {
      return;
    }

    const stockValue = item.stock * item.purchaseRate;
    const isLow = item.stock < 15;
    const statusTag = isLow ? `<span style="color: var(--rose); font-weight:700;"><i class="fa-solid fa-triangle-exclamation"></i> Low Stock</span>` : `<span style="color: var(--emerald); font-weight:700;">In Stock</span>`;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><code>${item.sku}</code></td>
      <td><strong>${item.name}</strong></td>
      <td><span class="badge-pro" style="background: var(--bg-card); color: var(--cyan); border: 1px solid var(--cyan);">${item.category}</span></td>
      <td>₹${item.mrp.toFixed(2)}</td>
      <td>₹${item.purchaseRate.toFixed(2)}</td>
      <td>${item.stock}</td>
      <td>${item.unit}</td>
      <td>₹${stockValue.toLocaleString('en-IN')}</td>
      <td>${statusTag}</td>
    `;
    tbody.appendChild(tr);
  });
}

function openAddProductModal() {
  document.getElementById('product-modal').classList.add('active');
}

function closeProductModal() {
  document.getElementById('product-modal').classList.remove('active');
}

function handleAddProduct(e) {
  e.preventDefault();
  const name = document.getElementById('p-name').value;
  const category = document.getElementById('p-cat').value;
  const unit = document.getElementById('p-unit').value;
  const mrp = parseFloat(document.getElementById('p-mrp').value);
  const purchaseRate = parseFloat(document.getElementById('p-purchase').value);
  const stock = parseInt(document.getElementById('p-stock').value);

  const sku = `PROD-${Math.floor(100 + Math.random() * 900)}`;

  currentInventory.unshift({ sku, name, category, mrp, purchaseRate, stock, unit });
  changeBusinessType();
  closeProductModal();
  alert(`✅ Product "${name}" added to inventory successfully!`);
}

// KHATA LEDGER
function renderKhataTable() {
  const tbody = document.getElementById('khata-tbody');
  tbody.innerHTML = '';

  khataList.forEach((k, index) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${k.name}</strong></td>
      <td>${k.phone}</td>
      <td>₹${k.credit.toLocaleString('en-IN')}</td>
      <td class="text-emerald">₹${k.paid.toLocaleString('en-IN')}</td>
      <td class="text-cyan" style="font-weight: 800;">₹${k.pending.toLocaleString('en-IN')}</td>
      <td>${k.lastTx}</td>
      <td>
        <button class="btn btn-sm btn-emerald" onclick="sendWhatsAppReminder('${k.phone}', '${k.name}', ${k.pending})">
          <i class="fa-brands fa-whatsapp"></i> Send Reminder
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function openAddUdharModal() {
  document.getElementById('udhar-modal').classList.add('active');
}

function closeUdharModal() {
  document.getElementById('udhar-modal').classList.remove('active');
}

function handleAddUdhar(e) {
  e.preventDefault();
  const name = document.getElementById('u-name').value;
  const phone = document.getElementById('u-phone').value;
  const amount = parseFloat(document.getElementById('u-amount').value);

  khataList.unshift({
    name,
    phone,
    credit: amount,
    paid: 0,
    pending: amount,
    lastTx: new Date().toISOString().split('T')[0]
  });

  renderKhataTable();
  closeUdharModal();
  updateKpiMetrics();
  alert(`✅ Credit Entry of ₹${amount} saved for ${name}!`);
}

function sendWhatsAppReminder(phone, name, amount) {
  const text = encodeURIComponent(`Dear ${name}, greeting from Rajesh Super Mart! Reminder for pending dues of ₹${amount.toLocaleString('en-IN')}. Pay via UPI: rajeshstore@upi. Thank you!`);
  window.open(`https://wa.me/${phone.replace(/[^0-9]/g, '')}?text=${text}`, '_blank');
}

// MODALS & TOOLS
function openCameraBarcodeScanner() {
  document.getElementById('camera-modal').classList.add('active');
}

function closeCameraModal() {
  document.getElementById('camera-modal').classList.remove('active');
}

function simulateBarcodeScan() {
  closeCameraModal();
  if (currentInventory.length > 0) {
    const randomItem = currentInventory[Math.floor(Math.random() * currentInventory.length)];
    cart.push({
      sku: randomItem.sku,
      name: randomItem.name,
      mrp: randomItem.mrp,
      price: randomItem.mrp,
      qty: 1,
      total: randomItem.mrp
    });
    renderCart();
    alert(`⚡ Barcode Scanned: ${randomItem.name} added to cart!`);
  }
}

function openOnlineStoreModal() {
  document.getElementById('ecomm-modal').classList.add('active');
}

function closeOnlineStoreModal() {
  document.getElementById('ecomm-modal').classList.remove('active');
}

function copyStoreLink() {
  const link = document.getElementById('store-link-input');
  link.select();
  document.execCommand('copy');
  alert('🔗 Digital Store link copied to clipboard!');
}

function shareStoreWhatsApp() {
  const link = document.getElementById('store-link-input').value;
  const text = encodeURIComponent(`Shop online 24/7 at Rajesh Super Mart! Click link to order: ${link}`);
  window.open(`https://wa.me/?text=${text}`, '_blank');
}

function openUpiQrModal() {
  const grandText = document.getElementById('cs-grand').textContent;
  document.getElementById('upi-qr-amount').textContent = grandText;
  document.getElementById('upi-modal').classList.add('active');
}

function closeUpiQrModal() {
  document.getElementById('upi-modal').classList.remove('active');
}

function confirmUpiPaymentReceived() {
  closeUpiQrModal();
  alert('💰 UPI Payment Received & Confirmed!');
  cart = [];
  renderCart();
}

function sendSupplierOrder(name) {
  const text = encodeURIComponent(`Hello ${name}, please dispatch stock order for Rajesh Super Mart, Darbhanga.`);
  window.open(`https://wa.me/919835011223?text=${text}`, '_blank');
}

function exportGstrReport() {
  alert('📊 Exporting GSTR-1 & GSTR-3B HSN Tax Report (Excel format)... File downloaded.');
}

function buyPlan(planName, price) {
  alert(`💳 Requesting License Key for "${planName}" (₹${price}/mo). Payment Gateway opening...`);
}

function toggleLanguage() {
  currentLanguage = currentLanguage === 'hi' ? 'en' : 'hi';
  document.getElementById('lbl-lang-toggle').textContent = currentLanguage === 'hi' ? 'Hindi (हिन्दी)' : 'English (English)';
  alert(`Language switched to ${currentLanguage === 'hi' ? 'Hindi' : 'English'}`);
}

function updateKpiMetrics() {
  let totalSales = 48950;
  let totalProfit = 12430;
  let totalUdhar = khataList.reduce((acc, k) => acc + k.pending, 0);
  let lowStockCount = currentInventory.filter(i => i.stock < 15).length;

  document.getElementById('kpi-sales').textContent = `₹${totalSales.toLocaleString('en-IN')}.00`;
  document.getElementById('kpi-profit').textContent = `₹${totalProfit.toLocaleString('en-IN')}.00`;
  document.getElementById('kpi-udhar').textContent = `₹${totalUdhar.toLocaleString('en-IN')}.00`;
  document.getElementById('kpi-lowstock').textContent = `${lowStockCount} Items`;
}

// CHARTS INITIALIZATION
function initCharts() {
  const ctxWeekly = document.getElementById('weeklyChart');
  const ctxPayment = document.getElementById('paymentChart');

  if (!ctxWeekly || !ctxPayment) return;

  if (weeklyChart) weeklyChart.destroy();
  if (paymentChart) paymentChart.destroy();

  weeklyChart = new Chart(ctxWeekly, {
    type: 'bar',
    data: {
      labels: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],
      datasets: [
        {
          label: 'Revenue (₹)',
          data: [32000, 41000, 38000, 45000, 52000, 68000, 48950],
          backgroundColor: '#06b6d4',
          borderRadius: 6
        },
        {
          label: 'Net Profit (₹)',
          data: [8000, 10500, 9200, 11000, 13500, 17000, 12430],
          backgroundColor: '#10b981',
          borderRadius: 6
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#94a3b8' } }
      },
      scales: {
        x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } },
        y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
      }
    }
  });

  paymentChart = new Chart(ctxPayment, {
    type: 'doughnut',
    data: {
      labels: ['UPI (GPay/PhonePe)', 'Cash', 'Udhar Khata', 'Card'],
      datasets: [{
        data: [55, 25, 15, 5],
        backgroundColor: ['#06b6d4', '#10b981', '#f59e0b', '#8b5cf6'],
        borderWidth: 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { color: '#94a3b8' } }
      }
    }
  });
}
