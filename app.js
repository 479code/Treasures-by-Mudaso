const sampleOrders = [
  { name: 'Tosin', items: '5 items · 10:00 AM · Abuja', status: 'Ready to pick', type: '' },
  { name: 'Blessing', items: '8 items · 12:30 PM · Kano', status: 'Packing', type: 'packing' },
  { name: 'Zainab', items: '6 items · 3:00 PM · Kaduna', status: 'Out for delivery', type: 'delivery' }
];

const storageKey = 'treasures-products';
let products = JSON.parse(localStorage.getItem(storageKey) || '[]');

function saveProducts() {
  localStorage.setItem(storageKey, JSON.stringify(products));
}

function brand() {
  return `<div class="brand"><div class="mark">TM</div><div><div class="brand-name">Treasures by Mudaso</div><div class="brand-sub">Thoughtful. Quality. Trusted.</div></div></div>`;
}

function nav(active) {
  return `<nav class="nav" aria-label="Main navigation">
    ${[['home','Home'],['shop','Shop'],['packages','Packages'],['owner','Owner']].map(([id, label]) => `<button class="${active === id ? 'active' : ''}" data-page="${id}">${label}</button>`).join('')}
  </nav>`;
}

function home() {
  return `<section class="screen">
    <header class="topbar">${brand()}<button class="round-button" aria-label="Open shopping bag">⌕</button></header>
    <section class="hero"><div class="eyebrow">Curated groceries</div><h1>Delivering care to your door.</h1><p>Thoughtfully selected essentials for a happier home.</p><button class="button gold" data-page="shop">Shop now</button></section>
    <section class="schedule"><small><b>Delivery days</b>Choose the day that suits you</small><span class="day">Mon</span><span class="day">Wed</span><span class="day">Fri</span></section>
    <div class="section-head"><h2>Shop by category</h2><button class="link" data-page="shop">See all</button></div>
    <div class="category-grid">
      <button class="category"><span>·</span><strong>Pantry</strong></button><button class="category"><span>◌</span><strong>Rice & grains</strong></button><button class="category"><span>□</span><strong>Household</strong></button><button class="category"><span>⌁</span><strong>Fresh essentials</strong></button>
    </div>
    <div class="section-head"><h2>Curated packages</h2><button class="link" data-page="packages">View all</button></div>
    <article class="package-card"><div><div class="eyebrow" style="color:var(--gold-soft)">Made with care</div><h3>Packages for every home</h3><p>Flexible essentials, beautifully prepared and ready when you are.</p></div><div class="package-badge">TM</div></article>
  </section>${nav('home')}`;
}

function owner() {
  return `<section class="screen">
    <header class="owner-hero"><div class="topbar">${brand()}<button class="round-button" aria-label="Owner settings">⚙</button></div><div class="eyebrow" style="color:var(--gold-soft)">Owner overview</div><h1>Today</h1><p>Everything that needs your attention, in one place.</p><div class="day-tabs"><button>Monday</button><button class="active">Wednesday</button><button>Friday</button></div></header>
    <div class="stats"><div class="stat"><b>8</b><span>orders to prepare</span></div><div class="stat"><b>2</b><span>awaiting payment</span></div><div class="stat"><b>3</b><span>low-stock items</span></div><div class="stat"><b>₦124.5k</b><span>expected today</span></div></div>
    <div class="section-head"><h2>Today’s orders</h2><button class="link">See all</button></div>
    <div class="order-list">${sampleOrders.map(o => `<article class="order"><div class="avatar">${o.name[0]}</div><div><strong>${o.name}</strong><small>${o.items}</small></div><span class="status ${o.type}">${o.status}</span></article>`).join('')}</div>
    <div class="alert"><b>Urgent order</b><span>A customer needs their order today. Please prioritise.</span></div>
    <div class="quick-actions"><button class="button" data-page="add-product">Add product</button><button class="button secondary" data-page="create-package">Create package</button><button class="button secondary" data-page="orders">View all orders</button></div>
  </section>${nav('owner')}`;
}

function placeholder(page) {
  const map = { shop: ['The shop is ready for products.', 'Once the owner adds items in Google Sheets, they will appear here automatically.'], packages: ['A home for every package.', 'The owner will add package names, contents, images and prices here.'], orders: ['All orders will live here.', 'This page will show payment, picking, packing and delivery status.'] };
  if (page === 'shop' && products.length) return `<section class="screen"><header class="topbar">${brand()}<button class="round-button" data-page="owner" aria-label="Go to owner area">⌕</button></header><div class="eyebrow">Your catalogue</div><h1>Shop essentials</h1><p class="muted">Items added by the owner are shown below.</p><div class="product-list">${products.map(product => `<article class="product"><div class="product-initial">${product.name[0]}</div><div><strong>${product.name}</strong><small>${product.category} · ${product.stock} in stock</small></div><b>₦${Number(product.price).toLocaleString()}</b></article>`).join('')}</div></section>${nav('shop')}`;
  return `<section class="screen placeholder"><div class="placeholder-box"><div class="eyebrow">Framework ready</div><h1>${map[page][0]}</h1><p class="muted">${map[page][1]}</p><button class="button" data-page="owner">Go to owner area</button></div></section>${nav(page)}`;
}

function addProduct() {
  return `<section class="screen"><header class="topbar">${brand()}<button class="round-button" data-page="owner" aria-label="Back to owner area">×</button></header><div class="eyebrow">Owner area</div><h1>Add a product</h1><p class="muted">This information will later be saved to the Products tab in Google Sheets.</p><form id="product-form" class="form-card"><label>Product name<input required name="name" placeholder="e.g. Basmati rice, 5kg" /></label><label>Category<select required name="category"><option value="">Choose a category</option><option>Pantry</option><option>Rice & grains</option><option>Household</option><option>Fresh essentials</option></select></label><div class="form-row"><label>Price (₦)<input required name="price" type="number" min="0" placeholder="0" /></label><label>Stock quantity<input required name="stock" type="number" min="0" placeholder="0" /></label></div><label>Photo link <span>(optional)</span><input name="photo" type="url" placeholder="https://…" /></label><label class="toggle"><input name="available" type="checkbox" checked /><span>Available to customers</span></label><button class="button" type="submit">Save product</button></form></section>${nav('owner')}`;
}

function createPackage() {
  return `<section class="screen"><header class="topbar">${brand()}<button class="round-button" data-page="owner" aria-label="Back to owner area">×</button></header><div class="eyebrow">Owner area</div><h1>Create a package</h1><p class="muted">Packages will be built from the owner’s product list after the Sheet connection is added.</p><div class="placeholder-box"><h2>Coming next</h2><p class="muted">The structure is ready; first add a few products, then package creation can select their contents and quantities.</p><button class="button" data-page="add-product">Add a first product</button></div></section>${nav('owner')}`;
}

function render(page = 'home') { document.querySelector('#app').innerHTML = page === 'home' ? home() : page === 'owner' ? owner() : page === 'add-product' ? addProduct() : page === 'create-package' ? createPackage() : placeholder(page); }
document.addEventListener('click', (event) => { const page = event.target.closest('[data-page]')?.dataset.page; if (page) render(page); });
document.addEventListener('submit', (event) => {
  if (event.target.id !== 'product-form') return;
  event.preventDefault();
  const values = Object.fromEntries(new FormData(event.target));
  products.push({ ...values, available: event.target.available.checked });
  saveProducts();
  render('shop');
});
render();
if ('serviceWorker' in navigator) window.addEventListener('load', () => navigator.serviceWorker.register('sw.js'));
