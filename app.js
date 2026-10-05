const initialProducts = [
  {
    id: 1,
    name: "Everyday canvas tote",
    category: "Carry",
    price: 38,
    stock: 24,
    tag: "Bestseller",
    image:
      "https://images.unsplash.com/photo-1590874103328-eac38a683ce7?auto=format&fit=crop&w=760&q=80",
  },
  {
    id: 2,
    name: "Sunday stoneware mug",
    category: "Home",
    price: 32,
    stock: 6,
    tag: "Small batch",
    image:
      "https://images.unsplash.com/photo-1514228742587-6b1558fcca3d?auto=format&fit=crop&w=760&q=80",
  },
  {
    id: 3,
    name: "Object no. 04 candle",
    category: "Objects",
    price: 46,
    stock: 18,
    tag: "",
    image:
      "https://images.unsplash.com/photo-1603006905003-be475563bc59?auto=format&fit=crop&w=760&q=80",
  },
  {
    id: 4,
    name: "Market day basket",
    category: "Carry",
    price: 64,
    stock: 4,
    tag: "Low stock",
    image:
      "https://images.unsplash.com/photo-1591561954557-26941169b49e?auto=format&fit=crop&w=760&q=80",
  },
  {
    id: 5,
    name: "Soft form throw",
    category: "Home",
    price: 88,
    stock: 12,
    tag: "",
    image:
      "https://images.unsplash.com/photo-1600210492486-724fe5c67fb0?auto=format&fit=crop&w=760&q=80",
  },
  {
    id: 6,
    name: "Still life bud vase",
    category: "Objects",
    price: 42,
    stock: 0,
    tag: "Sold out",
    image:
      "https://images.unsplash.com/photo-1578500494198-246f612d3b3d?auto=format&fit=crop&w=760&q=80",
  },
  {
    id: 7,
    name: "The daily journal",
    category: "Objects",
    price: 24,
    stock: 31,
    tag: "New",
    image:
      "https://images.unsplash.com/photo-1544947950-fa07a98d237f?auto=format&fit=crop&w=760&q=80",
  },
  {
    id: 8,
    name: "Good company tray",
    category: "Home",
    price: 54,
    stock: 9,
    tag: "",
    image:
      "https://images.unsplash.com/photo-1603199506016-b9a594b593c0?auto=format&fit=crop&w=760&q=80",
  },
];

const loadData = (key, fallback) => {
  try {
    const saved = localStorage.getItem(key);
    return saved ? JSON.parse(saved) : fallback;
  } catch {
    return fallback;
  }
};

let products = loadData(
  "zenvora-products",
  loadData("fieldnotes-products", initialProducts),
);
let bag = loadData("zenvora-bag", loadData("fieldnotes-bag", []));
let activeFilter = "All";
let toastTimer;

const money = (value) =>
  new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(
    value,
  );
const escapeHtml = (value) =>
  String(value).replace(
    /[&<>"']/g,
    (character) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[
        character
      ],
  );
const statusFor = (stock) =>
  stock === 0 ? "Out of stock" : stock < 8 ? "Low stock" : "In stock";
const saveProducts = () =>
  localStorage.setItem("zenvora-products", JSON.stringify(products));
const saveBag = () => localStorage.setItem("zenvora-bag", JSON.stringify(bag));

function renderIcons() {
  if (window.lucide) window.lucide.createIcons();
}

function renderProducts() {
  const grid = document.querySelector("#product-grid");
  const search =
    document
      .querySelector(".search-toggle")
      .dataset.query?.trim()
      .toLowerCase() || "";
  const sort = document.querySelector("#sort-select").value;
  let visible = products.filter(
    (product) =>
      (activeFilter === "All" || product.category === activeFilter) &&
      product.name.toLowerCase().includes(search),
  );
  if (sort === "price-low") visible.sort((a, b) => a.price - b.price);
  if (sort === "price-high") visible.sort((a, b) => b.price - a.price);
  if (sort === "name") visible.sort((a, b) => a.name.localeCompare(b.name));
  document.querySelector("#product-total").textContent = String(
    products.length,
  ).padStart(2, "0");
  grid.innerHTML = visible.length
    ? visible
        .map(
          (product) => `
    <article class="product-card">
      <div class="product-image-wrap">
        <img class="product-image" src="${escapeHtml(product.image)}" alt="${escapeHtml(product.name)}" loading="lazy" />
        ${product.tag ? `<span class="product-tag">${escapeHtml(product.tag)}</span>` : ""}
        <button class="quick-add" data-add-to-bag="${product.id}" ${product.stock < 1 ? "disabled" : ""}><span>${product.stock < 1 ? "Out of stock" : "Add to bag"}</span><i data-lucide="plus"></i></button>
      </div>
      <div class="product-meta"><span class="product-name">${escapeHtml(product.name)}</span><span class="product-price">${money(product.price)}</span></div>
      <div class="product-category">${escapeHtml(product.category)}</div>
    </article>`,
        )
        .join("")
    : '<p class="empty-state">No pieces found. Try another search.</p>';
  renderIcons();
}

function renderInventory() {
  const query = document
    .querySelector("#inventory-search")
    .value.trim()
    .toLowerCase();
  const stockFilter = document.querySelector("#stock-filter").value;
  const filtered = products.filter((product) => {
    const matchesQuery =
      product.name.toLowerCase().includes(query) ||
      product.category.toLowerCase().includes(query);
    const status = statusFor(product.stock);
    const matchesStock =
      stockFilter === "all" ||
      (stockFilter === "low" && status === "Low stock") ||
      (stockFilter === "out" && status === "Out of stock") ||
      (stockFilter === "healthy" && status === "In stock");
    return matchesQuery && matchesStock;
  });
  document.querySelector("#metric-products").textContent = String(
    products.length,
  ).padStart(2, "0");
  document.querySelector("#metric-categories").textContent = new Set(
    products.map((product) => product.category),
  ).size;
  document.querySelector("#metric-value").textContent = money(
    products.reduce(
      (total, product) => total + product.price * product.stock,
      0,
    ),
  );
  document.querySelector("#metric-low").textContent = String(
    products.filter((product) => product.stock > 0 && product.stock < 8).length,
  ).padStart(2, "0");
  document.querySelector("#metric-out").textContent = String(
    products.filter((product) => product.stock === 0).length,
  ).padStart(2, "0");
  document.querySelector("#table-product-count").textContent =
    `(${filtered.length})`;
  document.querySelector("#inventory-table-body").innerHTML = filtered
    .map((product) => {
      const status = statusFor(product.stock);
      const statusClass =
        status === "Low stock" ? "low" : status === "Out of stock" ? "out" : "";
      return `<tr>
      <td><div class="table-product"><img src="${escapeHtml(product.image)}" alt="" loading="lazy" /><span>${escapeHtml(product.name)}</span></div></td>
      <td>${escapeHtml(product.category)}</td><td>${money(product.price)}</td><td class="stock-cell">${product.stock} units</td>
      <td><span class="status-pill ${statusClass}">${status}</span></td>
      <td><div class="row-actions"><button class="row-action" data-adjust-stock="${product.id}" data-change="-1" aria-label="Remove one ${escapeHtml(product.name)} from stock" title="Remove one unit"><i data-lucide="minus"></i></button><button class="row-action" data-adjust-stock="${product.id}" data-change="1" aria-label="Add one ${escapeHtml(product.name)} to stock" title="Add one unit"><i data-lucide="plus"></i></button></div></td>
    </tr>`;
    })
    .join("");
  document
    .querySelector("#empty-inventory")
    .classList.toggle("hidden", filtered.length > 0);
  renderIcons();
}

function renderBag() {
  const totalCount = bag.reduce((total, item) => total + item.quantity, 0);
  const subtotal = bag.reduce(
    (total, item) => total + item.price * item.quantity,
    0,
  );
  document.querySelector("#bag-count").textContent = totalCount;
  document.querySelector("#bag-panel-count").textContent = `(${totalCount})`;
  document.querySelector("#bag-subtotal").textContent = money(subtotal);
  document.querySelector("#bag-items").innerHTML = bag.length
    ? bag
        .map(
          (item) => `
    <div class="bag-row"><img src="${escapeHtml(item.image)}" alt="" /><div><div class="bag-row-name">${escapeHtml(item.name)}</div><div class="bag-row-price">${money(item.price)}</div><div class="quantity-control"><button data-bag-change="${item.id}" data-change="-1" aria-label="Remove one">−</button><span>${item.quantity}</span><button data-bag-change="${item.id}" data-change="1" aria-label="Add one">+</button></div></div><div class="bag-row-total">${money(item.price * item.quantity)}</div></div>`,
        )
        .join("")
    : '<p class="bag-empty">A good thing belongs in every day. Yours is empty for now.</p>';
  renderIcons();
}

function showToast(message) {
  const toast = document.querySelector("#toast");
  toast.textContent = message;
  toast.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove("show"), 2400);
}

function setView(view) {
  const inventory = view === "inventory";
  document.querySelector("#shop-view").classList.toggle("hidden", inventory);
  document.querySelector(".site-footer").classList.toggle("hidden", inventory);
  document
    .querySelector("#inventory-view")
    .classList.toggle("hidden", !inventory);
  if (inventory) renderInventory();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function closePanels() {
  document.querySelector("#overlay").classList.add("hidden");
  document.querySelector("#bag-panel").classList.add("hidden");
  document.querySelector("#product-modal").classList.add("hidden");
}

function openBag() {
  renderBag();
  document.querySelector("#overlay").classList.remove("hidden");
  document.querySelector("#bag-panel").classList.remove("hidden");
}

function openProductModal() {
  document.querySelector("#overlay").classList.remove("hidden");
  document.querySelector("#product-modal").classList.remove("hidden");
  document.querySelector('#product-form [name="name"]').focus();
}

document.querySelectorAll(".nav-link").forEach((link) =>
  link.addEventListener("click", (event) => {
    if (
      link.getAttribute("href") === "#shop" &&
      !document.querySelector("#inventory-view").classList.contains("hidden")
    ) {
      event.preventDefault();
      setView("shop");
    }
    document.querySelector(".main-nav").classList.remove("open");
  }),
);
document
  .querySelector("#inventory-toggle")
  .addEventListener("click", () => setView("inventory"));
document
  .querySelector("#back-to-shop")
  .addEventListener("click", () => setView("shop"));
document.querySelector("#bag-toggle").addEventListener("click", openBag);
document.querySelector("#overlay").addEventListener("click", closePanels);
document
  .querySelectorAll("[data-close-panel], #close-modal, #cancel-modal")
  .forEach((button) => button.addEventListener("click", closePanels));
document
  .querySelector("#mobile-menu")
  .addEventListener("click", () =>
    document.querySelector(".main-nav").classList.toggle("open"),
  );
document.querySelectorAll(".filter-chip").forEach((button) =>
  button.addEventListener("click", () => {
    activeFilter = button.dataset.filter;
    document
      .querySelectorAll(".filter-chip")
      .forEach((chip) => chip.classList.toggle("selected", chip === button));
    renderProducts();
  }),
);
document
  .querySelector("#sort-select")
  .addEventListener("change", renderProducts);
document.querySelector(".search-toggle").addEventListener("click", () => {
  const search = document.querySelector(".search-toggle");
  if (search.dataset.query !== undefined) {
    delete search.dataset.query;
    search.title = "Search";
    search.setAttribute("aria-label", "Search products");
    document
      .querySelector("#collection")
      .scrollIntoView({ behavior: "smooth" });
    renderProducts();
    return;
  }
  const query = window.prompt("Search the collection");
  if (query !== null) {
    search.dataset.query = query;
    search.title = query ? `Searching: ${query} (click to clear)` : "Search";
    search.setAttribute(
      "aria-label",
      query ? `Clear search for ${query}` : "Search products",
    );
    document
      .querySelector("#collection")
      .scrollIntoView({ behavior: "smooth" });
    renderProducts();
  }
});
document
  .querySelector("#inventory-search")
  .addEventListener("input", renderInventory);
document
  .querySelector("#stock-filter")
  .addEventListener("change", renderInventory);
document
  .querySelector("#add-product-button")
  .addEventListener("click", openProductModal);
document.querySelector("#product-form").addEventListener("submit", (event) => {
  event.preventDefault();
  const formData = new FormData(event.currentTarget);
  const name = String(formData.get("name")).trim();
  const image = String(formData.get("image")).trim();
  products.unshift({
    id: Date.now(),
    name,
    category: String(formData.get("category")),
    price: Number(formData.get("price")),
    stock: Number(formData.get("stock")),
    tag: "New",
    image:
      image ||
      "https://images.unsplash.com/photo-1490312278390-ab64016e0aa9?auto=format&fit=crop&w=760&q=80",
  });
  saveProducts();
  renderInventory();
  renderProducts();
  event.currentTarget.reset();
  closePanels();
  showToast(`${name} added to your catalog`);
});
document
  .querySelector("#inventory-table-body")
  .addEventListener("click", (event) => {
    const button = event.target.closest("[data-adjust-stock]");
    if (!button) return;
    const product = products.find(
      (item) => item.id === Number(button.dataset.adjustStock),
    );
    product.stock = Math.max(0, product.stock + Number(button.dataset.change));
    saveProducts();
    renderInventory();
    renderProducts();
  });
document.querySelector("#product-grid").addEventListener("click", (event) => {
  const button = event.target.closest("[data-add-to-bag]");
  if (!button) return;
  const product = products.find(
    (item) => item.id === Number(button.dataset.addToBag),
  );
  if (!product || product.stock < 1) return;
  const existing = bag.find((item) => item.id === product.id);
  if (existing) existing.quantity += 1;
  else bag.push({ ...product, quantity: 1 });
  product.stock -= 1;
  saveProducts();
  saveBag();
  renderProducts();
  renderBag();
  renderInventory();
  showToast(`${product.name} added to your bag`);
});
document.querySelector("#bag-items").addEventListener("click", (event) => {
  const button = event.target.closest("[data-bag-change]");
  if (!button) return;
  const id = Number(button.dataset.bagChange);
  const item = bag.find((entry) => entry.id === id);
  const product = products.find((entry) => entry.id === id);
  const change = Number(button.dataset.change);
  if (change > 0 && product.stock < 1) {
    showToast("No more stock available");
    return;
  }
  item.quantity += change;
  product.stock -= change;
  if (item.quantity < 1) bag = bag.filter((entry) => entry.id !== id);
  saveProducts();
  saveBag();
  renderBag();
  renderProducts();
  renderInventory();
});
document.querySelector("#checkout-button").addEventListener("click", () => {
  if (!bag.length) {
    showToast("Your bag is waiting for something good");
    return;
  }
  showToast("Checkout is ready to connect to your payment provider");
});
document
  .querySelector("#newsletter-form")
  .addEventListener("submit", (event) => {
    event.preventDefault();
    document.querySelector("#newsletter-message").textContent =
      "You’re on the list. See you soon.";
    event.currentTarget.reset();
  });

renderProducts();
renderBag();
renderIcons();
