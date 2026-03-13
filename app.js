/**
 * VibeGuard Facturation App
 * Logic for simulating the invoice entry flow with dynamic pricing.
 * Updated for Multi-Product Support & Supabase Integration.
 */

// --- Supabase Client & Configuration ---
// Note: Webhook URL updated to production URL for CORS and persistence
let supabaseClient = null;

async function initSupabase() {
    console.log('Initializing Supabase connection...');
    const env = await window.loadEnv();
    if (!env || !env.VITE_SUPABASE_URL || !env.VITE_SUPABASE_ANON_KEY) {
        console.error('Supabase credentials missing in .env or .env could not be loaded');
        return false;
    }
    supabaseClient = window.supabase.createClient(env.VITE_SUPABASE_URL, env.VITE_SUPABASE_ANON_KEY);
    state.webhookUrl = env.VITE_N8N_WEBHOOK_URL;
    console.log('Supabase client created.');
    return true;
}

// --- Application State ---
const state = {
    customers: [],
    products: [],
    selectedCustomerId: null,

    // Current Selection
    selectedProducts: [], // Array of { id, name, quantity, ht, ttc, tva, isStandard }

    cart: [], // Array of { id, clientId, clientName, productId, productName, quantity, unitPrice, totalPrice, isStandard }

    // Preview Controls
    previewSearchTerm: '',
    previewSortOrder: 'asc', // 'asc' or 'desc'

    isLoading: false,
    isAllExpanded: false,
    expandedClients: new Set()
};

// --- DOM Elements ---
const dom = {
    clientSelect: document.getElementById('clientSelect'),
    productSearch: document.getElementById('productSearch'),
    productDropdownList: document.getElementById('productDropdownList'),
    selectedProductsList: document.getElementById('selectedProductsList'),
    dateInput: document.getElementById('dateInput'),

    // Line Preview
    addBtn: document.getElementById('addBtn'),
    linePriceDisplay: document.getElementById('linePriceDisplay'),

    // Cart Display
    itemsContainer: document.getElementById('itemsContainer'),
    invoiceList: document.getElementById('invoiceList'),
    grandTotalDisplay: document.getElementById('grandTotalDisplay'),

    submitBtn: document.getElementById('submitBtn'),
    invoiceForm: document.getElementById('invoiceForm'),
    toast: document.getElementById('toast'),
    toggleAllBtn: document.getElementById('toggleAllBtn'),
    previewSearchInput: document.getElementById('previewSearchInput'),
    sortPreviewBtn: document.getElementById('sortPreviewBtn'),

    // Excel Import
    dropZone: document.getElementById('dropZone'),
    excelInput: document.getElementById('excelInput'),

    // Modal
    modal: document.getElementById('customModal'),
    modalTitle: document.getElementById('modalTitle'),
    modalMsg: document.getElementById('modalMessage'),
    modalCloseBtn: document.getElementById('modalCloseBtn'),

    // Authentication Elements
    authContainer: document.getElementById('authContainer'),
    appLayout: document.getElementById('appLayout'),
    loginBtn: document.getElementById('loginBtn'),
    loginError: document.getElementById('loginError'),
    logoutBtn: document.getElementById('logoutBtn'),
    userEmailDisplay: document.getElementById('userEmailDisplay')
};

// --- Initialization ---
async function init() {
    console.log('VibeGuard App Initialized 🚀');

    // Set default date to today
    dom.dateInput.valueAsDate = new Date();

    const isSupabaseReady = await initSupabase();
    if (isSupabaseReady) {
        // Authentifier ou demander la connexion
        const { data: { session } } = await supabaseClient.auth.getSession();

        if (session) {
            const email = session.user.email;
            if (email && (email.endsWith('@theoma.fr') || email.endsWith('@theoma.com'))) {
                // Déjà connecté avec un bon email
                if (dom.userEmailDisplay) dom.userEmailDisplay.textContent = email;
                showMainApp();
                await fetchData();
                populateSelects();
            } else {
                // Force logout of mismatched domain from previous sessions
                await supabaseClient.auth.signOut();
                showAuthScreen();
                dom.loginError.textContent = "Accès réservé aux adresses @theoma.fr ou @theoma.com.";
                dom.loginError.classList.remove('hidden');
            }
        } else {
            // Pas de session, on garde l'écran de login
            console.log("Aucune session trouvée. En attente de connexion...");

            // Écouter les changements d'état d'authentification (si login réussi plus tard)
            supabaseClient.auth.onAuthStateChange(async (event, session) => {
                if (event === 'SIGNED_IN' && session) {
                    const email = session.user.email;

                    // Vérification du domaine ("theoma.fr" ou "theoma.com")
                    if (email && (email.endsWith('@theoma.fr') || email.endsWith('@theoma.com'))) {
                        if (dom.userEmailDisplay) dom.userEmailDisplay.textContent = email;
                        showMainApp();
                        await fetchData();
                        populateSelects();
                    } else {
                        // Accès refusé : mauvais domaine
                        console.warn(`Tentative de connexion refusée pour l'email: ${email}`);
                        await supabaseClient.auth.signOut();
                        showAuthScreen();
                        dom.loginError.textContent = "Accès réservé aux adresses @theoma.fr ou @theoma.com.";
                        dom.loginError.classList.remove('hidden');
                    }
                } else if (event === 'SIGNED_OUT') {
                    showAuthScreen();
                }
            });
        }
    } else {
        const env = await window.loadEnv();
        let errorMsg = "⚠️ Problème de configuration Supabase";
        if (!env || Object.keys(env).length === 0) {
            errorMsg += "\n\nLe fichier 'env.local.txt' est introuvable ou illisible.";
        } else {
            if (!env.VITE_SUPABASE_URL) errorMsg += "\n- URL Supabase manquante";
            if (!env.VITE_SUPABASE_ANON_KEY) errorMsg += "\n- Clé API manquante";
        }
        alert(errorMsg + "\n\nOuvrez la console (F12) pour plus de détails.");
    }

    attachEventListeners();
}

function showMainApp() {
    dom.authContainer.classList.add('hidden');
    dom.appLayout.classList.remove('hidden');
}

function showAuthScreen() {
    dom.authContainer.classList.remove('hidden');
    dom.appLayout.classList.add('hidden');
}

async function handleLogin(e) {
    if (e) e.preventDefault();

    dom.loginBtn.disabled = true;
    dom.loginBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Redirection vers Microsoft...';
    dom.loginError.classList.add('hidden');

    try {
        const { data, error } = await supabaseClient.auth.signInWithOAuth({
            provider: 'azure',
            options: {
                scopes: 'email profile', // <--- Ajout explicite des scopes ici
                // Ensure redirect matches the current environment URL
                redirectTo: window.location.origin + window.location.pathname
            }
        });

        if (error) throw error;
        // La page va être redirigée vers la page de login Microsoft.
    } catch (err) {
        console.error("Login Error:", err);
        dom.loginError.textContent = "Erreur de connexion Microsoft. Vérifiez la configuration de l'application.";
        dom.loginError.classList.remove('hidden');
        dom.loginBtn.disabled = false;
        dom.loginBtn.innerHTML = '<i class="fa-brands fa-microsoft" style="font-size: 1.2rem;"></i> Se connecter avec Microsoft';
    }
}

async function handleLogout() {
    console.log("Logout button clicked");
    try {
        await supabaseClient.auth.signOut();
        console.log("Supabase signOut successful");
    } catch (err) {
        console.error("Logout error:", err);
    }
}

async function fetchData() {
    console.log('Fetching data from Supabase...');
    state.isLoading = true;
    try {
        // 1. Récupérer l'utilisateur actuel
        const { data: { user }, error: userErr } = await supabaseClient.auth.getUser();
        if (userErr || !user) throw new Error("Utilisateur non authentifié");

        // 2. Récupérer les catégories autorisées pour cet utilisateur
        const { data: allowedCats, error: catErr } = await supabaseClient
            .from('user_allowed_categories')
            .select('category_id')
            .eq('user_id', user.id);

        if (catErr) {
            console.error('Fetch Allowed Categories Error:', catErr);
            throw catErr;
        }

        const allowedCategoryIds = allowedCats.map(c => c.category_id);

        // Fetch Customers
        const { data: customers, error: customerErr } = await supabaseClient
            .from('customers')
            .select('id, name, siren, pennylane_id')
            .order('name');

        if (customerErr) {
            console.error('Fetch Customers Error:', customerErr);
            throw customerErr;
        }
        console.log('Successfully fetched', customers?.length, 'customers');
        state.customers = customers || [];

        // Fetch ALL Products (non archivés)
        const { data: allProducts, error: prodErr } = await supabaseClient
            .from('products')
            .select('*')
            .is('archived_at', null)
            .order('label');

        if (prodErr) {
            console.error('Fetch Products Error:', prodErr);
            throw prodErr;
        }

        // Mark products as allowed or not based on user's categories
        state.products = (allProducts || []).map(p => ({
            ...p,
            isAllowed: allowedCategoryIds.length === 0 ? false : allowedCategoryIds.includes(p.category_id)
        }));

        console.log(`Successfully fetched ${state.products.length} products (${state.products.filter(p=>p.isAllowed).length} allowed)`);

    } catch (err) {
        console.error('Detailed fetch error:', err);
        alert('Erreur SQL Supabase : ' + (err.message || 'Problème de connexion'));
    } finally {
        state.isLoading = false;
    }
}

function populateSelects() {
    // Populate Customers
    state.customers.forEach(customer => {
        const option = document.createElement('option');
        option.value = customer.id;
        option.textContent = customer.name;
        dom.clientSelect.appendChild(option);
    });

    // Init Dropdown (empty for now)
    dom.productDropdownList.innerHTML = '';
}

function attachEventListeners() {
    dom.clientSelect.addEventListener('change', async (e) => {
        state.selectedCustomerId = e.target.value;
        const productsArea = document.getElementById('productSelectionArea');
        if (state.selectedCustomerId) {
            productsArea.style.opacity = '1';
            productsArea.style.pointerEvents = 'auto';
        } else {
            productsArea.style.opacity = '0.5';
            productsArea.style.pointerEvents = 'none';
        }
        await updateSelectedProductsPrices();
    });

    // Custom Dropdown Logic
    dom.productSearch.addEventListener('click', () => {
        if (!state.selectedCustomerId) return;
        renderDropdownList(dom.productSearch.value);
        dom.productDropdownList.classList.remove('hidden');
    });

    dom.productSearch.addEventListener('input', (e) => {
        if (!state.selectedCustomerId) return;
        renderDropdownList(e.target.value);
        dom.productDropdownList.classList.remove('hidden');
    });

    // Hide dropdown when clicking outside
    document.addEventListener('click', (e) => {
        if (!dom.productSearch.contains(e.target) && !dom.productDropdownList.contains(e.target)) {
            dom.productDropdownList.classList.add('hidden');
        }
    });

    dom.addBtn.addEventListener('click', addItemToCart);
    dom.invoiceForm.addEventListener('submit', handleFormSubmit);

    // Auth Listeners
    if (dom.loginBtn) {
        dom.loginBtn.addEventListener('click', handleLogin);
    }
    if (dom.logoutBtn) {
        dom.logoutBtn.addEventListener('click', handleLogout);
    }

    // Excel Import
    if (dom.dropZone) {
        dom.dropZone.addEventListener('click', () => dom.excelInput.click());

        dom.dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dom.dropZone.classList.add('drag-over');
        });

        dom.dropZone.addEventListener('dragleave', () => {
            dom.dropZone.classList.remove('drag-over');
        });

        dom.dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dom.dropZone.classList.remove('drag-over');

            if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                dom.excelInput.files = e.dataTransfer.files;
                handleExcelImport({ target: dom.excelInput });
            }
        });
    }

    dom.excelInput.addEventListener('change', handleExcelImport);

    // Toggle All Accordions
    if (dom.toggleAllBtn) {
        dom.toggleAllBtn.addEventListener('click', toggleAllAccordions);
    }

    // Preview Controls
    if (dom.previewSearchInput) {
        dom.previewSearchInput.addEventListener('input', (e) => {
            state.previewSearchTerm = e.target.value.toLowerCase();
            renderCart();
        });
    }

    if (dom.sortPreviewBtn) {
        dom.sortPreviewBtn.addEventListener('click', () => {
            state.previewSortOrder = state.previewSortOrder === 'asc' ? 'desc' : 'asc';
            // Update icon
            const icon = dom.sortPreviewBtn.querySelector('i');
            if (state.previewSortOrder === 'asc') {
                icon.className = 'fa-solid fa-arrow-down-a-z';
            } else {
                icon.className = 'fa-solid fa-arrow-up-z-a';
            }
            renderCart();
        });
    }

    // Modal
    dom.modalCloseBtn.addEventListener('click', closeModal);
}

function toggleAllAccordions() {
    state.isAllExpanded = !state.isAllExpanded;

    const allItems = document.querySelectorAll('.client-group-items');
    const allIcons = document.querySelectorAll('.toggle-icon');

    if (state.isAllExpanded) {
        dom.toggleAllBtn.textContent = 'Tout réduire';
        dom.toggleAllBtn.title = "Tout réduire";

        // Update State
        const uniqueCustomers = new Set(state.cart.map(item => item.customerName || 'Client Inconnu'));
        state.expandedClients = uniqueCustomers;

        // Animate DOM
        allItems.forEach(el => el.classList.remove('collapsed'));
        allIcons.forEach(el => el.classList.remove('rotate'));
    } else {
        dom.toggleAllBtn.textContent = 'Tout développer';
        dom.toggleAllBtn.title = "Tout développer";

        // Update State
        state.expandedClients.clear();

        // Animate DOM
        allItems.forEach(el => el.classList.add('collapsed'));
        allIcons.forEach(el => el.classList.add('rotate'));
    }
}

// --- Logic ---

function resetLineSelection() {
    state.selectedProducts = [];
    dom.productSearch.value = "";
    dom.productDropdownList.classList.add('hidden');
    renderSelectedProductsList();
    calculateCurrentSelectionPrice();
}

function renderDropdownList(filterText = "") {
    dom.productDropdownList.innerHTML = '';

    const lowerFilter = filterText.toLowerCase();

    const sortedProducts = [...state.products]
        .filter(p => p.isAllowed) // ONLY ALLOWED PRODUCTS IN DROP-DOWN
        .sort((a, b) => a.label.localeCompare(b.label));

    const filteredProducts = sortedProducts.filter(p => p.label.toLowerCase().includes(lowerFilter));

    if (filteredProducts.length === 0) {
        dom.productDropdownList.innerHTML = '<div class="dropdown-item" style="color: var(--color-text-muted); pointer-events: none;">Aucun produit trouvé</div>';
        return;
    }

    filteredProducts.forEach(product => {
        const item = document.createElement('div');
        item.className = 'dropdown-item';

        // Highlight if already selected
        if (state.selectedProducts.find(sp => sp.id === product.id)) {
            item.classList.add('active');
            item.innerHTML = `<i class="fa-solid fa-check" style="color: var(--color-primary); margin-right: 8px;"></i>${product.label}`;
        } else {
            item.textContent = product.label;
        }

        item.addEventListener('click', (e) => {
            e.stopPropagation(); // prevent document click from firing immediately
            toggleProductSelection(product);
            dom.productSearch.value = "";
            dom.productDropdownList.classList.add('hidden');
            dom.productSearch.focus(); // Keep focus for easy multiple selection
        });

        dom.productDropdownList.appendChild(item);
    });
}

async function toggleProductSelection(product) {
    if (!state.selectedCustomerId) return;

    const existingIndex = state.selectedProducts.findIndex(p => p.id === product.id);

    if (existingIndex >= 0) {
        // Remove
        state.selectedProducts.splice(existingIndex, 1);
    } else {
        // Add
        state.selectedProducts.push({
            id: product.id,
            name: product.label,
            quantity: 1,
            ht: 0,
            originalHt: 0,
            ttc: 0,
            tva: 0,
            tvaRate: 0,
            isStandard: true,
            userModifiedPrice: false,
            isExceptional: false
        });
    }

    await updateSelectedProductsPrices();
}

async function updateSelectedProductsPrices() {
    if (!state.selectedCustomerId || state.selectedProducts.length === 0) {
        renderSelectedProductsList();
        calculateCurrentSelectionPrice();
        return;
    }

    dom.addBtn.disabled = true;
    dom.addBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i>';

    // Fetch rates for all selected products
    for (const sp of state.selectedProducts) {
        if (!sp.userModifiedPrice) {
            const rate = await getRate(state.selectedCustomerId, sp.id);
            sp.ht = rate.ht;
            sp.originalHt = rate.originalHt || rate.ht;
            sp.ttc = rate.ttc;
            sp.tva = rate.tva;
            sp.tvaRate = rate.tvaRate;
            sp.isStandard = rate.isStandard;
        }
    }

    renderSelectedProductsList();
    calculateCurrentSelectionPrice();
}

function renderSelectedProductsList() {
    dom.selectedProductsList.innerHTML = '<label>Détails des prestations sélectionnées</label>';

    if (state.selectedProducts.length === 0) {
        dom.selectedProductsList.classList.add('hidden');
        return;
    }

    dom.selectedProductsList.classList.remove('hidden');

    state.selectedProducts.forEach((sp, index) => {
        const row = document.createElement('div');
        row.className = 'selected-product-row';
        row.innerHTML = `
            <div class="sp-name">
                <div style="display: flex; align-items: center; gap: 8px;">
                    <i class="fa-solid fa-check-circle"></i>
                    <span>${sp.name}</span>
                    ${!sp.isStandard ? '<span class="specific-price" style="margin-left: 4px;">(Spé.)</span>' : ''}
                </div>
                <button type="button" class="sp-remove-btn" data-index="${index}" title="Retirer">
                    <i class="fa-solid fa-times"></i>
                </button>
            </div>
            <div class="sp-qty">
                <div style="display: flex; align-items: center; gap: 12px; flex: 1;">
                    <label>Prix HT (€)</label>
                    <input type="number" step="0.01" min="0" value="${sp.ht}" data-index="${index}" class="sp-price-input">
                </div>
                <div style="display: flex; align-items: center; gap: 12px;">
                    <label>Qté</label>
                    <input type="number" min="1" value="${sp.quantity}" data-index="${index}" class="sp-qty-input">
                </div>
            </div>
            <div class="sp-exceptional">
                <label class="checkbox-container">
                    <input type="checkbox" class="sp-exceptional-input" data-index="${index}" ${sp.isExceptional ? 'checked' : ''}>
                    <span class="checkmark"></span>
                    Prix exceptionnel ? (Ne sera pas enregistré comme tarif par défaut)
                </label>
            </div>
        `;
        dom.selectedProductsList.appendChild(row);
    });

    // Attach event listeners to quantity inputs
    const qtyInputs = dom.selectedProductsList.querySelectorAll('.sp-qty-input');
    qtyInputs.forEach(input => {
        input.addEventListener('input', (e) => {
            const idx = parseInt(e.target.dataset.index);
            const val = parseInt(e.target.value);
            if (val > 0) {
                state.selectedProducts[idx].quantity = val;
                calculateCurrentSelectionPrice();
            }
        });
    });

    // Attach event listeners to price inputs
    const priceInputs = dom.selectedProductsList.querySelectorAll('.sp-price-input');
    priceInputs.forEach(input => {
        input.addEventListener('input', (e) => {
            const idx = parseInt(e.target.dataset.index);
            const newHt = parseFloat(e.target.value);

            if (!isNaN(newHt) && newHt >= 0) {
                const sp = state.selectedProducts[idx];
                sp.ht = newHt;
                // Recalculate TVA and TTC based on the preserved tvaRate
                sp.tva = newHt * sp.tvaRate;
                sp.ttc = newHt + sp.tva;

                sp.isStandard = false;
                sp.userModifiedPrice = true;

                calculateCurrentSelectionPrice();
                renderSelectedProductsList(); // Added re-render to update the "Exceptionnel" status immediately if needed
            }
        });
    });

    // Attach event listeners to exceptional checkboxes
    const exceptionalInputs = dom.selectedProductsList.querySelectorAll('.sp-exceptional-input');
    exceptionalInputs.forEach(input => {
        input.addEventListener('change', (e) => {
            const idx = parseInt(e.target.dataset.index);
            state.selectedProducts[idx].isExceptional = e.target.checked;
            calculateCurrentSelectionPrice(); // Update the line preview tag
        });
    });

    // Attach event listeners to remove buttons
    const removeBtns = dom.selectedProductsList.querySelectorAll('.sp-remove-btn');
    removeBtns.forEach(btn => {
        btn.addEventListener('click', (e) => {
            const idx = parseInt(e.currentTarget.dataset.index);
            state.selectedProducts.splice(idx, 1);
            updateSelectedProductsPrices(); // Re-render and recalculate
        });
    });
}

async function getRate(clientId, productId) {
    const product = state.products.find(p => p.id === productId);
    if (!product) return { ht: 0, ttc: 0, tva: 0, isStandard: true };

    // Standard prices from product
    const stdHt = parseFloat(product.price_before_tax) || 0;
    const stdTtc = parseFloat(product.price) || 0;
    let stdTva = stdTtc - stdHt;
    if (stdTva < 0) stdTva = 0;

    // Calcul du taux de TVA standard pour l'appliquer au prix spécifique si nécessaire
    let tvaRate = 0;
    if (stdHt > 0) {
        tvaRate = stdTva / stdHt;
    }

    // 1. Chercher d'abord un prix spécifique (customer_pricings)
    const { data: specificPrice, error } = await supabaseClient
        .from('customer_pricings')
        .select('custom_price_ht')
        .eq('customer_id', clientId)
        .eq('product_id', productId)
        .maybeSingle();

    if (error) console.error('Error checking specific price:', error);

    if (specificPrice && specificPrice.custom_price_ht != null) {
        const customHt = parseFloat(specificPrice.custom_price_ht);
        const customTva = customHt * tvaRate;
        const customTtc = customHt + customTva;
        return { ht: customHt, ttc: customTtc, tva: customTva, isStandard: false, tvaRate, originalHt: stdHt };
    }

    // 2. Sinon, prendre le prix par défaut du produit
    return { ht: stdHt, ttc: stdTtc, tva: stdTva, isStandard: true, tvaRate, originalHt: stdHt };
}

function calculateCurrentSelectionPrice() {
    if (state.selectedProducts.length === 0 || !state.selectedCustomerId) {
        dom.linePriceDisplay.classList.add('hidden');
        dom.addBtn.disabled = true;
        dom.addBtn.innerHTML = '<i class="fa-solid fa-plus"></i> Ajouter la sélection';
        return;
    }

    let totalHt = 0;
    state.selectedProducts.forEach(sp => {
        totalHt += sp.ht * sp.quantity;
    });

    const display = dom.linePriceDisplay;
    display.classList.remove('hidden');

    const valueEl = display.querySelector('.value');
    const tagEl = display.querySelector('.tag');

    valueEl.textContent = `${formatCurrency(totalHt)} HT`;

    const hasSpecific = state.selectedProducts.some(sp => !sp.isStandard);
    const hasExceptional = state.selectedProducts.some(sp => sp.isExceptional);

    if (hasExceptional) {
        tagEl.textContent = "Exceptionnel";
        tagEl.style.backgroundColor = "var(--color-danger)";
        tagEl.style.color = "white";
    } else if (!hasSpecific) {
        tagEl.textContent = "Standard";
        tagEl.style.backgroundColor = "var(--color-primary)";
        tagEl.style.color = "white";
    } else {
        tagEl.textContent = "Spécifique";
        tagEl.style.backgroundColor = "var(--color-secondary)";
        tagEl.style.color = "var(--color-text-on-yellow)";
    }

    dom.addBtn.disabled = false;
    dom.addBtn.innerHTML = '<i class="fa-solid fa-plus"></i> Ajouter la sélection';
}

// --- Cart Management ---

function addItemToCart() {
    if (state.selectedProducts.length === 0 || !state.selectedCustomerId) return;

    const customer = state.customers.find(c => c.id === state.selectedCustomerId);

    state.selectedProducts.forEach(sp => {
        const qte = sp.quantity;
        const lineTotalHt = sp.ht * qte;
        const lineTotalTtc = sp.ttc * qte;
        const lineTotalTva = sp.tva * qte;

        const newItem = {
            id: String(Date.now() + Math.random()), // String ID
            customerId: state.selectedCustomerId,
            customerName: customer.name,
            pennylaneId: customer.pennylane_id,
            productId: sp.id,
            productName: sp.name,
            quantity: qte,
            unitPriceHt: sp.ht,
            unitPriceTtc: sp.ttc,
            unitPriceTva: sp.tva,
            totalPriceHt: lineTotalHt,
            totalPriceTtc: lineTotalTtc,
            totalPriceTva: lineTotalTva,
            isStandard: sp.isStandard,
            userModifiedPrice: sp.userModifiedPrice,
            isExceptional: sp.isExceptional
        };

        state.cart.push(newItem);
    });

    // Auto-expand the customer so the user sees the new item immediately
    state.expandedClients.add(customer.name);

    renderCart();
    resetLineSelection();
}

function removeCartItem(itemId) {
    state.cart = state.cart.filter(item => String(item.id) !== String(itemId));
    renderCart();
}

function toggleEditCartItem(itemId) {
    console.log("toggleEditCartItem called for:", itemId);
    const item = state.cart.find(i => String(i.id) === String(itemId));
    if (item) {
        item.isEditing = !item.isEditing;
        renderCart();
    } else {
        console.warn("Item not found for editing:", itemId);
    }
}

function updateCartItem(itemId, field, value) {
    const item = state.cart.find(i => String(i.id) === String(itemId));
    if (!item) return;

    if (field === 'quantity') {
        item.quantity = parseInt(value) || 1;
    } else if (field === 'unitPriceHt') {
        item.unitPriceHt = parseFloat(value) || 0;
        item.isStandard = false; // Mark as custom if edited
    } else if (field === 'isExceptional') {
        item.isExceptional = value;
    }

    // Recalculate item totals
    item.totalPriceHt = item.quantity * item.unitPriceHt;
    item.totalPriceTva = item.totalPriceHt * 0.20; // 20% TVA
    item.totalPriceTtc = item.totalPriceHt + item.totalPriceTva;

    renderCart();
}

function clearCart() {
    state.cart = [];
    renderCart();
}

// --- Modal ---

function showModal(title, message) {
    dom.modalTitle.textContent = title;
    dom.modalMsg.innerHTML = message; // Changed to innerHTML to support structured reports
    dom.modal.classList.remove('hidden');
}

function closeModal() {
    const modalContent = dom.modal.querySelector('.modal-content');
    dom.modal.classList.add('closing');
    modalContent.classList.add('closing');

    // Wait for animation to finish (300ms)
    setTimeout(() => {
        dom.modal.classList.add('hidden');
        dom.modal.classList.remove('closing');
        modalContent.classList.remove('closing');
    }, 300);
}

function renderCart() {
    dom.invoiceList.innerHTML = '';
    const itemCountEl = document.getElementById('itemCountDisplay');
    const wrapper = document.querySelector('.layout-wrapper');

    if (state.cart.length === 0) {
        dom.itemsContainer.classList.add('hidden');
        wrapper?.classList.remove('has-sidebar');
        dom.submitBtn.disabled = true;
        if (itemCountEl) itemCountEl.textContent = "0 prestation(s)";
        return;
    }

    dom.itemsContainer.classList.remove('hidden');
    wrapper?.classList.add('has-sidebar');
    dom.submitBtn.disabled = false;
    if (itemCountEl) itemCountEl.textContent = `${state.cart.length} prestation(s)`;

    let grandTotalHt = 0;
    let grandTotalTva = 0;
    let grandTotalTtc = 0;

    // Group items by customer
    const groupedItems = state.cart.reduce((groups, item) => {
        const customerName = item.customerName || 'Client Inconnu';
        if (!groups[customerName]) groups[customerName] = [];
        groups[customerName].push(item);
        return groups;
    }, {});

    // Filter and Sort Customers
    let customerNames = Object.keys(groupedItems);

    if (state.previewSearchTerm) {
        customerNames = customerNames.filter(name =>
            name.toLowerCase().includes(state.previewSearchTerm)
        );
    }

    customerNames.sort((a, b) => {
        if (state.previewSortOrder === 'asc') {
            return a.localeCompare(b);
        } else {
            return b.localeCompare(a);
        }
    });

    customerNames.forEach(customerName => {
        const customerItems = groupedItems[customerName];

        // Create customer group container
        const clientGroup = document.createElement('div');
        clientGroup.className = 'client-group';

        // Check individual expansion state
        const isExpanded = state.expandedClients.has(customerName);

        // Create customer header
        const clientHeader = document.createElement('div');
        clientHeader.className = 'client-group-header';
        clientHeader.innerHTML = `
            <div class="header-left">
                <i class="fa-solid fa-building"></i>
                <span>${customerName}</span>
            </div>
            <i class="fa-solid fa-chevron-down toggle-icon ${!isExpanded ? 'rotate' : ''}"></i>
        `;

        // Items container for this client
        const itemsList = document.createElement('ul');
        itemsList.className = `invoice-list client-group-items ${!isExpanded ? 'collapsed' : ''}`;

        clientHeader.onclick = () => {
            const isCollapsing = !itemsList.classList.contains('collapsed');
            if (isCollapsing) {
                state.expandedClients.delete(customerName);
            } else {
                state.expandedClients.add(customerName);
            }
            itemsList.classList.toggle('collapsed');
            clientHeader.querySelector('.toggle-icon').classList.toggle('rotate');
        };

        clientGroup.appendChild(clientHeader);

        customerItems.forEach(item => {
            grandTotalHt += item.totalPriceHt;
            grandTotalTva += item.totalPriceTva;
            grandTotalTtc += item.totalPriceTtc;

            const li = document.createElement('li');
            li.className = 'invoice-item';

            const isEditing = item.isEditing || false;

            if (isEditing) {
                li.innerHTML = `
                    <div class="item-info">
                        <h4>Édition : ${item.productName}</h4>
                        <div class="details-edit" style="display: flex; gap: 16px; align-items: flex-end; margin-top: 12px; margin-bottom: 8px;">
                            <div class="edit-group" style="display: flex; flex-direction: column; gap: 4px;">
                                <label style="margin:0; font-size: 0.75rem; color: var(--color-text-muted); font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px;">Qté</label>
                                <input type="number" class="cart-edit-input" value="${item.quantity}" 
                                    onchange="updateCartItem('${item.id}', 'quantity', this.value)"
                                    style="width: 70px; padding: 6px 10px;">
                            </div>
                            <div class="edit-group" style="display: flex; flex-direction: column; gap: 4px;">
                                <label style="margin:0; font-size: 0.75rem; color: var(--color-text-muted); font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px;">Prix HT (€)</label>
                                <input type="number" step="0.01" class="cart-edit-input price-input" value="${item.unitPriceHt}" 
                                    onchange="updateCartItem('${item.id}', 'unitPriceHt', this.value)"
                                    style="width: 100px; padding: 6px 10px;">
                            </div>
                        </div>
                        <div class="exceptional-edit" style="margin-top: 8px;">
                            <label class="checkbox-container">
                                <input type="checkbox" onchange="updateCartItem('${item.id}', 'isExceptional', this.checked)" ${item.isExceptional ? 'checked' : ''}>
                                <span class="checkmark"></span>
                                <span style="font-size: 0.8rem;">Prix exceptionnel ?</span>
                            </label>
                        </div>
                    </div>
                    <div class="item-actions">
                        <button class="edit-btn" onclick="toggleEditCartItem('${item.id}')" title="Valider">
                            <i class="fa-solid fa-check" style="color: var(--color-success); font-size: 1.2rem;"></i>
                        </button>
                    </div>
                `;
            } else {
                li.innerHTML = `
                    <div class="item-info">
                        <h4>${item.productName} ${item.isExceptional ? '<span class="status-tag exceptional">Exceptionnel</span>' : ''}</h4>
                        <div class="details">
                            <span>Qté: ${item.quantity} × ${formatCurrency(item.unitPriceHt)} HT</span>
                            ${!item.isStandard && !item.isExceptional ? '<span class="status-tag specific">Tarif Spécifique</span>' : ''}
                        </div>
                    </div>
                    <div class="item-price">
                        <div class="price-row">
                            <span class="label">HT :</span>
                            <span class="value">${formatCurrency(item.totalPriceHt)}</span>
                        </div>
                        <div class="price-row tva-row">
                            <span class="label">TVA :</span>
                            <span class="value">${formatCurrency(item.totalPriceTva)}</span>
                        </div>
                        <div class="price-row total-row">
                            <span class="label">TTC :</span>
                            <span class="value total">${formatCurrency(item.totalPriceTtc)}</span>
                        </div>
                    </div>
                    <div class="item-actions">
                        <button class="edit-btn" onclick="toggleEditCartItem('${item.id}')" title="Modifier">
                            <i class="fa-solid fa-pencil"></i>
                        </button>
                        <button class="delete-btn" onclick="removeCartItem('${item.id}')" title="Supprimer">
                            <i class="fa-solid fa-times"></i>
                        </button>
                    </div>
                `;
            }
            itemsList.appendChild(li);
        });

        clientGroup.appendChild(itemsList);
        dom.invoiceList.appendChild(clientGroup);
    });

    dom.grandTotalDisplay.innerHTML = `
        <div class="global-totals">
            <div class="total-line"><span>Total HT :</span> <span>${formatCurrency(grandTotalHt)}</span></div>
            <div class="total-line tva-line"><span>Total TVA :</span> <span>${formatCurrency(grandTotalTva)}</span></div>
            <div class="total-line grand-total"><span>Total TTC :</span> <span>${formatCurrency(grandTotalTtc)}</span></div>
        </div>
    `;

    // Met à jour le texte du bouton selon le nombre de clients (1 client = 1 brouillon)
    const numCustomers = Object.keys(groupedItems).length;
    const btnText = numCustomers > 1 ? 'les brouillons' : 'le brouillon';
    dom.submitBtn.innerHTML = `<i class="fa-solid fa-paper-plane"></i> Générer ${btnText} de facture sur Pennylane`;
}

window.removeCartItem = removeCartItem;
window.toggleEditCartItem = toggleEditCartItem;
window.updateCartItem = updateCartItem;

function formatCurrency(amount) {
    return new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(amount);
}

// --- Form Submission ---

async function handleFormSubmit(e) {
    e.preventDefault();
    if (state.isLoading || state.cart.length === 0) return;

    state.isLoading = true;
    const btnContent = dom.submitBtn.innerHTML;
    dom.submitBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Envoi en cours...';
    dom.submitBtn.disabled = true;

    try {
        // Group items by customer for mass processing
        const groupedInvoices = state.cart.reduce((acc, item) => {
            if (!acc[item.customerId]) acc[item.customerId] = { items: [], totalHt: 0, totalTtc: 0 };
            acc[item.customerId].items.push(item);
            acc[item.customerId].totalHt += item.totalPriceHt;
            acc[item.customerId].totalTtc += item.totalPriceTtc;
            return acc;
        }, {});

        const customerIds = Object.keys(groupedInvoices);
        console.log(`Processing ${customerIds.length} invoices...`);

        // Prepare n8n Webhook Payload
        const n8nPayload = {
            date_globale: dom.dateInput.value || new Date().toISOString().split('T')[0],
            brouillons: []
        };

        for (const customerId of customerIds) {
            const invoiceData = groupedInvoices[customerId];
            const firstItem = invoiceData.items[0];

            // Add to n8n payload
            n8nPayload.brouillons.push({
                customer_id: customerId,
                pennylane_id: firstItem.pennylaneId,
                customer_name: firstItem.customerName || 'Client Inconnu',
                lignes: invoiceData.items.map(item => ({
                    produit_nom: item.productName,
                    quantite: item.quantity,
                    prix_unitaire_ht: item.unitPriceHt,
                    tva: item.unitPriceTva,
                    is_standard: item.isStandard,
                    is_exceptional: item.isExceptional || false
                })),
                total_ht: invoiceData.totalHt,
                total_ttc: invoiceData.totalTtc
            });

            // 1. Log the invoice in Supabase (Optional but good for history)
            try {
                const { error } = await supabaseClient
                    .from('invoice_logs')
                    .insert({
                        customer_id: customerId,
                        total_ht: invoiceData.totalHt,
                        status: 'pending'
                    });
                if (error) console.warn(`Supabase log error for ${customerId}:`, error);
            } catch (e) { console.warn(e); }

            // 2. Save custom prices to customer_pricing table (only if not exceptional)
            for (const item of invoiceData.items) {
                if (item.userModifiedPrice && !item.isExceptional) {
                    try {
                        console.log(`Saving custom price for ${item.productName}: ${item.unitPriceHt}€`);

                        // Check if entry already exists
                        const { data: existingEntry, error: fetchError } = await supabaseClient
                            .from('customer_pricings')
                            .select('id')
                            .eq('customer_id', customerId)
                            .eq('product_id', item.productId)
                            .maybeSingle();

                        if (fetchError) {
                            console.error('Error fetching existing custom price:', fetchError);
                            continue;
                        }

                        if (existingEntry) {
                            // Update existing entry
                            const { error: updateError } = await supabaseClient
                                .from('customer_pricings')
                                .update({
                                    custom_price_ht: item.unitPriceHt,
                                    updated_at: new Date().toISOString()
                                })
                                .eq('id', existingEntry.id);

                            if (updateError) console.error('Error updating custom price:', updateError);
                        } else {
                            // Insert new entry
                            const { error: insertError } = await supabaseClient
                                .from('customer_pricings')
                                .insert({
                                    customer_id: customerId,
                                    product_id: item.productId,
                                    custom_price_ht: item.unitPriceHt
                                });

                            if (insertError) console.error('Error inserting custom price:', insertError);
                        }
                    } catch (e) {
                        console.error('Error in custom pricing logic:', e);
                    }
                }
            }
        }

        // Send to n8n Webhook
        if (state.webhookUrl) {
            try {
                console.log("Sending payload to n8n:", n8nPayload);
                const webhookRes = await fetch(state.webhookUrl, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(n8nPayload)
                });

                if (!webhookRes.ok) {
                    console.error("Webhook error:", await webhookRes.text());
                    // Throw error internally to jump to catch block, but avoid blocking the entire process if Supabase succeeded
                    throw new Error("Erreur lors de l'envoi au webhook n8n.");
                }
            } catch (webhookErr) {
                console.error("Erreur Webhook interceptée:", webhookErr);
                alert("La facture a bien été enregistrée pour vos tarifs, mais une erreur est survenue lors de l'envoi vers Pennylane (CORS ou Webhook indisponible).");
                // Reset loading state and UI here since the main try/catch won't catch this without blowing up the whole block
                state.isLoading = false;
                dom.submitBtn.innerHTML = btnContent;
                dom.submitBtn.disabled = false;
                return; // Stop execution here to avoid showing success toast
            }
        } else {
            console.warn("No webhook URL configured. Supabase logs only.");
        }

        // 2. Success
        showToast();

        // Reset everything
        // Reset everything
        clearCart();
        dom.invoiceForm.reset();
        resetLineSelection();
        dom.clientSelect.value = "";
        state.selectedCustomerId = null;
        dom.dateInput.valueAsDate = new Date();

    } catch (err) {
        console.error('Submission error:', err);
        alert('Erreur lors de l\'envoi de la facture. Vérifiez la console.');
    } finally {
        dom.submitBtn.innerHTML = btnContent;
        state.isLoading = false;
    }
}

// --- Excel Import Logic ---

async function handleExcelImport(e) {
    const file = e.target.files[0];
    if (!file) return;

    // Loading state for drop zone
    const dropZoneContent = dom.dropZone.querySelector('.drop-zone-content');
    const originalContent = dropZoneContent.innerHTML;

    dropZoneContent.innerHTML = '<i class="fa-solid fa-spinner fa-spin drop-icon"></i><span class="drop-text">Importation en cours...</span>';
    dom.dropZone.style.pointerEvents = 'none'; // Disable clicks during import

    console.log('Reading Excel file:', file.name);
    const reader = new FileReader();

    reader.onload = async (evt) => {
        try {
            const data = evt.target.result;
            const workbook = XLSX.read(data, { type: 'binary' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Convert to JSON (array of objects)
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            if (jsonData.length < 2) {
                alert("Le fichier Excel semble vide ou mal structuré.");
                return;
            }

            // 1. Loop through all data rows starting from index 1
            const columnMapping = {
                2: ["Bulletins de paie"],
                3: ["Bulletins de paie", "Formalités d'entrée de collaborateurs", "Formalités de sortie de collaborateurs"],
                4: ["DSN évènement"],
                5: ["Formalités de sortie de collaborateurs"],
                6: ["Formalités d'entrée de collaborateurs"],
                7: ["Contrat de travail sans clause spécifique"],
                8: ["Avenant"]
            };

            const STOP_WORDS = ["de", "la", "le", "du", "en", "un", "une", "au", "aux", "par", "pour", "sur", "avec", "dans", "et"];
            
            const cleanString = (str) => {
                if (!str) return "";
                return str.toLowerCase()
                    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
                    .replace(/[^a-z0-9\s]/g, ' ')
                    .trim();
            };

            const tokenize = (str) => {
                return cleanString(str).split(/\s+/).filter(t => t.length >= 2 && !STOP_WORDS.includes(t));
            };

            const isApprox = (s1, s2) => {
                const n1 = s1; // already clean via tokenize
                const n2 = s2; 
                if (n1 === n2) return true;
                
                // Special aliases
                if (n1 === "dat" && (n2 === "accident" || n2 === "travail")) return true;
                if (n1 === "co" && n2.startsWith("convention")) return true; 
                if ((n1 === "w" || n1 === "work" || n1 === "med") && (n2 === "travail" || n2 === "medecine")) {
                    if (n1 === "w" && n2 === "travail") return true;
                    if (n1 === "med" && n2 === "medecine") return true;
                }

                const t1 = n1.replace(/[sx]$/, '');
                const t2 = n2.replace(/[sx]$/, '');
                if (t1 === t2) return true;
                
                // Prefix matching (min 3 chars strictly)
                if (t1.length >= 3 && t2.startsWith(t1)) return true;
                if (t2.length >= 3 && t1.startsWith(t2)) return true;

                // Fuzzy for long words
                if (t1.length >= 5 && t2.length >= 5 && Math.abs(t1.length - t2.length) <= 1) {
                    let l = t1.length > t2.length ? t1 : t2;
                    let s = t1.length > t2.length ? t2 : t1;
                    let err = 0, si = 0;
                    for (let li = 0; li < l.length && err <= 1; li++) {
                        if (l[li] !== s[si]) {
                            err++;
                            if (l.length === s.length) si++;
                        } else {
                            si++;
                        }
                    }
                    return err <= 1;
                }
                return false;
            };

            const itemsToAdd = [];
            let customersFound = 0;
            const IMPORT_LOGIC_VERSION = "2.0";
            console.log(`[Import Debug] Starting process with version ${IMPORT_LOGIC_VERSION}`);

            const skippedRows = {
                noSiren: [],
                unknownCustomer: [],
                purelyEmpty: [],
                textualRecognized: [],
                textualForbidden: [],
                textualAutoAdded: [],
                textualUnrecognized: []
            };

            for (let r = 1; r < jsonData.length; r++) {
                // Update percentage for user feedback
                const percent = Math.round((r / (jsonData.length - 1)) * 100);
                dropZoneContent.innerHTML = `<i class="fa-solid fa-spinner fa-spin drop-icon"></i><span class="drop-text">Importation ${percent}%...</span>`;

                const row = jsonData[r];
                if (!row || row.length < 1) continue;

                const rowNum = r + 1; // Excel row number

                const clientSiren = String(row[1] || "").trim();
                if (!clientSiren) {
                    skippedRows.noSiren.push(rowNum);
                    continue;
                }

                // Match by SIREN instead of name
                const matchedCustomer = state.customers.find(c =>
                    String(c.siren || "").trim() === clientSiren
                );

                if (!matchedCustomer) {
                    const potentialName = String(row[0] || "").trim();
                    skippedRows.unknownCustomer.push({ row: rowNum, siren: clientSiren, name: potentialName });
                    continue;
                }

                customersFound++;
                let hasMappedColumnWithData = false;
                let rowValid = false;

                // 1. Process standard mapped columns (2-8) for numeric quantities
                for (const [colIndex, productLabels] of Object.entries(columnMapping)) {
                    const idx = parseInt(colIndex);
                    const cellValue = String(row[idx] || "").trim();
                    if (!cellValue || cellValue === "0") continue;

                    const quantity = parseFloat(cellValue.replace(',', '.'));
                    if (!isNaN(quantity) && quantity > 0) {
                        hasMappedColumnWithData = true;
                        
                        for (const label of productLabels) {
                            const matchedProduct = state.products.find(p =>
                                p.label.toLowerCase().trim() === label.toLowerCase().trim()
                            );

                            if (matchedProduct) {
                                rowValid = true;
                                const rate = await getRate(matchedCustomer.id, matchedProduct.id);
                                itemsToAdd.push({
                                    id: String(Date.now() + Math.random()), 
                                    customerId: matchedCustomer.id,
                                    customerName: matchedCustomer.name,
                                    productId: matchedProduct.id,
                                    productName: matchedProduct.label,
                                    quantity: quantity,
                                    unitPriceHt: rate.ht,
                                    unitPriceTtc: rate.ttc,
                                    unitPriceTva: rate.tva,
                                    totalPriceHt: rate.ht * quantity,
                                    totalPriceTtc: rate.ttc * quantity,
                                    totalPriceTva: rate.tva * quantity,
                                    isStandard: rate.isStandard,
                                    rowNum: rowNum
                                });
                            }
                        }
                    }
                }

                // 2. Specific check for "Autre" column (Index 9)
                const autreRaw = String(row[9] || "").trim();
                let autreValue = autreRaw;
                let quantityInAutre = 0;

                if (autreRaw && autreRaw !== "0") {
                    hasMappedColumnWithData = true;

                    // Handle multi-product separator (+ or et)
                    const subParts = autreRaw.split(/\s*(?:\+|\bet\b)\s*/i);

                    for (const rawPart of subParts) {
                        if (!rawPart.trim()) continue;

                        let currentVal = rawPart.trim();
                        let currentQty = 0;

                        // Extract quantity and unit (e.g. "15 mn", "2 h", "1.5 h")
                        const qtyMatch = currentVal.match(/^(\d+[\.,]?\d*)\s*(mn|min|h|hs|heures)?[\s,.]+(.*)/i);
                        if (qtyMatch) {
                            let rawQty = parseFloat(qtyMatch[1].replace(',', '.'));
                            const unit = (qtyMatch[2] || "").toLowerCase();
                            
                            if (unit.startsWith('m')) {
                                // Convert minutes to decimal segments if needed, but here we assume qty is the unit
                                currentQty = rawQty / 60; 
                            } else {
                                currentQty = rawQty;
                            }
                            currentVal = qtyMatch[3].trim();
                        }
                        
                        let bestProduct = null;
                        let maxMatchStrength = -1;
                        let bestConfidence = 0;
                        let bestProductCoverage = 0;
                        let candidates = [];
                        
                        const valTokens = tokenize(currentVal);

                        for (const product of state.products) {
                            const pWords = tokenize(product.label);
                            if (pWords.length === 0) continue;

                            let currentPWordIdx = 0;
                            let matchesFoundInSequence = 0;

                            for (const token of valTokens) {
                                for (let i = currentPWordIdx; i < pWords.length; i++) {
                                    if (isApprox(token, pWords[i])) {
                                        matchesFoundInSequence++;
                                        currentPWordIdx = i + 1; 
                                        break;
                                    }
                                }
                            }

                            if (matchesFoundInSequence > 0) {
                                const coverageProduct = (matchesFoundInSequence / pWords.length);
                                const coverageInput = (matchesFoundInSequence / valTokens.length);
                                
                                // Score = 100% if we cover the shorter set of semantic words
                                const confidence = (matchesFoundInSequence / Math.min(pWords.length, valTokens.length)) * 100;
                                
                                // Tie-breaker: prioritize products where we match more of the actual words
                                const strength = (confidence * 1000) + (coverageProduct * 100) + product.label.length; 
                                
                                candidates.push({ label: product.label, confidence, strength, matches: matchesFoundInSequence, coverage: coverageProduct, isAllowed: product.isAllowed });

                                if (strength > maxMatchStrength) {
                                    maxMatchStrength = strength;
                                    bestProduct = product;
                                    bestConfidence = confidence;
                                    bestProductCoverage = coverageProduct;
                                }
                            }
                        }

                        // RELIABILITY CRITERIA (v2.0):
                        // - Either very high confidence (>85%) 
                        // - Or good confidence (>60%) with either 50%+ product coverage OR 2+ semantic words match
                        const matches = bestProduct ? (candidates.find(c => c.label === bestProduct.label)?.matches || 0) : 0;
                        const isReliable = bestProduct && (
                            (bestConfidence >= 85) || 
                            (bestConfidence >= 60 && (bestProductCoverage >= 0.5 || matches >= 2))
                        );

                        if (rowNum === 58 || rowNum === 70 || rowNum === 2 || (bestProduct && !isReliable)) {
                            console.log(`[Import Debug] ROW ${rowNum} PART "${rawPart.trim()}":`, 
                                { winner: bestProduct?.label, reliable: !!isReliable, confidence: bestConfidence, matches, coverage: bestProductCoverage },
                                candidates.sort((a,b) => b.strength - a.strength).slice(0, 3)
                            );
                        }

                        if (isReliable) {
                            console.log(`[Import Debug] ROW ${rowNum} | Auto-Add: "${bestProduct.label}"`);
                            
                            const finalQty = (currentQty > 0) ? currentQty : 1;
                            const rate = await getRate(matchedCustomer.id, bestProduct.id);
                            itemsToAdd.push({
                                id: String(Date.now() + Math.random()), 
                                customerId: matchedCustomer.id,
                                customerName: matchedCustomer.name,
                                productId: bestProduct.id,
                                productName: bestProduct.label,
                                quantity: finalQty,
                                unitPriceHt: rate.ht,
                                unitPriceTtc: rate.ttc,
                                unitPriceTva: rate.tva,
                                totalPriceHt: rate.ht * finalQty,
                                totalPriceTtc: rate.ttc * finalQty,
                                totalPriceTva: rate.tva * finalQty,
                                isStandard: rate.isStandard,
                                rowNum: rowNum,
                                fromAutre: true
                            });
                            skippedRows.textualAutoAdded.push({ row: rowNum, product: bestProduct.label });
                        } else if (bestProduct) {
                            // Recognition without auto-add (unreliable or forbidden)
                            if (bestProduct.isAllowed) {
                                skippedRows.textualRecognized.push({ row: rowNum, product: bestProduct.label });
                            } else {
                                skippedRows.textualForbidden.push({ row: rowNum, product: bestProduct.label });
                            }
                        } else {
                            skippedRows.textualUnrecognized.push({ row: rowNum, content: rawPart });
                        }
                    }
                }

                // 3. Categorize skipped row
                if (!hasMappedColumnWithData) {
                    skippedRows.purelyEmpty.push(rowNum);
                }
            }

            // Removed the early return to allow showing the report even if nothing was added

            // 3. Update cart (Append instead of Replace)
            state.cart = [...state.cart, ...itemsToAdd];
            renderCart();

            // Success Analysis
            const title = itemsToAdd.length > 0 ? "Importation Terminée" : "Importation Échouée";
            const successfulRowNums = [...new Set(itemsToAdd.map(item => item.rowNum))];
            
            const autoAddedRows = skippedRows.textualAutoAdded.map(m => m.row);
            const recognizedInSuccess = successfulRowNums.filter(num => !autoAddedRows.includes(num) && skippedRows.textualRecognized.some(m => m.row === num));
            const forbiddenInSuccess = successfulRowNums.filter(num => !autoAddedRows.includes(num) && skippedRows.textualForbidden.some(m => m.row === num));
            const unrecognizedInSuccess = successfulRowNums.filter(num => !autoAddedRows.includes(num) && skippedRows.textualUnrecognized.some(m => m.row === num));

            let msg = `<div class="import-summary">
                <div class="summary-item success">
                    <div class="summary-icon"><i class="fa-solid fa-check-circle"></i></div>
                    <div class="summary-content">
                        <strong>${itemsToAdd.length} prestation(s) ajoutée(s)</strong>
                        <span>(${customersFound} clients identifiés)</span>
                        
                        ${skippedRows.textualAutoAdded.length > 0 ? `
                        <div class="summary-sub-container" style="margin-top: 8px; border-top: 1px dashed rgba(16, 185, 129, 0.4); padding-top: 8px;">
                            <div class="sub-reason" style="color: var(--color-success); font-weight: bold; font-size: 0.75rem;">Ajouts automatiques (via colonne "Autre") :</div>
                            <div class="sub-lines" style="color: var(--color-success); opacity: 0.9; font-size: 0.8rem;">Lignes : ${skippedRows.textualAutoAdded.map(m => m.row).sort((a,b) => a-b).join(', ')}<br><small>Produit(s) reconnu(s) et ajouté(s) automatiquement grâce à la valeur textuelle.</small></div>
                        </div>` : ''}

                        ${recognizedInSuccess.length > 0 ? `
                        <div class="summary-sub-container" style="margin-top: 8px; border-top: 1px dashed rgba(16, 185, 129, 0.2); padding-top: 8px;">
                            <div class="sub-reason" style="color: var(--color-success); opacity: 1; font-size: 0.7rem;">Commentaire dans la colonne "Autre" avec produit reconnu :</div>
                            <div class="sub-lines" style="color: var(--color-success); opacity: 0.8; font-size: 0.8rem;">Lignes : ${recognizedInSuccess.sort((a,b) => a-b).join(', ')}<br><small>Merci d'ajouter ce produit en tant que colonne dans l'excel.</small></div>
                        </div>` : ''}

                        ${forbiddenInSuccess.length > 0 ? `
                        <div class="summary-sub-container" style="margin-top: 8px; border-top: 1px dashed rgba(16, 185, 129, 0.2); padding-top: 8px;">
                            <div class="sub-reason" style="color: var(--color-success); opacity: 1; font-size: 0.7rem;">Produit reconnu (hors catégorie autorisée) :</div>
                            <div class="sub-lines" style="color: var(--color-success); opacity: 0.8; font-size: 0.8rem;">Lignes : ${forbiddenInSuccess.sort((a,b) => a-b).join(', ')}<br><small>Merci de contacter l'administrateur pour l'ajouter à vos accès.</small></div>
                        </div>` : ''}

                        ${unrecognizedInSuccess.length > 0 ? `
                        <div class="summary-sub-container" style="margin-top: 8px; border-top: 1px dashed rgba(16, 185, 129, 0.2); padding-top: 8px;">
                            <div class="sub-reason" style="color: var(--color-success); opacity: 1; font-size: 0.7rem;">Commentaire dans la colonne "Autre" avec produit non reconnu :</div>
                            <div class="sub-lines" style="color: var(--color-success); opacity: 0.8; font-size: 0.8rem;">Lignes : ${unrecognizedInSuccess.sort((a,b) => a-b).join(', ')}<br><small>Merci d'ajouter cette mention dans la colonne commentaire et d'ajouter le produit "Divers social" dans une colonne du excel.</small></div>
                        </div>` : ''}
                    </div>
                </div>`;

            if (skippedRows.noSiren.length > 0) {
                msg += `<div class="summary-item error">
                    <div class="summary-icon"><i class="fa-solid fa-triangle-exclamation"></i></div>
                    <div class="summary-content">
                        <strong>SIREN manquant</strong>
                        <span>Lignes : ${skippedRows.noSiren.join(', ')}</span>
                    </div>
                </div>`;
            }

            if (skippedRows.unknownCustomer.length > 0) {
                const sirenList = skippedRows.unknownCustomer.map(item => {
                    const nameStr = item.name ? ` (${item.name})` : '';
                    return `L:${item.row} - ${item.siren}${nameStr}`;
                }).join('<br>');
                
                msg += `<div class="summary-item error">
                    <div class="summary-icon"><i class="fa-solid fa-user-slash"></i></div>
                    <div class="summary-content">
                        <strong>SIREN non reconnus ou absents de la base</strong>
                        <div class="small-list" style="margin-top: 5px; line-height: 1.4;">${sirenList}</div>
                        <div style="margin-top: 8px; font-size: 0.75rem; opacity: 0.8; color: var(--color-danger);">
                            <i class="fa-solid fa-circle-info"></i> Vérifiez que ces clients sont bien synchronisés depuis Pennylane.
                        </div>
                    </div>
                </div>`;
            }

            // Ignored Rows Analysis
            const groupPurelyEmpty = skippedRows.purelyEmpty.sort((a,b) => a-b);
            const groupTextRecognized = skippedRows.textualRecognized.filter(m => !successfulRowNums.includes(m.row)).map(m => m.row).sort((a,b) => a-b);
            const groupTextForbidden = skippedRows.textualForbidden.filter(m => !successfulRowNums.includes(m.row)).map(m => m.row).sort((a,b) => a-b);
            const groupTextUnrecognized = skippedRows.textualUnrecognized.filter(m => !successfulRowNums.includes(m.row)).map(m => m.row).sort((a,b) => a-b);

            if (groupPurelyEmpty.length > 0 || groupTextRecognized.length > 0 || groupTextForbidden.length > 0 || groupTextUnrecognized.length > 0) {
                let sections = [];
                
                if (groupPurelyEmpty.length > 0) {
                    sections.push(`
                        <div class="summary-sub-item">
                            <div class="sub-reason">Lignes à 0 ou vides</div>
                            <div class="sub-lines">Lignes : ${groupPurelyEmpty.join(', ')}</div>
                        </div>`);
                }

                if (groupTextRecognized.length > 0) {
                    sections.push(`
                        <div class="summary-sub-item">
                            <div class="sub-reason">Commentaire dans la colonne "Autre" avec produit reconnu</div>
                            <div class="sub-lines">Lignes : ${groupTextRecognized.join(', ')}<br><small>Merci d'ajouter ce produit en tant que colonne dans l'excel.</small></div>
                        </div>`);
                }

                if (groupTextForbidden.length > 0) {
                    sections.push(`
                        <div class="summary-sub-item">
                            <div class="sub-reason">Produit reconnu (hors catégorie autorisée)</div>
                            <div class="sub-lines">Lignes : ${groupTextForbidden.join(', ')}<br><small>Merci de contacter l'administrateur pour l'ajouter à vos accès.</small></div>
                        </div>`);
                }
                
                if (groupTextUnrecognized.length > 0) {
                    sections.push(`
                        <div class="summary-sub-item">
                            <div class="sub-reason">Commentaire dans la colonne "Autre" avec produit non reconnu</div>
                            <div class="sub-lines">Lignes : ${groupTextUnrecognized.join(', ')}<br><small>Merci d'ajouter cette mention dans la colonne commentaire et d'ajouter le produit "Divers social" dans une colonne du excel.</small></div>
                        </div>`);
                }

                msg += `<div class="summary-item warning">
                    <div class="summary-icon"><i class="fa-solid fa-circle-exclamation"></i></div>
                    <div class="summary-content">
                        <strong>LIGNES IGNORÉES</strong>
                        <div class="summary-sub-container">
                            ${sections.join('')}
                        </div>
                    </div>
                </div>`;
            }

            msg += `</div>`;

            showModal(title, msg);

        } catch (err) {
            console.error('Excel Import Error:', err);
            showModal("Erreur d'Importation", "Erreur lors de la lecture du fichier Excel.");
        } finally {
            dom.excelInput.value = ''; // Reset input
            dropZoneContent.innerHTML = originalContent;
            dom.dropZone.style.pointerEvents = 'auto'; // Re-enable clicks
        }
    };

    reader.readAsBinaryString(file);
}

function showToast() {
    dom.toast.classList.remove('hidden');
    setTimeout(() => {
        dom.toast.classList.add('hidden');
    }, 3000);
}

// Start
document.addEventListener('DOMContentLoaded', init);
