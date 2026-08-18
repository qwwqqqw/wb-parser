let config = {};
let categoriesData = [];

window.addEventListener('pywebviewready', function () {
    loadConfig();
    setupEventListeners();
    fetchCategories();
});

async function loadConfig() {
    try {
        config = await pywebview.api.get_config();
        applyConfigToUI();
    } catch (e) {
        console.error("Failed to load config", e);
    }
}

function applyConfigToUI() {
    document.body.className = `theme-${config.theme}`;
    document.getElementById('theme-switch').checked = config.theme === 'light';

    const modeRadios = document.getElementsByName('mode');
    for (let r of modeRadios) {
        if (r.value === config.mode) r.checked = true;
    }
    toggleSellerInput();

    document.getElementById('seller-url').value = config.seller_url || '';
    document.getElementById('price-min').value = config.price_min || 0;
    document.getElementById('price-max').value = config.price_max || 100000;
    document.getElementById('items-count').value = config.items_count || 100;
    if (config.merge_files !== undefined) {
        document.getElementById('merge-files').checked = config.merge_files;
    }
    if (config.save_images !== undefined) {
        document.getElementById('save-images').checked = config.save_images;
    }
}

async function saveConfig() {
    config.theme = document.getElementById('theme-switch').checked ? 'light' : 'dark';

    const modeRadios = document.getElementsByName('mode');
    for (let r of modeRadios) {
        if (r.checked) config.mode = r.value;
    }

    config.seller_url = document.getElementById('seller-url').value;
    config.price_min = parseInt(document.getElementById('price-min').value) || 0;
    config.price_max = parseInt(document.getElementById('price-max').value) || 100000;
    config.items_count = parseInt(document.getElementById('items-count').value) || 100;
    config.merge_files = document.getElementById('merge-files').checked;
    config.save_images = document.getElementById('save-images').checked;

    config.selected_categories = getSelectedCategories();

    await pywebview.api.save_config(config);
}

function toggleSellerInput() {
    const isSeller = document.querySelector('input[name="mode"]:checked').value === 'seller';
    document.getElementById('seller-url-group').style.display = isSeller ? 'flex' : 'none';
}

function setupEventListeners() {
    document.getElementById('theme-switch').addEventListener('change', (e) => {
        document.body.className = e.target.checked ? 'theme-light' : 'theme-dark';
        saveConfig();
    });
    const modeRadios = document.getElementsByName('mode');
    for (let r of modeRadios) {
        r.addEventListener('change', () => {
            toggleSellerInput();
            saveConfig();
        });
    }

    const inputs = ['seller-url', 'price-min', 'price-max', 'items-count', 'merge-files'];
    inputs.forEach(id => {
        document.getElementById(id).addEventListener('change', saveConfig);
    });

    document.getElementById('btn-start').addEventListener('click', startParsing);
    document.getElementById('btn-stop').addEventListener('click', stopParsing);
    document.getElementById('btn-refresh-cat').addEventListener('click', fetchCategories);

    document.getElementById('btn-open-excel').addEventListener('click', () => {
        const btn = document.getElementById('btn-open-excel');
        if (btn._filepath) {
            pywebview.api.open_excel(btn._filepath);
        }
    });
}

async function fetchCategories() {
    const treeContainer = document.getElementById('categories-tree');
    treeContainer.innerHTML = '<div class="loading">Загрузка категорий...</div>';

    try {
        categoriesData = await pywebview.api.get_categories();
        renderTree(categoriesData, treeContainer);
        restoreCategorySelection();
    } catch (e) {
        treeContainer.innerHTML = `<div style="color: var(--danger-color)">Ошибка загрузки: ${e}</div>`;
    }
}

function renderTree(data, container) {
    const ul = document.createElement('ul');

    data.forEach(node => {
        const li = document.createElement('li');
        li.className = 'tree-node';

        const labelDiv = document.createElement('div');
        labelDiv.className = 'tree-label';

        const hasChildren = node.children && node.children.length > 0;
        const toggleBtn = document.createElement('span');
        toggleBtn.className = 'toggle-btn';
        toggleBtn.textContent = hasChildren ? '▶' : '';

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.value = node.id;
        checkbox.dataset.name = node.name;
        checkbox.className = 'cat-checkbox';

        const textSpan = document.createElement('span');
        textSpan.textContent = node.name;

        labelDiv.appendChild(toggleBtn);
        labelDiv.appendChild(checkbox);
        labelDiv.appendChild(textSpan);
        li.appendChild(labelDiv);

        if (hasChildren) {
            const childrenContainer = document.createElement('div');
            childrenContainer.className = 'tree-children';
            childrenContainer.style.display = 'none'; // collapsed by default
            renderTree(node.children, childrenContainer);
            li.appendChild(childrenContainer);

            toggleBtn.addEventListener('click', () => {
                const isHidden = childrenContainer.style.display === 'none';
                childrenContainer.style.display = isHidden ? 'block' : 'none';
                toggleBtn.textContent = isHidden ? '▼' : '▶';
            });

            checkbox.addEventListener('change', (e) => {
                const childCheckboxes = childrenContainer.querySelectorAll('input[type="checkbox"]');
                childCheckboxes.forEach(cb => cb.checked = e.target.checked);
                saveConfig();
            });
        } else {
            checkbox.addEventListener('change', saveConfig);
        }

        ul.appendChild(li);
    });

    container.innerHTML = '';
    container.appendChild(ul);
}

function restoreCategorySelection() {
    if (!config.selected_categories) return;
    const checkboxes = document.querySelectorAll('.cat-checkbox');
    checkboxes.forEach(cb => {
        if (config.selected_categories.includes(parseInt(cb.value))) {
            cb.checked = true;
            let parent = cb.closest('.tree-children');
            while (parent) {
                parent.style.display = 'block';
                const toggleBtn = parent.previousElementSibling.querySelector('.toggle-btn');
                if (toggleBtn) toggleBtn.textContent = '▼';
                parent = parent.parentElement.closest('.tree-children');
            }
        }
    });
}

function getSelectedCategories() {
    const checkboxes = document.querySelectorAll('.cat-checkbox:checked');
    const ids = Array.from(checkboxes).map(cb => parseInt(cb.value));
    return ids;
}

async function startParsing() {
    document.getElementById('btn-start').style.display = 'none';
    document.getElementById('btn-stop').style.display = 'block';
    document.getElementById('btn-open-excel').style.display = 'none';
    document.getElementById('log-container').innerHTML = '';

    await saveConfig();
    await pywebview.api.start_parsing(config);
}

async function stopParsing() {
    await pywebview.api.stop_parsing();
}

window.addLog = function (message) {
    const container = document.getElementById('log-container');
    const entry = document.createElement('div');
    entry.className = 'log-entry';

    const time = new Date().toLocaleTimeString();
    entry.textContent = `[${time}] ${message}`;

    container.appendChild(entry);
    container.scrollTop = container.scrollHeight;
}

window.parsingFinished = function (filepath) {
    document.getElementById('btn-start').style.display = 'block';
    document.getElementById('btn-stop').style.display = 'none';
    if (filepath) {
        const btnOpen = document.getElementById('btn-open-excel');
        btnOpen.style.display = 'block';
        btnOpen._filepath = filepath;
    }
}


