!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>فاکتورساز هوشمند ایتا</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        body { font-family: Tahoma, sans-serif; padding: 15px; background: #f9f9f9; }
        textarea { width: 100%; height: 100px; margin-bottom: 10px; border-radius: 8px; border: 1px solid #ccc; padding: 8px; box-sizing: border-box; }
        .section { background: white; padding: 15px; margin-bottom: 15px; border-radius: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        h4 { margin-top: 0; color: #444; }
        button { width: 100%; padding: 12px; background: #007bff; color: white; border: none; border-radius: 5px; font-weight: bold; cursor: pointer; }
        button:hover { background: #0056b3; }
        .output-box { background: #e8f5e9; padding: 15px; border: 2px dashed #4caf50; min-height: 100px; white-space: pre-wrap; margin-top: 10px; border-radius: 8px; }
    </style>
</head>
<body>

    <script>
        // کنترل امنیتی
        let password = prompt("رمز عبور را وارد کنید:");
        if (password !== "1122") { 
            alert("رمز اشتباه است!");
            document.body.innerHTML = "<h2 style='text-align:center; color:red; margin-top:50px;'>دسترسی غیرمجاز.</h2>";
            throw new Error("Invalid Password");
        }
    </script>

    <h2 style="text-align: center;">فاکتورساز هوشمند ایتا</h2>
<div style="display:flex; gap:6px; margin-bottom:15px;">
    <button onclick="showPage('invoicePage')">فاکتور</button>
    <button onclick="showPage('pricesPage')" style="background:#198754;">بانک قیمت‌ها</button>
    <button onclick="showPage('updatePage')" style="background:#e69500;">تغییر قیمت‌ها</button>
<button onclick="showPage('savedInvoicesPage')">
    فاکتورهای ذخیره‌شده
</button>
</div>

<div id="invoicePage">

    <div class="section">
    <h4>۲. سفارش مشتری</h4>

    <div style="display:flex;gap:10px;flex-wrap:wrap;align-items:stretch;">
        <div style="flex:1 1 250px;min-width:0;">
            <label>متن سفارش</label>
            <textarea id="orderInput"
                placeholder="نام محصول و تعداد را وارد کن"
                style="width:100%;box-sizing:border-box;min-height:180px;"></textarea>
                <button type="button" onclick="prepareOrderRows()"
    style="margin:8px 0;background:#198754;">
    آماده‌سازی و اصلاح ردیف‌های سفارش
</button>

<div id="orderRows"></div>
        </div>

       
    </div>

    <button onclick="generateInvoice()">ساخت فاکتور</button>
</div>

    <div class="section">
        <h4>۳. خروجی فاکتور</h4>
        <div class="output-box" id="output">فاکتور آماده نمایش...</div>
        <button onclick="copyInvoice()" style="margin-top:10px; background-color: #28a745;">کپی فاکتور</button>
    </div>
</div>

<div id="pricesPage" style="display:none;">
    <div class="section">
        <h4>بانک دائمی قیمت محصولات</h4>
        <div style="border:1px solid #ddd;border-radius:10px;padding:12px;margin:12px 0;">
    <h4>مدیریت و افزودن محصولات</h4>

    <label>آپلود فایل اکسل:</label>
    <input type="file" id="excelInput"
        accept=".xlsx, .xls, .csv"
        onchange="handleExcel(event)"
        style="margin-bottom:5px;">

    <div id="uploadStatus"
        style="font-size:13px;font-weight:bold;margin-bottom:10px;">
    </div>

    <label>افزودن یا واردکردن گروهی محصولات:</label>
    <textarea id="priceListInput"
        placeholder="مداد مشکی 1000
پاک کن 500"
        style="height:130px;"></textarea>

    <button onclick="addManualPrices()" style="background:#198754;width:100%;">
        افزودن محصولات به بانک قیمت‌ها
    </button>
</div>
        <input
            type="text"
            id="priceSearch"
            placeholder="جستجوی محصول..."
            oninput="renderPriceList()"
            style="width:100%; padding:12px; box-sizing:border-box; margin-bottom:10px;"
        >
                <select id="priceSort" onchange="renderPriceList()"
                style="width:100%;padding:10px;margin-bottom:10px;">
            <option value="excel">ترتیب اکسل</option>
            <option value="alpha">ترتیب الفبایی</option>
        </select>
        <div id="priceCount"></div>
        <button onclick="clearAllProducts()"
        style="background:#dc3545;width:100%;margin:10px 0;">
    پاک‌کردن همهٔ محصولات
</button>
        <button onclick="copyPrices()"
        style="background:#198754;width:100%;margin:10px 0;">
    📋 کپی محصولات و قیمت‌ها
</button>
<button onclick="sharePrices()"
        style="background:#087e9c;width:100%;margin:8px 0;">
    📤 اشتراک‌گذاری محصولات
</button>
        <div id="priceList"></div>
    </div>
</div>

<div id="updatePage" style="display:none;">
    <div class="section">
        <h4>به‌روزرسانی قیمت محصولات</h4>
        <p style="font-size:13px;color:#666;">
            نام محصول و قیمت جدید را در هر خط وارد کن.
        </p>
        <textarea id="updatePriceInput"
            placeholder="کرم نارگیل 450000&#10;کرم خیار 380000"
            style="height:150px;"></textarea>

        <button onclick="updatePrices()" style="background:#e69500;">
            بررسی و اعمال تغییرات
        </button>
    
<button onclick="refreshBasalamProducts()"
    style="background:#198754;margin-top:10px;">
    دریافت و تازه‌سازی محصولات باسلام
</button>

<div id="basalamRefreshReport"
     style="margin:8px 0;padding:8px;">
</div>
    
        <button onclick="previewBasalamPrices()"
    style="background:#087e9c;margin-top:10px;">
    بررسی قیمت‌ها در باسلام
</button>

<div id="basalamPriceReport" class="output-box"></div>

        <div id="updateReport" class="output-box"></div>
    </div>
</div>
<div id="savedInvoicesPage" style="display:none;">
    <div class="section">
        <h4>۱۰ فاکتور آخر</h4>
        <div id="savedInvoiceList"></div>
    </div>
</div>
    <script>
      let masterDatabase = {};
let priceDatabaseReady = false;
let priceSaveQueue = Promise.resolve();

async function loadPriceDatabase() {
    try {
        const response = await fetch("/api/faktorsaz/prices");
        const data = await response.json();

        if (!response.ok || !data.ok) {
            throw new Error("دریافت بانک قیمت از سرور ناموفق بود.");
        }

        const serverPrices = data.prices || {};

        // اگر بانک سرور خالی است، اطلاعات قبلی مرورگر را منتقل کن
        let oldPrices = {};

        try {
            oldPrices = JSON.parse(
                localStorage.getItem("eitaa_price_database") || "{}"
            );
        } catch(e) {
            oldPrices = {};
        }

        if (
            Object.keys(serverPrices).length === 0 &&
            Object.keys(oldPrices).length > 0
        ) {
            const migrationResponse = await fetch(
                "/api/faktorsaz/prices",
                {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json"
                    },
                    body: JSON.stringify(oldPrices)
                }
            );

            const migrationResult = await migrationResponse.json();

            if (!migrationResponse.ok || !migrationResult.ok) {
                throw new Error("انتقال بانک قیمت‌های قبلی ناموفق بود.");
            }

            masterDatabase = oldPrices;

        } else {
            masterDatabase = serverPrices;
        }

        priceDatabaseReady = true;

        localStorage.setItem(
            "eitaa_price_database",
            JSON.stringify(masterDatabase)
        );

        renderPriceList();

        console.log("بانک قیمت‌ها از سرور دریافت شد.");

    } catch(error) {
        console.error(error);

        alert(
            "دریافت بانک قیمت‌ها ناموفق بود:\n" +
            error.message
        );
    }
}


function savePriceDatabase() {

    if (!priceDatabaseReady) {
        alert("هنوز بانک قیمت‌ها از سرور دریافت نشده است.");
        return;
    }

    renderPriceList();

    const snapshot = JSON.stringify(masterDatabase);

    priceSaveQueue = priceSaveQueue.then(async () => {

        const response = await fetch(
            "/api/faktorsaz/prices",
            {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: snapshot
            }
        );

        const result = await response.json();

        if (!response.ok || !result.ok) {
            throw new Error(
                result.error || "ذخیره بانک قیمت‌ها ناموفق بود."
            );
        }

        // نگهداری یک نسخه پشتیبان در مرورگر
        localStorage.setItem(
            "eitaa_price_database",
            snapshot
        );

    }).catch(error => {
        alert("خطا در ذخیره بانک قیمت‌ها:\n" + error.message);
    });

    return priceSaveQueue;
}
function addManualPrices() {
    const input = document.getElementById("priceListInput").value;

    if (!input.trim()) {
        alert("ابتدا نام و قیمت محصولات را وارد کن.");
        return;
    }

    let added = 0;
    let invalid = 0;

    input.split("\n").forEach(line => {
        const parts = line.trim().split(/\s+/);

        if (parts.length < 2) {
            if (line.trim()) invalid++;
            return;
        }

        const price = parseInt(toEnglishDigits(parts.pop()), 10);
        const name = parts.join(" ").trim();

        if (!name || !Number.isFinite(price) || price <= 0) {
            invalid++;
            return;
        }

        masterDatabase[name] = price;
        added++;
    });

    if (added > 0) {
        savePriceDatabase();
    }

    alert(
        added + " محصول به بانک قیمت‌ها اضافه یا به‌روزرسانی شد." +
        (invalid ? "\n" + invalid + " خط نامعتبر بود." : "")
    );
}
function showPage(pageId) {
    ["invoicePage", "pricesPage", "updatePage", "savedInvoicesPage"].forEach(id => {
        document.getElementById(id).style.display =
            id === pageId ? "block" : "none";
    });

    if (pageId === "pricesPage") renderPriceList();
    if (pageId === "savedInvoicesPage") renderSavedInvoices();
}
function renderPriceList() {
    const list = document.getElementById("priceList");
    if (!list) return;

    const query = document.getElementById("priceSearch")
        .value.trim().toLowerCase();

    const sortMode = document.getElementById("priceSort").value;

    let items = Object.entries(masterDatabase)
        .filter(([name]) => name.toLowerCase().includes(query));

    if (sortMode === "alpha") {
        items.sort((a, b) => a[0].localeCompare(b[0], "fa"));
    }

    document.getElementById("priceCount").innerText =
        "تعداد محصولات: " + Object.keys(masterDatabase).length;

    list.innerHTML = "";

    items.forEach(([name, price]) => {
        const row = document.createElement("div");
        row.style.cssText =
            "padding:12px;border-bottom:1px solid #ddd;display:flex;justify-content:space-between;gap:8px;";

        const title = document.createElement("span");
        title.textContent = name;

        const value = document.createElement("strong");
        value.textContent = Number(price).toLocaleString("fa-IR") + " تومان";
const editBtn = document.createElement("button");
editBtn.textContent = "ویرایش";
editBtn.style.cssText =
    "background:#198754;padding:4px 8px;margin-right:4px;font-size:12px;width:auto;min-width:0;";

editBtn.onclick = () => editProduct(name);
        const deleteBtn = document.createElement("button");
deleteBtn.textContent = "حذف";
deleteBtn.style.cssText =
    "background:#dc3545;padding:4px 8px;margin-right:6px;font-size:12px;width:auto;min-width:0;";

deleteBtn.onclick = () => deleteProduct(name);

row.append(title, value, editBtn, deleteBtn);
list.appendChild(row);
    });
}     
function searchInvoiceProducts() {
    const searchInput = document.getElementById("invoiceProductSearch");
    const results = document.getElementById("invoiceProductResults");
    const query = normalizeName(searchInput.value);

    results.innerHTML = "";

    if (!query) return;

    const matches = Object.entries(masterDatabase)
        .filter(([name]) => normalizeName(name).includes(query))
        .slice(0, 30);

    if (matches.length === 0) {
        results.textContent = "محصولی پیدا نشد.";
        return;
    }

    matches.forEach(({ name, price }) => {
        const item = document.createElement("button");
        item.type = "button";
        item.style.cssText =
            "display:block;width:100%;text-align:right;background:white;color:#222;" +
            "padding:10px;border:0;border-bottom:1px solid #eee;border-radius:0;";

        item.textContent =
            name + " — " + Number(price).toLocaleString("fa-IR") + " تومان";

        item.onclick = () => {
            const orderInput = document.getElementById("orderInput");
            const line = name + " 1";

            orderInput.value = orderInput.value.trim()
                ? orderInput.value.trimEnd() + "\n" + line
                : line;

            searchInput.value = "";
            results.innerHTML = "";
            orderInput.focus();
        };

        results.appendChild(item);
    });
}
function deleteProduct(name) {
    if (!confirm(`محصول «${name}» از بانک قیمت‌ها حذف شود؟`)) {
        return;
    }

    delete masterDatabase[name];

    savePriceDatabase();

    alert("محصول حذف شد.");
}   
function editProduct(oldName) {
    const newName = prompt("نام جدید محصول:", oldName);
    if (newName === null) return;

    const cleanName = newName.trim();

    if (!cleanName) {
        alert("نام محصول نمی‌تواند خالی باشد.");
        return;
    }

    const oldPrice = masterDatabase[oldName];

    const newPriceText = prompt(
        "قیمت جدید محصول:",
        oldPrice
    );

    if (newPriceText === null) return;

    const newPrice = parseInt(toEnglishDigits(newPriceText), 10);

    if (!Number.isFinite(newPrice) || newPrice <= 0) {
        alert("قیمت واردشده معتبر نیست.");
        return;
    }

    if (cleanName !== oldName && masterDatabase.hasOwnProperty(cleanName)) {
        alert("محصولی با این نام از قبل وجود دارد.");
        return;
    }

    const updatedDatabase = {};

Object.entries(masterDatabase).forEach(([name, price]) => {
    if (name === oldName) {
        updatedDatabase[cleanName] = newPrice;
    } else {
        updatedDatabase[name] = price;
    }
});

masterDatabase = updatedDatabase;

    savePriceDatabase();

    alert("محصول با موفقیت ویرایش شد.");
}
function clearAllProducts() {
    const count = Object.keys(masterDatabase).length;

    if (count === 0) {
        alert("بانک قیمت‌ها خالی است.");
        return;
    }

    if (!confirm(`همهٔ ${count} محصول از بانک قیمت‌ها پاک شوند؟ این کار قابل برگشت نیست.`)) {
        return;
    }

    masterDatabase = {};
    savePriceDatabase();

    alert("همهٔ محصولات از بانک قیمت‌ها پاک شدند.");
}
async function copyPrices() {
    const sortMode = document.getElementById("priceSort").value;

    let items = Object.entries(masterDatabase);

    if (sortMode === "alpha") {
        items.sort((a, b) => a[0].localeCompare(b[0], "fa"));
    }

    const rows = [
        ["name", "price"],
        ...items.map(([name, price]) => [name, price])
    ];

    const text = rows.map(row => row.join("\t")).join("\n");

    try {
        if (navigator.clipboard && window.isSecureContext) {
            await navigator.clipboard.writeText(text);
        } else {
            throw new Error("Clipboard unavailable");
        }

        alert("اطلاعات آماده انتقال به Excel است.");

    } catch (e) {
        const area = document.createElement("textarea");
        area.value = text;
        area.style.position = "fixed";
        area.style.left = "0";
        area.style.top = "0";
        area.style.width = "100%";
        area.style.height = "200px";
        area.style.zIndex = "99999";

        document.body.appendChild(area);
        area.focus();
        area.select();
        area.setSelectionRange(0, text.length);

        const success = document.execCommand("copy");
        area.remove();

        if (success) {
            alert("محصولات و قیمت‌ها کپی شدند.");
        } else {
            alert("کپی خودکار انجام نشد. باید متن را دستی انتخاب و کپی کنی.");
        }
    }
}
async function sharePrices() {
    const sortMode = document.getElementById("priceSort").value;

    let items = Object.entries(masterDatabase).filter(
    ([name]) => !["name", "price"].includes(name.trim().toLowerCase())
);

    if (sortMode === "alpha") {
        items.sort((a, b) => a[0].localeCompare(b[0], "fa"));
    }

    const text = items
    .map(([name, price]) =>
        `${name}\t${Number(price).toLocaleString("fa-IR")}`
    )
    .join("\n");

    if (!text) {
        alert("محصولی برای اشتراک‌گذاری وجود ندارد.");
        return;
    }

    if (navigator.share) {
        try {
            await navigator.share({
                title: "فهرست محصولات کرشمه",
                text: text
            });
        } catch (e) {
            if (e.name !== "AbortError") {
                alert("اشتراک‌گذاری انجام نشد.");
            }
        }
    } else {
        alert("اشتراک‌گذاری مستقیم در این مرورگر پشتیبانی نمی‌شود.");
    }
}
function normalizeName(name) {
    return name
        .replace(/ي/g, "ی")
        .replace(/ك/g, "ک")
        .replace(/\s+/g, " ")
        .trim()
        .toLowerCase();
}

let pendingPriceUpdates = [];

function normalizeDigits(text) {
    return String(text)
        .replace(/[۰-۹]/g, d => "۰۱۲۳۴۵۶۷۸۹".indexOf(d))
        .replace(/[٠-٩]/g, d => "٠١٢٣٤٥٦٧٨٩".indexOf(d));
}

function productSimilarity(a, b) {
    a = normalizeName(a);
    b = normalizeName(b);

    if (a === b) return 1;

    if (a.includes(b) || b.includes(a)) {
        return Math.min(a.length, b.length) / Math.max(a.length, b.length) + 0.3;
    }

    const wordsA = new Set(a.split(/\s+/));
    const wordsB = new Set(b.split(/\s+/));

    const common = [...wordsA].filter(w => wordsB.has(w)).length;

    return common / Math.max(wordsA.size, wordsB.size);
}

function updatePrices() {
    const input = document.getElementById("updatePriceInput").value;
    const report = document.getElementById("updateReport");

    let updated = [];
    let notFound = [];
    let invalid = [];
    let duplicates = [];
    let unchanged = [];

    pendingPriceUpdates = [];

    const seen = new Set();
    const keys = Object.keys(masterDatabase);

    input.split("\n").forEach(line => {
        line = line.trim();
        if (!line) return;

        let parts = normalizeDigits(line).split(/\s+/);

        if (parts[parts.length - 1] === "تومان") {
            parts.pop();
        }

        const priceText = parts.pop().replace(/[٬,،]/g, "");
        const price = Number(priceText);
        const name = parts.join(" ").trim();

        if (!name || !Number.isSafeInteger(price) || price <= 0) {
            invalid.push(line);
            return;
        }

        const normalized = normalizeName(name);

        if (seen.has(normalized)) {
            duplicates.push(name);
            return;
        }

        seen.add(normalized);

        // ابتدا جستجوی نام دقیق
        const exactKey = keys.find(
            key => normalizeName(key) === normalized
        );

        if (exactKey) {
            const oldPrice = Number(masterDatabase[exactKey]);

            if (oldPrice === price) {
                unchanged.push(exactKey);
                return;
            }

            masterDatabase[exactKey] = price;

            updated.push({
                name: exactKey,
                oldPrice,
                newPrice: price
            });

            return;
        }

        // جستجوی محصولات مشابه
        const candidates = keys
            .map(key => ({
                name: key,
                score: productSimilarity(name, key)
            }))
            .filter(item => item.score >= 0.5)
            .sort((a, b) => b.score - a.score)
            .slice(0, 5);

        if (candidates.length) {
            pendingPriceUpdates.push({
                inputName: name,
                price,
                candidates
            });
        } else {
            notFound.push(name);
        }
    });

    if (updated.length) {
        savePriceDatabase();
    }

    renderUpdateReport({
        updated,
        notFound,
        invalid,
        duplicates,
        unchanged
    });
}


// نمایش گزارش و انتخاب محصول مشابه
function renderUpdateReport(data) {
    const report = document.getElementById("updateReport");
    report.innerHTML = "";

    const title = document.createElement("h4");
    title.textContent = "گزارش به‌روزرسانی قیمت‌ها";
    report.appendChild(title);

    function addLine(text) {
        const p = document.createElement("p");
        p.textContent = text;
        report.appendChild(p);
    }

    addLine("✅ قیمت‌های دقیقاً منطبق: " + data.updated.length);
    addLine("🔍 محصولات نیازمند تأیید: " + pendingPriceUpdates.length);
    addLine("➖ بدون تغییر: " + data.unchanged.length);
    addLine("❌ پیدا نشده: " + data.notFound.length);
    addLine("⚠️ خطوط نامعتبر: " + data.invalid.length);
    addLine("🔁 موارد تکراری: " + data.duplicates.length);

    if (data.updated.length) {
        const heading = document.createElement("h4");
        heading.textContent = "تغییرات انجام‌شده:";
        report.appendChild(heading);

        data.updated.forEach(item => {
            addLine(
                "✅ " + item.name +
                " | قبلی: " + item.oldPrice.toLocaleString("fa-IR") +
                " | جدید: " + item.newPrice.toLocaleString("fa-IR")
            );
        });
    }

    if (pendingPriceUpdates.length) {
        const heading = document.createElement("h4");
        heading.textContent = "انتخاب محصولات مشابه:";
        report.appendChild(heading);

        pendingPriceUpdates.forEach((item, index) => {
            const box = document.createElement("div");
            box.style.cssText =
                "background:#fff;padding:10px;margin:10px 0;border:1px solid #e0a000;border-radius:8px;";

            const name = document.createElement("p");
            name.textContent =
                "نام واردشده: " + item.inputName +
                " | قیمت جدید: " + item.price.toLocaleString("fa-IR");

            const select = document.createElement("select");
            select.dataset.index = index;
            select.style.cssText =
                "width:100%;padding:10px;margin:5px 0;";

            const first = document.createElement("option");
            first.value = "";
            first.textContent = "محصول صحیح را انتخاب کن...";
            select.appendChild(first);

            item.candidates.forEach(candidate => {
                const option = document.createElement("option");
                option.value = candidate.name;
                option.textContent =
                    candidate.name +
                    " (قیمت فعلی: " +
                    Number(masterDatabase[candidate.name]).toLocaleString("fa-IR") +
                    ")";
                select.appendChild(option);
            });

            box.append(name, select);
            report.appendChild(box);
        });

        const applyButton = document.createElement("button");
        applyButton.textContent = "تأیید انتخاب‌ها و اعمال قیمت‌ها";
        applyButton.style.background = "#198754";
        applyButton.onclick = applySimilarPrices;
        report.appendChild(applyButton);
    }

    if (data.notFound.length) {
        addLine("محصولات پیدا نشده: " + data.notFound.join("، "));
    }

    if (data.invalid.length) {
        addLine("خطوط نامعتبر: " + data.invalid.join("، "));
    }

    if (data.duplicates.length) {
        addLine("موارد تکراری: " + data.duplicates.join("، "));
    }
}


// اعمال قیمت فقط برای محصولاتی که تأیید شده‌اند
function applySimilarPrices() {
    const report = document.getElementById("updateReport");
    const selects = report.querySelectorAll("select[data-index]");

    let count = 0;
    let details = [];

    selects.forEach(select => {
        const index = Number(select.dataset.index);
        const item = pendingPriceUpdates[index];
        const selectedName = select.value;

        if (!selectedName || !item) return;

        const oldPrice = Number(masterDatabase[selectedName]);

        if (oldPrice === item.price) {
            details.push("➖ " + selectedName + ": قیمت تغییر نکرد.");
            count++;
            return;
        }

        masterDatabase[selectedName] = item.price;

        details.push(
            "✅ " + selectedName +
            " | قبلی: " + oldPrice.toLocaleString("fa-IR") +
            " | جدید: " + item.price.toLocaleString("fa-IR")
        );

        count++;
    });

    if (count > 0) {
        savePriceDatabase();
    }

    pendingPriceUpdates = [];

    report.innerHTML = "";

    const title = document.createElement("h4");
    title.textContent = "نتیجه تأیید محصولات مشابه";
    report.appendChild(title);

    const result = document.createElement("p");
    result.textContent = "تعداد موارد تأییدشده: " + count;
    report.appendChild(result);

    details.forEach(detail => {
        const p = document.createElement("p");
        p.textContent = detail;
        report.appendChild(p);
    });

    if (count === 0) {
        const p = document.createElement("p");
        p.textContent = "هیچ محصولی برای اعمال انتخاب نشده است.";
        report.appendChild(p);
    }
}

async function refreshBasalamProducts() {
    const report = document.getElementById("basalamRefreshReport");

    report.textContent = "در حال دریافت محصولات از باسلام...";

    try {
        const response = await fetch(
            "http://127.0.0.1:5000/api/basalam/vendor-products"
        );

        const data = await response.json();

        if (!response.ok || !data.ok) {
            throw new Error(data.error || "دریافت محصولات ناموفق بود");
        }

        report.textContent =
            "دریافت موفق بود؛ تعداد محصولات: " +
            Number(data.count).toLocaleString("fa-IR") +
            " محصول. اکنون می‌توانی قیمت‌ها را بررسی کنی.";

    } catch (error) {
        report.textContent = "خطا: " + error.message;
    }
}


async function previewBasalamPrices() {
    const report = document.getElementById("basalamPriceReport");

    report.innerHTML = "در حال بررسی قیمت‌ها...";

    try {
        const products = Object.entries(masterDatabase).map(
            ([name, price]) => ({
                name: name,
                price_toman: Number(price)
            })
        );

        const savedMappings = JSON.parse(
            localStorage.getItem("basalam_product_name_map") || "{}"
        );

        const response = await fetch(
            "http://127.0.0.1:5000/api/basalam/price-update/preview-bank",
            {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    products: products,
                    mappings: savedMappings
                })
            }
        );

        const data = await response.json();

        if (!response.ok || !data.ok) {
            throw new Error(data.error || "خطا در بررسی قیمت‌ها");
        }

        report.innerHTML = "";

        const title = document.createElement("h4");
        title.textContent = "گزارش مقایسه قیمت‌ها با باسلام";
        report.appendChild(title);

        const summary = document.createElement("p");
        summary.textContent =
            "تعداد محصولات بانک: " + products.length +
            " | موارد منطبق: " + (data.matched || 0) +
            " | پیدا نشده: " + (data.not_found || 0);

        report.appendChild(summary);

        let number = 0;

        // نمایش محصولات منطبق و موارد تکراری
        (data.items || []).forEach(item => {
            number++;

            const p = document.createElement("p");

            if (item.status === "duplicate") {
                p.textContent =
                    number + " ـ " + item.title +
                    " | " + item.message;
            } else {
                p.textContent =
                    number + " ـ " +
                    item.title +
                    " | قیمت باسلام: " +
                    Number(item.old_price).toLocaleString("fa-IR") +
                    " ریال | قیمت جدید: " +
                    Number(item.new_price).toLocaleString("fa-IR") +
                    " ریال";
            }

            report.appendChild(p);
        });

        // محصولات نیازمند تطبیق دستی
        const unmatched = data.not_found_names || [];
        const selectedMappings = {};

        if (unmatched.length) {
            const heading = document.createElement("h4");
            heading.textContent = "محصولات نیازمند تطبیق دستی";
            report.appendChild(heading);

            unmatched.forEach(item => {
                number++;

                const box = document.createElement("div");

                box.style.cssText =
                    "border:1px solid #e0a000;padding:10px;margin:10px 0;border-radius:8px;background:#fff;";

                const name = document.createElement("p");
                name.textContent = number + " ـ " + item.name;

                const select = document.createElement("select");
                select.style.cssText =
                    "width:100%;padding:10px;margin:8px 0;";

                const first = document.createElement("option");
                first.value = "";
                first.textContent =
                    "محصول صحیح در باسلام را انتخاب کن";

                select.appendChild(first);

                (item.suggestions || []).forEach(candidate => {
                    const option = document.createElement("option");

                    option.value = candidate.id;
                    option.textContent =
                        candidate.title +
                        " | شباهت: " + candidate.score + "%";

                    select.appendChild(option);
                });

                // نمایش انتخاب قبلی، در صورت وجود
                const previous = savedMappings[item.name];

                if (previous) {
                    const exists = Array.from(select.options).some(
                        option => String(option.value) === String(previous.id)
                    );

                    if (exists) {
                        select.value = previous.id;
                    }
                }

                select.onchange = () => {
                    if (select.value) {
                        selectedMappings[item.name] = {
                            id: select.value,
                            title: select.options[select.selectedIndex].text
                        };
                    } else {
                        delete selectedMappings[item.name];
                    }
                };

                box.append(name, select);
                report.appendChild(box);
            });

            const saveButton = document.createElement("button");
            saveButton.textContent = "ذخیره انتخاب‌های تطبیق";
            saveButton.style.background = "#198754";

            saveButton.onclick = () => {
                const oldMappings = JSON.parse(
                    localStorage.getItem("basalam_product_name_map") || "{}"
                );

                Object.assign(oldMappings, selectedMappings);

                localStorage.setItem(
                    "basalam_product_name_map",
                    JSON.stringify(oldMappings)
                );

                alert(
                    "تعداد انتخاب‌های جدید ذخیره‌شده: " +
                    Object.keys(selectedMappings).length
                );
            };

            report.appendChild(saveButton);

        } else {
            const message = document.createElement("p");
            message.textContent =
                "همه محصولات پیدا شده‌اند یا تطبیق داده شده‌اند.";

            report.appendChild(message);
        }
                // دکمه اعمال قیمت‌های جدید در باسلام
        const updateItems = (data.items || []).filter(item =>
            item.status === "matched" &&
            item.changed === true
        );

        if (updateItems.length > 0) {
            const updateButton = document.createElement("button");

            updateButton.textContent =
                "اعمال قیمت‌های جدید در باسلام (" +
                updateItems.length + " محصول)";

            updateButton.style.cssText =
                "display:block;width:100%;padding:13px;margin:15px 0;" +
                "background:#d9534f;color:white;font-weight:bold;";

            updateButton.onclick = async () => {
                const confirmed = confirm(
                    "قیمت " + updateItems.length +
                    " محصول در باسلام تغییر کند؟\n" +
                    "این عملیات قیمت محصولات را واقعاً تغییر می‌دهد."
                );

                if (!confirmed) return;

                updateButton.disabled = true;
                updateButton.textContent = "در حال ارسال قیمت‌ها...";

                try {
                    const updateResponse = await fetch(
                        "http://127.0.0.1:5000/api/basalam/price-update/apply-bank",
                        {
                            method: "POST",
                            headers: {
                                "Content-Type": "application/json"
                            },
                            body: JSON.stringify({
                                confirm_update:
                                    "اعمال قیمت‌های انتخاب‌شده در باسلام",
                                items: updateItems
                            })
                        }
                    );

                    const result = await updateResponse.json();

                    if (!updateResponse.ok || !result.ok) {
                        throw new Error(
                            result.error || "اعمال قیمت‌ها ناموفق بود."
                        );
                    }

                    alert(
                        result.message ||
                        "درخواست تغییر قیمت ارسال شد.\n" +
                        "تعداد محصولات: " + result.submitted
                    );

                    updateButton.textContent = "درخواست ارسال شد";

                } catch (error) {
                    alert("خطا: " + error.message);
                    updateButton.disabled = false;
                    updateButton.textContent =
                        "اعمال قیمت‌های جدید در باسلام (" +
                        updateItems.length + " محصول)";
                }
            };

            report.appendChild(updateButton);
        }

    } catch (error) {
        report.textContent = "خطا: " + error.message;
    }
}


        function handleExcel(e) {
            const file = e.target.files[0];
            const statusEl = document.getElementById('uploadStatus');
            if (!file) return;
            
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, {type: 'array'});
                    const sheet = workbook.Sheets[workbook.SheetNames[0]];
                    const json = XLSX.utils.sheet_to_json(sheet, {header: 1});
                    json.forEach(row => {
    if (row[0] && row[1]) {
        masterDatabase[row[0].toString().trim()] =
            parseInt(toEnglishDigits(row[1]), 10);
    }
});
                    savePriceDatabase();
                    statusEl.innerText = "فایل اکسل با موفقیت بارگذاری شد!";
                    statusEl.style.color = "green";
                } catch (err) { 
                    statusEl.innerText = "خطا در خواندن فایل اکسل!";
                    statusEl.style.color = "red";
                }
            };
            reader.readAsArrayBuffer(file);
        }
let editingInvoicePrices = null;
let editingInvoiceId = null;
function toEnglishDigits(value) {
    return String(value)
        .replace(/[۰-۹]/g, d => "۰۱۲۳۴۵۶۷۸۹".indexOf(d))
        .replace(/[٠-٩]/g, d => "٠١٢٣٤٥٦٧٨٩".indexOf(d))
        .replace(/[٬،,]/g, "")
        .trim();
}
function prepareOrderRows() {
    const text = document.getElementById("orderInput").value;
    const container = document.getElementById("orderRows");

    const lines = text.split(/\r?\n/).filter(line => line.trim());

    container.innerHTML = "";

    lines.forEach(line => {
        let clean = normalizeDigits(line).trim();
        clean = clean.replace(/\s*عدد\s*$/u, "").trim();

        const match = clean.match(/[:：]?\s*(\d+)\s*$/);
        const qty = match ? parseInt(match[1], 10) : 1;

        const name = match
            ? clean.slice(0, match.index).replace(/[:：]\s*$/, "").trim()
            : clean;

        if (name) {
            addOrderRow(container, name, qty);
        }
    });

    if (lines.length === 0) {
        alert("متن سفارش خالی است؛ اول سفارش را در کادر وارد کن.");
    } else {
        alert("تعداد ردیف‌های آماده‌شده: " + container.querySelectorAll("div").length);
    }
}

function addOrderRow(container, initialName, initialQty) {
    const row = document.createElement("div");

    row.style.cssText =
        "border:1px solid #ddd;border-radius:8px;padding:10px;margin:8px 0;background:#fff;";

    const nameLabel = document.createElement("div");
    nameLabel.textContent = initialName;
    const foundInitially = Object.keys(masterDatabase).some(
    name => normalizeName(name) === normalizeName(initialName)
);

row.style.background = foundInitially ? "#e8f5e9" : "#ffebee";
row.style.border = foundInitially
    ? "1px solid #4caf50"
    : "1px solid #f44336";

nameLabel.style.color = foundInitially ? "#198754" : "#dc3545";

if (!foundInitially) {
    nameLabel.textContent = "⚠️ پیدا نشد: " + initialName;
}
    nameLabel.style.cssText = "font-weight:bold;margin-bottom:8px;";

    const qtyInput = document.createElement("input");
    qtyInput.type = "number";
    qtyInput.min = "1";
    qtyInput.value = initialQty;
    qtyInput.style.cssText = "width:80px;padding:8px;box-sizing:border-box;";

    const search = document.createElement("input");
    search.type = "text";
    search.placeholder = "برای اصلاح، نام محصول را جست‌وجو کن";
    search.style.cssText =
        "width:100%;padding:10px;box-sizing:border-box;margin-top:8px;";

    const results = document.createElement("div");
    results.style.cssText = "max-height:150px;overflow-y:auto;";

    const rowData = {
        name: initialName,
        qty: initialQty
    };

    function syncOrderText() {
        const rows = [...container.children]
    .filter(el => el._orderData)
    .map(el => {
        const data = el._orderData;
        return `${data.name}: ${data.qty} عدد`;
    });

        document.getElementById("orderInput").value = rows.join("\n");
    }

    qtyInput.onchange = () => {
        rowData.qty = Math.max(
            1,
            parseInt(qtyInput.value, 10) || 1
        );

        qtyInput.value = rowData.qty;
        syncOrderText();
    };

    search.oninput = () => {
        results.innerHTML = "";

        const query = normalizeName(search.value);

        if (!query) return;

        const words = query.split(/\s+/).filter(Boolean);

const matches = Object.entries(masterDatabase)
    .map(([name, price]) => {
        const normalizedName = normalizeName(name);
        const score = words.filter(word =>
            normalizedName.includes(word)
        ).length;

        return { name, price, score };
    })
    .filter(item => item.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, 20);

        matches.forEach(({ name, price }) => {
            const option = document.createElement("button");

            option.type = "button";

            option.style.cssText =
                "display:block;width:100%;text-align:right;background:#fff;color:#222;" +
                "padding:9px;border:0;border-bottom:1px solid #eee;border-radius:0;";

            option.textContent =
                `${name} — ${Number(price).toLocaleString("fa-IR")} تومان`;

            option.onclick = () => {
                rowData.name = name;
nameLabel.textContent = "✓ " + name;

row.style.background = "#e8f5e9";
row.style.border = "1px solid #4caf50";
nameLabel.style.color = "#198754";

                search.value = "";
                results.innerHTML = "";

                syncOrderText();
            };

            results.appendChild(option);
        });

        if (!results.children.length) {
            results.textContent = "محصولی پیدا نشد.";
        }
    };

    row._orderData = rowData;

    row.append(nameLabel, qtyInput, search, results);

container.appendChild(row);

if (!foundInitially) {
    search.value = initialName;
    search.oninput();
}
}
        function generateInvoice() {
            // ترکیب دیتابیس فعلی با متن دستی (بدون حذف اطلاعات قبلی)
            
            const orderText = document.getElementById('orderInput').value;
const orders = orderText.split('\n');

let invoice = "";
let total = 0;
let invoiceItems = [];

            orders.forEach(line => {
                if (!line.trim()) return;
                
                let parts = line.trim().split(/\s+/);
                let qty = 1; // مقدار پیش‌فرض
                let isExplicit = false;

                // ۱. حذف کلمه "عدد" اگر انتهای خط باشد
                if (parts[parts.length - 1] === "عدد") {
                    parts.pop(); 
                }

                // ۲. بررسی وجود عدد در انتها (تعداد)
                let lastWord = parts[parts.length - 1];
                if (!isNaN(parseInt(lastWord))) {
                    qty = parseInt(parts.pop());
                    isExplicit = true;
                }
                let name = parts.join(' ');
                
                // ۳. جستجوی هوشمند در دیتابیس
                let foundKey = Object.keys(masterDatabase).find(k => k.includes(name) || name.includes(k));
                let price = foundKey
    ? (
        editingInvoicePrices &&
        Object.prototype.hasOwnProperty.call(editingInvoicePrices, foundKey)
            ? editingInvoicePrices[foundKey]
            : masterDatabase[foundKey]
      )
    : 0;
                let foundName = foundKey || name;
                
                // ۴. محاسبه قیمت (مدیریت NaN با ضرب ساده)
                let lineTotal = price * qty;
invoiceItems.push({
    name: foundName,
    qty: qty,
    explicit: isExplicit,
    price: Number(price),
    lineTotal: lineTotal
});
                // ۵. فرمت خروجی (فقط در صورتی که تعداد دستی وارد شده باشد "عدد" را بنویس)
                if (qty > 1) {
    invoice += `${foundName} ${qty} عدد: ${lineTotal}\n`;
} else {
    invoice += `${foundName}: ${lineTotal}\n`;
}
                
                total += lineTotal;
            });

            invoice += `\nجمع: ${total}`;
document.getElementById('output').innerText = invoice;

let savedInvoices = [];

try {
    savedInvoices = JSON.parse(
        localStorage.getItem("eitaa_saved_invoices") || "[]"
    );
} catch(e) {
    savedInvoices = [];
}

const newInvoice = {
    id: editingInvoiceId || Date.now(),
    date: new Date().toLocaleString("fa-IR"),
    orderText: orderText,
    items: invoiceItems,
    total: total,
    invoiceText: invoice
};

const existingIndex = savedInvoices.findIndex(
    inv => inv.id === editingInvoiceId
);

if (editingInvoiceId && existingIndex !== -1) {
    savedInvoices[existingIndex] = newInvoice;
} else {
    savedInvoices.unshift(newInvoice);
}

savedInvoices = savedInvoices.slice(0, 10);

localStorage.setItem(
    "eitaa_saved_invoices",
    JSON.stringify(savedInvoices)
);

editingInvoiceId = null;
editingInvoicePrices = null;
        }
function editSavedInvoice(id) {
    const savedInvoices = JSON.parse(
        localStorage.getItem("eitaa_saved_invoices") || "[]"
    );

    const saved = savedInvoices.find(inv => inv.id === id);

    if (!saved) {
        alert("فاکتور پیدا نشد.");
        return;
    }

    editingInvoiceId = saved.id;
    editingInvoicePrices = {};

    saved.items.forEach(item => {
        editingInvoicePrices[item.name] = Number(item.price);
    });

    document.getElementById("orderInput").value = saved.orderText;
    document.getElementById("output").innerText = saved.invoiceText;

    showPage("invoicePage");
}
function renderSavedInvoices() {
    const list = document.getElementById("savedInvoiceList");
    if (!list) return;

    let savedInvoices = [];

    try {
        savedInvoices = JSON.parse(
            localStorage.getItem("eitaa_saved_invoices") || "[]"
        );
    } catch(e) {
        savedInvoices = [];
    }

    list.innerHTML = "";

    if (savedInvoices.length === 0) {
        list.innerHTML = "<p>هنوز فاکتوری ذخیره نشده است.</p>";
        return;
    }

    savedInvoices.forEach(inv => {
        const box = document.createElement("div");
        box.style.cssText =
            "padding:12px;margin:10px 0;border:1px solid #ddd;border-radius:10px;";

        const info = document.createElement("div");
        info.innerText =
            "تاریخ: " + inv.date +
            "\nجمع: " + Number(inv.total).toLocaleString("fa-IR");

        const btn = document.createElement("button");
        btn.innerText = "بازکردن و ویرایش";
        btn.style.background = "#198754";
        btn.onclick = () => editSavedInvoice(inv.id);

        box.append(info, btn);
        list.appendChild(box);
    });
}
        async function copyInvoice() {
    const textToCopy = document.getElementById("output").innerText;

    if (!textToCopy.trim()) {
        alert("هنوز فاکتوری برای کپی وجود ندارد.");
        return;
    }

    try {
        if (navigator.clipboard && window.isSecureContext) {
            await navigator.clipboard.writeText(textToCopy);
        } else {
            throw new Error("Clipboard unavailable");
        }

        alert("فاکتور با موفقیت کپی شد! 🌷");

    } catch (e) {
        const area = document.createElement("textarea");
        area.value = textToCopy;
        area.style.position = "fixed";
        area.style.top = "0";
        area.style.left = "0";
        area.style.width = "100%";
        area.style.height = "200px";
        area.style.zIndex = "999999";

        document.body.appendChild(area);
        area.focus();
        area.select();
        area.setSelectionRange(0, textToCopy.length);

        const success = document.execCommand("copy");
        area.remove();

        if (success) {
            alert("فاکتور با موفقیت کپی شد! 🌷");
        } else {
            alert("کپی خودکار انجام نشد. متن فاکتور را انتخاب و دستی کپی کن.");
        }
    }
}
loadPriceDatabase();
    </script>
</body>
</html>
