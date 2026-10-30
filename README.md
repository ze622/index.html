!DOOCTYPE html>
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
</div>

<div id="invoicePage">
    <div class="section">
        <h4>۱. لیست قیمت‌ها (اکسل یا متن)</h4>
        <input type="file" id="excelInput" accept=".xlsx, .xls, .csv" onchange="handleExcel(event)" style="margin-bottom: 5px;">
        <div id="uploadStatus" style="font-size: 13px; font-weight: bold; margin-bottom: 10px;"></div>
        <textarea id="priceListInput" placeholder="مداد مشکی 1000
پاک کن 500"></textarea>
    </div>

    <div class="section">
        <h4>۲. سفارش مشتری</h4>
        <textarea id="orderInput" placeholder="مداد 2
پاک کن"></textarea>
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
        <input
            type="text"
            id="priceSearch"
            placeholder="جستجوی محصول..."
            oninput="renderPriceList()"
            style="width:100%; padding:12px; box-sizing:border-box; margin-bottom:10px;"
        >
        <div id="priceCount"></div>
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

        <div id="updateReport" class="output-box"></div>
    </div>
</div>
    <script>
        let masterDatabase = {};

try {
    masterDatabase = JSON.parse(
        localStorage.getItem("eitaa_price_database") || "{}"
    );
} catch(e) {
    masterDatabase = {};
}

function savePriceDatabase() {
    localStorage.setItem(
        "eitaa_price_database",
        JSON.stringify(masterDatabase)
    );
    renderPriceList();
}

function showPage(pageId) {
    ["invoicePage", "pricesPage", "updatePage"].forEach(id => {
        document.getElementById(id).style.display =
            id === pageId ? "block" : "none";
    });

    if (pageId === "pricesPage") renderPriceList();
}
function renderPriceList() {
    const list = document.getElementById("priceList");
    if (!list) return;

    const query = document.getElementById("priceSearch")
        .value.trim().toLowerCase();

    const items = Object.entries(masterDatabase)
        .filter(([name]) => name.toLowerCase().includes(query))
        .sort((a, b) => a[0].localeCompare(b[0], "fa"));

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

        row.append(title, value);
        list.appendChild(row);
    });
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

        const normalized = normalizeName(name);

        if (seen.has(normalized)) {
            duplicates.push(name);
            return;
        }

        seen.add(normalized);

        const foundKey = Object.keys(masterDatabase).find(
            key => normalizeName(key) === normalized
        );

        if (!foundKey) {
            const candidates = Object.keys(masterDatabase).filter(
                key =>
                    normalizeName(key).includes(normalized) ||
                    normalized.includes(normalizeName(key))
            );

            notFound.push({
                name,
                candidates
            });
            return;
        }

        const oldPrice = Number(masterDatabase[foundKey]);

        if (oldPrice === price) {
            return;
        }

        masterDatabase[foundKey] = price;

        updated.push({
            name: foundKey,
            oldPrice,
            newPrice: price
        });
    });

    savePriceDatabase();

    let html = `<b>گزارش به‌روزرسانی قیمت‌ها</b><hr>`;

    html += `✅ قیمت‌های تغییرکرده: ${updated.length}<br>`;
    html += `🔍 محصولات نیازمند بررسی: ${notFound.length}<br>`;
    html += `⚠️ خطوط نامعتبر: ${invalid.length}<br>`;
    html += `🔁 موارد تکراری: ${duplicates.length}<br><hr>`;

    updated.forEach(item => {
        html += `✅ ${item.name}<br>`;
        html += `قیمت قبلی: ${item.oldPrice.toLocaleString("fa-IR")}<br>`;
        html += `قیمت جدید: ${item.newPrice.toLocaleString("fa-IR")}<br><br>`;
    });

    notFound.forEach(item => {
        html += `🔍 ${item.name}<br>`;

        if (item.candidates.length) {
            html += `موارد مشابه: ${item.candidates.join("، ")}<br>`;
        } else {
            html += `در بانک قیمت‌ها پیدا نشد.<br>`;
        }

        html += `<br>`;
    });

    invalid.forEach(item => {
        html += `⚠️ ${item}<br>`;
    });

    duplicates.forEach(item => {
        html += `🔁 ${item}<br>`;
    });

    report.innerHTML = html;
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
                    json.forEach(row => { if (row[0] && row[1]) masterDatabase[row[0].toString().trim()] = parseInt(row[1]); });
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

        function generateInvoice() {
            // ترکیب دیتابیس فعلی با متن دستی (بدون حذف اطلاعات قبلی)
            const manualText = document.getElementById('priceListInput').value;
            manualText.split('\n').forEach(line => {
                const parts = line.trim().split(/\s+/);
                if (parts.length >= 2) {
                    const price = parseInt(parts.pop());
                    const name = parts.join(' ');
                    masterDatabase[name] = price;
                }
            });
savePriceDatabase();
            const orders = document.getElementById('orderInput').value.split('\n');
            let invoice = "";
            let total = 0;

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
                let price = foundKey ? masterDatabase[foundKey] : 0;
                let foundName = foundKey || name;
                
                // ۴. محاسبه قیمت (مدیریت NaN با ضرب ساده)
                let lineTotal = price * qty;

                // ۵. فرمت خروجی (فقط در صورتی که تعداد دستی وارد شده باشد "عدد" را بنویس)
                if (isExplicit) {
                    invoice += `${foundName} ${qty} عدد: ${lineTotal}\n`;
                } else {
                    invoice += `${foundName}: ${lineTotal}\n`;
                }
                
                total += lineTotal;
            });

            invoice += `\nجمع: ${total}`;
            document.getElementById('output').innerText = invoice;
        }

        function copyInvoice() {
            const textToCopy = document.getElementById('output').innerText;
            navigator.clipboard.writeText(textToCopy).then(() => alert("فاکتور کپی شد!"));
        }
    </script>
</body>
</html>
