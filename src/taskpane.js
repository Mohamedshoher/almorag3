/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// Store mistakes globally
let globalMistakes = {
    spaces: [],
    spelling: [],
    grammar: [],
    punctuation: [],
    style: []
};
let operationHistory = [];

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // Event Listeners
        document.getElementById("run-check").onclick = runFullCheck;
        document.getElementById("fetch-models").onclick = fetchAvailableModels;
        document.getElementById("btn-remove-tashkeel").onclick = handleRemoveTashkeel;
        document.getElementById("btn-tashkeel").onclick = handleTashkeel;
        document.getElementById("btn-fast-fix").onclick = handleFastAutoFix;
        document.getElementById("btn-reverse-brackets").onclick = handleReverseBrackets;
        document.getElementById("btn-remove-empty-lines").onclick = handleRemoveEmptyLines;
        document.getElementById("btn-quran-search").onclick = openQuranSearch;



        // Remove Brackets Button
        document.getElementById("btn-remove-brackets").onclick = handleRemoveAllBrackets;


        // Settings Modal Logic
        const modal = document.getElementById("settings-modal");
        const settingsBtn = document.getElementById("settings-btn");
        const closeBtn = document.getElementById("close-settings");

        if (settingsBtn) {
            settingsBtn.onclick = () => modal.classList.remove("hidden");
        }
        if (closeBtn) {
            closeBtn.onclick = () => modal.classList.add("hidden");
        }

        // Quran Modal Logic
        const quranModal = document.getElementById("quran-modal");
        const closeQuran = document.getElementById("close-quran");
        if (closeQuran) {
            closeQuran.onclick = () => quranModal.classList.add("hidden");
        }

        // Alminasa Modal Logic
        const alminasaModal = document.getElementById("alminasa-modal");
        const closeAlminasa = document.getElementById("close-alminasa");
        if (document.getElementById("btn-alminasa")) {
            document.getElementById("btn-alminasa").onclick = () => alminasaModal.classList.remove("hidden");
        }
        if (closeAlminasa) {
            closeAlminasa.onclick = () => alminasaModal.classList.add("hidden");
        }

        // Turath Modal Logic
        const turathModal = document.getElementById("turath-modal");
        const closeTurath = document.getElementById("close-turath");
        if (document.getElementById("btn-turath")) {
            document.getElementById("btn-turath").onclick = () => turathModal.classList.remove("hidden");
        }
        if (closeTurath) {
            closeTurath.onclick = () => turathModal.classList.add("hidden");
        }

        // Google Search Logic
        const googleModal = document.getElementById("google-modal");
        const closeGoogle = document.getElementById("close-google");
        if (document.getElementById("btn-google-search")) {
            document.getElementById("btn-google-search").onclick = handleGoogleSearch;
        }
        if (closeGoogle) {
            closeGoogle.onclick = () => googleModal.classList.add("hidden");
        }

        // Browser Controls Logic
        const browserBack = document.getElementById("browser-back");
        const browserForward = document.getElementById("browser-forward");
        const browserRefresh = document.getElementById("browser-refresh");
        const browserGo = document.getElementById("browser-go");
        const browserUrlInput = document.getElementById("browser-url");
        const googleIframe = document.getElementById("google-iframe");

        if (browserBack) browserBack.onclick = () => { try { googleIframe.contentWindow.history.back(); } catch (e) { console.log(e); } };
        if (browserForward) browserForward.onclick = () => { try { googleIframe.contentWindow.history.forward(); } catch (e) { console.log(e); } };
        if (browserRefresh) browserRefresh.onclick = () => { googleIframe.src = googleIframe.src; };
        if (browserGo) browserGo.onclick = () => navigateToUrl();
        if (browserUrlInput) {
            browserUrlInput.onkeydown = (e) => {
                if (e.key === "Enter") navigateToUrl();
            };
        }

        function navigateToUrl() {
            let value = browserUrlInput.value.trim();
            if (!value) return;

            // Check if it's a URL or search query
            if (value.startsWith('http://') || value.startsWith('https://')) {
                googleIframe.src = value;
            } else if (value.includes('.') && !value.includes(' ')) {
                googleIframe.src = 'https://' + value;
            } else {
                // Perform Google search
                googleIframe.src = `https://www.google.com/search?q=${encodeURIComponent(value)}&igu=1`;
            }
        }

        // History Modal Logic
        const historyBtn = document.getElementById("history-btn");
        const historyModal = document.getElementById("history-modal");
        const closeHistory = document.getElementById("close-history");

        if (historyBtn) {
            historyBtn.onclick = () => {
                renderHistory();
                historyModal.classList.remove("hidden");
            };
        }
        if (closeHistory) {
            closeHistory.onclick = () => historyModal.classList.add("hidden");
        }

        // Results Modal Logic
        const resultsModal = document.getElementById("results-modal");
        const closeResults = document.getElementById("close-results");
        const exportPdfModal = document.getElementById("export-pdf-modal");

        if (closeResults) {
            closeResults.onclick = () => resultsModal.classList.add("hidden");
        }
        if (exportPdfModal) {
            exportPdfModal.onclick = exportToPDF;
        }

        // Export PDF Logic (Legacy Support if needed)
        const exportBtn = document.getElementById("export-pdf");
        if (exportBtn) {
            exportBtn.onclick = exportToPDF;
        }

    }
});

// Accordion Logic
window.toggleSection = (category) => {
    const content = document.getElementById(`content-${category}`);
    const section = document.getElementById(`section-${category}`);

    if (content.classList.contains('hidden')) {
        content.classList.remove('hidden');
        section.classList.add('open');
    } else {
        content.classList.add('hidden');
        section.classList.remove('open');
    }
};

// --- Quran Search Functions ---

function openQuranSearch() {
    const quranModal = document.getElementById("quran-modal");
    quranModal.classList.remove("hidden");
}

// --- Google Search Function ---
async function handleGoogleSearch() {
    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        let query = selection.text ? selection.text.trim() : "";
        if (!query) {
            showFeedback("⚠️ يرجى تحديد نص للبحث عنه في جوجل.", "warning");
            return;
        }

        const googleModal = document.getElementById("google-modal");
        const googleIframe = document.getElementById("google-iframe");
        const browserUrlInput = document.getElementById("browser-url");

        const searchUrl = `https://www.google.com/search?q=${encodeURIComponent(query)}&igu=1`;

        if (browserUrlInput) browserUrlInput.value = query;
        googleModal.classList.remove("hidden");
        googleIframe.src = searchUrl;
    });
}

window.handleGoogleSearch = handleGoogleSearch;

// --- Spacing Review Function (Based on Microsoft Word Arabic Standards) ---

async function handleSpacingReview() {
    const messageArea = document.getElementById("message-area");
    const resultsArea = document.getElementById("results-area");
    const progressArea = document.getElementById("progress-area");
    const progressFill = document.getElementById("progress-fill");
    const progressText = document.getElementById("progress-text");

    // Reset UI
    document.querySelectorAll('.category-section').forEach(el => el.classList.add('hidden'));
    document.querySelectorAll('.section-content').forEach(el => el.classList.add('hidden'));
    document.querySelectorAll('ul[id^="list-"]').forEach(el => el.innerHTML = '');
    document.querySelectorAll('.count-badge').forEach(el => el.innerText = '0');
    messageArea.innerText = "";
    globalMistakes = { spaces: [], spelling: [], grammar: [], punctuation: [], style: [] };

    progressArea.classList.remove("hidden");
    progressText.innerText = "جاري البدء في الفحص الدقيق...";
    progressFill.style.width = "10%";
    resultsArea.classList.remove("hidden");

    await Word.run(async (context) => {
        let range = context.document.getSelection();
        range.load("text");
        await context.sync();

        if (!range.text || range.text.trim().length === 0) {
            range = context.document.body;
            range.load("text");
            await context.sync();
        }

        const fullText = range.text;
        if (!fullText || fullText.trim().length === 0) {
            messageArea.innerText = "المستند فارغ.";
            progressArea.classList.add("hidden");
            return;
        }

        let spacingIssues = [];
        const lines = fullText.split(/[\r\n]+/);

        // 1. القواعد العامة لعلامات الترقيم (، ؛ . : ! ؟)
        const marks = ['،', '؛', '.', ':', '!', '؟'];
        const openingBrackets = ['(', '[', '{', '﴿', '«', '‹', '<'];
        const closingBrackets = [')', ']', '}', '﴾', '»', '›', '>'];

        progressText.innerText = "فحص علامات الترقيم والمسافات...";
        progressFill.style.width = "40%";

        for (let line of lines) {
            // أ. مسافة قبل علامة الترقيم (خطأ)
            for (let mark of marks) {
                const regexBefore = new RegExp(`\\s+${mark}`, 'g');
                let match;
                while ((match = regexBefore.exec(line)) !== null) {
                    spacingIssues.push({
                        error: match[0],
                        correction: mark,
                        reason: `علامة "${mark}" يجب أن تلتصق بما قبلها مباشرة.`,
                        category: "beforePunctuation"
                    });
                }
            }

            // ب. عدم وجود مسافة بعد علامة الترقيم (خطأ) إلا إذا كان وراءها قوس إغلاق
            for (let mark of marks) {
                // نستخدم regex يبحث عن العلامة متبوعة بحرف ليس مسافة وليس قوس إغلاق وليس علامة ترقيم أخرى
                const regexAfter = new RegExp(`${mark}([^\\s\\)\\]\\}﴾»›>،؛\\.:!؟\\d])`, 'g');
                let match;
                while ((match = regexAfter.exec(line)) !== null) {
                    spacingIssues.push({
                        error: match[0],
                        correction: mark + ' ' + match[1],
                        reason: `يجب ترك مسافة واحدة بعد علامة "${mark}".`,
                        category: "afterPunctuation"
                    });
                }
            }

            // ج. مسافات زائدة (أكثر من مسافة)
            const multiSpaceRegex = / {2,}/g;
            let spaceMatch;
            while ((spaceMatch = multiSpaceRegex.exec(line)) !== null) {
                spacingIssues.push({
                    error: spaceMatch[0],
                    correction: " ",
                    reason: "مسافات زائدة، المتفق عليه مسافة واحدة فقط بين الكلمات.",
                    category: "multipleSpaces"
                });
            }

            // د. واو العطف (يجب أن تلتصق بما بعدها)
            const wawRegex = /و\s+([\u0600-\u06FF])/g;
            let wawMatch;
            while ((wawMatch = wawRegex.exec(line)) !== null) {
                spacingIssues.push({
                    error: wawMatch[0],
                    correction: "و" + wawMatch[1],
                    reason: "واو العطف يجب أن تلتصق بالكلمة التي بعدها مباشرة.",
                    category: "multipleSpaces"
                });
            }
        }

        // 2. فحص الأقواس
        progressText.innerText = "فحص دقة المسافات حول الأقواس...";
        progressFill.style.width = "70%";

        // أ. مسافة بعد قوس الافتتاح (خطأ)
        for (let b of openingBrackets) {
            const regex = new RegExp(`\\${b}\\s+`, 'g');
            let match;
            while ((match = regex.exec(fullText)) !== null) {
                spacingIssues.push({
                    error: match[0],
                    correction: b,
                    reason: "لا يجوز ترك مسافة بعد قوس الافتتاح.",
                    category: "afterOpenBracket"
                });
            }
        }

        // ب. مسافة قبل قوس الإغلاق (خطأ)
        for (let b of closingBrackets) {
            const regex = new RegExp(`\\s+\\${b}`, 'g');
            let match;
            while ((match = regex.exec(fullText)) !== null) {
                spacingIssues.push({
                    error: match[0],
                    correction: b,
                    reason: "لا يجوز ترك مسافة قبل قوس الإغلاق.",
                    category: "beforeCloseBracket"
                });
            }
        }

        // ج. فقدان المسافة قبل قوس الافتتاح
        for (let b of openingBrackets) {
            const regex = new RegExp(`([^\\s\\(\\[\\{﴿«\\/])${b}`, 'g');
            let match;
            while ((match = regex.exec(fullText)) !== null) {
                spacingIssues.push({
                    error: match[0],
                    correction: match[1] + ' ' + b,
                    reason: "يجب ترك مسافة قبل قوس الافتتاح.",
                    category: "beforeOpenBracket"
                });
            }
        }

        progressFill.style.width = "100%";
        progressArea.classList.add("hidden");

        // Categorize and display results
        const categories = {
            multipleSpaces: { name: 'مسافات زائدة وتنسيق كلمات', icon: '⚠️', issues: [] },
            beforePunctuation: { name: 'التصاق بما قبلها (ترقيم)', icon: '❌', issues: [] },
            afterPunctuation: { name: 'مسافة بعد الترقيم', icon: '⚡', issues: [] },
            beforeOpenBracket: { name: 'مسافة قبل الأقواس', icon: '(', issues: [] },
            afterOpenBracket: { name: 'مسافة داخل الأقواس (بداية)', icon: '(', issues: [] },
            beforeCloseBracket: { name: 'مسافة داخل الأقواس (نهاية)', icon: ')', issues: [] },
            afterCloseBracket: { name: 'مسافة بعد الأقواس', icon: ')', issues: [] }
        };

        // تصفية التكرارات الناتجة عن الـ regex
        const uniqueIssues = [];
        const seen = new Set();
        spacingIssues.forEach(issue => {
            const key = `${issue.error}-${issue.correction}-${issue.reason}`;
            if (!seen.has(key)) {
                seen.add(key);
                uniqueIssues.push(issue);
            }
        });

        uniqueIssues.forEach(issue => {
            if (categories[issue.category]) {
                categories[issue.category].issues.push(issue);
            }
        });

        let totalIssues = 0;
        for (const [key, cat] of Object.entries(categories)) {
            if (cat.issues.length > 0) {
                totalIssues += cat.issues.length;
                renderSpacingCategory(cat.name, cat.issues, cat.icon);
            }
        }

        if (totalIssues > 0) {
            globalMistakes.spaces = uniqueIssues;
            messageArea.innerHTML = `<div class="success-msg">تم العثور على ${totalIssues} ملاحظة في المسافات والترقيم!</div>`;
        } else {
            messageArea.innerHTML = "<div class='success-msg'>✨ مراجعة مثالية! جميع المسافات وعلامات الترقيم مطابقة للقواعد.</div>";
        }
    });
}


function renderSpacingCategory(categoryName, issues, icon) {
    const section = document.getElementById('section-spaces');
    const list = document.getElementById('list-spaces');
    const badge = document.getElementById('count-spaces');

    section.classList.remove("hidden");

    const currentCount = parseInt(badge.innerText) || 0;
    badge.innerText = currentCount + issues.length;

    const categoryHeader = document.createElement("li");
    categoryHeader.style.background = "linear-gradient(135deg, #667eea 0%, #764ba2 100%)";
    categoryHeader.style.color = "white";
    categoryHeader.style.padding = "10px 15px";
    categoryHeader.style.borderRadius = "8px";
    categoryHeader.style.fontWeight = "bold";
    categoryHeader.style.marginTop = "10px";
    categoryHeader.style.marginBottom = "5px";
    categoryHeader.innerHTML = `${icon} ${categoryName} (${issues.length})`;
    list.appendChild(categoryHeader);

    issues.forEach((mistake) => {
        const li = document.createElement("li");
        const errorEscaped = mistake.error.replace(/'/g, "\\'");
        const correctionEscaped = mistake.correction.replace(/'/g, "\\'");

        li.innerHTML = `
            <div class="correction-card">
                <div class="correction-header">
                    <span class="error-text">${mistake.error}</span>
                    <span class="arrow">←</span>
                    <span class="suggestion-text">${mistake.correction}</span>
                </div>
                <div class="reason-text">${mistake.reason}</div>
                <div class="actions-row">
                    <button class="icon-btn select-btn" onclick="highlightText('${errorEscaped}', 'spaces')">
                        👁️ تحديد
                    </button>
                    <button class="icon-btn apply-btn" onclick="applyCorrection('${errorEscaped}', '${correctionEscaped}', this)">
                        ✓ تطبيق
                    </button>
                </div>
            </div>
        `;
        list.appendChild(li);
    });
}

async function fetchAvailableModels() {
    const apiKey = document.getElementById("api-key").value.trim();
    const modelsList = document.getElementById("models-list");

    if (!apiKey) {
        modelsList.innerHTML = "<span style='color: red;'>الرجاء إدخال المفتاح أولاً</span>";
        return;
    }

    modelsList.innerHTML = "جاري التحميل...";

    try {
        const response = await fetch(`https://generativelanguage.googleapis.com/v1/models?key=${apiKey}`);
        const data = await response.json();

        if (data.models) {
            const supportedModels = data.models
                .filter(m => m.supportedGenerationMethods?.includes('generateContent'))
                .map(m => m.name.replace('models/', ''));

            if (supportedModels.length > 0) {
                modelsList.innerHTML = `<strong>النماذج المتاحة:</strong><br>${supportedModels.join('<br>')}`;
            } else {
                modelsList.innerHTML = "<span style='color: orange;'>لا توجد نماذج متاحة</span>";
            }
        } else {
            modelsList.innerHTML = "<span style='color: red;'>خطأ في جلب النماذج</span>";
        }
    } catch (error) {
        modelsList.innerHTML = `<span style='color: red;'>خطأ: ${error.message}</span>`;
    }
}

// --- Tashkeel Functions ---

async function handleRemoveTashkeel() {
    await Word.run(async (context) => {
        let range = context.document.getSelection();
        range.load("text");
        await context.sync();

        if (!range.text || range.text.trim().length === 0) {
            range = context.document.body;
            range.load("text");
            await context.sync();
        }

        const diacritics = "[ًٌٍَُِّْ]";
        const searchResults = range.search(diacritics, { matchWildcards: true });
        searchResults.load("items");
        await context.sync();

        for (let i = searchResults.items.length - 1; i >= 0; i--) {
            searchResults.items[i].insertText("", Word.InsertLocation.replace);
        }

        await context.sync();
    });
}

async function handleTashkeel() {
    const apiKey = document.getElementById("api-key").value.trim();
    const messageArea = document.getElementById("message-area");
    const progressArea = document.getElementById("progress-area");
    const progressFill = document.getElementById("progress-fill");
    const progressText = document.getElementById("progress-text");

    if (!apiKey) {
        messageArea.innerText = "الرجاء إدخال مفتاح API في الإعدادات.";
        return;
    }

    progressArea.classList.remove("hidden");
    progressText.innerText = "جاري تشكيل النص...";
    progressFill.style.width = "30%";

    await Word.run(async (context) => {
        let range = context.document.getSelection();
        range.load("text");
        await context.sync();

        if (!range.text || range.text.trim().length === 0) {
            range = context.document.body;
            range.load("text");
            await context.sync();
        }

        const text = range.text;
        if (!text || text.trim().length === 0) {
            progressArea.classList.add("hidden");
            return;
        }

        const model = document.getElementById("model-select").value.trim();
        const url = `https://generativelanguage.googleapis.com/v1/models/${model}:generateContent?key=${apiKey}`;

        const prompt = `
        أعد كتابة النص العربي التالي مع إضافة التشكيل الكامل (الحركات) عليه بدقة لغوية عالية.
        النص: "${text}"
        
        أرجع النص المشكول فقط بدون أي مقدمات أو شروحات أو تنسيق Markdown.
        `;

        try {
            const response = await fetch(url, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    contents: [{ parts: [{ text: prompt }] }],
                    generationConfig: { temperature: 0.2 }
                })
            });

            const data = await response.json();

            if (response.ok) {
                let tashkeelText = data.candidates?.[0]?.content?.parts?.[0]?.text;

                if (tashkeelText) {
                    tashkeelText = tashkeelText.trim().replace(/```(\w+)?/g, '').replace(/```/g, '');
                    range.insertText(tashkeelText, Word.InsertLocation.replace);
                    await context.sync();
                    progressFill.style.width = "100%";
                    setTimeout(() => progressArea.classList.add("hidden"), 1000);
                    messageArea.innerHTML = `<div class="success-msg">✅ تم تشكيل النص بنجاح!</div>`;
                } else {
                    messageArea.innerText = "فشل في الحصول على نتيجة من الذكاء الاصطناعي.";
                    progressArea.classList.add("hidden");
                }
            } else {
                const errorMsg = data.error?.message || "حدث خطأ غير معروف في الخدمة.";
                messageArea.innerHTML = `<div style="color: #ef4444; background: #fee2e2; padding: 10px; border-radius: 8px;">❌ خطأ من الخادم: ${errorMsg}</div>`;
                progressArea.classList.add("hidden");
            }
        } catch (e) {
            console.error(e);
            messageArea.innerText = "فشل في الاتصال: تأكد من الإنترنت ومفتاح API.";
            progressArea.classList.add("hidden");
        }
    });
}

// --- Main Check Function ---

async function runFullCheck() {
    const apiKey = document.getElementById("api-key").value.trim();
    const messageArea = document.getElementById("message-area");
    const resultsArea = document.getElementById("results-area");
    const progressArea = document.getElementById("progress-area");
    const progressFill = document.getElementById("progress-fill");
    const progressText = document.getElementById("progress-text");

    if (!apiKey) {
        messageArea.innerHTML = `<div class="success-msg" style="background: #fff5f5; color: #e53e3e;">⚠️ يرجى إدخال مفتاح API في الإعدادات أولاً.</div>`;
        return;
    }

    // Reset UI
    document.querySelectorAll('.category-section').forEach(el => el.classList.add('hidden'));
    document.querySelectorAll('.section-content').forEach(el => el.classList.add('hidden'));
    document.querySelectorAll('ul[id^="list-"]').forEach(el => el.innerHTML = '');
    document.querySelectorAll('.count-badge').forEach(el => el.innerText = '0');
    messageArea.innerText = "";
    globalMistakes = { spaces: [], spelling: [], grammar: [], punctuation: [], style: [] };

    progressArea.classList.remove("hidden");
    progressText.innerText = "جاري تحضير النص للمراجعة...";
    progressFill.style.width = "5%";

    // إظهار النافذة المنبثقة فوراً
    document.getElementById("results-modal").classList.remove("hidden");
    resultsArea.classList.add("hidden"); // إخفاء منطقة النتائج في البداية لإظهار شريط التقدم فقط

    await Word.run(async (context) => {
        let range = context.document.getSelection();
        range.load(["text", "isEmpty"]);
        await context.sync();

        // التبديل لكامل المستند إذا لم يكن هناك تحديد
        if (range.isEmpty || range.text.trim().length === 0) {
            range = context.document.body;
            range.load("text");
            await context.sync();
            progressText.innerText = "جاري مراجعة كامل المستند...";
        } else {
            progressText.innerText = "جاري مراجعة التحديد...";
        }

        const fullText = range.text;
        if (!fullText || fullText.trim().length === 0) {
            messageArea.innerText = "المستند أو التحديد فارغ.";
            progressArea.classList.add("hidden");
            return;
        }

        // 1. فحص المسافات (محلي وسريع)
        progressFill.style.width = "10%";
        const spacingErrors = [];
        // سنستخدم نفس منطق الـ Regex السريع لفحص الأخطاء بدلاً من البحث في وورد لتجنب البطء
        const lines = fullText.split(/[\r\n]+/);
        lines.forEach(line => {
            const multiSpaceRegex = / {2,}/g;
            let match;
            while ((match = multiSpaceRegex.exec(line)) !== null) {
                spacingErrors.push({
                    error: match[0],
                    correction: " ",
                    reason: "مسافة زائدة"
                });
            }
        });

        if (spacingErrors.length > 0) {
            globalMistakes.spaces = spacingErrors;
            renderMistakes('spaces', spacingErrors);
        }

        // 2. التحليل الذكي عبر API
        const chunkSize = 4000; // تقليل حجم القطعة لضمان استجابة أسرع وأدق
        const chunks = [];
        for (let i = 0; i < fullText.length; i += chunkSize) {
            chunks.push(fullText.substring(i, i + chunkSize));
        }

        let totalAiErrors = 0;
        const batchSize = 2;

        for (let i = 0; i < chunks.length; i += batchSize) {
            const batch = chunks.slice(i, i + batchSize);
            const currentProgress = 15 + (((i + batchSize) / chunks.length) * 85);
            progressFill.style.width = `${Math.min(currentProgress, 100)}%`;
            progressText.innerText = `تحليل ذكي: جاري معالجة الجزء ${Math.min(i + batchSize, chunks.length)} من ${chunks.length}...`;

            const promises = batch.map(chunk => analyzeChunk(chunk, apiKey));
            const results = await Promise.all(promises);

            for (const result of results) {
                if (result && result.error) {
                    messageArea.innerHTML = `<div class="success-msg" style="background: #fff5f5; color: #e53e3e;">❌ خطأ API: ${result.message}</div>`;
                    progressArea.classList.add("hidden");
                    return;
                }

                if (result) {
                    for (const [category, mistakes] of Object.entries(result)) {
                        if (mistakes && mistakes.length > 0 && globalMistakes[category]) {
                            totalAiErrors += mistakes.length;
                            mistakes.forEach(m => globalMistakes[category].push(m));
                            renderMistakes(category, mistakes);
                        }
                    }
                }
            }
        }

        progressArea.classList.add("hidden");

        const total = totalAiErrors + globalMistakes.spaces.length;
        resultsArea.classList.remove("hidden"); // إظهار منطقة النتائج دائماً عند الانتهاء

        if (total > 0) {
            messageArea.innerHTML = `<div class="success-msg">✅ اكتملت المراجعة! تم العثور على ${total} ملاحظة.</div>`;
        } else {
            const successHtml = "<div class='success-msg' style='margin-top: 20px;'>✨ المستند سليم ولم يتم العثور على أخطاء!</div>";
            messageArea.innerHTML = successHtml;
            // إظهار رسالة النجاح داخل النافذة المنبثقة أيضاً إذا كانت فارغة
            resultsArea.innerHTML = successHtml;
        }
    });
}


async function analyzeChunk(text, apiKey) {
    let model = document.getElementById("model-select").value.trim();
    // تحسين اسم النموذج إذا لزم الأمر
    if (model === "gemini-2.5-flash") model = "gemini-2.5-flash";
    else if (model === "gemini-2.0-flash") model = "gemini-2.0-flash";
    else if (model === "gemini-1.5-flash") model = "gemini-1.5-flash";

    // استخدام v1 بدلاً من v1beta لضمان الاستقرار مع النماذج الحالية
    const url = `https://generativelanguage.googleapis.com/v1/models/${model}:generateContent?key=${apiKey}`;

    const prompt = `
    حلل النص التالي واستخرج جميع الأخطاء الإملائية والنحوية والترقيمية: "${text}"
    
    مهم: ركز جداً على الهمزات (أ، إ، آ، ء) والتاء المربوطة والهاء والياء والألف المقصورة.
    
    يجب أن تعيد النتيجة بصيغة JSON حصراً كما يلي:
    {
        "spelling": [{"error": "كلمة خطأ", "correction": "تصحيح", "reason": "سبب"}],
        "grammar": [],
        "punctuation": [],
        "style": []
    }
    تجنب كتابة أي نص خارج الـ JSON.
    `;

    try {
        const response = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                contents: [{ parts: [{ text: prompt }] }],
                generationConfig: { temperature: 0.2 }
            })
        });

        if (!response.ok) {
            const errData = await response.json();
            return { error: true, message: errData.error?.message || "خطأ في نموذج الذكاء الاصطناعي" };
        }

        const data = await response.json();
        let content = data.candidates?.[0]?.content?.parts?.[0]?.text;
        if (!content) return { error: true, message: "لم تصل استجابة من الذكاء الاصطناعي" };

        content = content.replace(/```json/g, '').replace(/```/g, '').trim();
        return JSON.parse(content);
    } catch (e) {
        console.error("AI Error:", e);
        return { error: true, message: "فشل الاتصال: تأكد من مفتاح API والإنترنت" };
    }
}

function renderMistakes(category, mistakes) {
    const section = document.getElementById(`section-${category}`);
    const list = document.getElementById(`list-${category}`);
    const badge = document.getElementById(`count-${category}`);

    section.classList.remove("hidden");

    const currentCount = parseInt(badge.innerText) || 0;
    badge.innerText = currentCount + mistakes.length;

    mistakes.forEach((mistake) => {
        const li = document.createElement("li");
        const errorEscaped = mistake.error.replace(/'/g, "\\'");
        const correctionEscaped = mistake.correction.replace(/'/g, "\\'");

        li.innerHTML = `
            <div class="correction-card">
                <div class="correction-header">
                    <span class="error-text">${mistake.error}</span>
                    <span class="arrow">←</span>
                    <span class="suggestion-text">${mistake.correction}</span>
                </div>
                <div class="reason-text">${mistake.reason}</div>
                <div class="actions-row">
                    <button class="icon-btn select-btn" onclick="highlightText('${errorEscaped}', '${category}')">
                        👁️ تحديد
                    </button>
                    <button class="icon-btn apply-btn" onclick="applyCorrection('${errorEscaped}', '${correctionEscaped}', this)">
                        ✓ تطبيق
                    </button>
                </div>
            </div>
        `;
        list.appendChild(li);
    });
}

window.highlightText = async (text, category) => {
    await Word.run(async (context) => {
        const results = context.document.body.search(text, { matchCase: false, matchWholeWord: false });
        results.load("items");
        await context.sync();

        if (results.items.length > 0) {
            const foundRange = results.items[0];

            if (category === 'punctuation') {
                try {
                    let expanded = foundRange.expand(Word.RangeUnit.word);
                    let before = expanded.getRange(Word.RangeLocation.start).getRange(Word.RangeLocation.before).expand(Word.RangeUnit.word);
                    let after = expanded.getRange(Word.RangeLocation.end).getRange(Word.RangeLocation.after).expand(Word.RangeUnit.word);
                    let finalRange = before.expandTo(after);
                    finalRange.select();
                } catch (e) {
                    foundRange.select();
                }
            } else {
                foundRange.select();
            }
            await context.sync();
        }
    });
};

window.applyCorrection = async (error, correction, btn) => {
    let success = false;
    await Word.run(async (context) => {
        const results = context.document.body.search(error, { matchCase: false, matchWholeWord: false });
        results.load("items");
        await context.sync();
        if (results.items.length > 0) {
            results.items[0].insertText(correction, Word.InsertLocation.replace);
            await context.sync();
            success = true;

            // Add to history
            operationHistory.push({
                error: error,
                correction: correction,
                timestamp: new Date().toLocaleTimeString('ar-EG'),
                status: 'تم القبول'
            });
        }
    });

    if (success) {
        btn.innerText = "تم";
        btn.disabled = true;
        btn.closest('li').style.background = "#e6fffa";
    }
};

window.renderHistory = () => {
    const list = document.getElementById("history-list");
    if (operationHistory.length === 0) {
        list.innerHTML = '<p class="empty-msg">لا توجد عمليات سابقة.</p>';
        return;
    }

    list.innerHTML = operationHistory.slice().reverse().map((item, index) => `
        <div class="history-item">
            <div class="item-text">
                <strong>تعديل:</strong> <span class="error-text">${item.error}</span> ← <span class="suggestion-text">${item.correction}</span>
            </div>
            <div class="reason-text">الوقت: ${item.timestamp}</div>
            <button class="undo-btn" onclick="undoCorrection(${operationHistory.length - 1 - index})">التراجع عن التعديل</button>
        </div>
    `).join('');
};

window.undoCorrection = async (index) => {
    const item = operationHistory[index];
    await Word.run(async (context) => {
        const results = context.document.body.search(item.correction, { matchCase: false, matchWholeWord: false });
        results.load("items");
        await context.sync();
        if (results.items.length > 0) {
            results.items[0].insertText(item.error, Word.InsertLocation.replace);
            await context.sync();

            // Remove from history
            operationHistory.splice(index, 1);
            renderHistory();

            // Alert user
            const messageArea = document.getElementById("message-area");
            messageArea.innerHTML = `<div class="success-msg" style="background: #fff5f5; color: #e53e3e;">تم التراجع عن التعديل بنجاح.</div>`;
        }
    });
};

async function exportToPDF() {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({
        orientation: 'p',
        unit: 'mm',
        format: 'a4'
    });

    // Note: Proper Arabic support in jsPDF requires embedding a font.
    // For this demo, we'll create a structured report.
    doc.setFontSize(22);
    doc.text("Report: Linguistic Review", 105, 20, { align: "center" });
    doc.setFontSize(12);
    doc.text(`Generated on: ${new Date().toLocaleString()}`, 105, 30, { align: "center" });

    let yPos = 40;
    const categories = {
        spelling: "Spelling Errors",
        grammar: "Grammar Errors",
        punctuation: "Punctuation Issues",
        style: "Style Suggestions",
        spaces: "Spacing Issues"
    };

    for (const [key, label] of Object.entries(categories)) {
        const mistakes = globalMistakes[key];
        if (mistakes && mistakes.length > 0) {
            doc.setFontSize(16);
            doc.setTextColor(37, 99, 235);
            doc.text(label, 20, yPos);
            yPos += 10;

            doc.setFontSize(10);
            doc.setTextColor(0, 0, 0);

            mistakes.forEach((m, idx) => {
                if (yPos > 270) {
                    doc.addPage();
                    yPos = 20;
                }
                doc.text(`${idx + 1}. [Original]: ${m.error} -> [Suggested]: ${m.correction}`, 25, yPos);
                yPos += 6;
                doc.setFontSize(8);
                doc.setTextColor(100, 116, 139);
                doc.text(`Reason: ${m.reason}`, 30, yPos);
                doc.setFontSize(10);
                doc.setTextColor(0, 0, 0);
                yPos += 8;
            });
            yPos += 5;
        }
    }

    doc.save("Linguistic_Report.pdf");
}

window.applyBatch = async (category) => {
    const mistakes = globalMistakes[category];
    if (!mistakes || mistakes.length === 0) return;

    await Word.run(async (context) => {
        if (category === 'spaces') {
            const results = context.document.body.search(" {2,}", { matchWildcards: true });
            results.load("items");
            await context.sync();
            for (const item of results.items) {
                item.insertText(" ", Word.InsertLocation.replace);
            }
        } else {
            for (const m of mistakes) {
                const results = context.document.body.search(m.error, { matchCase: false, matchWholeWord: false });
                results.load("items");
                await context.sync();
                for (const item of results.items) {
                    item.insertText(m.correction, Word.InsertLocation.replace);
                }
            }
        }
        await context.sync();
    });

    const list = document.getElementById(`list-${category}`);
    list.querySelectorAll('.apply-btn').forEach(btn => {
        btn.innerText = "تم";
        btn.disabled = true;
        btn.closest('li').style.background = "#e6fffa";
    });
};

// --- New Text Cleaner Functions (Optimized/Batched) ---

async function handleReverseBrackets() {
    await Word.run(async (context) => {
        let range = context.document.getSelection();
        range.load("text");
        await context.sync();

        // إذا لم يكن هناك نص محدد، نستخدم كامل المستند
        let targetRange = range;
        if (!range.text || range.text.trim().length === 0) {
            targetRange = context.document.body;
            targetRange.load("text");
            await context.sync();
        }

        const text = targetRange.text;
        if (!text || text.trim().length === 0) {
            document.getElementById("message-area").innerText = "المستند فارغ.";
            return;
        }

        // خريطة شاملة لجميع أنواع الأقواس
        const bracketMap = {
            '(': ')', ')': '(',
            '[': ']', ']': '[',
            '{': '}', '}': '{',
            '﴿': '﴾', '﴾': '﴿',
            '«': '»', '»': '«',
            '‹': '›', '›': '‹',
            '<': '>', '>': '<',
            '〔': '〕', '〕': '〔',
            '【': '】', '】': '【',
            '〖': '〗', '〗': '〖',
            '〚': '〛', '〛': '〚',
            '⟨': '⟩', '⟩': '⟨'
        };

        let result = "";
        let flipCount = 0;
        let stats = {};

        // 1. عكس الأقواس وتحصيل الإحصائيات
        for (let char of text) {
            if (bracketMap[char]) {
                result += bracketMap[char];
                flipCount++;
                stats[char] = (stats[char] || 0) + 1;
            } else {
                result += char;
            }
        }

        // 2. تحسين الدقة: معالجة المسافات حول الأقواس (اختياري ولكن يزيد الدقة)
        // إزالة المسافات بعد القوس الافتتاحي وقبل القوس الختامي
        // (نص) -> (نص)
        result = result.replace(/([(\[{\+«‹<〔【〖〚⟨])\s+/g, '$1');
        result = result.replace(/\s+([)\]\}+»›>〕】〗〛⟩])/g, '$1');

        // 3. التحقق من التوازن (Balance Check)
        const stack = [];
        const openingBrackets = '([{﴿«‹<〔【〖〚⟨';
        const closingBrackets = ')]}﴾»›>〕】〗〛⟩';
        const pairMap = {
            ')': '(', ']': '[', '}': '{', '﴾': '﴿', '»': '«', '›': '‹', '>': '<', '〕': '〔', '】': '【', '〗': '〖', '〛': '〚', '⟩': '⟨'
        };
        let unbalanced = false;

        for (let char of result) {
            if (openingBrackets.includes(char)) {
                stack.push(char);
            } else if (closingBrackets.includes(char)) {
                if (stack.length === 0 || stack.pop() !== pairMap[char]) {
                    unbalanced = true;
                    break;
                }
            }
        }
        if (stack.length > 0) unbalanced = true;

        if (flipCount > 0) {
            targetRange.insertText(result, Word.InsertLocation.replace);
            await context.sync();

            let balanceWarning = unbalanced ? `<div style="color: #f59e0b; margin-top: 5px; font-size: 0.8rem;">⚠️ تنبيه: تم رصد عدم توازن في الأقواس (قوس مفقود أو زائد).</div>` : "";

            document.getElementById("message-area").innerHTML = `
                <div class="success-msg">
                    ✅ تم تصحيح ${flipCount} قوس بنجاح!
                    <div style="font-size: 0.8rem; margin-top: 5px; font-weight: normal;">
                        تم ضبط الاتجاهات وتنظيف المسافات الداخلية.
                    </div>
                    ${balanceWarning}
                </div>
            `;

            operationHistory.push({
                error: "تصحيح أقواس",
                correction: "تم معالجة " + flipCount + " قوس",
                timestamp: new Date().toLocaleTimeString('ar-EG'),
                status: 'تم التنفيذ'
            });
        } else {
            document.getElementById("message-area").innerText = "لم يتم العثور على أي أقواس في النص لمعالجتها.";
        }
    });
}

async function handleOrnateBrackets() {
    await Word.run(async (context) => {
        const range = getRange(context);
        // Pattern: (text) -> ﴿text﴾
        const foundItems = range.search("\\([!)(]*\\)", { matchWildcards: true });
        foundItems.load("items");
        await context.sync();

        if (foundItems.items.length === 0) {
            document.getElementById("message-area").innerText = "لم يتم العثور على أقواس عادية.";
            return;
        }

        let operations = [];
        foundItems.items.forEach(item => {
            let startSearch = item.search("(", { matchWildcards: false });
            let endSearch = item.search(")", { matchWildcards: false });
            startSearch.load("items");
            endSearch.load("items");
            operations.push({ start: startSearch, end: endSearch });
        });

        await context.sync(); // Load all secondary search results

        let changesCount = 0;
        operations.forEach(op => {
            if (op.start.items.length > 0) {
                op.start.items[0].insertText("﴿", Word.InsertLocation.replace);
                changesCount++;
            }
            if (op.end.items.length > 0) {
                const lastIdx = op.end.items.length - 1;
                op.end.items[lastIdx].insertText("﴾", Word.InsertLocation.replace);
            }
        });

        await context.sync(); // Sync after all replacements
        document.getElementById("message-area").innerText = `تم زخرفة ${changesCount} قوساً.`;
    });
}

async function handleRemoveEmptyLines() {
    await Word.run(async (context) => {
        let range = context.document.getSelection();
        range.load("text");
        await context.sync();

        let target = range;
        // إذا لم يكن هناك تحديد، نطبق على كامل المستند
        if (!range.text || range.text.trim().length === 0) {
            target = context.document.body;
        }

        const paragraphs = target.paragraphs;
        paragraphs.load("text");
        await context.sync();

        let deleteCount = 0;
        // البدء من النهاية لتجنب تغيير الفهرس أثناء الحذف
        for (let i = paragraphs.items.length - 1; i >= 0; i--) {
            // إذا كان السطر فارغاً تماماً أو يحتوي على مسافات فقط
            if (paragraphs.items[i].text.trim() === "") {
                // ملاحظة: لا يمكن حذف الفقرة الأخيرة إذا كانت هي الوحيدة في المستند
                if (paragraphs.items.length > 1) {
                    paragraphs.items[i].delete();
                    deleteCount++;
                }
            }
        }

        await context.sync();

        if (deleteCount > 0) {
            document.getElementById("message-area").innerHTML = `<div class="success-msg">✅ تم حذف ${deleteCount} سطر فارغ بنجاح!</div>`;

            operationHistory.push({
                error: "أسطر فارغة",
                correction: "تم الحذف",
                timestamp: new Date().toLocaleTimeString('ar-EG'),
                status: 'تم التنظيف'
            });
        } else {
            document.getElementById("message-area").innerText = "لا توجد أسطر فارغة لحذفها.";
        }
    });
}

function getRange(context) {
    // Helper to get selection or whole body if selection is empty/point
    const selection = context.document.getSelection();
    // We can't synchronously check selection length.
    // So we default to searching the selection. If the user selected nothing (insertion point),
    // search on selection might return nothing or just search that point?
    // Word behavior: Searching an insertion point usually searches nothing.
    // We want: If selection is collapsed, search BODY.
    // But we need to load 'selection' prop 'text' or 'isEmpty' which costs a sync.
    // Optimization: Just allow searching selection. If user wants body, select all (Ctrl+A).
    // PREVIOUS BEHAVIOR: explicitly checked body. Users prefer "Do what I mean".
    // Let's add that check.

    // BUT we cannot await inside this helper if we want to use it inline.
    // So we will just use body directly for now as per "Fast" request? 
    // No, context-sensitive is better.
    // Let's do: return context.document.body; (As per user request "Make it work like Word addin" -> usually operates on document).
    // Actually, let's stick to Body for "Clean All" actions. It's safer for "Text Cleaner".
}

async function handleWrapText(openBracket, closeBracket) {
    await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.load("text");
        await context.sync();

        if (range.text && range.text.trim().length > 0) {
            const originalText = range.text;
            const newText = openBracket + originalText + closeBracket;
            range.insertText(newText, Word.InsertLocation.replace);

            // Re-select to show result
            range.select();
            await context.sync();

            operationHistory.push({
                error: originalText,
                correction: newText,
                timestamp: new Date().toLocaleTimeString('ar-EG'),
                status: 'تم التقويس'
            });
        } else {
            const messageArea = document.getElementById("message-area");
            if (messageArea) {
                messageArea.innerHTML = `<div class="success-msg" style="color: #f59e0b;">⚠️ يرجى تحديد نص لتقويسه أولاً.</div>`;
            }
        }
    });
}

// Ensure the function is global for HTML access
window.handleWrapText = handleWrapText;

async function handleRemoveAllBrackets() {
    await Word.run(async (context) => {
        let range = context.document.getSelection();
        range.load("text");
        await context.sync();

        // إذا لم يكن هناك نص محدد، نستخدم كامل المستند
        let targetRange = range;
        if (!range.text || range.text.trim().length === 0) {
            targetRange = context.document.body;
            targetRange.load("text");
            await context.sync();
        }

        const text = targetRange.text;
        if (!text || text.trim().length === 0) {
            document.getElementById("message-area").innerText = "المستند فارغ.";
            return;
        }

        // إزالة جميع أنواع الأقواس وعلامات التنصيص
        const cleanedText = text
            // الأقواس الشائعة
            .replace(/\(/g, '').replace(/\)/g, '')     // ( )
            .replace(/\[/g, '').replace(/\]/g, '')     // [ ]
            .replace(/\{/g, '').replace(/\}/g, '')     // { }
            .replace(/﴿/g, '').replace(/﴾/g, '')       // ﴿ ﴾
            .replace(/«/g, '').replace(/»/g, '')       // « »
            .replace(/</g, '').replace(/>/g, '')       // < >
            // علامات التنصيص - جميع الأنواع
            .replace(/"/g, '')                          // " علامة عادية
            .replace(/'/g, '')                          // ' علامة عادية
            .replace(/“/g, '').replace(/”/g, '')       // “ ” علامات مزدوجة ذكية
            .replace(/‘/g, '').replace(/’/g, '')       // ‘ ’ علامات مفردة ذكية
            .replace(/‚/g, '').replace(/„/g, '')       // ‚ „
            .replace(/‹/g, '').replace(/›/g, '')       // ‹ ›
            .replace(/〔/g, '').replace(/〕/g, '')      // 〔 〕
            .replace(/【/g, '').replace(/】/g, '');     // 【 】

        const removedCount = text.length - cleanedText.length;

        if (removedCount > 0) {
            targetRange.insertText(cleanedText, Word.InsertLocation.replace);
            await context.sync();

            document.getElementById("message-area").innerHTML = `
                <div class="success-msg">
                    ✅ تم حذف ${removedCount} قوس بنجاح!
                </div>
            `;

            operationHistory.push({
                error: "حذف أقواس",
                correction: `تم حذف ${removedCount} قوس`,
                timestamp: new Date().toLocaleTimeString('ar-EG'),
                status: 'تم التنفيذ'
            });
        } else {
            document.getElementById("message-area").innerText = "لم يتم العثور على أي أقواس في النص.";
        }
    });
}

async function handleFastAutoFix() {
    const progressArea = document.getElementById("progress-area");
    const progressFill = document.getElementById("progress-fill");
    const progressText = document.getElementById("progress-text");
    const messageArea = document.getElementById("message-area");

    progressArea.classList.remove("hidden");
    progressText.innerText = "جاري الإصلاح التلقائي الفوري...";
    progressFill.style.width = "50%";

    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        // Use selection or body
        let targetRange = (selection.text && selection.text.trim().length > 0) ? selection : context.document.body;
        targetRange.load("text");
        await context.sync();

        let text = targetRange.text;
        if (!text || text.trim().length === 0) {
            messageArea.innerText = "لا يوجد نص للمعالجة.";
            progressArea.classList.add("hidden");
            return;
        }

        const originalText = text;

        // --- Fast Local Regex Replacements (Offline & Instant) ---
        text = text
            .replace(/ {2,}/g, ' ')                        // 1. Multiple spaces -> Single
            .replace(/\s+([،؛.:!؟)])/g, '$1')              // 2. Remove space BEFORE punctuation
            .replace(/([،؛.:!؟])(?![ \s\)\d\u0660-\u0669])/g, '$1 ') // 3. Add space AFTER punctuation (if not followed by digit/space)
            .replace(/([(\[﴿«])\s+/g, '$1')                 // 5. Remove space AFTER opening brackets
            .replace(/\s+([)\]﴾»])/g, '$1');                // 6. Remove space BEFORE closing brackets

        if (text !== originalText) {
            targetRange.insertText(text, Word.InsertLocation.replace);
            await context.sync();

            progressFill.style.width = "100%";
            messageArea.innerHTML = `<div class="success-msg">⚡ تم الإصلاح التلقائي الشامل فوراً (بدون إنترنت)!</div>`;

            operationHistory.push({
                error: "إصلاح تلقائي للمسافات",
                correction: "تم تنظيف النص بالكامل",
                timestamp: new Date().toLocaleTimeString('ar-EG'),
                status: 'إصلاح سريع'
            });
        } else {
            messageArea.innerHTML = `<div class="success-msg" style="background: #f0f9ff; color: #0369a1;">✨ النص سليم بالفعل، لم تكن هناك حاجة لتعديلات.</div>`;
        }

        setTimeout(() => progressArea.classList.add("hidden"), 500);
    });
}

window.handleFastAutoFix = handleFastAutoFix;


async function handleLocalReview() {
    const messageArea = document.getElementById("message-area");
    const resultsArea = document.getElementById("results-area");
    const progressArea = document.getElementById("progress-area");
    const progressFill = document.getElementById("progress-fill");
    const progressText = document.getElementById("progress-text");

    // Reset UI
    document.querySelectorAll('.category-section').forEach(el => el.classList.add('hidden'));
    document.querySelectorAll('.section-content').forEach(el => el.classList.add('hidden'));
    document.querySelectorAll('ul[id^="list-"]').forEach(el => el.innerHTML = '');
    document.querySelectorAll('.count-badge').forEach(el => el.innerText = '0');
    messageArea.innerText = "";
    globalMistakes = { spelling: [], grammar: [], punctuation: [], style: [], spaces: [] };

    progressArea.classList.remove("hidden");
    progressText.innerText = "جاري تشغيل محرك التدقيق المحلي العبقري...";
    progressFill.style.width = "15%";
    resultsArea.classList.remove("hidden");

    await Word.run(async (context) => {
        let range = context.document.getSelection();
        range.load("text");
        await context.sync();

        if (!range.text || range.text.trim().length === 0) {
            range = context.document.body;
            range.load("text");
            await context.sync();
        }

        const text = range.text;
        if (!text || text.trim().length === 0) {
            messageArea.innerText = "المستند فارغ.";
            progressArea.classList.add("hidden");
            return;
        }

        const localMistakes = {
            spelling: [],
            grammar: [],
            style: []
        };

        // --- مكتبة القواعد المحلية الضخمة ---
        const rules = [
            // 1. الأخطاء الإملائية المشهورة (Spelling)
            { find: /ان شاء الله/g, replace: "إن شاء الله", reason: "فصل 'إن' الشرطية عن فعل المشيئة", cat: "spelling" },
            { find: /\bانشاء الله\b/g, replace: "إن شاء الله", reason: "الإنشاء يعني الإيجاد، والمقصود هنا المشيئة", cat: "spelling" },
            { find: /\bأسم\b/g, replace: "اسم", reason: "همزة وصل في كلمة 'اسم'", cat: "spelling" },
            { find: /\bأبن\b/g, replace: "ابن", reason: "همزة وصل في كلمة 'ابن'", cat: "spelling" },
            { find: /\bأبنة\b/g, replace: "ابنة", reason: "همزة وصل في كلمة 'ابنة'", cat: "spelling" },
            { find: /\bأمرأة\b/g, replace: "امرأة", reason: "همزة وصل في كلمة 'امرأة'", cat: "spelling" },
            { find: /\bأثنان\b/g, replace: "اثنان", reason: "همزة وصل في كلمة 'اثنان'", cat: "spelling" },
            { find: /\bأثنتان\b/g, replace: "اثنتان", reason: "همزة وصل في كلمة 'اثنتان'", cat: "spelling" },
            { find: /\bإستمارة\b/g, replace: "استمارة", reason: "همزة وصل (مصدر خماسي أو سداسي)", cat: "spelling" },
            { find: /\bإستخدام\b/g, replace: "استخدام", reason: "همزة وصل (مصدر سداسي)", cat: "spelling" },
            { find: /\bإستقبال\b/g, replace: "استقبال", reason: "همزة وصل (مصدر سداسي)", cat: "spelling" },
            { find: /\bإستقالة\b/g, replace: "استقالة", reason: "همزة وصل (مصدر سداسي)", cat: "spelling" },
            { find: /([^أإآ])اذا /g, replace: "$1إذا ", reason: "همزة قطع في 'إذا'", cat: "spelling" },
            { find: /([^أإآ])ان /g, replace: "$1إن ", reason: "همزة قطع في 'إن' أو 'أن'", cat: "spelling" },
            { find: /([^أإآ])الى /g, replace: "$1إلى ", reason: "همزة قطع في حرف 'إلى'", cat: "spelling" },
            { find: /\bشئ\b/g, replace: "شيء", reason: "الهمزة متطرفة بعد باء ساكنة تكتب على السطر", cat: "spelling" },
            { find: /\bدفئ\b/g, replace: "دفء", reason: "الهمزة متطرفة بعد ساكن تكتب على السطر", cat: "spelling" },
            { find: /\bبطئ\b/g, replace: "بطيء", reason: "تكتب الياء ثم الهمزة على السطر", cat: "spelling" },
            { find: /\bمسئول\b/g, replace: "مسؤول", reason: "الرسم الإملائي الأصح للهمزة المتوسطة المضمومة", cat: "spelling" },
            { find: /\bشؤون\b/g, replace: "شؤون", reason: "تنبيه: تكتب شؤون (رسم مصري) أو شئون، والأولى شؤون", cat: "spelling" },
            { find: /\bالذى\b/g, replace: "الذي", reason: "تكتب الياء تحتها نقطتان", cat: "spelling" },
            { find: /\bالتى\b/g, replace: "التي", reason: "تكتب الياء تحتها نقطتان", cat: "spelling" },
            { find: /\bهذى\b/g, replace: "هذه", reason: "الهاء المربوطة لا تنقط في كلمة 'هذه'", cat: "spelling" },

            // 2. الأخطاء النحوية المشهورة (Grammar)
            { find: /اللهم صلي/g, replace: "اللهم صلِّ", reason: "فعل أمر للمفرد المذكر يبنى على حذف حرف العلة", cat: "grammar" },
            { find: /صل الله عليه/g, replace: "صلى الله عليه", reason: "فعل ماضٍ مسند لاسم الجلالة (لا يحذف حرف العلة)", cat: "grammar" },
            { find: /لم يشاء/g, replace: "لم يشأ", reason: "جزم الفعل المعتل الوسط (التقاء ساكنين)", cat: "grammar" },
            { find: /لا تنسى /g, replace: "لا تنسَ ", reason: "لا الناهية تجزم الفعل المضارع بحذف حرف العلة", cat: "grammar" },
            { find: /لم ينمو/g, replace: "لم ينمُ", reason: "جزم المضارع المعتل الآخر بحذف الواو", cat: "grammar" },
            { find: /لم يدعو/g, replace: "لم يدعُ", reason: "جزم المضارع المعتل الآخر بحذف الواو", cat: "grammar" },
            { find: /لن ينمو/g, replace: "لن ينمو", reason: "فتح الواو في النصب (سليم)", cat: "style" }, // مجرد مثال
            { find: /\bالغير /g, replace: "غير الـ", reason: "كلمة 'غير' لا تدخل عليها (ال) التعريف، بل تدخل على ما بعدها", cat: "grammar" },

            // 3. أخطاء الأسلوب والتعبيرات الشائعة (Style)
            { find: /\bمبروك\b/g, replace: "مبارك", reason: "مبارك من البركة، أما مبروك فمن بروك الناقة", cat: "style" },
            { find: /\bبناءا على\b/g, replace: "بناءً على", reason: "الهمزة المسبوقة بألف لا ترسم بعدها ألف تنوين", cat: "style" },
            { find: /\bسويا\b/g, replace: "معاً", reason: "سوياً تعني الاستواء والاعتدال، ومعاً تعني الاجتماع", cat: "style" },
            { find: /\bكافة /g, replace: "كلمة (كافة) يفضل أن تأتي في نهاية الجملة فتقول: المواضيع كافة", cat: "style" },
            { find: /\bلماذا لا تقم\b/g, replace: "لماذا لا تقوم", reason: "تنبيه: 'لا' هنا نافية وليست جازمة", cat: "grammar" },
            { find: /اعتذر منه/g, replace: "اعتذر إليه", reason: "الفعل 'اعتذر' يتعدى بـ (إلى) للشخص وبـ (عن) للخطأ", cat: "style" },
            { find: /أجاب على/g, replace: "أجاب عن", reason: "الفعل 'أجاب' يتعدى بـ (عن)", cat: "style" }
        ];

        progressText.innerText = "تحليل النص ومطابقة القواعد...";
        progressFill.style.width = "50%";

        const lines = text.split(/[\r\n]+/);
        lines.forEach(line => {
            rules.forEach(rule => {
                let match;
                // إعادة ضبط الـ search index للـ regex العالمي
                rule.find.lastIndex = 0;
                while ((match = rule.find.exec(line)) !== null) {
                    localMistakes[rule.cat].push({
                        error: match[0],
                        correction: (typeof rule.replace === 'string' && rule.replace.includes('$'))
                            ? match[0].replace(rule.find, rule.replace)
                            : rule.replace,
                        reason: rule.reason,
                        category: rule.cat
                    });
                    // منع اللانهائية في حال كان الـ regex غير عالمي
                    if (!rule.find.global) break;
                }
            });
        });

        // فحص التاء المربوطة والهاء في نهايات الكلمات (منطق ذكي)
        const checkTaa = (line) => {
            // كلمات شائعة تنتهي بالهاء بدلا من التاء
            const commonTaaFixes = [
                { reg: /مدرسه\b/g, corr: "مدرسة", res: "تاء مربوطة تنطق هاء عند الوقف" },
                { reg: /مكتبه\b/g, corr: "مكتبة", res: "تاء مربوطة" },
                { reg: /قصه\b/g, corr: "قصة", res: "تاء مربوطة" },
                { reg: /جامعه\b/g, corr: "جامعة", res: "تاء مربوطة" }
            ];
            commonTaaFixes.forEach(f => {
                let match;
                while ((match = f.reg.exec(line)) !== null) {
                    localMistakes.spelling.push({ error: match[0], correction: f.corr, reason: f.res, category: "spelling" });
                }
            });
        };
        lines.forEach(checkTaa);

        progressFill.style.width = "90%";

        // تصفية التكرارات
        const finalMistakes = {};
        for (const [cat, list] of Object.entries(localMistakes)) {
            const seen = new Set();
            finalMistakes[cat] = list.filter(m => {
                const key = `${m.error}-${m.correction}-${m.reason}`;
                if (seen.has(key)) return false;
                seen.add(key);
                return true;
            });
        }

        progressFill.style.width = "100%";
        progressArea.classList.add("hidden");

        // عرض النتائج
        let total = 0;
        for (const [cat, mistakes] of Object.entries(finalMistakes)) {
            if (mistakes.length > 0) {
                total += mistakes.length;
                globalMistakes[cat] = mistakes;
                renderMistakes(cat, mistakes);
            }
        }

        if (total > 0) {
            messageArea.innerHTML = `<div class="success-msg">✅ تم التدقيق المحلي الشامل! وجدنا ${total} ملاحظة لغوية مشهورة.</div>`;
        } else {
            messageArea.innerHTML = "<div class='success-msg'>✨ مراجعة رائعة! النص سليم من الأخطاء اللغوية المشهورة محلياً.</div>";
        }
    });
}



// --- Intelligent Indexing System (v2.0) ---

/**
 * يضيف النص المحدد إلى قائمة الفهرسة
 */
async function handleAddToIndex() {
    const messageArea = document.getElementById("message-area");
    const progressArea = document.getElementById("progress-area");

    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();
            range.load(["text", "isEmpty"]);
            await context.sync();

            if (range.isEmpty || !range.text || range.text.trim().length === 0) {
                showFeedback("⚠️ يرجى تحديد نص لإضافته للفهرس أولاً.", "warning");
                return;
            }

            // التأكد من عدم وجود وسم مسبق في نفس المنطقة (اختياري)
            const existingControls = range.getContentControls();
            existingControls.load("items/tag");
            await context.sync();

            if (existingControls.items.some(cc => cc.tag === "SMART_INDEX_ITEM")) {
                showFeedback("⚠️ هذا النص مضاف للفهرس مسبقاً.", "info");
                return;
            }

            // التأكد من عدم تجاوز العدد الأقصى (200)
            const allControls = context.document.contentControls.getByTag("SMART_INDEX_ITEM");
            allControls.load("items");
            await context.sync();
            if (allControls.items.length >= 200) {
                showFeedback("⚠️ تم الوصول للحد الأقصى (200 نص). يرجى حذف بعض العناصر.", "error");
                return;
            }

            const cleanText = range.text.trim().replace(/[\r\n]/g, " ");
            const uniqueID = "ID" + Math.random().toString(36).substring(2, 9).toUpperCase();

            // إنشاء Content Control
            const cc = range.insertContentControl();
            cc.tag = "SMART_INDEX_ITEM";
            cc.title = uniqueID; // نستخدم العنوان كـ ID فريد ثابت
            cc.appearance = Word.ContentControlAppearance.hidden;
            cc.color = "#2563eb";

            // إنشاء بوكمارك ثابت فوراً لضمان الدقة لاحقاً
            const bookmarkName = `IDX_${uniqueID}`;
            cc.getRange().insertBookmark(bookmarkName);

            await context.sync();

            showFeedback(`✅ تمت إضافة "${cleanText.substring(0, 20)}..." بنجاح.`, "success");

            // تحديث القائمة في الواجهة
            await refreshIndexList();
        });
    } catch (error) {
        console.error("Add Index Error:", error);
        showFeedback("❌ حدث خطأ أثناء الإضافة للفهرس.", "error");
    }
}

/**
 * تحديث قائمة العناصر المفهرسة في واجهة الإضافة
 */
async function refreshIndexList() {
    const listElement = document.getElementById("indexed-list");
    const container = document.getElementById("indexed-items-container");
    const countElement = document.getElementById("indexed-count");

    if (!listElement) return;

    try {
        await Word.run(async (context) => {
            const contentControls = context.document.contentControls.getByTag("SMART_INDEX_ITEM");
            contentControls.load(["items/text", "items/title"]);
            await context.sync();

            const items = contentControls.items;
            countElement.innerText = items.length;

            if (items.length > 0) {
                container.classList.remove("hidden");
                listElement.innerHTML = "";

                items.forEach((item, index) => {
                    const li = document.createElement("li");
                    li.className = "index-item-row";
                    li.style.background = index % 2 === 0 ? "#ffffff" : "#f1f5f9";

                    const textSpan = document.createElement("span");
                    textSpan.className = "index-item-text";
                    textSpan.innerText = item.text.length > 30 ? item.text.substring(0, 30) + "..." : item.text;
                    textSpan.title = item.text;
                    textSpan.onclick = () => jumpToIndexItem(index);

                    const actionBtns = document.createElement("div");
                    actionBtns.style.display = "flex";
                    actionBtns.style.gap = "8px";

                    const goBtn = document.createElement("button");
                    goBtn.className = "icon-btn";
                    goBtn.innerHTML = "📍";
                    goBtn.title = "انتقال للنص";
                    goBtn.style.padding = "2px 6px";
                    goBtn.onclick = (e) => { e.stopPropagation(); jumpToIndexItem(index); };

                    const delBtn = document.createElement("button");
                    delBtn.className = "icon-btn";
                    delBtn.innerHTML = "🗑️";
                    delBtn.title = "حذف من الفهرس";
                    delBtn.style.color = "#ef4444";
                    delBtn.style.padding = "2px 6px";
                    delBtn.onclick = (e) => { e.stopPropagation(); deleteIndexItem(index); };

                    actionBtns.appendChild(goBtn);
                    actionBtns.appendChild(delBtn);

                    li.appendChild(textSpan);
                    li.appendChild(actionBtns);
                    listElement.appendChild(li);
                });
            } else {
                container.classList.add("hidden");
                listElement.innerHTML = "";
            }
        });
    } catch (error) {
        console.error("Refresh List Error:", error);
    }
}

/**
 * الانتقال إلى موقع العنصر في المستند
 */
async function jumpToIndexItem(index) {
    await Word.run(async (context) => {
        const controls = context.document.contentControls.getByTag("SMART_INDEX_ITEM");
        controls.load("items");
        await context.sync();
        if (controls.items[index]) {
            controls.items[index].select();
            context.document.getActiveView().focus();
        }
        await context.sync();
    });
}

/**
 * حذف عنصر واحد من الفهرس
 */
async function deleteIndexItem(index) {
    try {
        await Word.run(async (context) => {
            const controls = context.document.contentControls.getByTag("SMART_INDEX_ITEM");
            controls.load("items");
            await context.sync();

            if (controls.items[index]) {
                const item = controls.items[index];
                item.load("title");
                await context.sync();

                // محاولة حذف البوكمارك المرتبط (اختياري، وورد سيحذفه غالباً مع النص)
                // لكن للأمان نبقيه أو نحذفه لاحقاً
                item.delete(false); // false يعني لا تحذف النص، فقط الـ Control
            }
            await context.sync();
            await refreshIndexList();
        });
    } catch (error) {
        console.error("Delete Item Error:", error);
    }
}

/**
 * مسح كل الفهرس
 */
async function handleClearIndex() {
    if (!confirm("هل أنت متأكد من مسح جميع علامات الفهرس؟ (لن يتم مسح نصوص المستند)")) return;

    try {
        await Word.run(async (context) => {
            const controls = context.document.contentControls.getByTag("SMART_INDEX_ITEM");
            controls.load("items");
            await context.sync();

            for (let i = controls.items.length - 1; i >= 0; i--) {
                controls.items[i].delete(false);
            }
            await context.sync();
            showFeedback("✅ تم تفريغ قائمة الفهرس.", "success");
            await refreshIndexList();
        });
    } catch (error) {
        console.error("Clear Index Error:", error);
    }
}

/**
 * توليد المعاينة النهائية للفهرس وحساب الصفحات
 */
/**
 * توليد المعاينة النهائية للفهرس وحساب الصفحات
 */
async function handleGenerateIndex() {
    showProgress("جاري فحص المستند واستخراج الصفحات...", 10);

    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            const controls = context.document.contentControls.getByTag("SMART_INDEX_ITEM");

            // المرحلة 1: تحميل المجموعة أولاً
            controls.load("items");
            await context.sync();

            if (controls.items.length === 0) {
                hideProgress();
                showFeedback("⚠️ لا توجد نصوص مضافة للفهرس حالياً.", "warning");
                return;
            }

            // المرحلة 2: تحميل الخصائص لكل عنصر بشكل صريح لضمان توفرها (حل مشكلة property not available)
            for (let i = 0; i < controls.items.length; i++) {
                controls.items[i].load(["text", "title"]);
            }
            await context.sync();

            const entries = [];
            const startRange = body.getRange("Start");

            showProgress("جاري معالجة العناصر وحساب المواقع...", 40);

            for (let i = 0; i < controls.items.length; i++) {
                const item = controls.items[i];

                // التأكد من وجود ID صالح للبوكمارك
                let itemID = item.title;
                if (!itemID || !/^[A-Z0-9]+$/.test(itemID.replace("ID", ""))) {
                    itemID = "ID" + Math.random().toString(36).substring(2, 9).toUpperCase();
                    item.title = itemID;
                }

                const itemRange = item.getRange();
                const bookmarkName = `IDX_${itemID}`;

                // إدراج البوكمارك للتتبع
                itemRange.insertBookmark(bookmarkName);

                const distRange = startRange.expandTo(itemRange);
                distRange.load("pageCount");

                entries.push({
                    text: item.text ? item.text.trim().replace(/[\r\n]/g, " ") : "نص غير معروف",
                    id: itemID,
                    distRange: distRange
                });
            }

            showProgress("جاري المزامنة مع بيانات التخطيط...", 80);
            await context.sync();

            // فرز العناصر حسب رقم الصفحة
            entries.sort((a, b) => (a.distRange.pageCount || 0) - (b.distRange.pageCount || 0));

            const finalData = entries.map(e => ({
                text: e.text,
                bookmark: `IDX_${e.id}`,
                page: e.distRange.pageCount || 1
            }));

            renderIndexPreview(finalData);
            hideProgress();
            window.lastGeneratedIndex = finalData;
        });
    } catch (error) {
        console.error("Generate Error Details:", error);
        hideProgress();
        let errorMsg = "❌ فشل توليد الفهرس: " + (error.message || "خطأ غير معروف");
        // خطأ مشهور يحدث عندما لا يكون المستند في وضع Print Layout
        if (error.code === "BaseNotVisible" || (error.message && error.message.includes("visible"))) {
            errorMsg = "❌ فشل التوليد: يرجى التأكد من أن المستند في وضع 'تخطيط الطباعة' (Print Layout).";
        }
        showFeedback(errorMsg, "error");
    }
}

/**
 * رسم بطاقة المعاينة في واجهة المستخدم
 */
function renderIndexPreview(data) {
    const messageArea = document.getElementById("message-area");

    let html = `
        <div class="index-preview-card">
            <div class="index-preview-header" style="background: linear-gradient(135deg, #1e293b 0%, #334155 100%); color: white; padding: 12px;">
                <h3 style="margin: 0; font-size: 1.1rem; color: white;">📋 معاينة الفهرس (${data.length} عنصر)</h3>
            </div>
            <div class="index-preview-table-container" style="max-height: 200px; overflow-y: auto;">
                <table style="width: 100%; border-collapse: collapse; font-size: 0.9rem;">
                    <thead style="position: sticky; top: 0; background: #f8fafc; box-shadow: 0 1px 0 #e2e8f0;">
                        <tr>
                            <th style="padding: 10px; text-align: right;">المادة</th>
                            <th style="padding: 10px; text-align: center; width: 60px;">الصفحة</th>
                        </tr>
                    </thead>
                    <tbody>
    `;

    data.forEach(entry => {
        html += `
            <tr style="border-bottom: 1px solid #f1f5f9;">
                <td style="padding: 8px 10px; color: #1e293b;">${entry.text}</td>
                <td style="padding: 8px 10px; text-align: center; font-weight: bold; color: #2563eb;">${entry.page}</td>
            </tr>
        `;
    });

    html += `
                    </tbody>
                </table>
            </div>
            <div style="padding: 12px; display: flex; gap: 8px; background: #f8fafc;">
                <button class="primary-button" style="flex: 2; margin: 0;" onclick="insertIndexInDoc()">إدراج الفهرس في المستند</button>
                <button class="secondary-btn" style="flex: 1; margin: 0; padding: 10px;" onclick="handleGenerateIndex()">🔄 تحديث</button>
            </div>
        </div>
    `;

    messageArea.innerHTML = html;
}

/**
 * إدراج الفهرس الفعلي في المستند باستخدام PAGEREF لضمان الحداثة
 */
async function insertIndexInDoc() {
    if (!window.lastGeneratedIndex || window.lastGeneratedIndex.length === 0) {
        showFeedback("⚠️ يرجى توليد المعاينة أولاً.", "warning");
        return;
    }

    try {
        await Word.run(async (context) => {
            const body = context.document.body;

            // محاولة إيجاد حاوية الفهرس السابقة لتحديثها
            let existingContainer = context.document.contentControls.getByTag("FINAL_INDEX_CONTAINER");
            existingContainer.load("items");
            await context.sync();

            let container;
            if (existingContainer.items.length > 0) {
                container = existingContainer.items[0];
                container.cannotDelete = false;
                container.clear();
            } else {
                // إضافة فاصل صفحات وعنوان
                body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
                const titlePara = body.insertParagraph("فهرس المواضيع", Word.InsertLocation.end);
                titlePara.font.name = "Cairo";
                titlePara.font.size = 18;
                titlePara.font.bold = true;
                titlePara.alignment = Word.Alignment.center;
                titlePara.spacingAfter = 20;

                container = body.insertParagraph("", Word.InsertLocation.end).insertContentControl();
                container.tag = "FINAL_INDEX_CONTAINER";
                container.title = "الفهرس التلقائي";
                container.appearance = Word.ContentControlAppearance.hidden;
            }

            // إنشاء جدول الفهرس
            const rowCount = window.lastGeneratedIndex.length;
            const table = container.insertTable(rowCount, 2, Word.InsertLocation.start);
            table.widthBase = "Percent";
            table.width = 100;

            // تحميل الصفوف للتمكن من تعديلها
            table.load("rows/items");
            await context.sync();

            // تنسيق الحدود بشكل صحيح (استخدام borders)
            table.borders.insideHorizontal.color = "#e2e8f0";
            table.borders.insideVertical.color = "#ffffff";
            table.borders.outside.color = "#ffffff";

            for (let i = 0; i < rowCount; i++) {
                const entry = window.lastGeneratedIndex[i];
                const row = table.rows.items[i];

                // تحميل الخلايا لكل صف
                row.load("cells/items");
                await context.sync();

                row.shadingColor = (i % 2 === 0) ? "#FFFFFF" : "#F8FAFC";

                // رقم الصفحة (على اليمين في الجداول العربية، لكننا سنمشي حسب التصميم)
                const cellPage = row.cells.items[0];
                cellPage.width = 15;
                const pPage = cellPage.paragraphs.getFirst();
                pPage.alignment = Word.Alignment.left;
                pPage.font.name = "Cairo";
                pPage.font.bold = true;
                pPage.insertField(`PAGEREF ${entry.bookmark} \\h`, Word.InsertLocation.start);

                // النص
                const cellText = row.cells.items[1];
                const pText = cellText.paragraphs.getFirst();
                pText.alignment = Word.Alignment.right;
                pText.font.name = "Cairo";
                pText.font.size = 11;

                const link = pText.insertHyperlink(entry.text, "#" + entry.bookmark, Word.HyperlinkType.internal);
                link.font.color = "#000000";
                link.font.underline = false;
            }

            await context.sync();
            showFeedback("✅ تم إدراج الفهرس بنجاح! ملاحظة: لتحديث الأرقام لاحقاً، يمكن الضغط على Ctrl+A ثم F9 في وورد.", "success");
        });
    } catch (error) {
        console.error("Insert Doc Error Details:", error);
        showFeedback("❌ فشل إدراج الفهرس: " + (error.message || "خطأ غير معروف"), "error");
    }
}

// --- Helpers ---

function showFeedback(msg, type = "info") {
    const area = document.getElementById("message-area");
    if (!area) return;

    let bgColor = "#f0f9ff";
    let textColor = "#0369a1";

    if (type === "success") { bgColor = "#ecfdf5"; textColor = "#059669"; }
    else if (type === "error") { bgColor = "#fef2f2"; textColor = "#dc2626"; }
    else if (type === "warning") { bgColor = "#fffbeb"; textColor = "#d97706"; }

    area.innerHTML = `<div class="success-msg" style="background: ${bgColor}; color: ${textColor}; border: 1px solid currentColor; border-radius: 8px; padding: 12px; margin-top: 10px; animation: slideDown 0.3s ease-out;">${msg}</div>`;
}

function showProgress(text, width) {
    const area = document.getElementById("progress-area");
    const fill = document.getElementById("progress-fill");
    const txt = document.getElementById("progress-text");

    if (area) area.classList.remove("hidden");
    if (fill) fill.style.width = width + "%";
    if (txt) txt.innerText = text;
}

function hideProgress() {
    const area = document.getElementById("progress-area");
    if (area) area.classList.add("hidden");
}

window.handleAddToIndex = handleAddToIndex;
window.handleGenerateIndex = handleGenerateIndex;
window.handleClearIndex = handleClearIndex;
window.insertIndexInDoc = insertIndexInDoc;
window.refreshIndexList = refreshIndexList;
window.deleteIndexItem = deleteIndexItem;
window.jumpToIndexItem = jumpToIndexItem;


window.handleLocalReview = handleLocalReview;


