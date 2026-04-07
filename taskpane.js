let timer;
let timeLeft;
let isRunning = false;
let wordHistory = []; 
let currentAudio = null; 
let archiveCharts = []; 

let currentGlobalDocName = "未命名文档"; 
let selectedArchiveDoc = null; 
let showAllDocs = false; 

// 🚀 Prompt 专属全局变量 🚀
let customPrompts = [];
let activePromptId = null;
let showAllPrompts = false;

Office.onReady(async (info) => {
    if (info.host === Office.HostType.Word) {
        try {
            if (Office.context.document.url) {
                currentGlobalDocName = Office.context.document.url.split('\\').pop().split('/').pop() || "未命名文档";
                currentGlobalDocName = decodeURIComponent(currentGlobalDocName);
            }
        } catch(e) {}
        
        // 初始化存储数据
        renderTable();
        initPrompts();
        loadStoredApiKeys(); // 🔒 安全强化：初始化时加载已存的 Key

        const slider = document.getElementById("time-slider");
        const display = document.getElementById("time-display");
        const startBtn = document.getElementById("start-btn");
        const giveUpBtn = document.getElementById("give-up-btn");
        const stopAudioBtn = document.getElementById("stop-audio-btn");
        const aiInput = document.getElementById("ai-input");
        const aiSendBtn = document.getElementById("ai-send-btn");
        
        // 🔒 安全强化：绑定保存 Key 的按钮
        const saveApiBtn = document.getElementById("save-api-keys");
        if(saveApiBtn) saveApiBtn.onclick = saveApiKeysToStorage;

        slider.oninput = function() {
            if (!isRunning) {
                display.innerText = `${this.value}:00`;
                timeLeft = this.value * 60;
            }
        };

        startBtn.onclick = () => startTimer();
        giveUpBtn.onclick = () => stopTimer(true);

        const audioToggles = document.querySelectorAll(".audio-toggle");
        audioToggles.forEach(toggle => {
            toggle.onclick = function() {
                const type = this.getAttribute("data-type");
                playAudio(type, this);
            };
        });
        stopAudioBtn.onclick = () => stopAllAudio();

        aiSendBtn.onclick = () => handleAiChat();
        aiInput.addEventListener('keydown', function(e) {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault(); 
                handleAiChat();
            }
        });
    }
});

// ==================== 🔒 安全强化：Key 管理逻辑 ====================

function saveApiKeysToStorage() {
    const dsKey = document.getElementById("ds-key-input").value.trim();
    const glmKey = document.getElementById("glm-key-input").value.trim();
    
    if (dsKey) localStorage.setItem('writer_ds_key', dsKey);
    if (glmKey) localStorage.setItem('writer_glm_key', glmKey);
    
    alert("API 设置已安全保存至本地设备。");
}

function loadStoredApiKeys() {
    const dsKey = localStorage.getItem('writer_ds_key');
    const glmKey = localStorage.getItem('writer_glm_key');
    
    if (dsKey && document.getElementById("ds-key-input")) {
        document.getElementById("ds-key-input").value = dsKey;
    }
    if (glmKey && document.getElementById("glm-key-input")) {
        document.getElementById("glm-key-input").value = glmKey;
    }
}

// ==================== 页面导航 ====================
function switchTab(tabName) {
    ['write', 'prompt', 'archive'].forEach(name => {
        const tabEl = document.getElementById(`tab-${name}`);
        const bodyEl = document.getElementById(`${name === 'write' ? 'app' : name}-body`);
        if(tabEl) tabEl.classList.remove('active');
        if(bodyEl) bodyEl.style.display = 'none';
    });

    const activeTab = document.getElementById(`tab-${tabName}`);
    const activeBody = document.getElementById(`${tabName === 'write' ? 'app' : tabName}-body`);
    if(activeTab) activeTab.classList.add('active');
    if(activeBody) activeBody.style.display = 'flex';

    if (tabName === 'archive') renderArchive();
    if (tabName === 'prompt') renderPromptList();
}

// ==================== 计时与历史 ====================
function startTimer() {
    isRunning = true;
    const minutes = document.getElementById("time-slider").value;
    timeLeft = timeLeft || minutes * 60;
    document.getElementById("controls-wrapper").style.display = "none";
    document.getElementById("give-up-btn").style.display = "inline-block";
    const displayElement = document.getElementById("time-display");
    displayElement.style.transform = "scale(1)"; 
    displayElement.style.color = "#c8c6c4"; 
    timer = setInterval(() => {
        timeLeft--;
        updateDisplay();
        if (timeLeft <= 0) completeSession();
    }, 1000);
}

function stopTimer(isGiveUp) {
    clearInterval(timer);
    isRunning = false;
    document.getElementById("controls-wrapper").style.display = "flex";
    document.getElementById("give-up-btn").style.display = "none";
    const displayElement = document.getElementById("time-display");
    displayElement.style.transform = "scale(0.6)"; 
    displayElement.style.color = "#605e5c"; 
    if (isGiveUp) {
        const minutes = document.getElementById("time-slider").value;
        timeLeft = minutes * 60;
        updateDisplay();
    }
}

function updateDisplay() {
    const mins = Math.floor(timeLeft / 60);
    const secs = timeLeft % 60;
    document.getElementById("time-display").innerText = `${mins}:${secs.toString().padStart(2, '0')}`;
}

async function completeSession() {
    stopTimer(false);
    await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text"); 
        await context.sync();
        const text = body.text;
        const chineseChars = text.match(/[\u4e00-\u9fa5]/g) || [];
        const englishWords = text.replace(/[\u4e00-\u9fa5]/g, ' ').trim().split(/\s+/).filter(word => word.length > 0);
        const totalWordCount = chineseChars.length + englishWords.length;
        const duration = document.getElementById("time-slider").value;
        saveRecord(totalWordCount, duration, currentGlobalDocName);
    });
}

function saveRecord(count, duration, docName) {
    const now = new Date();
    const timeStr = now.toLocaleTimeString('zh-CN', { hour12: false, hour: '2-digit', minute: '2-digit' });
    const dateStr = now.toLocaleDateString('zh-CN');
    wordHistory.unshift({ time: timeStr, count: count, duration: duration });
    if (wordHistory.length > 5) wordHistory.pop();
    renderTable();
    const db = JSON.parse(localStorage.getItem('writerCompanionDB') || '[]');
    db.push({ docName, date: dateStr, time: timeStr, duration, count, timestamp: now.getTime() });
    localStorage.setItem('writerCompanionDB', JSON.stringify(db));
}

function renderTable() {
    const listContainer = document.getElementById('history-list');
    if(!listContainer) return;
    listContainer.innerHTML = ''; 
    if (wordHistory.length === 0) {
        listContainer.innerHTML = '<div style="font-size: 12px; color: #999; text-align: center; margin-top: 10px;">暂无记录</div>';
        return;
    }
    wordHistory.forEach(record => {
        const item = document.createElement('div');
        item.style.cssText = "padding: 8px 12px; font-size: 12px; display: flex; justify-content: space-between;";
        item.innerHTML = `<span style="color: #8a8886;">🕒 ${record.duration}m at ${record.time}</span><span style="color: #605e5c; font-weight: 600;">${record.count} 字</span>`;
        listContainer.appendChild(item);
    });
}

// ==================== 档案逻辑 ====================
function renderArchive() {
    const db = JSON.parse(localStorage.getItem('writerCompanionDB') || '[]');
    let docLastUsed = {};
    db.forEach(r => { if (!docLastUsed[r.docName] || r.timestamp > docLastUsed[r.docName]) docLastUsed[r.docName] = r.timestamp || 0; });
    let sortedDocs = Object.keys(docLastUsed).sort((a, b) => docLastUsed[b] - docLastUsed[a]);
    if (sortedDocs.includes(currentGlobalDocName)) sortedDocs = [currentGlobalDocName, ...sortedDocs.filter(d => d !== currentGlobalDocName)];
    
    const listContainer = document.getElementById('archive-doc-list');
    if(!listContainer) return;
    listContainer.innerHTML = '';

    if (sortedDocs.length === 0) {
        document.getElementById('archive-card-container').innerHTML = '<div style="color: #999; margin-top: 50px; font-size: 13px;">暂无档案</div>';
        return;
    }
    if (!selectedArchiveDoc || !sortedDocs.includes(selectedArchiveDoc)) selectedArchiveDoc = sortedDocs[0];
    
    let displayDocs = showAllDocs ? sortedDocs : sortedDocs.slice(0, 5);
    displayDocs.forEach(docName => {
        const item = document.createElement('div');
        item.className = `doc-list-item ${docName === selectedArchiveDoc ? 'active' : ''}`;
        item.innerText = docName;
        item.onclick = () => { selectedArchiveDoc = docName; renderArchive(); };
        listContainer.appendChild(item);
    });
    if (!showAllDocs && sortedDocs.length > 5) {
        const moreBtn = document.createElement('div');
        moreBtn.className = 'doc-list-item';
        moreBtn.innerText = '...'; 
        moreBtn.onclick = () => { showAllDocs = true; renderArchive(); };
        listContainer.appendChild(moreBtn);
    }
    renderArchiveCard(selectedArchiveDoc, db);
}

let deleteArchiveConfirmTimer = null;
function deleteDocRecords(docName) {
    const btn = document.getElementById('del-btn-archive');
    if (btn && btn.innerText === '删除') {
        btn.innerText = '确认删除？';
        btn.style.color = '#a4262c'; 
        clearTimeout(deleteArchiveConfirmTimer);
        deleteArchiveConfirmTimer = setTimeout(() => {
            if (btn) { btn.innerText = '删除'; btn.style.color = '#a19f9d'; }
        }, 3000);
        return;
    }
    let db = JSON.parse(localStorage.getItem('writerCompanionDB') || '[]');
    db = db.filter(r => r.docName !== docName);
    localStorage.setItem('writerCompanionDB', JSON.stringify(db));
    if (selectedArchiveDoc === docName) selectedArchiveDoc = null;
    renderArchive();
}

function renderArchiveCard(docName, db) {
    const container = document.getElementById('archive-card-container');
    if(!container) return;
    container.innerHTML = '';
    archiveCharts.forEach(c => c.destroy());
    archiveCharts = [];
    const records = db.filter(r => r.docName === docName);
    if (records.length === 0) return;
    const todayStr = new Date().toLocaleDateString('zh-CN');
    const card = document.createElement('div');
    card.style.cssText = "aspect-ratio: 3/4; width: 100%; background: white; border: 1px solid #e1dfdd; border-radius: 6px; padding: 15px; box-sizing: border-box; display: flex; flex-direction: column; box-shadow: 0 2px 5px rgba(0,0,0,0.05); position: relative;";
    card.innerHTML = `<div style="font-size: 13px; color: #605e5c; font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; margin-bottom: 10px;" title="${docName}">📄 ${docName}</div><div style="font-size: 11px; color: #8a8886; margin-bottom: 4px;">今日趋势</div><div style="flex: 1; min-height: 0; margin-bottom: 10px;"><canvas id="L_archive"></canvas></div><div style="font-size: 11px; color: #8a8886; margin-bottom: 4px;">每日字数</div><div style="flex: 1; min-height: 0; margin-bottom: 5px;"><canvas id="B_archive"></canvas></div><div style="text-align: right; margin-top: auto;"><button id="del-btn-archive" onclick="deleteDocRecords('${docName}')" style="background: transparent; border: none; color: #a19f9d; font-size: 10px; cursor: pointer; text-decoration: underline; padding: 0; transition: color 0.3s;">删除</button></div>`;
    container.appendChild(card);
    setTimeout(() => {
        const todayData = records.filter(r => r.date === todayStr);
        const ctxL = document.getElementById(`L_archive`).getContext('2d');
        archiveCharts.push(new Chart(ctxL, { type: 'line', data: { labels: todayData.map(r => r.time), datasets: [{ data: todayData.map(r => r.count), borderColor: '#605e5c', fill: true, tension: 0.3 }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } } }));
        let daily = {}; 
        records.forEach(r => { if(!daily[r.date] || r.count > daily[r.date]) daily[r.date] = r.count; });
        const ctxB = document.getElementById(`B_archive`).getContext('2d');
        archiveCharts.push(new Chart(ctxB, { type: 'bar', data: { labels: Object.keys(daily), datasets: [{ data: Object.values(daily), backgroundColor: '#605e5c', maxBarThickness: 30 }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } } }));
    }, 0);
}

// ==================== 🚀 Prompt 核心逻辑 ====================

function initPrompts() {
    const defaultPrompts = [
        { id: 'p_default_1', title: '基础润色', content: '你是一个资深的文字编辑和排版专家。请帮我润色和优化发给你的文字，使其表达更流畅、专业。直接输出润色后的结果，不要加上废话。' }
    ];
    customPrompts = JSON.parse(localStorage.getItem('writerPrompts')) || defaultPrompts;
    if (customPrompts.length === 0) customPrompts = defaultPrompts; 
    
    activePromptId = localStorage.getItem('writerActivePrompt') || customPrompts[0].id;
    if (!customPrompts.find(p => p.id === activePromptId)) activePromptId = customPrompts[0].id;
    
    updatePromptCapsule();
}

function updatePromptCapsule() {
    const activeP = customPrompts.find(p => p.id === activePromptId) || customPrompts[0];
    const cap = document.getElementById('active-prompt-capsule');
    if(cap) cap.innerText = activeP.title;
}

function savePromptsData() {
    localStorage.setItem('writerPrompts', JSON.stringify(customPrompts));
    localStorage.setItem('writerActivePrompt', activePromptId);
    updatePromptCapsule();
}

function renderPromptList() {
    const listContainer = document.getElementById('prompt-list');
    if(!listContainer) return;
    listContainer.innerHTML = '';
    
    let displayPrompts = showAllPrompts ? customPrompts : customPrompts.slice(0, 5);
    
    displayPrompts.forEach(p => {
        const item = document.createElement('div');
        item.className = `doc-list-item ${p.id === activePromptId ? 'active' : ''}`;
        item.innerText = p.title;
        item.onclick = () => {
            activePromptId = p.id;
            savePromptsData();
            renderPromptList(); 
        };
        listContainer.appendChild(item);
    });

    if (!showAllPrompts && customPrompts.length > 5) {
        const moreBtn = document.createElement('div');
        moreBtn.className = 'doc-list-item';
        moreBtn.innerText = '...'; 
        moreBtn.onclick = () => { showAllPrompts = true; renderPromptList(); };
        listContainer.appendChild(moreBtn);
    }
    renderPromptCard(activePromptId);
}

function renderPromptCard(id) {
    const container = document.getElementById('prompt-card-container');
    if(!container) return;
    container.innerHTML = '';
    const p = customPrompts.find(p => p.id === id);
    if (!p) return;

    const card = document.createElement('div');
    card.style.cssText = "aspect-ratio: 3/4; width: 100%; background: white; border: 1px solid #e1dfdd; border-radius: 6px; padding: 15px; box-sizing: border-box; display: flex; flex-direction: column; box-shadow: 0 2px 5px rgba(0,0,0,0.05); position: relative;";

    card.innerHTML = `
        <div style="margin-bottom: 10px;">
            <input type="text" id="edit-prompt-title" value="${p.title}" placeholder="提示词标题" onblur="saveCurrentPromptEdit('${id}')" style="width: 100%; border: none; background: transparent; font-size: 13px; color: #605e5c; font-weight: 600; outline: none; padding: 0; font-family: inherit;">
        </div>
        <div style="flex: 1; min-height: 0; display: flex; flex-direction: column;">
            <textarea id="edit-prompt-content" placeholder="输入你想给 AI 的系统级设定..." onblur="saveCurrentPromptEdit('${id}')" style="flex: 1; border: none; background: transparent; color: #8a8886; font-size: 12px; resize: none; outline: none; line-height: 1.6; font-family: inherit; padding: 0;">${p.content}</textarea>
        </div>
        <div style="text-align: right; margin-top: 10px;">
            <button id="del-btn-prompt" onclick="deletePrompt('${id}')" style="background: transparent; border: none; color: #a19f9d; font-size: 10px; cursor: pointer; text-decoration: underline; padding: 0; transition: color 0.3s;">删除</button>
        </div>
    `;
    container.appendChild(card);
}

function saveCurrentPromptEdit(id) {
    const p = customPrompts.find(x => x.id === id);
    if (p) {
        const newTitle = document.getElementById('edit-prompt-title').value.trim() || '未命名提示词';
        const newContent = document.getElementById('edit-prompt-content').value.trim();
        p.title = newTitle;
        p.content = newContent;
        savePromptsData();
        const activeItem = document.querySelector('#prompt-list .active');
        if(activeItem) activeItem.innerText = newTitle;
    }
}

function addNewPrompt() {
    const newId = 'p_' + Date.now();
    customPrompts.unshift({ id: newId, title: '新提示词', content: '你是一个专业的...' });
    activePromptId = newId;
    savePromptsData();
    renderPromptList();
}

let deletePromptConfirmTimer = null;
function deletePrompt(id) {
    const btn = document.getElementById('del-btn-prompt');
    if (customPrompts.length <= 1) {
        if (btn) { btn.innerText = '须至少保留一个'; setTimeout(() => { if (btn) btn.innerText = '删除'; }, 2000); }
        return;
    }
    if (btn && btn.innerText === '删除') {
        btn.innerText = '确认删除？';
        btn.style.color = '#a4262c'; 
        clearTimeout(deletePromptConfirmTimer);
        deletePromptConfirmTimer = setTimeout(() => { if (btn) { btn.innerText = '删除'; btn.style.color = '#a19f9d'; } }, 3000);
        return;
    }
    customPrompts = customPrompts.filter(p => p.id !== id);
    activePromptId = customPrompts[0].id;
    savePromptsData();
    renderPromptList();
}

// ==================== 音频 ====================
function stopAllAudio() {
    if (currentAudio) { currentAudio.audio.pause(); currentAudio = null; }
    document.querySelectorAll(".audio-toggle").forEach(el => el.style.opacity = "1");
    document.getElementById("stop-audio-btn").style.display = "none";
}
function playAudio(type, el) {
    stopAllAudio();
    const urls = { 'rainy': 'assets/rainy.mp3', 'sun': 'assets/sun.mp3', 'coffee': 'assets/coffee.mp3', 'river': 'assets/river.mp3', 'ocean': 'assets/ocean.mp3', 'fire': 'assets/fire.mp3' };
    const audio = new Audio(urls[type]);
    audio.loop = true; audio.play();
    currentAudio = { type, audio };
    document.querySelectorAll(".audio-toggle").forEach(e => e.style.opacity = "0.3");
    el.style.opacity = "1";
    document.getElementById("stop-audio-btn").style.display = "inline-flex";
}

// ==================== 🔒 安全强化：AI 接口逻辑 ====================

async function handleAiChat() {
    const inputEl = document.getElementById("ai-input");
    const text = inputEl.value.trim();
    if (!text) return;

    const provider = document.getElementById('api-provider').value;
    const isDeepThink = document.getElementById('deep-think-toggle').checked;

    // 🔒 从本地存储获取 Key，不再硬编码
    const storedDsKey = localStorage.getItem('writer_ds_key');
    const storedGlmKey = localStorage.getItem('writer_glm_key');
    
    // 检查 Key 是否存在
    if (provider === 'deepseek' && !storedDsKey) {
        addChatMessage("⚠️ 未检测到 DeepSeek Key，请在设置中输入并保存。", 'ai');
        return;
    }
    if (provider === 'glm' && !storedGlmKey) {
        addChatMessage("⚠️ 未检测到 GLM Key，请在设置中输入并保存。", 'ai');
        return;
    }

    addChatMessage(text, 'user');
    inputEl.value = ''; 
    inputEl.style.height = '20px';
    
    const loadingText = isDeepThink ? "思考中..." : "正在排版与润色...";
    addChatMessage(loadingText, 'ai', true);
    
    const activeP = customPrompts.find(p => p.id === activePromptId) || customPrompts[0];
    const systemPromptContent = activeP.content || "你是一个资深的文字编辑...";

    let apiUrl = '';
    let apiKey = (provider === 'deepseek') ? storedDsKey : storedGlmKey;
    let requestBody = {
        messages: [
            { role: "system", content: systemPromptContent }, 
            { role: "user", content: text }
        ]
    };

    if (provider === 'deepseek') {
        apiUrl = 'https://api.deepseek.com/chat/completions';
        requestBody.model = isDeepThink ? 'deepseek-reasoner' : 'deepseek-chat';
    } else {
        apiUrl = 'https://open.bigmodel.cn/api/paas/v4/chat/completions';
        requestBody.model = 'glm-4'; // 注意：通常是 glm-4，根据实际文档调整
        if (isDeepThink) requestBody.thinking = { type: "enabled" };
    }

    try {
        const response = await fetch(apiUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` },
            body: JSON.stringify(requestBody)
        });

        if (!response.ok) throw new Error(`连接失败 (HTTP ${response.status})`);

        const data = await response.json();
        removeLoadingMessage();
        if (data.choices && data.choices[0]) {
            addChatMessage(data.choices[0].message.content, 'ai');
        } else {
            throw new Error("接口未返回有效内容");
        }

    } catch (e) { 
        removeLoadingMessage(); 
        addChatMessage(`⚠️ 报错: ${e.message}\n请检查网络或确认 API Key 是否正确。`, 'ai'); 
    }
}

function addChatMessage(text, sender, isLoading = false) {
    const chatHistory = document.getElementById("chat-history");
    if(!chatHistory) return;
    const msgDiv = document.createElement('div');
    msgDiv.style.cssText = "padding: 8px 12px; border-radius: 12px; max-width: 85%; font-size: 13px; margin-bottom: 5px; word-break: break-word; white-space: pre-wrap;";
    if (sender === 'user') { 
        msgDiv.style.alignSelf = 'flex-end'; 
        msgDiv.style.background = '#605e5c'; 
        msgDiv.style.color = 'white'; 
        msgDiv.style.borderTopRightRadius = '2px';
    } else { 
        msgDiv.style.alignSelf = 'flex-start'; 
        msgDiv.style.background = '#edebe9'; 
        msgDiv.style.color = '#605e5c';
        msgDiv.style.borderTopLeftRadius = '2px';
        if (isLoading) msgDiv.id = 'ai-loading-msg'; 
    }
    msgDiv.innerText = text;
    chatHistory.appendChild(msgDiv);
    chatHistory.scrollTop = chatHistory.scrollHeight;
}

function removeLoadingMessage() {
    const el = document.getElementById('ai-loading-msg');
    if (el) el.remove();
}