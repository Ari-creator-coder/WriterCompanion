// 全局状态变量
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

// 🔒 密钥内存兜底机制（防止 Mac Word 禁用 localStorage）
let runtimeKeys = { deepseek: "", glm: "" };


Office.onReady(async (info) => {
    if (info.host === Office.HostType.Word) {
        try {
            if (Office.context.document.url) {
                currentGlobalDocName = Office.context.document.url.split('\\').pop().split('/').pop() || "未命名文档";
                currentGlobalDocName = decodeURIComponent(currentGlobalDocName);
            }
        } catch(e) {}
        
        // 初始化各个模块
        renderTable();
        initPrompts();
        setupInlineApiKeyUI(); // ⚡️ 激活全新的顶部栏 API Key 逻辑

        // 绑定倒计时控件
        const slider = document.getElementById("time-slider");
        const display = document.getElementById("time-display");
        const startBtn = document.getElementById("start-btn");
        const giveUpBtn = document.getElementById("give-up-btn");
        
        // 绑定音频控件
        const stopAudioBtn = document.getElementById("stop-audio-btn");
        const audioToggles = document.querySelectorAll(".audio-toggle");
        
        // 绑定 AI 聊天控件
        const aiInput = document.getElementById("ai-input");
        const aiSendBtn = document.getElementById("ai-send-btn");
        
        // 滑块事件
        if(slider) {
            slider.oninput = function() {
                if (!isRunning) {
                    display.innerText = `${this.value}:00`;
                    timeLeft = this.value * 60;
                }
            };
        }

        if(startBtn) startBtn.onclick = () => startTimer();
        if(giveUpBtn) giveUpBtn.onclick = () => stopTimer(true);

        audioToggles.forEach(toggle => {
            toggle.onclick = function() {
                const type = this.getAttribute("data-type");
                playAudio(type, this);
            };
        });
        if(stopAudioBtn) stopAudioBtn.onclick = () => stopAllAudio();

        if(aiSendBtn) aiSendBtn.onclick = () => handleAiChat();
        if(aiInput) {
            aiInput.addEventListener('keydown', function(e) {
                if ((e.key === 'Enter' || e.keyCode === 13) && !e.shiftKey) {
                    e.preventDefault(); 
                    handleAiChat();
                }
            });
        }
    }
});

// ==================== ⚡️ 全新：内联 API Key 逻辑（防沙盒拦截版） ====================

function getApiKey(provider) {
    if (runtimeKeys[provider]) return runtimeKeys[provider];
    try {
        let key = localStorage.getItem(provider === 'deepseek' ? 'writer_ds_key' : 'writer_glm_key');
        if (key) runtimeKeys[provider] = key;
        return key || "";
    } catch(e) { 
        return ""; 
    }
}

function setApiKey(provider, key) {
    runtimeKeys[provider] = key; // 内存兜底
    try {
        localStorage.setItem(provider === 'deepseek' ? 'writer_ds_key' : 'writer_glm_key', key);
    } catch(e) {
        console.warn("localStorage 被禁用，密钥已保存在当前会话内存中。");
    }
}

function setupInlineApiKeyUI() {
    const providerSelect = document.getElementById("api-provider");
    const keyInput = document.getElementById("inline-api-key");
    const container = document.getElementById("inline-api-container");
    const checkIcon = document.getElementById("api-save-check");

    if(!providerSelect || !keyInput) return; // 安全检查

    // 1. 初始化时读取
    updateInputDisplay();

    // 2. 切换时更新
    providerSelect.addEventListener("change", () => {
        updateInputDisplay();
    });

    // 3. 监听回车保存
    keyInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.keyCode === 13) {
            e.preventDefault();
            const currentProvider = providerSelect.value;
            const newKey = keyInput.value.trim();

            if(!newKey) {
                addChatMessage("⚠️ 密钥不能为空，请重新输入。", "ai");
                return;
            }
            
            // 保存逻辑
            setApiKey(currentProvider, newKey);

            // 绿色闪烁动画
            container.classList.remove("flash-success");
            void container.offsetWidth; // 触发重绘
            container.classList.add("flash-success");
            checkIcon.style.display = "inline";
            setTimeout(() => { checkIcon.style.display = "none"; }, 2000);
            
            keyInput.blur(); // 收起键盘
            
            // 💡 视觉确认：在聊天框明确通知用户保存成功！
            const modelName = currentProvider === 'deepseek' ? 'DeepSeek' : '智谱 GLM';
            addChatMessage(`✅ ${modelName} 密钥已保存至当前会话！请在下方输入文字开始创作。`, 'ai');
            
            const chatInput = document.getElementById("ai-input");
            if(chatInput) chatInput.focus();
        }
    });

    function updateInputDisplay() {
        const currentProvider = providerSelect.value;
        keyInput.value = getApiKey(currentProvider);
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
    
    try {
        const db = JSON.parse(localStorage.getItem('writerCompanionDB') || '[]');
        db.push({ docName, date: dateStr, time: timeStr, duration, count, timestamp: now.getTime() });
        localStorage.setItem('writerCompanionDB', JSON.stringify(db));
    } catch(e) {} // 忽略存储错误
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
    let db = [];
    try { db = JSON.parse(localStorage.getItem('writerCompanionDB') || '[]'); } catch(e){}
    
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
    try {
        let db = JSON.parse(localStorage.getItem('writerCompanionDB') || '[]');
        db = db.filter(r => r.docName !== docName);
        localStorage.setItem('writerCompanionDB', JSON.stringify(db));
    } catch(e) {}
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
    
    const safeName = docName.replace(/[^a-zA-Z0-9_\u4e00-\u9fa5]/g, '');
    const L_id = `L_${safeName}_${Date.now()}`;
    const B_id = `B_${safeName}_${Date.now()}`;

    const card = document.createElement('div');
    card.style.cssText = "aspect-ratio: 3/4; width: 100%; background: white; border: 1px solid #e1dfdd; border-radius: 6px; padding: 15px; box-sizing: border-box; display: flex; flex-direction: column; box-shadow: 0 2px 5px rgba(0,0,0,0.05); position: relative;";
    card.innerHTML = `<div style="font-size: 13px; color: #605e5c; font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; margin-bottom: 10px;" title="${docName}">📄 ${docName}</div><div style="font-size: 11px; color: #8a8886; margin-bottom: 4px;">今日趋势</div><div style="flex: 1; min-height: 0; margin-bottom: 10px;"><canvas id="${L_id}"></canvas></div><div style="font-size: 11px; color: #8a8886; margin-bottom: 4px;">每日字数</div><div style="flex: 1; min-height: 0; margin-bottom: 5px;"><canvas id="${B_id}"></canvas></div><div style="text-align: right; margin-top: auto;"><button id="del-btn-archive" onclick="deleteDocRecords('${docName}')" style="background: transparent; border: none; color: #a19f9d; font-size: 10px; cursor: pointer; text-decoration: underline; padding: 0; transition: color 0.3s;">删除</button></div>`;
    container.appendChild(card);
    
    setTimeout(() => {
        const todayData = records.filter(r => r.date === todayStr);
        const ctxLEl = document.getElementById(L_id);
        if(!ctxLEl) return; 
        const ctxL = ctxLEl.getContext('2d');
        archiveCharts.push(new Chart(ctxL, { type: 'line', data: { labels: todayData.map(r => r.time), datasets: [{ data: todayData.map(r => r.count), borderColor: '#605e5c', fill: true, tension: 0.3 }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } } }));
        
        let daily = {}; 
        records.forEach(r => { if(!daily[r.date] || r.count > daily[r.date]) daily[r.date] = r.count; });
        const ctxB = document.getElementById(B_id).getContext('2d');
        archiveCharts.push(new Chart(ctxB, { type: 'bar', data: { labels: Object.keys(daily), datasets: [{ data: Object.values(daily), backgroundColor: '#605e5c', maxBarThickness: 30 }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } } }));
    }, 0);
}

// ==================== 🚀 Prompt 核心逻辑 ====================

function initPrompts() {
    const defaultPrompts = [
        { id: 'p_default_1', title: '基础润色', content: '你是一个资深的文字编辑和排版专家。请帮我润色和优化发给你的文字，使其表达更流畅、专业。直接输出润色后的结果，不要加上废话。' }
    ];
    
    try {
        customPrompts = JSON.parse(localStorage.getItem('writerPrompts')) || defaultPrompts;
        activePromptId = localStorage.getItem('writerActivePrompt');
    } catch(e) {
        customPrompts = defaultPrompts;
    }
    
    if (customPrompts.length === 0) customPrompts = defaultPrompts; 
    if (!activePromptId || !customPrompts.find(p => p.id === activePromptId)) activePromptId = customPrompts[0].id;
    
    updatePromptCapsule();
}

function updatePromptCapsule() {
    const activeP = customPrompts.find(p => p.id === activePromptId) || customPrompts[0];
    const cap = document.getElementById('active-prompt-capsule');
    if(cap) cap.innerText = activeP.title;
}

function savePromptsData() {
    try {
        localStorage.setItem('writerPrompts', JSON.stringify(customPrompts));
        localStorage.setItem('writerActivePrompt', activePromptId);
    } catch(e) {}
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

// ==================== 🔒 进化：AI 接口与流式输出 ====================

async function handleAiChat() {
    const inputEl = document.getElementById("ai-input");
    const text = inputEl.value.trim();
    if (!text) return;

    const providerSelect = document.getElementById('api-provider');
    if(!providerSelect) return;
    const provider = providerSelect.value;
    
    const thinkToggle = document.getElementById('deep-think-toggle');
    const isDeepThink = thinkToggle ? thinkToggle.checked : false;

    // 💡 从强化的 getApiKey 获取密钥（即便跨域隔离也能读到内存）
    const apiKey = getApiKey(provider);
    
    if (!apiKey) {
        addChatMessage(`⚠️ 未检测到 ${provider === 'deepseek' ? 'DeepSeek' : '智谱 GLM'} 密钥。请在顶部文本框输入您的 sk- 开头的密钥，然后【按下回车键】保存。`, 'ai');
        return;
    }

    addChatMessage(text, 'user');
    inputEl.value = ''; 
    inputEl.style.height = '20px';
    
    const loadingText = "连接中...";
    addChatMessage(loadingText, 'ai', true);
    
    const activeP = customPrompts.find(p => p.id === activePromptId) || customPrompts[0];
    const systemPromptContent = activeP.content || "你是一个资深的文字编辑...";

    let apiUrl = '';
    
    let requestBody = {
        messages: [
            { role: "system", content: systemPromptContent }, 
            { role: "user", content: text }
        ],
        stream: true 
    };

    if (provider === 'deepseek') {
        apiUrl = 'https://api.deepseek.com/chat/completions';
        requestBody.model = isDeepThink ? 'deepseek-reasoner' : 'deepseek-chat';
    } else {
        apiUrl = 'https://open.bigmodel.cn/api/paas/v4/chat/completions';
        requestBody.model = 'glm-4'; 
    }

    try {
        const response = await fetch(apiUrl, {
            method: 'POST',
            headers: { 
                'Content-Type': 'application/json', 
                'Authorization': `Bearer ${apiKey}` 
            },
            body: JSON.stringify(requestBody)
        });

        if (!response.ok) {
            // 如果接口返回错误，把具体的错误码打印出来方便排查
            const errBody = await response.text();
            throw new Error(`HTTP ${response.status} - 接口拒绝连接。请检查密钥是否正确、是否欠费，或者网络是否被拦截。\n详情: ${errBody}`);
        }
        
        removeLoadingMessage();
        
        const streamMsgDiv = addChatMessage("", 'ai');
        let fullContent = "";

        const reader = response.body.getReader();
        const decoder = new TextDecoder("utf-8");
        let buffer = "";

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;

            buffer += decoder.decode(value, { stream: true });
            const lines = buffer.split('\n');
            
            buffer = lines.pop(); 

            for (const line of lines) {
                if (line.trim() === '') continue;
                if (line.startsWith('data: ')) {
                    const dataStr = line.slice(6).trim();
                    if (dataStr === '[DONE]') continue;
                    
                    try {
                        const data = JSON.parse(dataStr);
                        const delta = data.choices[0].delta;
                        
                        if (delta && delta.content) {
                            fullContent += delta.content;
                            streamMsgDiv.innerText = fullContent;
                            
                            const chatHistory = document.getElementById("chat-history");
                            chatHistory.scrollTop = chatHistory.scrollHeight;
                        }
                    } catch (e) {
                        // 忽略单个分片解析错误
                    }
                }
            }
        }

    } catch (e) { 
        removeLoadingMessage(); 
        addChatMessage(`⚠️ 发生错误: \n${e.message}`, 'ai'); 
    }
}

function addChatMessage(text, sender, isLoading = false) {
    const chatHistory = document.getElementById("chat-history");
    if(!chatHistory) return null;
    
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
    
    return msgDiv;
}

function removeLoadingMessage() {
    const el = document.getElementById('ai-loading-msg');
    if (el) el.remove();
}