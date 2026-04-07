// ==================== 全局状态变量 ====================
let timer;
let timeLeft;
let isRunning = false;
let wordHistory = []; 
let currentAudio = null; 
let archiveCharts = []; 

let currentGlobalDocName = "未命名文档"; 
let selectedArchiveDoc = null; 
let showAllDocs = false; 

// 🚀 Prompt 专属全局变量
let customPrompts = [];
let activePromptId = null;
let showAllPrompts = false;

// 🔒 密钥内存兜底
let runtimeKeys = { deepseek: "", glm: "" };

Office.onReady(async (info) => {
    if (info.host === Office.HostType.Word) {
        try {
            if (Office.context.document.url) {
                currentGlobalDocName = Office.context.document.url.split('\\').pop().split('/').pop() || "未命名文档";
                currentGlobalDocName = decodeURIComponent(currentGlobalDocName);
            }
        } catch(e) {}
        
        // 1. 初始化模块
        renderTable();
        initPrompts();
        setupInlineApiKeyUI(); 

        // 2. 绑定基础控件
        const slider = document.getElementById("time-slider");
        const display = document.getElementById("time-display");
        const startBtn = document.getElementById("start-btn");
        const giveUpBtn = document.getElementById("give-up-btn");
        const stopAudioBtn = document.getElementById("stop-audio-btn");
        const audioToggles = document.querySelectorAll(".audio-toggle");
        const aiInput = document.getElementById("ai-input");
        const aiSendBtn = document.getElementById("ai-send-btn");
        
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

        // 3. 绑定 AI 聊天逻辑 (双击回车)
        if(aiSendBtn) aiSendBtn.onclick = () => handleAiChat();
        if(aiInput) {
            let lastEnterTime = 0; 
            aiInput.addEventListener('keydown', function(e) {
                if (e.isComposing || e.keyCode === 229) return;
                if (e.key === 'Enter' && !e.shiftKey) {
                    const now = Date.now();
                    if (now - lastEnterTime < 400) {
                        e.preventDefault(); 
                        handleAiChat();
                        lastEnterTime = 0; 
                    } else {
                        lastEnterTime = now;
                    }
                }
            });
        }
    }
});

// ==================== ⚡️ API Key 核心逻辑 ====================

function getApiKey(provider) {
    if (runtimeKeys[provider]) return runtimeKeys[provider];
    try {
        let key = localStorage.getItem(provider === 'deepseek' ? 'writer_ds_key' : 'writer_glm_key');
        if (key) runtimeKeys[provider] = key;
        return key || "";
    } catch(e) { return ""; }
}

function setApiKey(provider, key) {
    runtimeKeys[provider] = key; 
    try {
        localStorage.setItem(provider === 'deepseek' ? 'writer_ds_key' : 'writer_glm_key', key);
    } catch(e) {}
}

function setupInlineApiKeyUI() {
    const chatHeader = document.getElementById("chat-header"); 
    const providerSelect = document.getElementById("api-provider");
    const keyInput = document.getElementById("inline-api-key");
    const container = document.getElementById("inline-api-container"); 
    const checkIcon = document.getElementById("api-save-check");

    if(!chatHeader || !providerSelect || !keyInput) return;

    function checkVisibility() {
        const key = getApiKey(providerSelect.value);
        container.style.display = (key && key.length > 0) ? "none" : "flex";
        keyInput.value = key;
    }

    checkVisibility();

    chatHeader.addEventListener("click", (e) => {
        if (e.target === providerSelect) return;
        container.style.display = (container.style.display === "none") ? "flex" : "none";
        if(container.style.display === "flex") keyInput.focus();
    });

    providerSelect.addEventListener("change", () => {
        checkVisibility();
    });

    keyInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.keyCode === 13) {
            e.preventDefault();
            const val = keyInput.value.trim();
            if(!val) return;
            setApiKey(providerSelect.value, val);
            
            container.classList.add("flash-success");
            checkIcon.style.display = "inline";
            setTimeout(() => {
                container.style.display = "none";
                container.classList.remove("flash-success");
                checkIcon.style.display = "none";
            }, 1000);
            keyInput.blur();
        }
    });
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
    const expectedEndTime = Date.now() + (timeLeft * 1000);

    document.getElementById("controls-wrapper").style.display = "none";
    document.getElementById("give-up-btn").style.display = "inline-block";
    const displayElement = document.getElementById("time-display");
    displayElement.style.transform = "scale(1)"; 
    displayElement.style.color = "#c8c6c4"; 
    
    if (timer) clearInterval(timer);
    timer = setInterval(() => {
        const now = Date.now();
        timeLeft = Math.max(0, Math.ceil((expectedEndTime - now) / 1000));
        updateDisplay();
        if (timeLeft <= 0) {
            completeSession();
        }
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
        timeLeft = document.getElementById("time-slider").value * 60;
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
        body.load("text"); await context.sync();
        const text = body.text;
        const count = (text.match(/[\u4e00-\u9fa5]/g) || []).length + text.replace(/[\u4e00-\u9fa5]/g, ' ').trim().split(/\s+/).filter(w => w.length > 0).length;
        saveRecord(count, document.getElementById("time-slider").value, currentGlobalDocName);
    });
}

function saveRecord(count, duration, docName) {
    const now = new Date();
    const timeStr = now.toLocaleTimeString('zh-CN', { hour12: false, hour: '2-digit', minute: '2-digit' });
    wordHistory.unshift({ time: timeStr, count: count, duration: duration });
    if (wordHistory.length > 5) wordHistory.pop();
    renderTable();
    try {
        let db = JSON.parse(localStorage.getItem('writerCompanionDB') || '[]');
        db.push({ docName, date: now.toLocaleDateString('zh-CN'), time: timeStr, duration, count, timestamp: now.getTime() });
        localStorage.setItem('writerCompanionDB', JSON.stringify(db));
    } catch(e) {}
}

function renderTable() {
    const list = document.getElementById('history-list');
    if(!list) return;
    list.innerHTML = wordHistory.length === 0 ? '<div style="font-size: 12px; color: #999; text-align: center; margin-top: 10px;">暂无记录</div>' : '';
    wordHistory.forEach(r => {
        const item = document.createElement('div');
        item.style.cssText = "padding: 8px 12px; font-size: 12px; display: flex; justify-content: space-between;";
        item.innerHTML = `<span>🕒 ${r.duration}m at ${r.time}</span><span style="font-weight: 600;">${r.count} 字</span>`;
        list.appendChild(item);
    });
}

// ==================== 档案管理 ====================
function renderArchive() {
    let db = []; try { db = JSON.parse(localStorage.getItem('writerCompanionDB') || '[]'); } catch(e){}
    let docs = [...new Set(db.map(r => r.docName))];
    const list = document.getElementById('archive-doc-list');
    if(!list) return;
    list.innerHTML = '';
    if (docs.length === 0) { document.getElementById('archive-card-container').innerHTML = '<div style="font-size: 12px; color: #999; text-align: center; width: 100%; margin-top: 20px;">暂无档案</div>'; return; }
    
    // 如果没有选中的文档，或者选中的文档被删除了，默认选中第一个
    if (!selectedArchiveDoc || !docs.includes(selectedArchiveDoc)) selectedArchiveDoc = docs[0];

    // 将当前选中的置顶
    let sortedDocs = [...docs];
    const activeIndex = sortedDocs.indexOf(selectedArchiveDoc);
    if(activeIndex > 0) {
        const activeItem = sortedDocs.splice(activeIndex, 1)[0];
        sortedDocs.unshift(activeItem);
    }

    const maxShow = 5;
    const toShow = showAllDocs ? sortedDocs : sortedDocs.slice(0, maxShow);

    toShow.forEach(name => {
        const item = document.createElement('div');
        item.className = `doc-list-item ${name === selectedArchiveDoc ? 'active' : ''}`;
        item.innerText = name;
        item.onclick = () => { selectedArchiveDoc = name; renderArchive(); };
        list.appendChild(item);
    });

    // 渲染省略号或收起按钮
    if (!showAllDocs && sortedDocs.length > maxShow) {
        const moreBtn = document.createElement('div');
        moreBtn.className = 'doc-list-item';
        moreBtn.style.textAlign = 'center';
        moreBtn.style.color = '#a19f9d';
        moreBtn.innerText = '...';
        moreBtn.onclick = () => { showAllDocs = true; renderArchive(); };
        list.appendChild(moreBtn);
    } else if (showAllDocs && sortedDocs.length > maxShow) {
        const lessBtn = document.createElement('div');
        lessBtn.className = 'doc-list-item';
        lessBtn.style.textAlign = 'center';
        lessBtn.style.color = '#a19f9d';
        lessBtn.innerText = '收起 ∧';
        lessBtn.onclick = () => { showAllDocs = false; renderArchive(); };
        list.appendChild(lessBtn);
    }

    renderArchiveCard(selectedArchiveDoc, db);
}

function renderArchiveCard(docName, db) {
    const container = document.getElementById('archive-card-container');
    if(!container) return; container.innerHTML = '';
    archiveCharts.forEach(c => c.destroy()); archiveCharts = [];
    const records = db.filter(r => r.docName === docName);
    const card = document.createElement('div');
    card.style.cssText = "aspect-ratio: 3/4; width: 92%; background: white; border: 1px solid #e1dfdd; border-radius: 6px; padding: 15px; display: flex; flex-direction: column;";
    const L_id = `L_${Date.now()}`; const B_id = `B_${Date.now()}`;
    
    // 💡【核心修改】浅灰色、文字“删除”、放置于右下角
    card.innerHTML = `
        <div style="font-size: 16px; font-weight: 600; color: #323130; margin-bottom: 10px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${docName}">📄 ${docName}</div>
        <div style="flex:1; position: relative;"><canvas id="${L_id}"></canvas></div>
        <div style="flex:1; position: relative;"><canvas id="${B_id}"></canvas></div>
        <div style="text-align: right; margin-top: 5px;">
            <button onclick="deleteArchive('${docName}')" style="background: transparent; border: none; color: #d2d0ce; cursor: pointer; font-size: 12px; padding: 0;" title="删除">删除</button>
        </div>
    `;
    container.appendChild(card);
    setTimeout(() => {
        const ctxL = document.getElementById(L_id).getContext('2d');
        archiveCharts.push(new Chart(ctxL, { type: 'line', data: { labels: records.map(r => r.time), datasets: [{ data: records.map(r => r.count), borderColor: '#605e5c' }] }, options: { responsive: true, maintainAspectRatio: false } }));
        const ctxB = document.getElementById(B_id).getContext('2d');
        archiveCharts.push(new Chart(ctxB, { type: 'bar', data: { labels: records.map(r => r.date), datasets: [{ data: records.map(r => r.count), backgroundColor: '#605e5c' }] }, options: { responsive: true, maintainAspectRatio: false } }));
    }, 0);
}

// 档案删除逻辑
window.deleteArchive = function(docName) {
    if(confirm('确定删除吗？')) {
        let db = []; try { db = JSON.parse(localStorage.getItem('writerCompanionDB') || '[]'); } catch(e){}
        db = db.filter(r => r.docName !== docName);
        localStorage.setItem('writerCompanionDB', JSON.stringify(db));
        selectedArchiveDoc = null; // 重置选中项
        renderArchive();
    }
};

// ==================== 提示词管理 ====================
function initPrompts() {
    const def = [{ id: 'p_1', title: '基础润色', content: '你是一个资深的文字编辑。帮我润色文字，直接输出结果。' }];
    try { customPrompts = JSON.parse(localStorage.getItem('writerPrompts')) || def; activePromptId = localStorage.getItem('writerActivePrompt'); } catch(e) { customPrompts = def; }
    if (!activePromptId || !customPrompts.find(p => p.id === activePromptId)) activePromptId = customPrompts[0].id;
    updatePromptCapsule();
}

function updatePromptCapsule() {
    const p = customPrompts.find(p => p.id === activePromptId) || customPrompts[0];
    const cap = document.getElementById('active-prompt-capsule');
    if(cap) cap.innerText = p.title;
}

function renderPromptList() {
    const list = document.getElementById('prompt-list'); if(!list) return;
    list.innerHTML = '';
    
    // 将当前使用的提示词置顶
    let sortedPrompts = [...customPrompts];
    const activeIndex = sortedPrompts.findIndex(p => p.id === activePromptId);
    if(activeIndex > 0) {
        const activeItem = sortedPrompts.splice(activeIndex, 1)[0];
        sortedPrompts.unshift(activeItem);
    }

    const maxShow = 5;
    const toShow = showAllPrompts ? sortedPrompts : sortedPrompts.slice(0, maxShow);

    toShow.forEach(p => {
        const item = document.createElement('div');
        item.className = `doc-list-item ${p.id === activePromptId ? 'active' : ''}`;
        item.innerText = p.title;
        item.onclick = () => { activePromptId = p.id; savePromptsData(); renderPromptList(); };
        list.appendChild(item);
    });

    // 渲染省略号或收起按钮
    if (!showAllPrompts && sortedPrompts.length > maxShow) {
        const moreBtn = document.createElement('div');
        moreBtn.className = 'doc-list-item';
        moreBtn.style.textAlign = 'center';
        moreBtn.style.color = '#a19f9d';
        moreBtn.innerText = '...';
        moreBtn.onclick = () => { showAllPrompts = true; renderPromptList(); };
        list.appendChild(moreBtn);
    } else if (showAllPrompts && sortedPrompts.length > maxShow) {
        const lessBtn = document.createElement('div');
        lessBtn.className = 'doc-list-item';
        lessBtn.style.textAlign = 'center';
        lessBtn.style.color = '#a19f9d';
        lessBtn.innerText = '收起 ∧';
        lessBtn.onclick = () => { showAllPrompts = false; renderPromptList(); };
        list.appendChild(lessBtn);
    }

    renderPromptCard(activePromptId);
}

function renderPromptCard(id) {
    const container = document.getElementById('prompt-card-container');
    if(!container) return; container.innerHTML = '';
    const p = customPrompts.find(x => x.id === id);
    const card = document.createElement('div');
    card.style.cssText = "aspect-ratio: 3/4; width: 92%; background: white; border: 1px solid #e1dfdd; border-radius: 6px; padding: 15px; display: flex; flex-direction: column;";
    
    // 💡【核心修改】浅灰色、文字“删除”、放置于右下角
    card.innerHTML = `
        <input id="edit-prompt-title" value="${p.title}" onblur="saveCurrentPromptEdit('${id}')" style="width:100%; border:none; font-weight:600; outline:none; font-size: 16px; color: #323130;">
        <textarea id="edit-prompt-content" onblur="saveCurrentPromptEdit('${id}')" style="flex:1; border:none; resize:none; outline:none; font-size:12px; margin-top:10px; color: #605e5c;">${p.content}</textarea>
        <div style="text-align: right; margin-top: 5px;">
            <button onclick="deletePrompt('${id}')" style="background: transparent; border: none; color: #d2d0ce; cursor: pointer; font-size: 12px; padding: 0;" title="删除">删除</button>
        </div>
    `;
    container.appendChild(card);
}

// 提示词删除逻辑
window.deletePrompt = function(id) {
    if(confirm('确定删除吗？')) {
        customPrompts = customPrompts.filter(p => p.id !== id);
        // 如果删光了，给一个默认兜底兜住
        if(customPrompts.length === 0) {
            customPrompts = [{ id: 'p_1', title: '基础润色', content: '你是一个资深的文字编辑。帮我润色文字，直接输出结果。' }];
        }
        activePromptId = customPrompts[0].id;
        savePromptsData();
        renderPromptList();
    }
};

function saveCurrentPromptEdit(id) {
    const p = customPrompts.find(x => x.id === id);
    if(p) {
        p.title = document.getElementById('edit-prompt-title').value;
        p.content = document.getElementById('edit-prompt-content').value;
        savePromptsData(); renderPromptList();
    }
}

function savePromptsData() {
    try { localStorage.setItem('writerPrompts', JSON.stringify(customPrompts)); localStorage.setItem('writerActivePrompt', activePromptId); } catch(e) {}
    updatePromptCapsule();
}

function addNewPrompt() {
    const nid = 'p_' + Date.now(); customPrompts.unshift({ id: nid, title: '新提示词', content: '设定...' });
    activePromptId = nid; savePromptsData(); renderPromptList();
}

// ==================== 音频逻辑 ====================
function stopAllAudio() {
    if (currentAudio) { currentAudio.audio.pause(); currentAudio = null; }
    document.querySelectorAll(".audio-toggle").forEach(el => el.style.opacity = "1");
    document.getElementById("stop-audio-btn").style.display = "none";
}
function playAudio(type, el) {
    stopAllAudio();
    const urls = { 'rainy': 'assets/rainy.mp3', 'sun': 'assets/sun.mp3', 'coffee': 'assets/coffee.mp3', 'river': 'assets/river.mp3', 'ocean': 'assets/ocean.mp3', 'fire': 'assets/fire.mp3' };
    const audio = new Audio(urls[type]); audio.loop = true; audio.play();
    currentAudio = { type, audio };
    document.querySelectorAll(".audio-toggle").forEach(e => e.style.opacity = "0.3");
    el.style.opacity = "1"; document.getElementById("stop-audio-btn").style.display = "inline-flex";
}

// ==================== 🔒 核心 AI 聊天逻辑 ====================

async function handleAiChat() {
    const inputEl = document.getElementById("ai-input");
    const text = inputEl.value.trim();
    if (!text) return;

    const providerSelect = document.getElementById('api-provider');
    const provider = providerSelect.value;
    const isDeepThink = document.getElementById('deep-think-toggle').checked;

    const apiKey = getApiKey(provider);
    if (!apiKey) {
        addChatMessage(`⚠️ 请先配置 ${provider} 的密钥。`, 'ai');
        return;
    }

    addChatMessage(text, 'user');
    inputEl.value = ''; inputEl.style.height = '20px';
    
    // 💡【核心修改】改成“思考中...”
    addChatMessage("思考中...", 'ai', true);
    
    const activeP = customPrompts.find(p => p.id === activePromptId) || customPrompts[0];
    let requestBody = {
        messages: [{ role: "system", content: activeP.content }, { role: "user", content: text }],
        stream: true 
    };

    let apiUrl = '';
    if (provider === 'deepseek') {
        apiUrl = 'https://api.deepseek.com/chat/completions';
        requestBody.model = isDeepThink ? 'deepseek-reasoner' : 'deepseek-chat';
    } else {
        apiUrl = 'https://open.bigmodel.cn/api/paas/v4/chat/completions';
        requestBody.model = 'glm-5'; 
        if (isDeepThink) requestBody.thinking = { type: "enabled" };
    }

    try {
        const response = await fetch(apiUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` },
            body: JSON.stringify(requestBody)
        });

        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        
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
                if (line.startsWith('data: ')) {
                    const dataStr = line.slice(6).trim();
                    if (dataStr === '[DONE]') continue;
                    try {
                        const data = JSON.parse(dataStr);
                        const delta = data.choices[0].delta;
                        if (delta && delta.content) {
                            fullContent += delta.content;
                            streamMsgDiv.innerText = fullContent;
                            const hist = document.getElementById("chat-history");
                            hist.scrollTop = hist.scrollHeight;
                        }
                    } catch (e) {}
                }
            }
        }
    } catch (e) { 
        removeLoadingMessage(); addChatMessage(`⚠️ 错误: ${e.message}`, 'ai'); 
    }
}

function addChatMessage(text, sender, isLoading = false) {
    const chatHistory = document.getElementById("chat-history");
    if(!chatHistory) return null;
    const msgDiv = document.createElement('div');
    msgDiv.style.cssText = "padding: 8px 12px; border-radius: 12px; max-width: 85%; font-size: 13px; margin-bottom: 5px; word-break: break-word; white-space: pre-wrap;";
    if (sender === 'user') { 
        msgDiv.style.alignSelf = 'flex-end'; msgDiv.style.background = '#605e5c'; msgDiv.style.color = 'white'; 
    } else { 
        msgDiv.style.alignSelf = 'flex-start'; msgDiv.style.background = '#edebe9'; msgDiv.style.color = '#605e5c';
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
