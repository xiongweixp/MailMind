(() => {
    let lastRange = null, insertionRange = null, popup = null, requestId = null, nextRequestId = 0;
    let labels = {};
    const texts = {
        'zh-CN': ['描述你想写的内容...', '发送', '生成中...', '插入', '取消', '生成失败，请重试', '关闭'],
        'zh-TW': ['描述你想寫的內容...', '傳送', '產生中...', '插入', '取消', '產生失敗，請重試', '關閉'],
        'en-US': ['Describe what you want to write...', 'Send', 'Generating...', 'Insert', 'Cancel', 'Generation failed. Try again.', 'Close']
    };
    document.addEventListener('selectionchange', () => {
        const selection = window.getSelection();
        if (selection.rangeCount && document.body.contains(selection.anchorNode) && document.body.contains(selection.focusNode))
            lastRange = selection.getRangeAt(0).cloneRange();
    });
    function close(notify = true) {
        if (notify && requestId) chrome.webview.postMessage({ type: 'composeAiCancel' });
        requestId = null;
        popup?.remove(); popup = null;
    }
    window.mailmindAiReset = () => { close(); lastRange = null; insertionRange = null; };
    function button(text, title, color, action) {
        const b = document.createElement('button');
        b.type = 'button'; b.textContent = text; b.title = title; b.setAttribute('aria-label', title);
        Object.assign(b.style, { border: '0', borderRadius: '7px', background: color, color: 'white', width: '34px', height: '34px', cursor: 'pointer', fontSize: '19px' });
        b.addEventListener('click', action);
        return b;
    }
    window.mailmindAiOpen = language => {
        close();
        labels = texts[language] || texts['en-US'];
        insertionRange = lastRange && document.body.contains(lastRange.startContainer) ? lastRange.cloneRange() : document.createRange();
        if (!lastRange || !document.body.contains(lastRange.startContainer)) insertionRange.setStart(document.body, 0);
        insertionRange.collapse(true);
        let rect = insertionRange.getBoundingClientRect();
        if (!rect.height) rect = document.body.getBoundingClientRect();
        popup = document.createElement('div');
        popup.contentEditable = 'false';
        Object.assign(popup.style, { position: 'fixed', zIndex: '2147483647', width: 'min(340px, calc(100vw - 24px))', boxSizing: 'border-box', padding: '12px', border: '1px solid #cbd5e1', borderRadius: '12px', background: '#fff', color: '#1e293b', boxShadow: '0 8px 30px #0003', font: '14px sans-serif' });
        document.documentElement.appendChild(popup);
        popup.style.left = Math.max(8, Math.min(rect.left, window.innerWidth - popup.offsetWidth - 8)) + 'px';
        popup.style.top = Math.max(8, Math.min(rect.bottom + 8, window.innerHeight - 160)) + 'px';
        const input = document.createElement('textarea');
        input.placeholder = labels[0]; input.setAttribute('aria-label', labels[0]);
        Object.assign(input.style, { width: '100%', height: '70px', resize: 'vertical', border: '0', outline: 'none', boxSizing: 'border-box', font: 'inherit', color: 'inherit' });
        const status = document.createElement('div');
        Object.assign(status.style, { color: '#b42318', whiteSpace: 'pre-wrap', marginBottom: '4px' });
        const controls = document.createElement('div');
        Object.assign(controls.style, { display: 'flex', justifyContent: 'flex-end', gap: '8px' });
        const cancel = button('×', labels[6], '#94a3b8', () => close());
        const result = document.createElement('div');
        Object.assign(result.style, { whiteSpace: 'pre-wrap', overflowY: 'auto', maxHeight: '220px', marginBottom: '10px' });
        const insert = button('✓', labels[3], '#228b51', () => {
            if (!document.body.contains(insertionRange.startContainer)) { close(); return; }
            const selection = window.getSelection();
            document.body.focus(); selection.removeAllRanges(); selection.addRange(insertionRange);
            document.execCommand('insertText', false, result.textContent);
            close();
        });
        insert.disabled = true; insert.style.opacity = '0.5';
        const send = button('➤', labels[1], '#2863cf', () => {
            const prompt = input.value.trim();
            if (!prompt || requestId) return;
            requestId = String(++nextRequestId);
            result.textContent = labels[2];
            popup.replaceChildren(result, controls);
            controls.replaceChildren(cancel);
            chrome.webview.postMessage({ type: 'composeAiGenerate', prompt, id: requestId });
        });
        controls.append(cancel, send); popup.append(input, status, controls); input.focus();
        window.mailmindAiProgress = (id, text) => {
            if (!popup || id !== requestId || !text) return;
            result.textContent = text;
            result.scrollTop = result.scrollHeight;
        };
        window.mailmindAiResult = (id, text, error) => {
            if (!popup || id !== requestId) return;
            requestId = null;
            if (error) {
                status.style.color = '#b42318'; status.textContent = labels[5] + ': ' + text;
                controls.replaceChildren(cancel, send); popup.replaceChildren(input, status, controls); return;
            }
            result.textContent = text;
            insert.disabled = false; insert.style.opacity = '1';
            controls.replaceChildren(insert, button('×', labels[4], '#c33d3d', () => close()));
            popup.style.top = Math.max(8, Math.min(parseFloat(popup.style.top), window.innerHeight - popup.offsetHeight - 8)) + 'px';
        };
    };
})();
