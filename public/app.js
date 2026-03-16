// --- Session Management ---
let sessionId = sessionStorage.getItem('sessionId');
if (!sessionId) {
  sessionId = crypto.randomUUID();
  sessionStorage.setItem('sessionId', sessionId);
}

// --- DOM References ---
const chatMessages = document.getElementById('chatMessages');
const chatInput = document.getElementById('chatInput');
const sendBtn = document.getElementById('sendBtn');
const fileInput = document.getElementById('fileInput');
const dropzone = document.getElementById('dropzone');
const fileList = document.getElementById('fileList');
const uploadProgress = document.getElementById('uploadProgress');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const newChatBtn = document.getElementById('newChatBtn');
const dragOverlay = document.getElementById('dragOverlay');
const sidebar = document.getElementById('sidebar');
const sidebarToggle = document.getElementById('sidebarToggle');
const sidebarClose = document.getElementById('sidebarClose');

let isSending = false;

// --- Markdown Renderer (lightweight regex-based) ---
function renderMarkdown(text) {
  let html = text;

  // Code blocks (```...```)
  html = html.replace(/```(\w*)\n([\s\S]*?)```/g, (_m, lang, code) => {
    const escaped = escapeHtml(code.trim());
    return `<pre><code class="language-${lang}">${escaped}</code></pre>`;
  });

  // Inline code
  html = html.replace(/`([^`]+)`/g, '<code>$1</code>');

  // Tables
  html = html.replace(/((?:^\|.+\|$\n?)+)/gm, (tableBlock) => {
    const lines = tableBlock.trim().split('\n');
    if (lines.length < 2) return tableBlock;

    const parseRow = (line) =>
      line.split('|').slice(1, -1).map(cell => cell.trim());

    const headerCells = parseRow(lines[0]);
    // Check if second line is separator
    const isSep = /^\|[\s\-:|]+\|$/.test(lines[1]);
    if (!isSep) return tableBlock;

    let table = '<table><thead><tr>';
    for (const h of headerCells) {
      table += `<th>${h}</th>`;
    }
    table += '</tr></thead><tbody>';

    for (let i = 2; i < lines.length; i++) {
      const cells = parseRow(lines[i]);
      table += '<tr>';
      for (const c of cells) {
        table += `<td>${c}</td>`;
      }
      table += '</tr>';
    }
    table += '</tbody></table>';
    return table;
  });

  // Headers
  html = html.replace(/^### (.+)$/gm, '<h3>$1</h3>');
  html = html.replace(/^## (.+)$/gm, '<h2>$1</h2>');
  html = html.replace(/^# (.+)$/gm, '<h1>$1</h1>');

  // Bold & Italic
  html = html.replace(/\*\*\*(.+?)\*\*\*/g, '<strong><em>$1</em></strong>');
  html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/\*(.+?)\*/g, '<em>$1</em>');

  // Unordered lists
  html = html.replace(/^[\-\*] (.+)$/gm, '<li>$1</li>');
  html = html.replace(/((?:<li>.*<\/li>\n?)+)/g, '<ul>$1</ul>');

  // Ordered lists
  html = html.replace(/^\d+\. (.+)$/gm, '<li>$1</li>');

  // Line breaks (double newline = paragraph)
  html = html.replace(/\n{2,}/g, '</p><p>');
  html = html.replace(/\n/g, '<br>');

  // Wrap in paragraph if not starting with block element
  if (!html.startsWith('<')) {
    html = `<p>${html}</p>`;
  }

  return html;
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

// --- Chat Functions ---
function appendMessage(role, content, isHtml = false) {
  // Remove welcome message
  const welcome = chatMessages.querySelector('.welcome-message');
  if (welcome) welcome.remove();

  const div = document.createElement('div');
  div.className = `message ${role}`;

  if (isHtml) {
    div.innerHTML = content;
  } else if (role === 'assistant') {
    div.innerHTML = renderMarkdown(content);
  } else {
    div.textContent = content;
  }

  chatMessages.appendChild(div);
  chatMessages.scrollTop = chatMessages.scrollHeight;
  return div;
}

function showStatus(text) {
  const pill = document.createElement('div');
  pill.className = 'status-pill';
  pill.textContent = text;
  chatMessages.appendChild(pill);
  chatMessages.scrollTop = chatMessages.scrollHeight;
  return pill;
}

function removeStatus(pill) {
  if (pill && pill.parentNode) pill.parentNode.removeChild(pill);
}

async function sendMessage(text) {
  if (isSending || !text.trim()) return;

  isSending = true;
  sendBtn.disabled = true;
  chatInput.disabled = true;

  appendMessage('user', text);
  chatInput.value = '';
  chatInput.style.height = 'auto';

  let assistantDiv = null;
  let accum = '';
  let statusPill = null;

  try {
    const response = await fetch('/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ message: text, sessionId }),
    });

    if (!response.ok) {
      throw new Error(`Server error: ${response.status}`);
    }

    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split('\n');
      buffer = lines.pop() || '';

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed.startsWith('data: ')) continue;

        let data;
        try {
          data = JSON.parse(trimmed.slice(6));
        } catch {
          continue;
        }

        if (data.type === 'token') {
          if (statusPill) {
            removeStatus(statusPill);
            statusPill = null;
          }
          accum += data.content;
          if (!assistantDiv) {
            assistantDiv = appendMessage('assistant', '');
          }
          assistantDiv.innerHTML = renderMarkdown(accum);
          chatMessages.scrollTop = chatMessages.scrollHeight;
        } else if (data.type === 'status') {
          if (statusPill) removeStatus(statusPill);
          statusPill = showStatus(data.content);
        } else if (data.type === 'error') {
          if (statusPill) {
            removeStatus(statusPill);
            statusPill = null;
          }
          appendMessage('error', data.content);
        } else if (data.type === 'done') {
          if (statusPill) {
            removeStatus(statusPill);
            statusPill = null;
          }
        }
      }
    }
  } catch (err) {
    appendMessage('error', `Failed to send message: ${err.message}`);
  }

  isSending = false;
  sendBtn.disabled = false;
  chatInput.disabled = false;
  chatInput.focus();
}

// --- File Management ---
async function loadFiles() {
  try {
    const res = await fetch('/api/files');
    const data = await res.json();
    renderFileList(data.files);
  } catch {
    // Silently fail
  }
}

function renderFileList(files) {
  fileList.innerHTML = '';
  for (const name of files) {
    const li = document.createElement('li');
    li.className = 'file-item';

    const span = document.createElement('span');
    span.className = 'file-name';
    span.textContent = name;

    const btn = document.createElement('button');
    btn.className = 'file-delete';
    btn.textContent = '\u00D7';
    btn.title = 'Delete file';
    btn.addEventListener('click', () => deleteFile(name));

    li.appendChild(span);
    li.appendChild(btn);
    fileList.appendChild(li);
  }
}

async function uploadFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls', 'csv'].includes(ext)) {
    appendMessage('error', 'Only .xlsx, .xls, and .csv files are supported.');
    return;
  }

  uploadProgress.hidden = false;
  progressFill.style.width = '0%';
  progressText.textContent = 'Uploading...';

  const formData = new FormData();
  formData.append('file', file);

  try {
    // Simulate progress with a simple animation
    progressFill.style.width = '60%';

    const res = await fetch('/api/upload', {
      method: 'POST',
      body: formData,
    });

    progressFill.style.width = '100%';
    const data = await res.json();

    if (!data.ok) {
      throw new Error(data.error || 'Upload failed');
    }

    progressText.textContent = 'Done!';
    setTimeout(() => {
      uploadProgress.hidden = true;
    }, 1000);

    await loadFiles();

    // Auto-send summary request
    sendMessage(`I've uploaded **${data.filename}**. Can you summarise its structure?`);
  } catch (err) {
    progressText.textContent = `Error: ${err.message}`;
    progressFill.style.width = '0%';
    setTimeout(() => {
      uploadProgress.hidden = true;
    }, 3000);
  }
}

async function deleteFile(filename) {
  try {
    await fetch(`/api/files/${encodeURIComponent(filename)}`, { method: 'DELETE' });
    await loadFiles();
  } catch {
    // Silently fail
  }
}

// --- New Chat ---
async function newChat() {
  try {
    await fetch(`/api/session/${encodeURIComponent(sessionId)}`, { method: 'DELETE' });
  } catch {
    // Silently fail
  }

  sessionId = crypto.randomUUID();
  sessionStorage.setItem('sessionId', sessionId);

  chatMessages.innerHTML = `
    <div class="welcome-message">
      <h2>Welcome to Excel Chat</h2>
      <p>Upload an Excel or CSV file to get started. Ask questions about your data in natural language.</p>
    </div>
  `;
}

// --- Event Listeners ---

// Send message
sendBtn.addEventListener('click', () => sendMessage(chatInput.value));

chatInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    sendMessage(chatInput.value);
  }
});

// Auto-grow textarea
chatInput.addEventListener('input', () => {
  chatInput.style.height = 'auto';
  chatInput.style.height = Math.min(chatInput.scrollHeight, 5 * 1.5 * 14 + 20) + 'px';
});

// File upload
dropzone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', () => {
  if (fileInput.files.length > 0) {
    uploadFile(fileInput.files[0]);
    fileInput.value = '';
  }
});

// Drag and drop on dropzone
dropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropzone.classList.add('drag-over');
});
dropzone.addEventListener('dragleave', () => {
  dropzone.classList.remove('drag-over');
});
dropzone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropzone.classList.remove('drag-over');
  if (e.dataTransfer.files.length > 0) {
    uploadFile(e.dataTransfer.files[0]);
  }
});

// Global drag overlay
let dragCounter = 0;
document.addEventListener('dragenter', (e) => {
  e.preventDefault();
  dragCounter++;
  dragOverlay.classList.add('active');
});
document.addEventListener('dragleave', (e) => {
  e.preventDefault();
  dragCounter--;
  if (dragCounter <= 0) {
    dragCounter = 0;
    dragOverlay.classList.remove('active');
  }
});
document.addEventListener('dragover', (e) => e.preventDefault());
document.addEventListener('drop', (e) => {
  e.preventDefault();
  dragCounter = 0;
  dragOverlay.classList.remove('active');
  if (e.dataTransfer.files.length > 0) {
    uploadFile(e.dataTransfer.files[0]);
  }
});

// New chat
newChatBtn.addEventListener('click', newChat);

// Sidebar toggle (mobile)
sidebarToggle.addEventListener('click', () => sidebar.classList.toggle('open'));
sidebarClose.addEventListener('click', () => sidebar.classList.remove('open'));

// --- Init ---
loadFiles();
