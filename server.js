import 'dotenv/config';
import express from 'express';
import multer from 'multer';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const UPLOADS_DIR = path.resolve(__dirname, 'uploads');

// --- Configuration ---
const DEEPSEEK_API_KEY = process.env.DEEPSEEK_API_KEY;
if (!DEEPSEEK_API_KEY) {
  console.error('DEEPSEEK_API_KEY environment variable is required');
  process.exit(1);
}

const DEEPSEEK_BASE_URL = process.env.DEEPSEEK_BASE_URL || 'https://api.deepseek.com';
const DEEPSEEK_MODEL = process.env.DEEPSEEK_MODEL || 'deepseek-chat';
const PORT = parseInt(process.env.PORT || '3000', 10);
const MAX_FILE_SIZE_MB = parseInt(process.env.MAX_FILE_SIZE_MB || '50', 10);
const SESSION_TTL_MINUTES = parseInt(process.env.SESSION_TTL_MINUTES || '120', 10);
const MAX_HISTORY_MESSAGES = parseInt(process.env.MAX_HISTORY_MESSAGES || '50', 10);
const MAX_TOOL_ITERATIONS = parseInt(process.env.MAX_TOOL_ITERATIONS || '10', 10);

const SYSTEM_PROMPT = `You are an expert data analyst assistant. The user has uploaded Excel workbooks that you can query using the available tools.

WORKFLOW:
1. When the user asks about their data, first use 'list_workbooks' to see what files are available (if you haven't already).
2. Use 'get_sheet_summary' to understand the structure of relevant sheets before querying.
3. Use 'query_sheet' for filtering, aggregation, and analysis. Use 'read_sheet_data' when you need raw row data.
4. Always explain your findings clearly. Format numbers nicely (commas, rounding). Use tables in markdown when presenting tabular results.

RULES:
- Never fabricate data. Only report what the tools return.
- If a query is ambiguous, ask the user to clarify which file or column they mean.
- If a tool returns an error, explain the issue to the user and suggest a fix.
- When the user first uploads a file, proactively summarise its structure.`;

// --- Ensure uploads directory ---
if (!fs.existsSync(UPLOADS_DIR)) {
  fs.mkdirSync(UPLOADS_DIR, { recursive: true });
}

// --- Express Setup ---
const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// --- Multer Setup ---
const storage = multer.diskStorage({
  destination: UPLOADS_DIR,
  filename: (_req, file, cb) => {
    cb(null, file.originalname);
  },
});

const upload = multer({
  storage,
  limits: { fileSize: MAX_FILE_SIZE_MB * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (['.xlsx', '.xls', '.csv'].includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error('Only .xlsx, .xls, and .csv files are allowed'));
    }
  },
});

// --- Session Store ---
const sessions = new Map();

function getSession(sessionId) {
  let session = sessions.get(sessionId);
  if (!session) {
    session = { messages: [], lastActive: new Date() };
    sessions.set(sessionId, session);
  }
  session.lastActive = new Date();
  return session;
}

function trimHistory(session) {
  if (session.messages.length > MAX_HISTORY_MESSAGES) {
    session.messages = session.messages.slice(-MAX_HISTORY_MESSAGES);
  }
}

// Session cleanup every 30 minutes
setInterval(() => {
  const cutoff = Date.now() - SESSION_TTL_MINUTES * 60 * 1000;
  for (const [id, session] of sessions) {
    if (session.lastActive.getTime() < cutoff) {
      sessions.delete(id);
    }
  }
}, 30 * 60 * 1000);

// --- MCP Client ---
let mcpClient = null;
let mcpTools = [];
let mcpRestartCount = 0;
const MAX_MCP_RESTARTS = 3;

async function startMcpClient() {
  const transport = new StdioClientTransport({
    command: 'node',
    args: [path.join(__dirname, 'mcp-server.js')],
  });

  const client = new Client({ name: 'excel-chat-client', version: '1.0.0' }, {});
  await client.connect(transport);

  const { tools } = await client.listTools();
  mcpTools = tools;
  mcpClient = client;
  mcpRestartCount = 0;

  // Handle transport close for restart
  transport.onclose = async () => {
    console.error('MCP server process exited');
    mcpClient = null;
    if (mcpRestartCount < MAX_MCP_RESTARTS) {
      mcpRestartCount++;
      console.log(`Attempting MCP restart (${mcpRestartCount}/${MAX_MCP_RESTARTS})...`);
      await new Promise(r => setTimeout(r, 2000));
      try {
        await startMcpClient();
        console.log('MCP server restarted successfully');
      } catch (err) {
        console.error('MCP restart failed:', err.message);
      }
    }
  };

  return client;
}

// Convert MCP tools to OpenAI function-calling format
function getToolDefinitions() {
  return mcpTools.map(tool => ({
    type: 'function',
    function: {
      name: tool.name,
      description: tool.description,
      parameters: tool.inputSchema,
    },
  }));
}

// --- API Endpoints ---

// Upload file
app.post('/api/upload', (req, res) => {
  upload.single('file')(req, res, (err) => {
    if (err) {
      return res.status(400).json({ ok: false, error: err.message });
    }
    if (!req.file) {
      return res.status(400).json({ ok: false, error: 'No file provided' });
    }
    res.json({ ok: true, filename: req.file.filename, originalName: req.file.originalname });
  });
});

// List files
app.get('/api/files', (_req, res) => {
  if (!fs.existsSync(UPLOADS_DIR)) {
    return res.json({ files: [] });
  }
  const files = fs.readdirSync(UPLOADS_DIR).filter(f =>
    /\.(xlsx|xls|csv)$/i.test(f)
  );
  res.json({ files });
});

// Delete file
app.delete('/api/files/:filename', (req, res) => {
  const filepath = path.resolve(UPLOADS_DIR, req.params.filename);
  if (!filepath.startsWith(UPLOADS_DIR)) {
    return res.status(400).json({ ok: false, error: 'Invalid filename' });
  }
  try {
    if (fs.existsSync(filepath)) {
      fs.unlinkSync(filepath);
    }
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Delete session
app.delete('/api/session/:sessionId', (req, res) => {
  sessions.delete(req.params.sessionId);
  res.json({ ok: true });
});

// Chat endpoint (SSE)
app.post('/api/chat', async (req, res) => {
  const { message, sessionId } = req.body;
  if (!message || !sessionId) {
    return res.status(400).json({ error: 'message and sessionId required' });
  }

  // SSE headers
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders();

  const session = getSession(sessionId);
  session.messages.push({ role: 'user', content: message });
  trimHistory(session);

  try {
    await processChat(session, res);
  } catch (err) {
    sendSSE(res, { type: 'error', content: `Error: ${err.message}` });
  }

  sendSSE(res, { type: 'done' });
  res.end();
});

function sendSSE(res, data) {
  res.write(`data: ${JSON.stringify(data)}\n\n`);
}

async function processChat(session, res) {
  let iterations = 0;

  while (iterations < MAX_TOOL_ITERATIONS) {
    iterations++;

    const messages = [
      { role: 'system', content: SYSTEM_PROMPT },
      ...session.messages,
    ];

    const result = await callDeepSeek(messages, res);

    if (result.toolCalls && result.toolCalls.length > 0) {
      // Add assistant message with tool calls to history
      session.messages.push({
        role: 'assistant',
        content: result.content || null,
        tool_calls: result.toolCalls,
      });

      // Execute each tool call via MCP
      for (const tc of result.toolCalls) {
        sendSSE(res, { type: 'status', content: 'Reading Excel data...' });

        let toolResult;
        try {
          if (!mcpClient) {
            throw new Error('MCP server is not available');
          }
          const mcpResult = await mcpClient.callTool({
            name: tc.function.name,
            arguments: JSON.parse(tc.function.arguments),
          });
          toolResult = mcpResult.content.map(c => c.text).join('\n');
        } catch (err) {
          toolResult = `Error calling tool: ${err.message}`;
        }

        session.messages.push({
          role: 'tool',
          tool_call_id: tc.id,
          content: toolResult,
        });
      }

      trimHistory(session);
      // Loop again to get the LLM's response with tool results
      continue;
    }

    // No tool calls — final response
    if (result.content) {
      session.messages.push({ role: 'assistant', content: result.content });
      trimHistory(session);
    }
    break;
  }

  if (iterations >= MAX_TOOL_ITERATIONS) {
    const msg = "I've reached the maximum number of data lookups for this query. Please try a more specific question.";
    sendSSE(res, { type: 'token', content: msg });
    session.messages.push({ role: 'assistant', content: msg });
  }
}

async function callDeepSeek(messages, res) {
  const body = {
    model: DEEPSEEK_MODEL,
    messages,
    tools: getToolDefinitions(),
    stream: true,
  };

  let response;
  let retries = 0;
  const maxRetries = 1;

  while (true) {
    try {
      response = await fetch(`${DEEPSEEK_BASE_URL}/v1/chat/completions`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${DEEPSEEK_API_KEY}`,
        },
        body: JSON.stringify(body),
      });

      if (response.status === 429 && retries < maxRetries) {
        retries++;
        await new Promise(r => setTimeout(r, 2000));
        continue;
      }
      if (response.status >= 500 && retries < maxRetries) {
        retries++;
        await new Promise(r => setTimeout(r, 2000));
        continue;
      }
      if (!response.ok) {
        throw new Error(`DeepSeek API error: ${response.status} ${response.statusText}`);
      }
      break;
    } catch (err) {
      if (retries < maxRetries && (err.code === 'ECONNRESET' || err.code === 'ETIMEDOUT')) {
        retries++;
        await new Promise(r => setTimeout(r, 2000));
        continue;
      }
      throw err;
    }
  }

  // Parse SSE stream from DeepSeek
  let contentAccum = '';
  const toolCallsAccum = {};
  let finishReason = null;

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
      if (!trimmed || !trimmed.startsWith('data: ')) continue;
      const data = trimmed.slice(6);
      if (data === '[DONE]') continue;

      let parsed;
      try {
        parsed = JSON.parse(data);
      } catch {
        continue;
      }

      const choice = parsed.choices?.[0];
      if (!choice) continue;

      if (choice.finish_reason) {
        finishReason = choice.finish_reason;
      }

      const delta = choice.delta;
      if (!delta) continue;

      // Text content
      if (delta.content) {
        contentAccum += delta.content;
        sendSSE(res, { type: 'token', content: delta.content });
      }

      // Tool calls
      if (delta.tool_calls) {
        for (const tc of delta.tool_calls) {
          const idx = tc.index;
          if (!toolCallsAccum[idx]) {
            toolCallsAccum[idx] = { id: '', function: { name: '', arguments: '' } };
          }
          if (tc.id) toolCallsAccum[idx].id = tc.id;
          if (tc.function?.name) toolCallsAccum[idx].function.name += tc.function.name;
          if (tc.function?.arguments) toolCallsAccum[idx].function.arguments += tc.function.arguments;
        }
      }
    }
  }

  const toolCalls = Object.keys(toolCallsAccum).length > 0
    ? Object.values(toolCallsAccum)
    : null;

  return { content: contentAccum, toolCalls, finishReason };
}

// --- Start Server ---
async function main() {
  console.log('Starting MCP server...');
  await startMcpClient();
  console.log(`MCP connected. ${mcpTools.length} tools available: ${mcpTools.map(t => t.name).join(', ')}`);

  app.listen(PORT, () => {
    console.log(`Excel Chat server running on http://localhost:${PORT}`);
  });
}

main().catch(err => {
  console.error('Startup error:', err);
  process.exit(1);
});
