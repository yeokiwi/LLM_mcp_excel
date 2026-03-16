# Excel Chat — LLM-Powered Excel Query Interface with MCP

## Project Overview

Build a web-based chat application that lets users upload Excel files and query their contents using natural language. The system uses **DeepSeek LLM** as the reasoning engine, connected to an **MCP (Model Context Protocol) server** that reads and parses Excel files via **SheetJS**. The architecture is a three-tier pipeline: **Browser UI → Chat Server (MCP Client) → MCP Server (SheetJS)**.

---

## System Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                        Browser (Frontend)                           │
│  ┌───────────────────────────────────────────────────────────────┐  │
│  │  Chat UI (HTML/CSS/JS)                                       │  │
│  │  • Message thread (user + assistant bubbles)                 │  │
│  │  • File upload dropzone (drag & drop + click-to-browse)      │  │
│  │  • Input bar with send button                                │  │
│  │  • Active files sidebar listing uploaded workbooks            │  │
│  │  • SSE streaming for real-time token display                 │  │
│  └───────────────────────────────────────────────────────────────┘  │
└──────────────────────────────┬──────────────────────────────────────┘
                               │ HTTP REST + SSE
                               ▼
┌─────────────────────────────────────────────────────────────────────┐
│                    Chat Server (Node.js + Express)                   │
│  ┌───────────────────────────────────────────────────────────────┐  │
│  │  Responsibilities:                                           │  │
│  │  • Serve static frontend files                               │  │
│  │  • Accept file uploads (multer) → store in ./uploads/        │  │
│  │  • Maintain per-session conversation history (in-memory Map) │  │
│  │  • Forward user queries to DeepSeek API with tool definitions│  │
│  │  • Execute MCP tool calls by proxying to the MCP Server      │  │
│  │  • Stream DeepSeek responses back to frontend via SSE        │  │
│  │  • Implement the MCP Client (JSON-RPC over stdio)            │  │
│  └───────────────────────────────────────────────────────────────┘  │
└──────────────────────────────┬──────────────────────────────────────┘
                               │ JSON-RPC 2.0 over stdio
                               ▼
┌─────────────────────────────────────────────────────────────────────┐
│                    MCP Server (Node.js child process)                │
│  ┌───────────────────────────────────────────────────────────────┐  │
│  │  Responsibilities:                                           │  │
│  │  • Expose MCP tools via JSON-RPC over stdio transport        │  │
│  │  • Use SheetJS (xlsx) to read .xlsx / .xls / .csv files      │  │
│  │  • Provide structured data back to the MCP client            │  │
│  └───────────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Technology Stack

| Layer            | Technology                                           |
| ---------------- | ---------------------------------------------------- |
| Frontend         | Vanilla HTML + CSS + JavaScript (no framework)       |
| Chat Server      | Node.js 20+, Express, multer, eventsource (SSE)     |
| LLM              | DeepSeek Chat API (`deepseek-chat` model)            |
| MCP Client       | `@modelcontextprotocol/sdk` (Client class)           |
| MCP Server       | `@modelcontextprotocol/sdk` (Server class)           |
| Excel Parsing    | SheetJS (`xlsx` npm package)                         |
| Transport        | MCP stdio transport (child_process spawn)            |

---

## Project Structure

```
excel-chat/
├── package.json
├── server.js                  # Express chat server + MCP client
├── mcp-server.js              # MCP server process (SheetJS tools)
├── uploads/                   # Uploaded Excel files (gitignored)
├── public/
│   ├── index.html             # Chat UI
│   ├── style.css              # Dark-themed styling
│   └── app.js                 # Frontend logic (SSE, file upload, chat)
└── README.md
```

---

## MCP Server (`mcp-server.js`)

### Transport

- Use `StdioServerTransport` from `@modelcontextprotocol/sdk/server/stdio.js`.
- The server is launched as a child process by the chat server and communicates over stdin/stdout using JSON-RPC 2.0.

### MCP Tools to Expose

The MCP server must register the following tools via `server.setRequestHandler(ListToolsRequestSchema, ...)` and `server.setRequestHandler(CallToolRequestSchema, ...)`:

#### Tool 1: `list_workbooks`

- **Description**: "List all uploaded Excel workbooks available for querying."
- **Input Schema**: (none — no parameters)
- **Behaviour**: Read the `uploads/` directory, filter for `.xlsx`, `.xls`, `.csv` extensions, return a JSON array of filenames.
- **Return**: `{ "content": [{ "type": "text", "text": "<JSON array of filenames>" }] }`

#### Tool 2: `get_sheet_names`

- **Description**: "Get the names of all sheets in a given workbook."
- **Input Schema**:
  ```json
  {
    "type": "object",
    "properties": {
      "filename": { "type": "string", "description": "Name of the uploaded workbook file" }
    },
    "required": ["filename"]
  }
  ```
- **Behaviour**: Use `XLSX.readFile(path)` from SheetJS. Return `workbook.SheetNames`.
- **Return**: `{ "content": [{ "type": "text", "text": "<JSON array of sheet names>" }] }`

#### Tool 3: `read_sheet_data`

- **Description**: "Read data from a specific sheet. Returns the first N rows as JSON (default 100). Supports optional column filtering."
- **Input Schema**:
  ```json
  {
    "type": "object",
    "properties": {
      "filename": { "type": "string", "description": "Name of the uploaded workbook file" },
      "sheet": { "type": "string", "description": "Sheet name. Omit to use the first sheet." },
      "max_rows": { "type": "number", "description": "Maximum rows to return. Default 100." },
      "columns": {
        "type": "array",
        "items": { "type": "string" },
        "description": "Optional list of column headers to include. Omit to return all columns."
      }
    },
    "required": ["filename"]
  }
  ```
- **Behaviour**:
  1. `XLSX.readFile(filepath)` to get the workbook.
  2. Select the target sheet (`sheet` param or `SheetNames[0]`).
  3. `XLSX.utils.sheet_to_json(worksheet)` to convert to an array of row objects.
  4. If `columns` is specified, map each row to include only those keys.
  5. Slice to `max_rows`.
  6. Return the JSON array as text content.
- **Return**: `{ "content": [{ "type": "text", "text": "<JSON array of row objects>" }] }`

#### Tool 4: `get_sheet_summary`

- **Description**: "Get a structural summary of a sheet: column headers, row count, data types per column, and a sample of the first 5 rows."
- **Input Schema**:
  ```json
  {
    "type": "object",
    "properties": {
      "filename": { "type": "string", "description": "Name of the uploaded workbook file" },
      "sheet": { "type": "string", "description": "Sheet name. Omit to use the first sheet." }
    },
    "required": ["filename"]
  }
  ```
- **Behaviour**:
  1. Parse the sheet to JSON.
  2. Extract column names from the first row's keys.
  3. Count total rows.
  4. For each column, infer predominant data type by sampling (number, string, date, boolean, empty).
  5. Include the first 5 rows as a sample.
  6. Return a summary object:
     ```json
     {
       "filename": "sales.xlsx",
       "sheet": "Sheet1",
       "total_rows": 1542,
       "columns": [
         { "name": "Date", "type": "date", "non_empty": 1540 },
         { "name": "Revenue", "type": "number", "non_empty": 1542 }
       ],
       "sample_rows": [ ... ]
     }
     ```
- **Return**: `{ "content": [{ "type": "text", "text": "<JSON summary>" }] }`

#### Tool 5: `query_sheet`

- **Description**: "Filter and aggregate data from a sheet. Supports filtering rows by column value conditions and basic aggregations (sum, average, count, min, max) on numeric columns."
- **Input Schema**:
  ```json
  {
    "type": "object",
    "properties": {
      "filename": { "type": "string", "description": "Name of the uploaded workbook file" },
      "sheet": { "type": "string", "description": "Sheet name. Omit to use the first sheet." },
      "filters": {
        "type": "array",
        "description": "Array of filter conditions. Each is { column, operator, value }. Operators: eq, neq, gt, gte, lt, lte, contains, startsWith, endsWith.",
        "items": {
          "type": "object",
          "properties": {
            "column": { "type": "string" },
            "operator": { "type": "string", "enum": ["eq", "neq", "gt", "gte", "lt", "lte", "contains", "startsWith", "endsWith"] },
            "value": { "type": ["string", "number", "boolean"] }
          },
          "required": ["column", "operator", "value"]
        }
      },
      "aggregations": {
        "type": "array",
        "description": "Array of aggregation operations. Each is { column, function }.",
        "items": {
          "type": "object",
          "properties": {
            "column": { "type": "string" },
            "function": { "type": "string", "enum": ["sum", "average", "count", "min", "max"] }
          },
          "required": ["column", "function"]
        }
      },
      "group_by": { "type": "string", "description": "Optional column name to group results by before aggregation." },
      "sort_by": { "type": "string", "description": "Column name to sort results by." },
      "sort_order": { "type": "string", "enum": ["asc", "desc"], "description": "Sort direction. Default asc." },
      "max_rows": { "type": "number", "description": "Maximum result rows to return. Default 50." }
    },
    "required": ["filename"]
  }
  ```
- **Behaviour**:
  1. Parse the sheet to JSON rows.
  2. Apply `filters` sequentially (AND logic). For each filter, compare `row[column]` using the given operator. Type-coerce as needed (e.g., parse strings to numbers for numeric comparisons).
  3. If `group_by` is specified, group the filtered rows by that column's value.
  4. If `aggregations` are specified, compute each aggregation per group (or over the full filtered set if no grouping). Return the aggregation results.
  5. If no aggregations, return the filtered (and optionally sorted/sliced) rows.
  6. Apply `sort_by` and `sort_order` to the result set.
  7. Slice to `max_rows`.
- **Return**: `{ "content": [{ "type": "text", "text": "<JSON result>" }] }`

### SheetJS Usage Notes

- Import as `const XLSX = require('xlsx');` (CommonJS) or use `import * as XLSX from 'xlsx'` if using ESM.
- File paths: resolve filenames against the `uploads/` directory. Validate that the resolved path stays within `uploads/` (path traversal guard).
- Handle errors gracefully: if a file doesn't exist, a sheet name is invalid, or a column doesn't exist, return a clear error message in the MCP response content rather than crashing.

### Error Handling

- Every tool handler must be wrapped in try/catch.
- On error, return: `{ "content": [{ "type": "text", "text": "Error: <descriptive message>" }], "isError": true }`
- Common errors to handle: file not found, invalid sheet name, invalid column name in filters/aggregations, parse errors for corrupted files.

---

## Chat Server (`server.js`)

### Express Server Setup

- Serve `public/` as static files.
- Use `multer` for multipart file uploads to `./uploads/`.
- Listen on port `3000` (configurable via `PORT` env var).

### Endpoints

#### `POST /api/upload`

- Accept multipart form data with a field named `file`.
- Use multer with file filter: only accept `.xlsx`, `.xls`, `.csv` extensions.
- Max file size: 50 MB.
- On success, return: `{ "ok": true, "filename": "<stored filename>", "originalName": "<original name>" }`
- Store files with their original name. If a file with the same name exists, overwrite it.

#### `GET /api/files`

- Return JSON array of files currently in `uploads/`: `{ "files": ["sales.xlsx", "inventory.csv", ...] }`

#### `DELETE /api/files/:filename`

- Delete the named file from `uploads/`.
- Return: `{ "ok": true }`

#### `POST /api/chat`

- **Request body**: `{ "message": "<user text>", "sessionId": "<uuid>" }`
- **Response**: SSE stream (`Content-Type: text/event-stream`).
- Processing flow:
  1. Retrieve or initialise conversation history for the `sessionId` from an in-memory `Map<string, Array>`.
  2. Append the user message to the conversation history.
  3. Build the DeepSeek API request (see below).
  4. Call the DeepSeek API with streaming enabled.
  5. Stream each token chunk to the client as `data: {"type":"token","content":"..."}\n\n`.
  6. If DeepSeek returns a tool call (`tool_calls` in the response), do NOT stream the tool call JSON to the user. Instead:
     a. Send `data: {"type":"status","content":"Reading Excel data..."}\n\n` to show a loading indicator.
     b. Forward the tool call to the MCP server using the MCP client.
     c. Collect the MCP tool result.
     d. Append the assistant's tool-call message and the tool result message to the conversation history.
     e. Call DeepSeek again with the updated conversation (including tool result) and resume streaming the new response.
     f. Repeat if DeepSeek issues another tool call (loop up to 10 iterations max as a safety limit).
  7. When the final response is fully streamed, send `data: {"type":"done"}\n\n` and close the stream.

#### `DELETE /api/session/:sessionId`

- Clear the conversation history for the given session.
- Return: `{ "ok": true }`

### MCP Client Integration

- On server startup, spawn `mcp-server.js` as a child process using `StdioClientTransport` from `@modelcontextprotocol/sdk/client/stdio.js`.
- Create an MCP `Client` instance, connect it to the transport.
- On startup, call `client.listTools()` to discover available tools and cache the tool list.
- To execute a tool call: `client.callTool({ name: toolName, arguments: toolArgs })`.
- If the MCP server process crashes, detect the exit event and attempt to restart it (max 3 retries with 2-second backoff).

### DeepSeek API Integration

- **Base URL**: `https://api.deepseek.com` (configurable via `DEEPSEEK_BASE_URL` env var).
- **API Key**: Read from `DEEPSEEK_API_KEY` environment variable. Fail fast on startup if not set.
- **Model**: `deepseek-chat` (configurable via `DEEPSEEK_MODEL` env var).
- **Request format**: Use the OpenAI-compatible chat completions endpoint (`/v1/chat/completions`).
- **System prompt** to include with every request:

```
You are an expert data analyst assistant. The user has uploaded Excel workbooks that you can query using the available tools.

WORKFLOW:
1. When the user asks about their data, first use 'list_workbooks' to see what files are available (if you haven't already).
2. Use 'get_sheet_summary' to understand the structure of relevant sheets before querying.
3. Use 'query_sheet' for filtering, aggregation, and analysis. Use 'read_sheet_data' when you need raw row data.
4. Always explain your findings clearly. Format numbers nicely (commas, rounding). Use tables in markdown when presenting tabular results.

RULES:
- Never fabricate data. Only report what the tools return.
- If a query is ambiguous, ask the user to clarify which file or column they mean.
- If a tool returns an error, explain the issue to the user and suggest a fix.
- When the user first uploads a file, proactively summarise its structure.
```

- **Tool definitions**: Convert the MCP tool list into the OpenAI function-calling format:
  ```json
  {
    "type": "function",
    "function": {
      "name": "<tool.name>",
      "description": "<tool.description>",
      "parameters": "<tool.inputSchema>"
    }
  }
  ```
- **Streaming**: Set `stream: true` in the API request. Parse SSE chunks from DeepSeek, extract `choices[0].delta.content` for text tokens and `choices[0].delta.tool_calls` for tool call fragments.
- **Tool call assembly**: DeepSeek streams tool calls in fragments. Accumulate fragments by `tool_calls[i].index`, merging `function.name` and `function.arguments` strings. When the stream ends with `finish_reason: "tool_calls"`, parse the assembled arguments JSON and dispatch to the MCP client.
- **Conversation history format**: Maintain the standard OpenAI message format:
  - `{ role: "user", content: "..." }`
  - `{ role: "assistant", content: "...", tool_calls: [...] }` (when assistant invoked tools)
  - `{ role: "tool", tool_call_id: "...", content: "..." }` (tool results)
  - `{ role: "assistant", content: "..." }` (final answer)

### Conversation History Management

- Store in-memory as `Map<string, { messages: Array, lastActive: Date }>`.
- On each request, update `lastActive`.
- Run a cleanup interval every 30 minutes: evict sessions idle for more than 2 hours.
- Limit conversation history to the last 50 messages per session. When the limit is reached, keep the system message and the most recent 50 messages, discarding the oldest.

---

## Frontend (`public/`)

### Theme & Styling (`style.css`)

- **Dark theme** with the following colour palette:
  - Background: `#0d1117`
  - Panel/card background: `#161b22`
  - Border/separator: `#30363d`
  - Primary text: `#e6edf3`
  - Secondary text: `#8b949e`
  - Accent (user message bubble): `#1f6feb`
  - Accent (assistant message bubble): `#1c2333`
  - Success/active: `#3fb950`
  - Error: `#f85149`
- Font: `'Segoe UI', system-ui, -apple-system, sans-serif`, base size 14px.
- Layout: Two-column — narrow left sidebar (280px) for file management, main area for chat.
- Chat messages area should scroll to bottom on new messages.
- Input bar pinned to the bottom of the chat area: a textarea (auto-growing, max 5 lines) + a send button.
- File upload: a dropzone area at the top of the sidebar. Drag-and-drop overlay appears when files are dragged over the page. Also a click-to-browse fallback.
- Responsive: On viewports < 768px, sidebar collapses to a toggle.

### Frontend Logic (`app.js`)

#### Session Management

- Generate a `sessionId` (UUID v4) on first load and store it in `sessionStorage`.
- Use it for all `/api/chat` requests.

#### File Upload Flow

1. On file drop or selection, POST to `/api/upload` with FormData.
2. Show an upload progress indicator.
3. On success, add the file to the sidebar list.
4. Automatically send a system-like chat message: "I've uploaded **{filename}**. Can you summarise its structure?" — this triggers the LLM to call `get_sheet_summary`.

#### Chat Flow

1. User types a message and presses Enter (or clicks Send).
2. Append the user message bubble to the chat area.
3. POST to `/api/chat` with the message and sessionId.
4. Open an SSE connection on the response.
5. On `type: "token"` events, append to the current assistant message bubble, rendering incrementally. Parse markdown in the accumulated text for display (use a lightweight markdown renderer — either a simple regex-based one or import `marked` from CDN).
6. On `type: "status"` events, show a status pill/badge (e.g., "Reading Excel data...") below the assistant's in-progress message.
7. On `type: "done"`, finalise the message bubble and re-enable the input.
8. On error, display an error message in the chat.

#### Markdown Rendering

- Render assistant messages with basic markdown: **bold**, *italic*, `code`, code blocks with syntax highlighting (use `<pre><code>` blocks), tables, and lists.
- Tables should be styled to match the dark theme with alternating row colours.

#### File Sidebar

- List all uploaded files (fetched from `/api/files` on load).
- Each file entry shows: filename, a delete button (red X icon).
- Delete calls `DELETE /api/files/:filename` and removes from the list.

#### Clear Chat

- A "New Chat" button in the header.
- Calls `DELETE /api/session/:sessionId`, clears the chat area, generates a new sessionId.

---

## Configuration & Environment

Create a `.env.example` file documenting all environment variables:

```env
# Required
DEEPSEEK_API_KEY=sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

# Optional (with defaults)
DEEPSEEK_BASE_URL=https://api.deepseek.com
DEEPSEEK_MODEL=deepseek-chat
PORT=3000
MAX_FILE_SIZE_MB=50
SESSION_TTL_MINUTES=120
MAX_HISTORY_MESSAGES=50
MAX_TOOL_ITERATIONS=10
```

Use `dotenv` to load `.env` at startup.

---

## Error Handling & Edge Cases

1. **Large files**: If a sheet has more than 10,000 rows, `read_sheet_data` and `query_sheet` must respect `max_rows` to avoid flooding the LLM context. The `get_sheet_summary` tool should always return just the summary, never all rows.
2. **Corrupted files**: SheetJS may throw on corrupted files. Catch and return a user-friendly error.
3. **Tool call loops**: Cap the tool-call loop at `MAX_TOOL_ITERATIONS` (default 10). If exceeded, return a message to the user: "I've reached the maximum number of data lookups for this query. Please try a more specific question."
4. **DeepSeek API errors**: On 429 (rate limit), wait and retry once after 2 seconds. On 5xx, retry once. On persistent failure, stream an error message to the user.
5. **MCP server crash**: Detect child process exit, attempt restart, queue pending tool calls and retry after reconnection. If restart fails 3 times, return an error to the user.
6. **File type validation**: Reject non-Excel/CSV files at upload time with a clear error message.
7. **Path traversal**: In the MCP server, resolve file paths and verify they are under the `uploads/` directory before reading.
8. **Empty sheets**: If a sheet has no data rows, return a clear message ("Sheet 'X' is empty") rather than an empty array.

---

## Startup Sequence

1. Load environment variables from `.env`.
2. Validate `DEEPSEEK_API_KEY` is set.
3. Ensure `uploads/` directory exists (create if not).
4. Spawn the MCP server child process (`mcp-server.js`).
5. Wait for MCP client connection and list tools.
6. Start the Express server.
7. Log: `Excel Chat server running on http://localhost:3000`

---

## Dependencies (`package.json`)

```json
{
  "name": "excel-chat",
  "version": "1.0.0",
  "type": "module",
  "scripts": {
    "start": "node server.js",
    "dev": "node --watch server.js"
  },
  "dependencies": {
    "@modelcontextprotocol/sdk": "^1.12.1",
    "dotenv": "^16.4.7",
    "express": "^4.21.2",
    "multer": "^1.4.5-lts.1",
    "xlsx": "^0.18.5"
  }
}
```

> Note: No separate HTTP client library is needed. Use the built-in `fetch` (Node 20+) for DeepSeek API calls.

---

## Testing Checklist

After implementation, verify the following:

- [ ] Upload a `.xlsx` file → appears in sidebar, LLM auto-summarises structure.
- [ ] Ask "What files do I have?" → LLM calls `list_workbooks` and lists them.
- [ ] Ask "What sheets are in sales.xlsx?" → LLM calls `get_sheet_names`.
- [ ] Ask "Show me the first 10 rows of Sheet1" → LLM calls `read_sheet_data` with `max_rows: 10`.
- [ ] Ask "What's the total revenue?" → LLM calls `query_sheet` with a sum aggregation.
- [ ] Ask "Show me all rows where Region is 'Asia'" → LLM calls `query_sheet` with a filter.
- [ ] Ask "Group sales by product category and show the average price" → LLM uses `group_by` + aggregation.
- [ ] Upload a `.csv` file → same flow works.
- [ ] Upload a corrupt file → clear error message returned.
- [ ] Delete a file from sidebar → file removed, confirmed.
- [ ] Start a new chat → history cleared, new session.
- [ ] LLM streams responses token-by-token with visible typing effect.
- [ ] Tool-call status indicator appears while MCP is queried.
- [ ] Markdown tables in assistant responses render properly.
