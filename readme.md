# Excel Chat — LLM-Powered Excel Query Interface with MCP

A web-based chat application that lets you upload Excel files and query their contents using natural language. The system uses **DeepSeek LLM** as the reasoning engine, connected to an **MCP (Model Context Protocol) server** that reads and parses Excel files via **SheetJS**.

## Architecture

```
Browser (Chat UI)  →  Chat Server (Express + MCP Client)  →  MCP Server (SheetJS)
       ↑                        ↕                                    ↕
    SSE stream          DeepSeek API (LLM)                   Excel/CSV files
```

## Prerequisites

- **Node.js 20+**
- **DeepSeek API key** — get one at [platform.deepseek.com](https://platform.deepseek.com/)

## Installation

```bash
git clone <repo-url>
cd excel-chat
npm install
```

## Configuration

Copy the example environment file and set your API key:

```bash
cp .env.example .env
```

Edit `.env`:

```env
# Required
DEEPSEEK_API_KEY=sk-your-api-key-here

# Optional (defaults shown)
DEEPSEEK_BASE_URL=https://api.deepseek.com
DEEPSEEK_MODEL=deepseek-chat
PORT=3000
MAX_FILE_SIZE_MB=50
SESSION_TTL_MINUTES=120
MAX_HISTORY_MESSAGES=50
MAX_TOOL_ITERATIONS=10
```

## Running

```bash
npm start
```

Then open [http://localhost:3000](http://localhost:3000) in your browser.

For development with auto-restart on file changes:

```bash
npm run dev
```

## Usage

1. **Upload** — Drag and drop an `.xlsx`, `.xls`, or `.csv` file onto the sidebar (or click to browse). The assistant will automatically summarise the file's structure.
2. **Ask questions** — Type natural language queries in the chat input:
   - "What files do I have?"
   - "Show me the first 10 rows of Sheet1"
   - "What's the total revenue?"
   - "Show all rows where Region is 'Asia'"
   - "Group sales by product category and show the average price"
3. **Manage files** — Delete files from the sidebar using the × button.
4. **New chat** — Click "New Chat" to clear conversation history and start fresh.

## MCP Tools

The MCP server exposes five tools that the LLM can call:

| Tool | Description |
|------|-------------|
| `list_workbooks` | List all uploaded Excel/CSV files |
| `get_sheet_names` | Get sheet names in a workbook |
| `read_sheet_data` | Read rows from a sheet (with optional column filtering and row limit) |
| `get_sheet_summary` | Get column headers, row count, data types, and sample rows |
| `query_sheet` | Filter, aggregate, group, and sort data |

## Project Structure

```
├── server.js          # Express chat server + MCP client
├── mcp-server.js      # MCP server process (SheetJS tools)
├── package.json
├── .env.example
├── public/
│   ├── index.html     # Chat UI
│   ├── style.css      # Dark-themed styling
│   └── app.js         # Frontend logic
└── uploads/           # Uploaded files (created automatically)
```

## License

MIT
