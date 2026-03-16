import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  ListToolsRequestSchema,
  CallToolRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import * as XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const UPLOADS_DIR = path.resolve(__dirname, 'uploads');

// Path traversal guard
function safePath(filename) {
  const resolved = path.resolve(UPLOADS_DIR, filename);
  if (!resolved.startsWith(UPLOADS_DIR + path.sep) && resolved !== UPLOADS_DIR) {
    throw new Error('Invalid filename: path traversal detected');
  }
  return resolved;
}

function textContent(data) {
  return { content: [{ type: 'text', text: typeof data === 'string' ? data : JSON.stringify(data) }] };
}

function errorContent(msg) {
  return { content: [{ type: 'text', text: `Error: ${msg}` }], isError: true };
}

// Infer predominant data type for a column
function inferType(values) {
  const counts = { number: 0, string: 0, boolean: 0, date: 0, empty: 0 };
  for (const v of values) {
    if (v === null || v === undefined || v === '') {
      counts.empty++;
    } else if (typeof v === 'boolean') {
      counts.boolean++;
    } else if (typeof v === 'number') {
      counts.number++;
    } else if (v instanceof Date) {
      counts.date++;
    } else if (typeof v === 'string') {
      // Check if it looks like a date
      if (!isNaN(Date.parse(v)) && /\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(v)) {
        counts.date++;
      } else if (!isNaN(Number(v)) && v.trim() !== '') {
        counts.number++;
      } else {
        counts.string++;
      }
    } else {
      counts.string++;
    }
  }
  let max = 'string';
  let maxCount = -1;
  for (const [type, count] of Object.entries(counts)) {
    if (type !== 'empty' && count > maxCount) {
      maxCount = count;
      max = type;
    }
  }
  return max;
}

// Apply a single filter condition to a row
function applyFilter(row, filter) {
  const val = row[filter.column];
  const target = filter.value;

  switch (filter.operator) {
    case 'eq': return val == target;
    case 'neq': return val != target;
    case 'gt': return Number(val) > Number(target);
    case 'gte': return Number(val) >= Number(target);
    case 'lt': return Number(val) < Number(target);
    case 'lte': return Number(val) <= Number(target);
    case 'contains': return String(val).toLowerCase().includes(String(target).toLowerCase());
    case 'startsWith': return String(val).toLowerCase().startsWith(String(target).toLowerCase());
    case 'endsWith': return String(val).toLowerCase().endsWith(String(target).toLowerCase());
    default: return true;
  }
}

// Compute an aggregation over an array of numeric values
function computeAggregation(values, func) {
  const nums = values.map(Number).filter(n => !isNaN(n));
  if (nums.length === 0) return null;
  switch (func) {
    case 'sum': return nums.reduce((a, b) => a + b, 0);
    case 'average': return nums.reduce((a, b) => a + b, 0) / nums.length;
    case 'count': return nums.length;
    case 'min': return Math.min(...nums);
    case 'max': return Math.max(...nums);
    default: return null;
  }
}

const server = new Server(
  { name: 'excel-mcp-server', version: '1.0.0' },
  { capabilities: { tools: {} } }
);

const tools = [
  {
    name: 'list_workbooks',
    description: 'List all uploaded Excel workbooks available for querying.',
    inputSchema: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'get_sheet_names',
    description: 'Get the names of all sheets in a given workbook.',
    inputSchema: {
      type: 'object',
      properties: {
        filename: { type: 'string', description: 'Name of the uploaded workbook file' },
      },
      required: ['filename'],
    },
  },
  {
    name: 'read_sheet_data',
    description: 'Read data from a specific sheet. Returns the first N rows as JSON (default 100). Supports optional column filtering.',
    inputSchema: {
      type: 'object',
      properties: {
        filename: { type: 'string', description: 'Name of the uploaded workbook file' },
        sheet: { type: 'string', description: 'Sheet name. Omit to use the first sheet.' },
        max_rows: { type: 'number', description: 'Maximum rows to return. Default 100.' },
        columns: {
          type: 'array',
          items: { type: 'string' },
          description: 'Optional list of column headers to include. Omit to return all columns.',
        },
      },
      required: ['filename'],
    },
  },
  {
    name: 'get_sheet_summary',
    description: 'Get a structural summary of a sheet: column headers, row count, data types per column, and a sample of the first 5 rows.',
    inputSchema: {
      type: 'object',
      properties: {
        filename: { type: 'string', description: 'Name of the uploaded workbook file' },
        sheet: { type: 'string', description: 'Sheet name. Omit to use the first sheet.' },
      },
      required: ['filename'],
    },
  },
  {
    name: 'query_sheet',
    description: 'Filter and aggregate data from a sheet. Supports filtering rows by column value conditions and basic aggregations (sum, average, count, min, max) on numeric columns.',
    inputSchema: {
      type: 'object',
      properties: {
        filename: { type: 'string', description: 'Name of the uploaded workbook file' },
        sheet: { type: 'string', description: 'Sheet name. Omit to use the first sheet.' },
        filters: {
          type: 'array',
          description: 'Array of filter conditions. Each is { column, operator, value }. Operators: eq, neq, gt, gte, lt, lte, contains, startsWith, endsWith.',
          items: {
            type: 'object',
            properties: {
              column: { type: 'string' },
              operator: { type: 'string', enum: ['eq', 'neq', 'gt', 'gte', 'lt', 'lte', 'contains', 'startsWith', 'endsWith'] },
              value: { type: ['string', 'number', 'boolean'] },
            },
            required: ['column', 'operator', 'value'],
          },
        },
        aggregations: {
          type: 'array',
          description: 'Array of aggregation operations. Each is { column, function }.',
          items: {
            type: 'object',
            properties: {
              column: { type: 'string' },
              function: { type: 'string', enum: ['sum', 'average', 'count', 'min', 'max'] },
            },
            required: ['column', 'function'],
          },
        },
        group_by: { type: 'string', description: 'Optional column name to group results by before aggregation.' },
        sort_by: { type: 'string', description: 'Column name to sort results by.' },
        sort_order: { type: 'string', enum: ['asc', 'desc'], description: 'Sort direction. Default asc.' },
        max_rows: { type: 'number', description: 'Maximum result rows to return. Default 50.' },
      },
      required: ['filename'],
    },
  },
];

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return { tools };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    switch (name) {
      case 'list_workbooks': {
        if (!fs.existsSync(UPLOADS_DIR)) {
          return textContent([]);
        }
        const files = fs.readdirSync(UPLOADS_DIR).filter(f =>
          /\.(xlsx|xls|csv)$/i.test(f)
        );
        return textContent(files);
      }

      case 'get_sheet_names': {
        const filepath = safePath(args.filename);
        if (!fs.existsSync(filepath)) {
          return errorContent(`File not found: ${args.filename}`);
        }
        const workbook = XLSX.readFile(filepath);
        return textContent(workbook.SheetNames);
      }

      case 'read_sheet_data': {
        const filepath = safePath(args.filename);
        if (!fs.existsSync(filepath)) {
          return errorContent(`File not found: ${args.filename}`);
        }
        const workbook = XLSX.readFile(filepath);
        const sheetName = args.sheet || workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
          return errorContent(`Sheet not found: ${sheetName}. Available sheets: ${workbook.SheetNames.join(', ')}`);
        }
        let rows = XLSX.utils.sheet_to_json(worksheet);
        if (rows.length === 0) {
          return textContent(`Sheet '${sheetName}' is empty`);
        }
        if (args.columns && args.columns.length > 0) {
          rows = rows.map(row => {
            const filtered = {};
            for (const col of args.columns) {
              if (col in row) filtered[col] = row[col];
            }
            return filtered;
          });
        }
        const maxRows = args.max_rows || 100;
        rows = rows.slice(0, maxRows);
        return textContent(rows);
      }

      case 'get_sheet_summary': {
        const filepath = safePath(args.filename);
        if (!fs.existsSync(filepath)) {
          return errorContent(`File not found: ${args.filename}`);
        }
        const workbook = XLSX.readFile(filepath);
        const sheetName = args.sheet || workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
          return errorContent(`Sheet not found: ${sheetName}. Available sheets: ${workbook.SheetNames.join(', ')}`);
        }
        const rows = XLSX.utils.sheet_to_json(worksheet);
        if (rows.length === 0) {
          return textContent({ filename: args.filename, sheet: sheetName, total_rows: 0, columns: [], sample_rows: [] });
        }
        const columnNames = Object.keys(rows[0]);
        const columns = columnNames.map(col => {
          const values = rows.map(r => r[col]);
          const nonEmpty = values.filter(v => v !== null && v !== undefined && v !== '').length;
          return { name: col, type: inferType(values), non_empty: nonEmpty };
        });
        const sampleRows = rows.slice(0, 5);
        return textContent({
          filename: args.filename,
          sheet: sheetName,
          total_rows: rows.length,
          columns,
          sample_rows: sampleRows,
        });
      }

      case 'query_sheet': {
        const filepath = safePath(args.filename);
        if (!fs.existsSync(filepath)) {
          return errorContent(`File not found: ${args.filename}`);
        }
        const workbook = XLSX.readFile(filepath);
        const sheetName = args.sheet || workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
          return errorContent(`Sheet not found: ${sheetName}. Available sheets: ${workbook.SheetNames.join(', ')}`);
        }
        let rows = XLSX.utils.sheet_to_json(worksheet);
        if (rows.length === 0) {
          return textContent(`Sheet '${sheetName}' is empty`);
        }

        // Apply filters
        if (args.filters && args.filters.length > 0) {
          for (const filter of args.filters) {
            rows = rows.filter(row => applyFilter(row, filter));
          }
        }

        const maxRows = args.max_rows || 50;

        // Group by
        if (args.group_by) {
          const groups = {};
          for (const row of rows) {
            const key = String(row[args.group_by] ?? '(empty)');
            if (!groups[key]) groups[key] = [];
            groups[key].push(row);
          }

          if (args.aggregations && args.aggregations.length > 0) {
            let results = Object.entries(groups).map(([groupKey, groupRows]) => {
              const result = { [args.group_by]: groupKey };
              for (const agg of args.aggregations) {
                const values = groupRows.map(r => r[agg.column]);
                result[`${agg.function}_${agg.column}`] = computeAggregation(values, agg.function);
              }
              return result;
            });

            // Sort
            if (args.sort_by) {
              const order = args.sort_order === 'desc' ? -1 : 1;
              results.sort((a, b) => {
                const av = a[args.sort_by] ?? '';
                const bv = b[args.sort_by] ?? '';
                if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * order;
                return String(av).localeCompare(String(bv)) * order;
              });
            }

            return textContent(results.slice(0, maxRows));
          }

          // No aggregations but group_by: return groups
          let results = Object.entries(groups).map(([key, groupRows]) => ({
            [args.group_by]: key,
            count: groupRows.length,
            rows: groupRows.slice(0, 5),
          }));

          if (args.sort_by) {
            const order = args.sort_order === 'desc' ? -1 : 1;
            results.sort((a, b) => {
              const av = a[args.sort_by] ?? '';
              const bv = b[args.sort_by] ?? '';
              if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * order;
              return String(av).localeCompare(String(bv)) * order;
            });
          }

          return textContent(results.slice(0, maxRows));
        }

        // No group_by
        if (args.aggregations && args.aggregations.length > 0) {
          const result = {};
          for (const agg of args.aggregations) {
            const values = rows.map(r => r[agg.column]);
            result[`${agg.function}_${agg.column}`] = computeAggregation(values, agg.function);
          }
          result.filtered_row_count = rows.length;
          return textContent(result);
        }

        // Sort
        if (args.sort_by) {
          const order = args.sort_order === 'desc' ? -1 : 1;
          rows.sort((a, b) => {
            const av = a[args.sort_by] ?? '';
            const bv = b[args.sort_by] ?? '';
            if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * order;
            return String(av).localeCompare(String(bv)) * order;
          });
        }

        return textContent(rows.slice(0, maxRows));
      }

      default:
        return errorContent(`Unknown tool: ${name}`);
    }
  } catch (err) {
    return errorContent(err.message || String(err));
  }
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch(err => {
  console.error('MCP server error:', err);
  process.exit(1);
});
