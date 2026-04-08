from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse
from pathlib import Path
import pandas as pd
import re
import html
import uvicorn

app = FastAPI()
DATA_FILE = Path(__file__).parent / "Dummy Data.xlsx"

try:
    df = pd.read_excel(DATA_FILE, sheet_name=0, dtype=str)
except Exception as exc:
    raise RuntimeError(f"Failed to load Excel file: {DATA_FILE}\n{exc}")

df = df.fillna("").astype(str)
df.columns = [col.strip() for col in df.columns]
search_frame = df.apply(lambda row: " ".join(row.values.astype(str)).lower(), axis=1)

STOPWORDS = {
    'show', 'find', 'list', 'all', 'employees', 'employee', 'in', 'the', 'a', 'of', 'with',
    'by', 'for', 'from', 'how', 'many', 'is', 'are', 'to', 'and', 'or', 'on', 'what',
    'which', 'where', 'who', 'please'
}

COLUMN_SYNONYMS = {
    'Ukuran Seragam': ['ukuran seragam', 'seragam', 'size', 'ukuran'],
    'Lokasi': ['lokasi', 'location'],
    'Department': ['department', 'dept'],
    'Salary': ['salary', 'gaji', 'upah'],
    'Umur': ['umur', 'age'],
    'Jumlah Anak': ['jumlah anak', 'anak'],
    'Jabatan': ['jabatan', 'position', 'posisi'],
    'Divisi': ['divisi', 'division'],
    'Nama Asuransi': ['asuransi', 'insurance'],
    'Tempat Lahir': ['tempat lahir', 'birthplace'],
    'Tanggal Lahir': ['tanggal lahir', 'birth date', 'dob'],
    'Status Pernikahan': ['status pernikahan', 'marital status'],
}

TABLE_ROW_LIMIT = 100

@app.get("/", response_class=HTMLResponse)
async def home():
    return """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>create by thm</title>
    <style>
        :root {
            --bg: #071a3d;
            --panel: #0f2a5b;
            --card: #132f67;
            --text: #f3f6ff;
            --muted: #99a7d6;
            --accent: #ff8c42;
            --accent-soft: #ffb57f;
            --border: rgba(255, 140, 66, 0.15);
        }

        body { font-family: Arial, sans-serif; background: radial-gradient(circle at top, #0d264e 0%, var(--bg) 65%); color: var(--text); margin: 0; padding: 0; }
        .container { max-width: 900px; margin: 2rem auto; background: var(--panel); border-radius: 18px; box-shadow: 0 30px 80px rgba(0,0,0,0.35); padding: 2rem; border: 1px solid rgba(255,255,255,0.08); }
        h1 { margin-top: 0; color: var(--text); }
        #chatWindow { max-height: 560px; overflow-y: auto; border: 1px solid rgba(255,255,255,0.12); border-radius: 14px; padding: 1rem; background: rgba(15, 41, 91, 0.9); }
        .message { margin-bottom: 1rem; display: flex; }
        .message.user { justify-content: flex-end; }
        .message.bot { justify-content: flex-start; }
        .message.user .bubble { background: linear-gradient(135deg, var(--accent), #ffa65d); color: #081520; margin-left: auto; box-shadow: 0 12px 30px rgba(255,140,66,0.22); }
        .message.bot .bubble { background: rgba(22, 56, 116, 0.95); color: var(--text); margin-right: auto; border: 1px solid rgba(255,255,255,0.08); }
        .bubble { padding: 0.95rem 1rem; border-radius: 20px; max-width: 100%; line-height: 1.6; white-space: pre-wrap; }
        form { display: flex; margin-top: 1rem; }
        input[type=text] { flex: 1; padding: 0.95rem 1rem; border: 1px solid rgba(255,255,255,0.12); border-radius: 999px; outline: none; font-size: 1rem; background: rgba(255,255,255,0.06); color: var(--text); }
        input[type=text]::placeholder { color: var(--muted); }
        button { margin-left: 0.75rem; padding: 0 1.2rem; border: none; border-radius: 999px; background: var(--accent); color: #081520; font-size: 1rem; cursor: pointer; box-shadow: 0 12px 25px rgba(255,140,66,0.25); }
        button:hover { background: #ff9c5f; }
        button:disabled { opacity: 0.6; cursor: not-allowed; }
        table.chat-table { width: 100%; border-collapse: collapse; margin-top: 0.75rem; background: rgba(255,255,255,0.05); }
        table.chat-table th, table.chat-table td { border: 1px solid rgba(255,255,255,0.12); padding: 0.65rem 0.85rem; text-align: left; color: var(--text); }
        table.chat-table th { background: rgba(255,140,66,0.15); color: #fff; }
        .table-container { overflow-x: auto; }
        .download-link { display: inline-block; margin-top: 0.75rem; padding: 0.5rem 0.8rem; border-radius: 999px; background: var(--accent); color: #081520; text-decoration: none; }
    </style>
</head>
<body>
    <div class="container">
        <h1>HRIS Chatbot</h1>
        <div id="chatWindow"></div>
        <form id="chatForm">
            <input id="messageInput" type="text" placeholder="Ask the chatbot about the data..." autocomplete="off" />
            <button type="submit">Send</button>
        </form>
    </div>

    <script>
        const chatWindow = document.getElementById('chatWindow');
        const chatForm = document.getElementById('chatForm');
        const messageInput = document.getElementById('messageInput');

        function appendMessage(text, role, htmlContent = null) {
            const wrapper = document.createElement('div');
            wrapper.className = `message ${role}`;
            const bubble = document.createElement('div');
            bubble.className = 'bubble';
            if (htmlContent) {
                bubble.innerHTML = htmlContent;
            } else {
                bubble.textContent = text;
            }
            wrapper.appendChild(bubble);
            chatWindow.appendChild(wrapper);
            chatWindow.scrollTop = chatWindow.scrollHeight;
        }

        function appendDownloadLink(url) {
            const wrapper = document.createElement('div');
            wrapper.className = 'message bot';
            const bubble = document.createElement('div');
            bubble.className = 'bubble';
            bubble.innerHTML = `<a class="download-link" href="${url}" target="_blank">Download Excel file</a>`;
            wrapper.appendChild(bubble);
            chatWindow.appendChild(wrapper);
            chatWindow.scrollTop = chatWindow.scrollHeight;
        }

        async function sendMessage(message) {
            const response = await fetch('/chat', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ message })
            });
            return response.json();
        }

        chatForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const text = messageInput.value.trim();
            if (!text) return;
            appendMessage(text, 'user');
            messageInput.value = '';
            messageInput.disabled = true;
            const data = await sendMessage(text);
            if (data.html) {
                appendMessage(data.reply || 'Here is the result:', 'bot', data.html);
            } else {
                appendMessage(data.reply, 'bot');
            }
            if (data.download_url) {
                appendDownloadLink(data.download_url);
            }
            messageInput.disabled = false;
            messageInput.focus();
        });

        appendMessage('Hello! Ask me about the employee data and I will return results as a table.', 'bot');
    </script>
</body>
</html>
"""

@app.post('/chat')
async def chat(request: Request):
    payload = await request.json()
    user_message = payload.get('message', '').strip()
    if not user_message:
        return JSONResponse({'reply': 'Please send a message to start the chat.'})

    result = answer_question(user_message)
    return result


def normalize_text(value: str) -> str:
    return re.sub(r'\s+', ' ', value.strip().lower())


def parse_query_terms(query: str) -> list[str]:
    tokens = re.findall(r"\w+", query.lower())
    return [token for token in tokens if token not in STOPWORDS and len(token) > 1]


def infer_filter_column(query: str) -> str | None:
    query_lower = query.lower()
    numeric_columns = {
        'salary': 'Salary',
        'gaji': 'Salary',
        'umur': 'Umur',
        'age': 'Umur',
        'jumlah anak': 'Jumlah Anak'
    }
    for key, col in numeric_columns.items():
        if key in query_lower and col in df.columns:
            return col
    for col, syns in COLUMN_SYNONYMS.items():
        for syn in syns:
            if syn in query_lower and col in df.columns:
                return col
    for col in df.columns:
        if col.lower() in query_lower:
            return col
    return None


def parse_string_filter(query: str) -> pd.DataFrame | None:
    query_lower = query.lower()
    candidates = {**COLUMN_SYNONYMS}
    for col in df.columns:
        if col not in candidates:
            candidates[col] = [col.lower()]

    for col, syns in candidates.items():
        if not any(syn in query_lower for syn in syns):
            continue
        values = [str(val).strip() for val in df[col].dropna().unique() if str(val).strip()]
        for value in sorted(values, key=lambda x: -len(x)):
            value_lower = value.lower()
            if re.search(rf'\b{re.escape(value_lower)}\b', query_lower):
                return df[df[col].str.lower() == value_lower]
    return None


def infer_requested_columns(query: str) -> list[str]:
    query_lower = query.lower()
    requested = []
    for col, syns in COLUMN_SYNONYMS.items():
        for syn in syns:
            if syn in query_lower and col in df.columns:
                requested.append(col)
                break

    for col in df.columns:
        if col.lower() in query_lower and col not in requested:
            requested.append(col)

    requested = list(dict.fromkeys(requested))
    if requested:
        if 'Nama' in df.columns and 'Nama' not in requested:
            requested.insert(0, 'Nama')
        return requested

    filter_column = infer_filter_column(query)
    if filter_column and 'Nama' in df.columns:
        return ['Nama', filter_column]

    return list(df.columns)


def select_columns(rows: pd.DataFrame, query: str) -> pd.DataFrame:
    cols = infer_requested_columns(query)
    cols = [col for col in cols if col in rows.columns]
    if cols:
        return rows[cols]
    return rows


def filter_rows(query: str) -> pd.DataFrame:
    terms = parse_query_terms(query)
    if not terms:
        return df.head(TABLE_ROW_LIMIT)

    numeric_match = parse_numeric_filter(query)
    if numeric_match is not None:
        return numeric_match

    string_match = parse_string_filter(query)
    if string_match is not None:
        return string_match.head(TABLE_ROW_LIMIT)

    matched = df[search_frame.apply(lambda row_text: all(term in row_text for term in terms))]
    if not matched.empty:
        return matched.head(TABLE_ROW_LIMIT)

    filtered = df[search_frame.apply(lambda row_text: any(term in row_text for term in terms))]
    return filtered.head(TABLE_ROW_LIMIT)


def parse_numeric_filter(query: str) -> pd.DataFrame | None:
    numeric_columns = {
        'salary': 'Salary',
        'gaji': 'Salary',
        'umur': 'Umur',
        'age': 'Umur',
        'jumlah anak': 'Jumlah Anak'
    }
    query_lower = query.lower()
    for key, column in numeric_columns.items():
        if key in query_lower and column in df.columns:
            number_match = re.search(r'([0-9]+[0-9\.,]*)', query_lower.replace(',', ''))
            if not number_match:
                continue
            value = float(number_match.group(1).replace('.', ''))
            numeric_values = pd.to_numeric(df[column].str.replace(r'[^0-9.-]', '', regex=True), errors='coerce')
            if any(term in query_lower for term in ['above', 'more than', 'lebih dari', '>']):
                return df[numeric_values > value].head(TABLE_ROW_LIMIT)
            if any(term in query_lower for term in ['below', 'less than', 'kurang dari', '<']):
                return df[numeric_values < value].head(TABLE_ROW_LIMIT)
            return df[numeric_values == value].head(TABLE_ROW_LIMIT)
    return None


def rows_to_html(rows: pd.DataFrame) -> str:
    if rows.empty:
        return '<div>Tidak Ada data yang sesuai.</div>'

    headers = ''.join(f'<th>{html.escape(str(col))}</th>' for col in rows.columns)
    body_rows = []
    for _, row in rows.iterrows():
        cells = ''.join(f'<td>{html.escape(str(cell))}</td>' for cell in row.values)
        body_rows.append(f'<tr>{cells}</tr>')
    return (
        '<div class="table-container">'
        '<table class="chat-table">'
        f'<thead><tr>{headers}</tr></thead>'
        f'<tbody>{"".join(body_rows)}</tbody>'
        '</table>'
        '</div>'
    )


def answer_question(query: str) -> dict:
    query_lower = query.lower()
    if any(word in query_lower for word in ['hello', 'hi', 'hey', 'halo']):
        return {'reply': 'Hello! Ask me questions about the employee data and I will return a table with the matching rows.'}
    if 'help' in query_lower:
        return {'reply': 'Ask about employees, locations, departments, salary, age, or anything in the Excel sheet. I will return a table of matching rows.'}
    if 'thank' in query_lower:
        return {'reply': 'You are welcome! Ask another question about the data if you want more results.'}

    filtered_rows = filter_rows(query)
    if filtered_rows.empty:
        return {'reply': 'No matching rows were found for that question.', 'html': rows_to_html(filtered_rows)}

    selected_rows = select_columns(filtered_rows, query)
    return {
        'reply': f'Found {len(filtered_rows)} matching rows.',
        'html': rows_to_html(selected_rows)
    }


if __name__ == '__main__':
    uvicorn.run('main:app', host='0.0.0.0', port=8000, reload=True)
