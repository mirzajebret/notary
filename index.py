from flask import Flask, request, jsonify, render_template_string, send_file
import os
import sqlite3
from docx import Document
import pandas as pd
import webbrowser

app = Flask(__name__)
DATABASE = 'file_index.db'

def init_db():
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS files
                 (id INTEGER PRIMARY KEY, filename TEXT, content TEXT, path TEXT, category TEXT)''')
    conn.commit()
    conn.close()

def index_files(base_directory):
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    for root, dirs, files in os.walk(base_directory):
        for file in files:
            # Abaikan file sementara dan file non-Word/Excel
            if file.startswith("~$"):
                continue
            if file.endswith(".docx"):
                try:
                    doc = Document(os.path.join(root, file))
                    text = "\n".join([para.text for para in doc.paragraphs])
                    category = os.path.basename(root)
                    c.execute("INSERT INTO files (filename, content, path, category) VALUES (?, ?, ?, ?)",
                              (file, text, os.path.join(root, file), category))
                except Exception as e:
                    print(f"Error processing {file}: {e}")
            elif file.endswith(".xlsx"):
                try:
                    df = pd.read_excel(os.path.join(root, file))
                    text = df.to_string()
                    category = os.path.basename(root)
                    c.execute("INSERT INTO files (filename, content, path, category) VALUES (?, ?, ?, ?)",
                              (file, text, os.path.join(root, file), category))
                except Exception as e:
                    print(f"Error processing {file}: {e}")
    conn.commit()
    conn.close()

@app.route('/')
def index():
    return render_template_string('''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Superapp Kenotariatan</title>
        <link href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" rel="stylesheet">
    </head>
    <body>
        <div class="container mt-5">
            <h1>Superapp Kenotariatan</h1>
            <form id="searchForm">
                <div class="form-group">
                    <input type="text" class="form-control" id="searchQuery" placeholder="Masukkan keyword">
                </div>
                <button type="submit" class="btn btn-primary">Cari</button>
            </form>
            <div id="results" class="mt-4"></div>
        </div>
        <script>
            document.getElementById('searchForm').addEventListener('submit', function(event) {
                event.preventDefault();
                let query = document.getElementById('searchQuery').value;
                fetch(`/search?q=${query}`)
                    .then(response => response.json())
                    .then(data => {
                        let resultsDiv = document.getElementById('results');
                        resultsDiv.innerHTML = '';
                        if (data.length === 0) {
                            resultsDiv.innerHTML = '<p>No results found</p>';
                        } else {
                            let list = document.createElement('ul');
                            list.classList.add('list-group');
                            data.forEach(item => {
                                let listItem = document.createElement('li');
                                listItem.classList.add('list-group-item');
                                
                                let link = document.createElement('a');
                                link.href = `/open_file?path=${encodeURIComponent(item[1])}`;
                                link.textContent = `${item[0]} - ${item[2]}`;
                                link.target = '_blank';

                                listItem.appendChild(link);
                                list.appendChild(listItem);
                            });
                            resultsDiv.appendChild(list);
                        }
                    });
            });
        </script>
    </body>
    </html>
    ''')

@app.route('/search', methods=['GET'])
def search():
    query = request.args.get('q')
    conn = sqlite3.connect(DATABASE)
    c = conn.cursor()
    c.execute("SELECT filename, path, category FROM files WHERE content LIKE ?", ('%' + query + '%',))
    results = c.fetchall()
    conn.close()
    return jsonify(results)

@app.route('/open_file', methods=['GET'])
def open_file():
    file_path = request.args.get('path')
    if os.path.exists(file_path):
        return send_file(file_path)
    else:
        return "File not found", 404

if __name__ == '__main__':
    init_db()
    index_files('E:/Notaris/superpp/Notaris&PPAT')  # Ganti dengan path ke folder yang ditunjukkan pada gambar
    app.run(debug=True)
