"""MosheAI - Flask Web Server with Login"""

import json
import os
import functools
from pathlib import Path
from flask import (Flask, render_template, request, Response,
                   jsonify, send_file, session, redirect, url_for)

from engine.agent import MosheAIAgent
from engine.tools import list_outputs, OUTPUT_DIR

app = Flask(__name__)
app.secret_key = "mosheai-secret-2026-xk9"
app.config["JSON_AS_ASCII"] = False

# ── אישורים ──────────────────────────────
CREDENTIALS = {
    "Moshei1": "Admin2026"
}

agent = MosheAIAgent()


# ── decorator הגנה ────────────────────────
def login_required(f):
    @functools.wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper


# ── דפים ──────────────────────────────────
@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        if CREDENTIALS.get(username) == password:
            session["logged_in"] = True
            session["username"]  = username
            return redirect(url_for("index"))
        error = "שם משתמש או סיסמה שגויים"
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def index():
    api_key_set = bool(os.environ.get("ANTHROPIC_API_KEY"))
    username    = session.get("username", "")
    return render_template("index.html", api_key_set=api_key_set, username=username)


# ── API ───────────────────────────────────
@app.route("/api/chat", methods=["POST"])
@login_required
def chat():
    data    = request.get_json(force=True)
    message = (data.get("message") or "").strip()
    if not message:
        return jsonify({"error": "הודעה ריקה"}), 400

    def generate():
        for chunk in agent.stream_response(message):
            yield f"data: {json.dumps(chunk, ensure_ascii=False)}\n\n"

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"}
    )


@app.route("/api/memory")
@login_required
def get_memory():
    return jsonify(agent.get_memory_summary())


@app.route("/api/files")
@login_required
def get_files():
    return jsonify(list_outputs())


@app.route("/api/file/<path:filename>")
@login_required
def download_file(filename):
    path = OUTPUT_DIR / Path(filename).name
    if not path.exists():
        return jsonify({"error": "קובץ לא נמצא"}), 404
    return send_file(str(path), as_attachment=True, download_name=path.name)


@app.route("/api/file/preview/<path:filename>")
@login_required
def preview_file(filename):
    path = OUTPUT_DIR / Path(filename).name
    if not path.exists():
        return jsonify({"error": "קובץ לא נמצא"}), 404
    if path.suffix.lower() == ".png":
        return send_file(str(path), mimetype="image/png")
    return jsonify({"error": "תצוגה מקדימה זמינה רק לתמונות"}), 400


if __name__ == "__main__":
    print("\n" + "═" * 50)
    print("  🤖  MosheAI  -  מוכן לעבודה!")
    print("═" * 50)
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("  ⚠️  ANTHROPIC_API_KEY לא מוגדר!")
        print("     הגדר: set ANTHROPIC_API_KEY=sk-ant-...")
    else:
        print("  ✅  API Key מוגדר")
    print(f"  📁  פלטים: {OUTPUT_DIR}")
    print("  🌐  פתח: http://localhost:5000")
    print("  🔐  משתמש: Moshei1 | סיסמה: Admin2026")
    print("═" * 50 + "\n")
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)
