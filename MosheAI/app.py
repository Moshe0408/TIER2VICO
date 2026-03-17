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

# ── Config file for API key ──────────────
CONFIG_FILE = Path(__file__).parent / "config.json"

def load_config():
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}

def save_config(cfg: dict):
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")

# Load saved API key into env if not already set
_cfg = load_config()
if not os.environ.get("ANTHROPIC_API_KEY") and _cfg.get("api_key"):
    os.environ["ANTHROPIC_API_KEY"] = _cfg["api_key"]

agent = MosheAIAgent()


def _reinit_agent():
    global agent
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


@app.route("/api/settings", methods=["GET", "POST"])
@login_required
def settings():
    if request.method == "GET":
        key = os.environ.get("ANTHROPIC_API_KEY", "")
        masked = ("sk-ant-..." + key[-6:]) if len(key) > 10 else ""
        return jsonify({"api_key_set": bool(key), "masked": masked})

    data = request.get_json(force=True)
    key  = (data.get("api_key") or "").strip()
    if not key:
        return jsonify({"error": "מפתח ריק"}), 400
    if not key.startswith("sk-"):
        return jsonify({"error": "מפתח לא תקין (חייב להתחיל ב-sk-)"}), 400

    os.environ["ANTHROPIC_API_KEY"] = key
    cfg = load_config()
    cfg["api_key"] = key
    save_config(cfg)
    _reinit_agent()
    return jsonify({"ok": True, "message": "✅ API Key נשמר והסוכן אותחל מחדש!"})


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
    import sys
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    print("\n" + "=" * 50)
    print("  MosheAI  -  Ready!")
    print("=" * 50)
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("  WARNING: ANTHROPIC_API_KEY not set!")
        print("     set ANTHROPIC_API_KEY=sk-ant-...")
    else:
        print("  API Key: OK")
    print(f"  Outputs: {OUTPUT_DIR}")
    print("  URL: http://localhost:5000")
    print("  User: Moshei1 | Pass: Admin2026")
    print("=" * 50 + "\n")
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)
