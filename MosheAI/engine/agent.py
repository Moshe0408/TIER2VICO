"""ליבת הסוכן MosheAI - מתחבר ל-Claude API ומייצר תוצאות"""

import json
import anthropic

from . import memory as mem_module
from .tools import TOOLS_SCHEMA, run_tool

MODEL = "claude-opus-4-6"

SYSTEM = """אתה MosheAI - עוזר AI מקצועי חכם שעובד בעברית.

יכולותיך:
• יצירת מצגות PowerPoint מקצועיות ויפות
• כתיבת דוחות Word עם טבלאות ומבנה ברור
• יצירת גרפים וסטטיסטיקה מרשימים
• ניתוח נתונים והצגתם בצורה ברורה

עקרונות עבודה:
• תמיד ייצר תוצאות מוחשיות (קבצים אמיתיים)
• השתמש בכלים הזמינים לך
• ענה תמיד בעברית
• שאף לאיכות גבוהה ומקצועית
• למד מהטעויות הקודמות שלך
"""


class MosheAIAgent:
    def __init__(self):
        self.memory = mem_module.load()
        self.client = anthropic.Anthropic()

    def stream_response(self, user_message: str):
        """
        Generator — מניב dict-ים בפורמט SSE:
          {"type": "thinking",  "content": "..."}
          {"type": "text",      "content": "..."}
          {"type": "tool_start","tool": "...", "label": "..."}
          {"type": "tool_done", "tool": "...", "result": {...}}
          {"type": "done",      "outputs": [...]}
          {"type": "error",     "content": "..."}
        """
        messages    = [{"role": "user", "content": user_message}]
        all_outputs = []
        system      = SYSTEM

        lessons = mem_module.get_lessons(self.memory)
        if lessons:
            system += f"\n\n{lessons}"

        try:
            while True:
                with self.client.messages.stream(
                    model=MODEL,
                    max_tokens=8192,
                    thinking={"type": "adaptive"},
                    system=system,
                    tools=TOOLS_SCHEMA,
                    messages=messages
                ) as stream:

                    in_thinking = False
                    thinking_buf = ""

                    for event in stream:
                        if event.type == "content_block_start":
                            if event.content_block.type == "thinking":
                                in_thinking = True
                                thinking_buf = ""
                            elif event.content_block.type == "text":
                                in_thinking = False
                                if thinking_buf:
                                    yield {"type": "thinking_done", "summary": thinking_buf[:200]}
                                    thinking_buf = ""

                        elif event.type == "content_block_delta":
                            if event.delta.type == "thinking_delta":
                                thinking_buf += event.delta.thinking
                                yield {"type": "thinking", "content": event.delta.thinking}
                            elif event.delta.type == "text_delta":
                                yield {"type": "text", "content": event.delta.text}

                    response = stream.get_final_message()

                # כלים?
                tool_uses = [b for b in response.content if b.type == "tool_use"]
                if not tool_uses:
                    break

                messages.append({"role": "assistant", "content": response.content})

                tool_results = []
                for tu in tool_uses:
                    label = _tool_label(tu.name)
                    yield {"type": "tool_start", "tool": tu.name, "label": label}

                    result = run_tool(tu.name, tu.input, self.memory)

                    if result.get("path"):
                        all_outputs.append(result["path"])

                    yield {"type": "tool_done", "tool": tu.name, "result": result}

                    tool_results.append({
                        "type": "tool_result",
                        "tool_use_id": tu.id,
                        "content": json.dumps(result, ensure_ascii=False)
                    })

                messages.append({"role": "user", "content": tool_results})

            mem_module.record_success(self.memory, user_message, all_outputs)
            yield {"type": "done", "outputs": all_outputs}

        except anthropic.AuthenticationError:
            err = "❌ API Key שגוי או חסר. הגדר את ANTHROPIC_API_KEY."
            mem_module.record_error(self.memory, user_message, err)
            yield {"type": "error", "content": err}
        except Exception as e:
            err = str(e)
            mem_module.record_error(self.memory, user_message, err)
            yield {"type": "error", "content": f"שגיאה: {err}"}

    def get_memory_summary(self) -> dict:
        m = mem_module.load()
        return {
            "stats": m["stats"],
            "recent_sessions": m["sessions"][-8:],
            "recent_errors":   m["errors"][-5:]
        }


def _tool_label(name: str) -> str:
    return {
        "create_presentation": "יוצר מצגת PowerPoint",
        "create_word_report":  "כותב דוח Word",
        "create_chart":        "מייצר גרף סטטיסטי",
        "recall_memory":       "קורא זיכרון"
    }.get(name, name)
