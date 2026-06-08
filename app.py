"""
CV Filler – DIGITALL Format
Drag-and-drop / browse a candidate CV (PDF or DOCX), review the parsed data,
then click Generate to produce a filled DIGITALL-template CV.
"""
import os
import sys
import json
import re
import threading
import urllib.request
import urllib.error
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ── colours & fonts ─────────────────────────────────────────────────────────
BG        = "#F5F7FA"
PANEL_BG  = "#FFFFFF"
ACCENT    = "#21408C"    # DIGITALL navy
ACCENT2   = "#0078D4"    # Copilot blue
LABEL_FG  = "#465967"
SECTION_FG= "#869CAD"
BTN_BG    = "#21408C"
BTN_FG    = "#FFFFFF"
BTN_HOVER = "#1a3370"
COP_BG    = "#0078D4"    # Copilot button
COP_HOVER = "#005fa3"
FONT      = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 10, "bold")
FONT_H    = ("Segoe UI", 13, "bold")
FONT_SM   = ("Segoe UI", 9)


# ── config (persisted to config.json next to the script) ─────────────────────

_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

_DEFAULT_CONFIG = {
    "provider":          "clipboard",   # "clipboard" | "azure_openai" | "openai" | "gemini" | "groq"
    "azure_endpoint":    "",            # e.g. https://myresource.openai.azure.com
    "azure_deployment":  "gpt-4o",
    "azure_api_version": "2024-02-01",
    "openai_model":      "gpt-4o",
    "gemini_model":      "gemini-2.0-flash",   # FREE: gemini-2.0-flash or gemini-1.5-flash
    "groq_model":        "llama-3.3-70b-versatile",
    "api_key":           "",
}

def load_config():
    if os.path.exists(_CONFIG_PATH):
        try:
            with open(_CONFIG_PATH, encoding="utf-8") as f:
                data = json.load(f)
            cfg = dict(_DEFAULT_CONFIG)
            cfg.update(data)
            return cfg
        except Exception:
            pass
    return dict(_DEFAULT_CONFIG)

def save_config(cfg):
    with open(_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)


# ── AI API caller ─────────────────────────────────────────────────────────────

def call_ai_api(config, prompt):
    """
    Call Azure OpenAI or OpenAI with the given prompt.
    Returns the assistant's reply as a string.
    Raises RuntimeError on failure.
    """
    provider = config.get("provider", "clipboard")
    api_key  = config.get("api_key", "").strip()

    if not api_key:
        raise RuntimeError("No API key configured. Open Settings ⚙ to add one.")

    # Build URL, headers, and body depending on provider
    # Gemini and Groq both expose an OpenAI-compatible endpoint —
    # so all four providers share the same request/response format.

    if provider == "azure_openai":
        endpoint   = config.get("azure_endpoint", "").rstrip("/")
        deployment = config.get("azure_deployment", "gpt-4o")
        api_ver    = config.get("azure_api_version", "2024-02-01")
        if not endpoint:
            raise RuntimeError("Azure endpoint URL is not configured. Open Settings ⚙.")
        url     = f"{endpoint}/openai/deployments/{deployment}/chat/completions?api-version={api_ver}"
        headers = {"Content-Type": "application/json", "api-key": api_key}
        model   = deployment

    elif provider == "openai":
        url     = "https://api.openai.com/v1/chat/completions"
        headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
        model   = config.get("openai_model", "gpt-4o")

    elif provider == "gemini":
        # Google AI Studio — OpenAI-compatible endpoint (no extra libraries needed)
        url     = "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions"
        headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
        model   = config.get("gemini_model", "gemini-2.0-flash")

    elif provider == "groq":
        url     = "https://api.groq.com/openai/v1/chat/completions"
        headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
        model   = config.get("groq_model", "llama-3.3-70b-versatile")

    else:
        raise RuntimeError("Provider is set to 'clipboard'. Open Settings ⚙ to configure an API.")

    body = json.dumps({
        "model":       model,
        "messages":    [{"role": "user", "content": prompt}],
        "max_tokens":  4096,
        "temperature": 0.2,
    }).encode("utf-8")

    req = urllib.request.Request(url, data=body, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=90) as resp:
            result = json.loads(resp.read().decode("utf-8"))
        return result["choices"][0]["message"]["content"]
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")[:300]
        raise RuntimeError(f"API error {e.code}: {detail}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"Network error: {e.reason}\n\nCheck your internet connection and endpoint URL.")


# ── helpers ──────────────────────────────────────────────────────────────────

def _btn(parent, text, command, width=18, bg=None, hover=None):
    bg    = bg    or BTN_BG
    hover = hover or BTN_HOVER
    b = tk.Button(parent, text=text, command=command,
                  bg=bg, fg=BTN_FG, font=FONT_BOLD,
                  relief="flat", bd=0, padx=10, pady=6,
                  cursor="hand2", width=width)
    b.bind("<Enter>", lambda e: b.config(bg=hover))
    b.bind("<Leave>", lambda e: b.config(bg=bg))
    return b


def _label(parent, text, bold=False, fg=LABEL_FG):
    return tk.Label(parent, text=text, bg=PANEL_BG,
                    fg=fg, font=FONT_BOLD if bold else FONT)


def _entry(parent, width=60):
    e = tk.Entry(parent, font=FONT, relief="solid", bd=1,
                 highlightthickness=1, highlightcolor=ACCENT, width=width)
    return e


def _text(parent, height=4, width=70):
    t = tk.Text(parent, font=FONT, relief="solid", bd=1,
                highlightthickness=1, highlightcolor=ACCENT,
                height=height, width=width, wrap="word")
    return t


def _scrollable_frame(parent):
    """Return (outer_frame, inner_frame) where inner_frame scrolls."""
    outer = tk.Frame(parent, bg=PANEL_BG)
    canvas = tk.Canvas(outer, bg=PANEL_BG, highlightthickness=0)
    sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    inner = tk.Frame(canvas, bg=PANEL_BG)
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.configure(yscrollcommand=sb.set)
    canvas.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    return outer, inner


# ── list editor (reusable for languages, certs, etc.) ───────────────────────

class ListEditor(tk.Frame):
    """A vertically stacked list of single-line entries with add/remove."""
    def __init__(self, parent, placeholder="", **kwargs):
        super().__init__(parent, bg=PANEL_BG, **kwargs)
        self.placeholder = placeholder
        self._rows = []
        self._btn_frame = tk.Frame(self, bg=PANEL_BG)
        self._btn_frame.pack(anchor="w", pady=(4, 0))
        add_btn = tk.Button(self._btn_frame, text="+ Add", command=self.add_row,
                            bg="#E8F0FE", fg=ACCENT, font=FONT, relief="flat",
                            bd=0, padx=8, pady=3, cursor="hand2")
        add_btn.pack(side="left")

    def add_row(self, value=""):
        row_frame = tk.Frame(self, bg=PANEL_BG)
        row_frame.pack(fill="x", pady=2)
        e = tk.Entry(row_frame, font=FONT, relief="solid", bd=1,
                     highlightthickness=1, highlightcolor=ACCENT, width=65)
        e.insert(0, value)
        e.pack(side="left", fill="x", expand=True)
        rm = tk.Button(row_frame, text="✕", command=lambda rf=row_frame: self._remove(rf),
                       bg=PANEL_BG, fg="#cc4444", font=FONT, relief="flat",
                       bd=0, padx=4, cursor="hand2")
        rm.pack(side="left", padx=(4, 0))
        self._rows.append((row_frame, e))

    def _remove(self, row_frame):
        for i, (rf, e) in enumerate(self._rows):
            if rf is row_frame:
                rf.destroy()
                self._rows.pop(i)
                break

    def set_values(self, values):
        for rf, e in self._rows:
            rf.destroy()
        self._rows.clear()
        for v in values:
            self.add_row(v)

    def get_values(self):
        return [e.get().strip() for _, e in self._rows if e.get().strip()]


# ── work experience editor ───────────────────────────────────────────────────

class JobEntry(tk.LabelFrame):
    def __init__(self, parent, remove_callback, index=1, **kwargs):
        super().__init__(parent, text=f"  Job #{index}  ", bg=PANEL_BG,
                         fg=ACCENT, font=FONT_BOLD, relief="groove", bd=1, **kwargs)
        self._remove_cb = remove_callback
        self._build()

    def _build(self):
        fields = [
            ("Position / Job Title", "position"),
            ("Employer (leave blank for Confidential)", "employer"),
            ("Duration  (e.g. June 2021 – Present)", "duration"),
            ("Project Name (optional)", "project_name"),
            ("Location  (e.g. Sofia, Bulgaria)", "location"),
        ]
        self._vars = {}
        for label, key in fields:
            row = tk.Frame(self, bg=PANEL_BG)
            row.pack(fill="x", padx=10, pady=3)
            tk.Label(row, text=label, bg=PANEL_BG, fg=LABEL_FG, font=FONT, width=42, anchor="w").pack(side="left")
            e = tk.Entry(row, font=FONT, relief="solid", bd=1,
                         highlightthickness=1, highlightcolor=ACCENT, width=42)
            e.pack(side="left", fill="x", expand=True)
            self._vars[key] = e

        # Responsibilities
        r2 = tk.Frame(self, bg=PANEL_BG)
        r2.pack(fill="x", padx=10, pady=3)
        tk.Label(r2, text="Responsibilities (one per line)", bg=PANEL_BG,
                 fg=LABEL_FG, font=FONT, width=42, anchor="nw").pack(side="left", anchor="n")
        self._resp = tk.Text(r2, font=FONT, relief="solid", bd=1,
                             highlightthickness=1, highlightcolor=ACCENT,
                             height=5, width=42, wrap="word")
        self._resp.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(r2, command=self._resp.yview)
        sb.pack(side="left", fill="y")
        self._resp.config(yscrollcommand=sb.set)

        # Remove button
        rm = tk.Button(self, text="Remove this job", command=self._remove_cb,
                       bg="#FFF0F0", fg="#cc4444", font=FONT, relief="flat",
                       bd=0, padx=8, pady=3, cursor="hand2")
        rm.pack(anchor="e", padx=10, pady=(0, 8))

    def set_data(self, job):
        for key, widget in self._vars.items():
            widget.delete(0, tk.END)
            widget.insert(0, job.get(key, ""))
        self._resp.delete("1.0", tk.END)
        self._resp.insert("1.0", "\n".join(job.get("responsibilities", [])))

    def get_data(self):
        data = {k: v.get().strip() for k, v in self._vars.items()}
        data["employer"] = data["employer"] or "Confidential"
        resp_raw = self._resp.get("1.0", tk.END).strip()
        data["responsibilities"] = [l.strip() for l in resp_raw.splitlines() if l.strip()]
        return data


class WorkExperienceTab(tk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=PANEL_BG, **kwargs)
        self._jobs = []
        self._outer, self._inner = _scrollable_frame(self)
        self._outer.pack(fill="both", expand=True)
        _btn(self._inner, "+ Add Job", self.add_job, width=14).pack(anchor="w", padx=10, pady=8)

    def add_job(self, job_data=None):
        idx = len(self._jobs) + 1
        container = tk.Frame(self._inner, bg=PANEL_BG)
        container.pack(fill="x", padx=10, pady=4)

        def remove(c=container, j=None):
            self._jobs = [(jf, jd) for jf, jd in self._jobs if jf is not c]
            c.destroy()
            self._renumber()

        je = JobEntry(container, remove_callback=lambda c=container: remove(c), index=idx)
        je.pack(fill="x")
        if job_data:
            je.set_data(job_data)
        self._jobs.append((container, je))

    def _renumber(self):
        for i, (_, je) in enumerate(self._jobs, 1):
            je.config(text=f"  Job #{i}  ")

    def set_jobs(self, jobs):
        for container, _ in self._jobs:
            container.destroy()
        self._jobs.clear()
        for job in jobs:
            self.add_job(job)

    def get_jobs(self):
        return [je.get_data() for _, je in self._jobs]


# ── education editor ─────────────────────────────────────────────────────────

class EduEntry(tk.LabelFrame):
    def __init__(self, parent, remove_callback, index=1, **kwargs):
        super().__init__(parent, text=f"  Education #{index}  ", bg=PANEL_BG,
                         fg=ACCENT, font=FONT_BOLD, relief="groove", bd=1, **kwargs)
        self._remove_cb = remove_callback
        fields = [
            ("Degree", "degree"),
            ("Institute / University", "institute"),
            ("Location", "location"),
            ("Year(s)  (e.g. 2010 – 2014)", "year"),
        ]
        self._vars = {}
        for label, key in fields:
            row = tk.Frame(self, bg=PANEL_BG)
            row.pack(fill="x", padx=10, pady=3)
            tk.Label(row, text=label, bg=PANEL_BG, fg=LABEL_FG, font=FONT, width=28, anchor="w").pack(side="left")
            e = tk.Entry(row, font=FONT, relief="solid", bd=1,
                         highlightthickness=1, highlightcolor=ACCENT, width=50)
            e.pack(side="left", fill="x", expand=True)
            self._vars[key] = e
        rm = tk.Button(self, text="Remove", command=remove_callback,
                       bg="#FFF0F0", fg="#cc4444", font=FONT, relief="flat",
                       bd=0, padx=8, pady=3, cursor="hand2")
        rm.pack(anchor="e", padx=10, pady=(0, 8))

    def set_data(self, d):
        for k, w in self._vars.items():
            w.delete(0, tk.END)
            w.insert(0, d.get(k, ""))

    def get_data(self):
        return {k: v.get().strip() for k, v in self._vars.items()}


class EducationTab(tk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=PANEL_BG, **kwargs)
        self._entries = []
        self._outer, self._inner = _scrollable_frame(self)
        self._outer.pack(fill="both", expand=True)
        _btn(self._inner, "+ Add Education", self.add_entry, width=18).pack(anchor="w", padx=10, pady=8)

    def add_entry(self, data=None):
        idx = len(self._entries) + 1
        container = tk.Frame(self._inner, bg=PANEL_BG)
        container.pack(fill="x", padx=10, pady=4)

        def remove(c=container):
            self._entries = [(ef, ed) for ef, ed in self._entries if ef is not c]
            c.destroy()

        ee = EduEntry(container, remove_callback=lambda c=container: remove(c), index=idx)
        ee.pack(fill="x")
        if data:
            ee.set_data(data)
        self._entries.append((container, ee))

    def set_entries(self, edu_list):
        for c, _ in self._entries:
            c.destroy()
        self._entries.clear()
        for d in edu_list:
            self.add_entry(d)

    def get_entries(self):
        return [ee.get_data() for _, ee in self._entries]


# ── Copilot prompt builder ────────────────────────────────────────────────────

_COPILOT_PROMPT_TEMPLATE = """\
You are a CV data extraction and writing assistant for an IT staffing company. Extract all information from the CV text below and return it as a single JSON object. Return ONLY valid JSON - no explanation, no commentary, no markdown code block. Start your response with {{ and end with }}.

Use exactly this JSON schema:

{{
  "name": "Full name of the candidate",
  "job_title": "Professional title - use their most recent job title or the title they present themselves with",
  "summary": "WRITE an extensive, polished 4-6 sentence professional summary in third person (e.g. John is an experienced...). Do NOT copy the original summary verbatim - rewrite and expand it to be more professional and impactful. Include: (1) total years of experience and main field, (2) 3-4 key technical competencies or specialisations, (3) notable achievements, industries, or environments they worked in, (4) soft skills or working style if mentioned. Make it sound compelling for a client presentation. Even if the original summary is short or missing, write a full paragraph based on their work experience.",
  "technologies": {{
    "primary_expertise": "Their main professional domain(s), comma-separated. This is WHAT they do at a high level. Examples by role: IT admin: Windows System Administration, Microsoft 365, Intune, Endpoint Management | Developer: Java Backend Development, Microservices, REST APIs | Recruiter: IT Recruitment, Talent Acquisition, Headhunting | PM: IT Project Management, Agile Delivery | DevOps: CI/CD Pipelines, Kubernetes, Cloud Operations",
    "programming_languages": "ONLY actual programming and scripting languages, comma-separated (Python, PowerShell, JavaScript, Java, C#, SQL, Bash, TypeScript, etc.). Leave as EMPTY STRING if the person does not write code - for example for recruiters, managers, or IT support staff without scripting duties.",
    "methodologies": "HOW they work - processes, frameworks, and practices. NOT specific tools or software. Examples: IT support: ITIL, Incident Management, Change Management, L2/L3 Support, SLA Management | Developer: Agile, Scrum, Kanban, TDD, Code Review, CI/CD | Recruiter: Full-cycle Recruitment, Boolean Search, Competency-based Interviews, Talent Mapping | PM: Agile, Waterfall, Risk Management, Budget Planning",
    "other_skills": "Specific tools, platforms, software products, and technologies. Everything concrete that does not fit above. Examples: IT: Active Directory, Azure AD, Intune, SCCM, GPO, VMware, Windows Server | Dev: Spring Boot, Docker, Kubernetes, AWS, PostgreSQL, Redis, Git | Recruiter: LinkedIn Recruiter, Workday, SAP SuccessFactors, Greenhouse | PM: Jira, Confluence, MS Project, ServiceNow"
  }},
  "education": [
    {{
      "degree": "Full degree name and field of study",
      "institute": "University or institution name",
      "location": "City, Country",
      "year": "e.g. 2010 - 2014 or 2014"
    }}
  ],
  "languages": [
    {{"language": "Language name", "level": "Proficiency level (e.g. Native, Fluent, C1, B2, Professional working proficiency)"}}
  ],
  "certifications": [
    "Full certification, course, or training name including issuer and date if available"
  ],
  "work_experience": [
    {{
      "position": "Exact job title at this role",
      "employer": "Company name (write Confidential if not clearly stated)",
      "duration": "Date range, e.g. June 2021 - Present or 06/2021 - Present",
      "project_name": "Client project name if specifically mentioned, otherwise empty string",
      "location": "City, Country if mentioned, otherwise empty string",
      "responsibilities": [
        "Each responsibility or achievement as a separate string - preserve full detail",
        "Do NOT summarize or merge bullet points",
        "One logical point per entry"
      ]
    }}
  ]
}}

Important rules:
- List work experience in reverse chronological order (most recent first)
- WRITE a proper 4-6 sentence professional summary - this is the most important field
- Keep all responsibilities fully detailed - never summarize, merge, or shorten them
- Adapt the technologies section to this person's actual role type using the examples above
- If a field has no data, use empty string "" or empty array []
- Return ONLY the JSON object. No other text before or after.

CV TEXT:
---
{cv_text}
---"""

def build_copilot_prompt(cv_text):
    # Truncate very long CVs to stay within Copilot's context window
    max_chars = 12000
    if len(cv_text) > max_chars:
        cv_text = cv_text[:max_chars] + "\n\n[... CV truncated for length ...]"
    return _COPILOT_PROMPT_TEMPLATE.format(cv_text=cv_text)


def _clean_json_text(text):
    """
    Fix the most common ways Copilot produces invalid JSON:
      1. Smart/curly quotes  "..."  →  "..."
      2. Unescaped literal newlines / tabs inside string values
      3. Trailing commas before } or ]
      4. Stray BOM or zero-width characters
    """
    # 1. Smart double quotes → straight
    text = text.replace("“", '"').replace("”", '"')
    # Single curly quotes → straight apostrophe (safe inside JSON string values)
    text = text.replace("‘", "'").replace("’", "'")
    # En-dash / em-dash that sometimes appear in durations
    text = text.replace("–", "-").replace("—", "-")

    # 2. Replace literal newlines/tabs that appear INSIDE JSON string values.
    #    We walk character by character to track whether we're inside a string.
    result = []
    in_string = False
    i = 0
    while i < len(text):
        ch = text[i]
        if ch == '\\' and in_string:          # escaped character — keep as-is
            result.append(ch)
            i += 1
            if i < len(text):
                result.append(text[i])
            i += 1
            continue
        if ch == '"':
            in_string = not in_string
            result.append(ch)
            i += 1
            continue
        if in_string and ch == '\n':
            result.append('\\n')              # escape the literal newline
        elif in_string and ch == '\r':
            pass                              # drop carriage returns
        elif in_string and ch == '\t':
            result.append('\\t')
        else:
            result.append(ch)
        i += 1
    text = "".join(result)

    # 3. Trailing commas: ,} or ,]  (not valid JSON)
    text = re.sub(r",(\s*[}\]])", r"\1", text)

    # 4. Strip BOM / zero-width chars
    text = text.lstrip("﻿​‌‍")

    return text


def parse_copilot_response(raw_text):
    """Extract and parse JSON from Copilot's response text."""
    # Strip markdown code fences if Copilot wrapped it anyway
    text = re.sub(r"```(?:json)?\s*", "", raw_text).strip()
    text = re.sub(r"```\s*$", "", text).strip()

    # Find first { ... last }
    start = text.find("{")
    end   = text.rfind("}")
    if start == -1 or end == -1:
        raise ValueError(
            "No JSON object found in the pasted text.\n"
            "Make sure you copied Copilot's full response."
        )
    json_str = text[start : end + 1]

    # First attempt — parse as-is
    try:
        return json.loads(json_str)
    except json.JSONDecodeError:
        pass

    # Second attempt — apply auto-fixes for common Copilot quirks
    cleaned = _clean_json_text(json_str)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError as exc:
        # Show a helpful snippet around the problem location
        char = exc.pos or 0
        snippet = cleaned[max(0, char - 60) : char + 60].replace("\n", "↵")
        raise ValueError(
            f"Could not parse Copilot's response as JSON.\n\n"
            f"Error: {exc.msg} (near char {char})\n"
            f"Context: …{snippet}…\n\n"
            f"Try copying Copilot's response again — make sure you got the full JSON block."
        )


# ── Settings dialog ──────────────────────────────────────────────────────────

class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, config, on_save):
        super().__init__(parent)
        self.title("AI Settings")
        self.geometry("620x520")
        self.resizable(False, False)
        self.configure(bg=BG)
        self.grab_set()
        self._config  = dict(config)
        self._on_save = on_save
        self._build()

    def _build(self):
        # Header
        hdr = tk.Frame(self, bg=ACCENT, height=44)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="⚙  AI Provider Settings", bg=ACCENT, fg="white",
                 font=FONT_BOLD).pack(side="left", padx=16)

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=14)

        # Provider
        tk.Label(body, text="Provider", bg=BG, fg=LABEL_FG, font=FONT_BOLD).grid(
            row=0, column=0, sticky="w", pady=(0, 4))
        self._provider = tk.StringVar(value=self._config.get("provider", "clipboard"))
        prov_frame = tk.Frame(body, bg=BG)
        prov_frame.grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 12))
        for val, lbl in [
            ("clipboard",   "Clipboard only  (no API — paste into any AI manually)"),
            ("gemini",      "Google Gemini  ✨ FREE — aistudio.google.com"),
            ("groq",        "Groq  ✨ FREE — console.groq.com  (very fast)"),
            ("openai",      "OpenAI  (ChatGPT API — paid)"),
            ("azure_openai","Azure OpenAI  (Microsoft / work tenant — paid)"),
        ]:
            tk.Radiobutton(prov_frame, text=lbl, variable=self._provider, value=val,
                           bg=BG, fg=LABEL_FG, font=FONT, activebackground=BG,
                           command=self._on_provider_change).pack(anchor="w")

        # Dynamic fields frame
        self._fields_frame = tk.LabelFrame(body, text="  Connection details  ",
                                           bg=BG, fg=ACCENT, font=FONT_BOLD,
                                           relief="groove", bd=1)
        self._fields_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        body.columnconfigure(0, weight=1)
        self._draw_fields()

        # Buttons
        btn_row = tk.Frame(body, bg=BG)
        btn_row.grid(row=3, column=0, columnspan=2, sticky="e")
        _btn(btn_row, "Test connection", self._test, width=18,
             bg="#6c757d", hover="#5a6268").pack(side="left", padx=(0, 8))
        _btn(btn_row, "Save", self._save, width=10).pack(side="left", padx=(0, 6))
        tk.Button(btn_row, text="Cancel", command=self.destroy,
                  bg=BG, fg=LABEL_FG, font=FONT, relief="flat", bd=0,
                  padx=10, pady=6, cursor="hand2").pack(side="left")

        self._status = tk.Label(body, text="", bg=BG, font=FONT_BOLD)
        self._status.grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))

    def _draw_fields(self):
        for w in self._fields_frame.winfo_children():
            w.destroy()
        prov = self._provider.get()
        if prov == "clipboard":
            tk.Label(self._fields_frame,
                     text="No API key needed — uses clipboard to bridge any AI.",
                     bg=BG, fg=SECTION_FG, font=FONT).pack(padx=12, pady=10)
            return

        self._field_rows_offset = 0
        fields = []
        if prov == "gemini":
            # Show a help link label above the fields
            tk.Label(self._fields_frame,
                     text="Get your FREE API key → aistudio.google.com  (sign in with any Google account)",
                     bg=BG, fg="#1a7f37", font=FONT).grid(
                     row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(8, 2))
            fields = [
                ("Model",   "gemini_model", "gemini-2.0-flash  (recommended, free)"),
                ("API key", "api_key",      "Paste the key from AI Studio"),
            ]
            # Offset rows because we inserted a label at row 0
            self._field_rows_offset = 1
        elif prov == "groq":
            tk.Label(self._fields_frame,
                     text="Get your FREE API key → console.groq.com  (free sign-up, no credit card)",
                     bg=BG, fg="#1a7f37", font=FONT).grid(
                     row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(8, 2))
            fields = [
                ("Model",   "groq_model", "llama-3.3-70b-versatile  (recommended, free)"),
                ("API key", "api_key",    "Paste the key from Groq console"),
            ]
            self._field_rows_offset = 1
        elif prov == "azure_openai":
            fields = [
                ("Azure endpoint URL", "azure_endpoint",
                 "e.g. https://myresource.openai.azure.com"),
                ("Deployment name",    "azure_deployment",  "e.g. gpt-4o"),
                ("API version",        "azure_api_version", "e.g. 2024-02-01"),
                ("API key",            "api_key",           "Your Azure OpenAI key"),
            ]
        elif prov == "openai":
            fields = [
                ("Model",   "openai_model", "e.g. gpt-4o  or  gpt-4o-mini"),
                ("API key", "api_key",      "sk-..."),
            ]

        offset = getattr(self, "_field_rows_offset", 0)
        self._field_vars = {}
        for i, (lbl, key, hint) in enumerate(fields):
            row_base = offset + i * 2
            tk.Label(self._fields_frame, text=lbl, bg=BG, fg=LABEL_FG,
                     font=FONT_BOLD, anchor="w").grid(row=row_base, column=0,
                     sticky="w", padx=12, pady=(8, 2))
            var = tk.StringVar(value=self._config.get(key, ""))
            show = "*" if key == "api_key" else ""
            e = tk.Entry(self._fields_frame, textvariable=var, font=FONT,
                         relief="solid", bd=1, width=52, show=show)
            e.grid(row=row_base+1, column=0, sticky="ew", padx=12, pady=(0, 2))
            tk.Label(self._fields_frame, text=hint, bg=BG, fg=SECTION_FG,
                     font=FONT_SM).grid(row=row_base+1, column=1, sticky="w", padx=4)
            self._fields_frame.columnconfigure(0, weight=1)
            self._field_vars[key] = var

    def _on_provider_change(self):
        self._draw_fields()

    def _collect(self):
        cfg = dict(self._config)
        cfg["provider"] = self._provider.get()
        for key, var in getattr(self, "_field_vars", {}).items():
            cfg[key] = var.get().strip()
        return cfg

    def _test(self):
        cfg = self._collect()
        if cfg["provider"] == "clipboard":
            self._status.config(text="ℹ  Clipboard mode — no connection to test.", fg=SECTION_FG)
            return
        self._status.config(text="Testing…", fg=LABEL_FG)
        self.update()
        def run():
            try:
                reply = call_ai_api(cfg, 'Reply with exactly: {"status":"ok"}')
                if "ok" in reply.lower():
                    self._status.config(text="✔  Connection successful!", fg="#1a7f37")
                else:
                    self._status.config(text=f"✔  Connected — reply: {reply[:60]}", fg="#1a7f37")
            except Exception as ex:
                self._status.config(text=f"✖  {ex}", fg="#cc4444")
        threading.Thread(target=run, daemon=True).start()

    def _save(self):
        cfg = self._collect()
        self._on_save(cfg)
        self.destroy()


# ── Copilot / AI dialog window ────────────────────────────────────────────────

class CopilotDialog(tk.Toplevel):
    """Two-step dialog: copy prompt → paste response."""

    def __init__(self, parent, cv_text, on_data_ready):
        super().__init__(parent)
        self.title("Copilot Assistant")
        self.geometry("700x580")
        self.resizable(True, True)
        self.configure(bg=BG)
        self.grab_set()          # modal
        self._cv_text = cv_text
        self._on_data_ready = on_data_ready
        self._build()

    def _build(self):
        # Header
        hdr = tk.Frame(self, bg=COP_BG, height=48)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="🤖  Copilot Assistant", bg=COP_BG, fg="white",
                 font=("Segoe UI", 14, "bold")).pack(side="left", padx=16, pady=8)

        # ── Step 1 ────────────────────────────────────────────────────────────
        s1 = tk.LabelFrame(self, text="  Step 1 — Copy the prompt to Copilot  ",
                           bg=BG, fg=ACCENT, font=FONT_BOLD, relief="groove", bd=1)
        s1.pack(fill="x", padx=16, pady=(14, 6))

        if self._cv_text:
            step1_msg = (
                "Click the button below to copy the extraction prompt to your clipboard.\n"
                "Then open Copilot in Teams or your browser and paste it (Ctrl+V)."
            )
        else:
            step1_msg = (
                "⚠  The CV file could not be read automatically (possibly a scanned/image PDF).\n"
                "Please open the CV manually, select all text (Ctrl+A), copy it (Ctrl+C),\n"
                "then click \"Paste CV text\" below. After that, copy the prompt to Copilot."
            )

        tk.Label(s1, text=step1_msg, bg=BG, fg=LABEL_FG, font=FONT,
                 justify="left").pack(anchor="w", padx=12, pady=(8, 4))

        # Manual CV text paste area — shown when auto-extraction failed
        if not self._cv_text:
            paste_cv_row = tk.Frame(s1, bg=BG)
            paste_cv_row.pack(fill="x", padx=12, pady=(0, 4))
            self._cv_text_box = tk.Text(paste_cv_row, font=("Consolas", 9), relief="solid", bd=1,
                                        height=5, wrap="word", bg="#FFFEF0")
            self._cv_text_box.insert("1.0", "Paste the candidate CV text here…")
            self._cv_text_box.config(fg="#999999")
            self._cv_text_box.bind("<FocusIn>", self._on_cv_box_focus)
            sb_cv = ttk.Scrollbar(paste_cv_row, command=self._cv_text_box.yview)
            self._cv_text_box.config(yscrollcommand=sb_cv.set)
            self._cv_text_box.pack(side="left", fill="both", expand=True)
            sb_cv.pack(side="left", fill="y")

            # Both buttons in one row, same style
            btn_row = tk.Frame(s1, bg=BG)
            btn_row.pack(anchor="w", padx=12, pady=(4, 10))
            _btn(btn_row, "📋  Paste CV text from clipboard", self._paste_cv_text,
                 width=30, bg="#6c757d", hover="#5a6268").pack(side="left", padx=(0, 8))
            self._copy_btn = _btn(btn_row, "📋  Copy prompt to clipboard", self._copy_prompt,
                                  width=30, bg=COP_BG, hover=COP_HOVER)
            self._copy_btn.pack(side="left")
        else:
            self._cv_text_box = None
            btn_row = tk.Frame(s1, bg=BG)
            btn_row.pack(anchor="w", padx=12, pady=(4, 10))
            self._copy_btn = _btn(btn_row, "📋  Copy prompt to clipboard", self._copy_prompt,
                                  width=30, bg=COP_BG, hover=COP_HOVER)
            self._copy_btn.pack(side="left")

        self._copy_status = tk.Label(s1, text="", bg=BG, fg="#1a7f37", font=FONT_BOLD)
        self._copy_status.pack(anchor="w", padx=12, pady=(0, 6))

        # ── Step 2 ────────────────────────────────────────────────────────────
        s2 = tk.LabelFrame(self, text="  Step 2 — Paste Copilot's response here  ",
                           bg=BG, fg=ACCENT, font=FONT_BOLD, relief="groove", bd=1)
        s2.pack(fill="both", expand=True, padx=16, pady=(6, 6))

        tk.Label(s2, text=(
            "After Copilot replies with a JSON block, select all of its response,\n"
            "copy it (Ctrl+C), then click the button below."
        ), bg=BG, fg=LABEL_FG, font=FONT, justify="left").pack(anchor="w", padx=12, pady=(8, 4))

        # Paste area
        paste_row = tk.Frame(s2, bg=BG)
        paste_row.pack(fill="both", expand=True, padx=12, pady=(0, 4))
        self._response_box = tk.Text(paste_row, font=("Consolas", 9), relief="solid", bd=1,
                                      highlightthickness=1, highlightcolor=COP_BG,
                                      height=10, wrap="word", bg="#FAFEFF")
        sb = ttk.Scrollbar(paste_row, command=self._response_box.yview)
        self._response_box.config(yscrollcommand=sb.set)
        self._response_box.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")

        tk.Label(s2, text="💡 Tip: you can also paste directly into the box above (Ctrl+V) and then click Fill.",
                 bg=BG, fg=SECTION_FG, font=FONT_SM).pack(anchor="w", padx=12, pady=(0, 4))

        btn_row = tk.Frame(s2, bg=BG)
        btn_row.pack(fill="x", padx=12, pady=(0, 10))
        _btn(btn_row, "📥  Paste from clipboard", self._paste_from_clipboard,
             width=24, bg=COP_BG, hover=COP_HOVER).pack(side="left", padx=(0, 8))
        _btn(btn_row, "✔  Fill form with response", self._fill_form,
             width=24).pack(side="left")

        self._fill_status = tk.Label(s2, text="", bg=BG, font=FONT_BOLD)
        self._fill_status.pack(anchor="w", padx=12, pady=(0, 6))

    def _on_cv_box_focus(self, event):
        """Clear placeholder text on first click."""
        if self._cv_text_box and self._cv_text_box.get("1.0", tk.END).strip() == "Paste the candidate CV text here…":
            self._cv_text_box.delete("1.0", tk.END)
            self._cv_text_box.config(fg="#000000")

    def _paste_cv_text(self):
        """Paste clipboard content into the manual CV text box."""
        try:
            text = self.clipboard_get()
            if self._cv_text_box:
                self._cv_text_box.delete("1.0", tk.END)
                self._cv_text_box.insert("1.0", text)
                self._cv_text_box.config(fg="#000000")
                self._cv_text = text
        except tk.TclError:
            pass

    def _get_cv_text(self):
        """Return CV text — from extraction or from manual paste box."""
        if self._cv_text:
            return self._cv_text
        if self._cv_text_box:
            t = self._cv_text_box.get("1.0", tk.END).strip()
            if t and t != "Paste the candidate CV text here…":
                self._cv_text = t
                return t
        return ""

    def _copy_prompt(self):
        cv_text = self._get_cv_text()
        if not cv_text:
            self._copy_status.config(
                text="⚠  Please paste the CV text first (see above).", fg="#cc4444")
            return
        prompt = build_copilot_prompt(cv_text)
        self.clipboard_clear()
        self.clipboard_append(prompt)
        self.update()   # flush clipboard on Windows before dialog closes
        self._copy_status.config(
            text="✔  Prompt copied! Now open Copilot in Teams / browser and paste it.", fg="#1a7f37")

    def _paste_from_clipboard(self):
        try:
            text = self.clipboard_get()
            self._response_box.delete("1.0", tk.END)
            self._response_box.insert("1.0", text)
        except tk.TclError:
            self._fill_status.config(text="⚠  Clipboard is empty or unavailable.", fg="#cc4444")

    def _fill_form(self):
        raw = self._response_box.get("1.0", tk.END).strip()
        if not raw:
            self._fill_status.config(text="⚠  Nothing pasted yet — copy Copilot's response first.", fg="#cc4444")
            return
        try:
            data = parse_copilot_response(raw)
        except ValueError as exc:
            self._fill_status.config(text=f"⚠  {exc}", fg="#cc4444")
            return

        self._on_data_ready(data)
        self._fill_status.config(text="✔  Form filled! Review the tabs and click Generate CV.", fg="#1a7f37")
        self.after(1500, self.destroy)


# ── main application window ──────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CV Filler – DIGITALL Format")
        self.geometry("900x720")
        self.resizable(True, True)
        self.configure(bg=BG)
        try:
            self.iconbitmap(default="")
        except Exception:
            pass
        self._source_path = tk.StringVar(value="No file selected")
        self._raw_cv_text = ""
        self._config      = load_config()
        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────

    def _build_ui(self):
        # ── header bar ───────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=ACCENT, height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="CV Filler", bg=ACCENT, fg="white",
                 font=("Segoe UI", 18, "bold")).pack(side="left", padx=20)
        tk.Label(hdr, text="DIGITALL Format", bg=ACCENT, fg="#9AB5D0",
                 font=("Segoe UI", 11)).pack(side="left", padx=0)
        _btn(hdr, "⚙  Settings", self._open_settings,
             width=12, bg="#1a3370", hover="#0f2050").pack(side="right", padx=12, pady=8)

        # ── source file row ───────────────────────────────────────────────
        file_row = tk.Frame(self, bg=BG)
        file_row.pack(fill="x", padx=20, pady=(12, 4))
        tk.Label(file_row, text="Source CV:", bg=BG, fg=LABEL_FG,
                 font=FONT_BOLD).pack(side="left")
        tk.Label(file_row, textvariable=self._source_path, bg=BG, fg=ACCENT,
                 font=FONT).pack(side="left", padx=10)
        _btn(file_row, "Browse…", self._browse_source, width=10).pack(side="left", padx=4)
        _btn(file_row, "Parse CV", self._parse_cv, width=10).pack(side="left", padx=4)

        # ── AI row ────────────────────────────────────────────────────────
        self._ai_row = tk.Frame(self, bg="#EBF3FB", bd=0)
        self._ai_row.pack(fill="x", padx=20, pady=(0, 8))
        self._rebuild_ai_row()

        # ── notebook ─────────────────────────────────────────────────────
        style = ttk.Style()
        style.configure("TNotebook", background=BG)
        style.configure("TNotebook.Tab", padding=[12, 6], font=FONT_BOLD)
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        # Tab 1 – Summary
        t1_outer, t1 = _scrollable_frame(tk.Frame(nb, bg=PANEL_BG))
        nb.add(t1_outer.master, text="Summary")
        t1_outer.pack(fill="both", expand=True)

        def _lbl_entry(parent, label, row_offset=0):
            tk.Label(parent, text=label, bg=PANEL_BG, fg=LABEL_FG,
                     font=FONT_BOLD, anchor="w").grid(row=row_offset, column=0, sticky="w", padx=12, pady=(10, 2))
            e = _entry(parent, width=72)
            e.grid(row=row_offset+1, column=0, sticky="ew", padx=12, pady=(0, 4))
            parent.columnconfigure(0, weight=1)
            return e

        self._name   = _lbl_entry(t1, "Candidate Name  (used for output filename)", 0)
        self._title  = _lbl_entry(t1, "Job Title / Role  (shown at top of CV)", 2)

        tk.Label(t1, text="Summary  (professional overview paragraph)",
                 bg=PANEL_BG, fg=LABEL_FG, font=FONT_BOLD).grid(row=4, column=0, sticky="w", padx=12, pady=(10, 2))
        self._summary = _text(t1, height=6, width=72)
        self._summary.grid(row=5, column=0, sticky="ew", padx=12, pady=(0, 12))

        # Tab 2 – Technologies
        t2_outer, t2 = _scrollable_frame(tk.Frame(nb, bg=PANEL_BG))
        nb.add(t2_outer.master, text="Technologies")
        t2_outer.pack(fill="both", expand=True)

        tech_fields = [
            ("primary_expertise",    "Primary Expertise  (e.g. Endpoint Management, Intune)"),
            ("programming_languages","Programming Languages  (e.g. PowerShell, Python)"),
            ("methodologies",        "Methodologies  (e.g. Agile, CI/CD, ITIL)"),
            ("other_skills",         "Other Skills / Tools"),
        ]
        self._tech = {}
        for i, (key, lbl) in enumerate(tech_fields):
            tk.Label(t2, text=lbl, bg=PANEL_BG, fg=LABEL_FG,
                     font=FONT_BOLD).grid(row=i*2, column=0, sticky="w", padx=12, pady=(10, 2))
            t2.columnconfigure(0, weight=1)
            if key == "other_skills":
                w = _text(t2, height=3, width=72)
                w.grid(row=i*2+1, column=0, sticky="ew", padx=12, pady=(0, 4))
            else:
                w = _entry(t2, width=72)
                w.grid(row=i*2+1, column=0, sticky="ew", padx=12, pady=(0, 4))
            self._tech[key] = w

        # Tab 3 – Education
        self._edu_tab = EducationTab(tk.Frame(nb, bg=PANEL_BG))
        nb.add(self._edu_tab.master, text="Education")
        self._edu_tab.pack(fill="both", expand=True)

        # Tab 4 – Languages & Certs
        t4_outer, t4 = _scrollable_frame(tk.Frame(nb, bg=PANEL_BG))
        nb.add(t4_outer.master, text="Languages & Certs")
        t4_outer.pack(fill="both", expand=True)

        tk.Label(t4, text="Languages  (e.g. English: Professional working proficiency)",
                 bg=PANEL_BG, fg=LABEL_FG, font=FONT_BOLD).pack(anchor="w", padx=12, pady=(12, 2))
        self._langs = ListEditor(t4)
        self._langs.pack(fill="x", padx=12)

        tk.Label(t4, text="Courses / Certificates / Trainings",
                 bg=PANEL_BG, fg=LABEL_FG, font=FONT_BOLD).pack(anchor="w", padx=12, pady=(18, 2))
        self._certs = ListEditor(t4)
        self._certs.pack(fill="x", padx=12)

        # Tab 5 – Work Experience
        self._we_tab = WorkExperienceTab(tk.Frame(nb, bg=PANEL_BG))
        nb.add(self._we_tab.master, text="Work Experience")
        self._we_tab.pack(fill="both", expand=True)

        # ── bottom bar ────────────────────────────────────────────────────
        bot = tk.Frame(self, bg=BG)
        bot.pack(fill="x", padx=20, pady=10)
        self._status = tk.Label(bot, text="", bg=BG, fg=ACCENT, font=FONT)
        self._status.pack(side="left")

        # Confidential toggle
        self._confidential = tk.BooleanVar(value=True)
        conf_frame = tk.Frame(bot, bg=BG)
        conf_frame.pack(side="right", padx=(0, 12))
        tk.Checkbutton(
            conf_frame, text="Confidential", variable=self._confidential,
            bg=BG, fg=LABEL_FG, font=FONT, activebackground=BG,
            selectcolor=BG, cursor="hand2",
        ).pack(side="left")
        tk.Label(conf_frame, text="(hides name & employers)", bg=BG,
                 fg=SECTION_FG, font=FONT_SM).pack(side="left", padx=(2, 0))

        _btn(bot, "Generate CV", self._generate, width=14).pack(side="right")

    # ── AI row helpers ────────────────────────────────────────────────────

    def _rebuild_ai_row(self):
        for w in self._ai_row.winfo_children():
            w.destroy()
        inner = tk.Frame(self._ai_row, bg="#EBF3FB")
        inner.pack(fill="x", padx=8, pady=6)

        provider = self._config.get("provider", "clipboard")
        has_api  = provider != "clipboard" and bool(self._config.get("api_key", "").strip())

        if has_api:
            label = {"azure_openai": "Azure OpenAI (Copilot)", "openai": "OpenAI (ChatGPT)"}.get(provider, "AI")
            tk.Label(inner, text=f"🤖  {label} connected —",
                     bg="#EBF3FB", fg="#1a7f37", font=FONT_BOLD).pack(side="left")
            tk.Label(inner, text="load a CV then auto-fill the form in one click.",
                     bg="#EBF3FB", fg=LABEL_FG, font=FONT_SM).pack(side="left", padx=8)
            _btn(inner, "✨ Auto-fill with AI", self._autofill_with_api,
                 width=20, bg="#1a7f37", hover="#155a2e").pack(side="right")
        else:
            tk.Label(inner, text="🤖  Better results with AI:", bg="#EBF3FB",
                     fg=ACCENT2, font=FONT_BOLD).pack(side="left")
            tk.Label(inner, text="Use the clipboard bridge, or configure an API key in Settings ⚙.",
                     bg="#EBF3FB", fg=LABEL_FG, font=FONT_SM).pack(side="left", padx=8)
            _btn(inner, "Use AI ✨", self._open_ai_dialog,
                 width=14, bg=COP_BG, hover=COP_HOVER).pack(side="right")

    def _open_settings(self):
        def on_save(cfg):
            self._config = cfg
            save_config(cfg)
            self._rebuild_ai_row()
            self._status.config(text="✔  Settings saved.")
        SettingsDialog(self, self._config, on_save)

    def _autofill_with_api(self):
        path = self._source_path.get()
        if not path or path == "No file selected":
            messagebox.showwarning("No file", "Please browse to a CV file first.")
            return
        self._status.config(text="Extracting CV text…"); self.update()
        try:
            raw = self._load_raw_text()
        except Exception as ex:
            messagebox.showerror("Extraction failed", str(ex))
            self._status.config(text=""); return

        prompt = build_copilot_prompt(raw)
        self._status.config(text="Sending to AI — please wait…"); self.update()

        def run():
            try:
                reply  = call_ai_api(self._config, prompt)
                data   = parse_copilot_response(reply)
                self.after(0, lambda: self._on_api_result(data))
            except Exception as ex:
                self.after(0, lambda: self._on_api_error(str(ex)))

        threading.Thread(target=run, daemon=True).start()

    def _on_api_result(self, data):
        self._populate_form(data)
        self._status.config(text="✔  Form filled by AI — review all tabs before generating.")

    def _on_api_error(self, msg):
        messagebox.showerror("AI error", msg)
        self._status.config(text="")

    # ── actions ───────────────────────────────────────────────────────────

    def _browse_source(self):
        path = filedialog.askopenfilename(
            title="Select candidate CV",
            filetypes=[("CV files", "*.pdf *.docx *.doc"), ("All files", "*.*")]
        )
        if path:
            self._source_path.set(path)
            self._raw_cv_text = ""   # reset cached text

    def _load_raw_text(self):
        """Extract raw text from the selected CV file (cached)."""
        if self._raw_cv_text:
            return self._raw_cv_text
        path = self._source_path.get()
        if not path or path == "No file selected":
            return ""
        from parser import extract_text
        self._raw_cv_text = extract_text(path)
        return self._raw_cv_text

    def _parse_cv(self):
        path = self._source_path.get()
        if not path or path == "No file selected":
            messagebox.showwarning("No file", "Please browse to a CV file first.")
            return
        if not os.path.exists(path):
            messagebox.showerror("File not found", f"Cannot find:\n{path}")
            return

        self._status.config(text="Parsing…")
        self.update()
        try:
            from parser import parse_cv
            data = parse_cv(path)
        except Exception as ex:
            messagebox.showerror("Parse error", str(ex))
            self._status.config(text="")
            return

        self._populate_form(data)

        warning = data.get("_warning")
        if warning:
            messagebox.showwarning("Partial parse – please fill in manually", warning)
            self._status.config(text="⚠ Sections not detected — try the Copilot button for better results.")
        else:
            self._status.config(text="✔ CV parsed — review fields, or use Copilot ✨ for better extraction.")

    def _open_ai_dialog(self):
        path = self._source_path.get()
        if not path or path == "No file selected":
            messagebox.showwarning("No file", "Please browse to a CV file first.")
            return
        if not os.path.exists(path):
            messagebox.showerror("File not found", f"Cannot find:\n{path}")
            return

        self._status.config(text="Extracting CV text…")
        self.update()
        raw = ""
        try:
            raw = self._load_raw_text()
        except Exception:
            pass  # Let the dialog handle it — user can paste text manually

        self._status.config(text="")
        CopilotDialog(self, raw, on_data_ready=self._populate_form)

    def _populate_form(self, data):
        """Fill all form fields with parsed data."""
        def _set_entry(widget, val):
            widget.delete(0, tk.END)
            widget.insert(0, val or "")

        def _set_text(widget, val):
            widget.delete("1.0", tk.END)
            widget.insert("1.0", val or "")

        _set_entry(self._name,  data.get("name", ""))
        _set_entry(self._title, data.get("job_title", ""))
        _set_text(self._summary, data.get("summary", ""))

        tech = data.get("technologies", {})
        for key, widget in self._tech.items():
            val = tech.get(key, "")
            if isinstance(widget, tk.Text):
                _set_text(widget, val)
            else:
                _set_entry(widget, val)

        self._edu_tab.set_entries(data.get("education", []))

        langs = data.get("languages", [])
        if langs and isinstance(langs[0], dict):
            lang_vals = [f"{l['language']}: {l['level']}" for l in langs]
        else:
            lang_vals = [str(l) for l in langs]
        self._langs.set_values(lang_vals)

        certs = data.get("certifications", [])
        self._certs.set_values([str(c) for c in certs])

        self._we_tab.set_jobs(data.get("work_experience", []))

        self._status.config(text="✔ Form filled — review all tabs before generating.")

    def _collect_data(self):
        """Collect all form fields into a data dict."""
        tech = {}
        for key, widget in self._tech.items():
            if isinstance(widget, tk.Text):
                tech[key] = widget.get("1.0", tk.END).strip()
            else:
                tech[key] = widget.get().strip()

        # Parse language strings "Language: Level"
        langs = []
        for raw in self._langs.get_values():
            if ":" in raw:
                parts = raw.split(":", 1)
                langs.append({"language": parts[0].strip(), "level": parts[1].strip()})
            else:
                langs.append({"language": raw.strip(), "level": "Fluent"})

        return {
            "name":          self._name.get().strip(),
            "job_title":     self._title.get().strip(),
            "summary":       self._summary.get("1.0", tk.END).strip(),
            "technologies":  tech,
            "education":     self._edu_tab.get_entries(),
            "languages":     langs,
            "certifications":self._certs.get_values(),
            "work_experience":self._we_tab.get_jobs(),
        }

    def _generate(self):
        data = self._collect_data()
        if not data.get("job_title") and not data.get("name"):
            messagebox.showwarning("Empty form", "Please parse a CV or fill in the fields first.")
            return

        # Suggest filename
        confidential = self._confidential.get()
        name = data.get("name", "Candidate")
        parts = name.strip().split()
        initials = ".".join(p[0].upper() for p in parts if p) + "." if parts else "XX"
        suffix = " - Confidential" if confidential else ""
        default_name = f"{initials} DIGITALL CV{suffix}.docx"

        out_path = filedialog.asksaveasfilename(
            title="Save generated CV",
            initialfile=default_name,
            defaultextension=".docx",
            filetypes=[("Word document", "*.docx")],
        )
        if not out_path:
            return

        # Find template
        template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template.docx")
        if not os.path.exists(template_path):
            messagebox.showerror("Template missing",
                                 f"Cannot find template.docx next to this script.\n"
                                 f"Expected: {template_path}")
            return

        self._status.config(text="Generating…")
        self.update()
        try:
            from builder import build_cv
            build_cv(template_path, data, out_path, confidential=confidential)
        except Exception as ex:
            import traceback
            messagebox.showerror("Generation failed", traceback.format_exc())
            self._status.config(text="")
            return

        self._status.config(text=f"✔ Saved: {os.path.basename(out_path)}")
        if messagebox.askyesno("Done!", f"CV saved to:\n{out_path}\n\nOpen it now?"):
            os.startfile(out_path)


# ── entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
