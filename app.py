import os, io, json, sqlite3, smtplib
from email.message import EmailMessage
from datetime import datetime
from flask import Flask, request, redirect, url_for, render_template, flash
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
from docx import Document
from pypdf import PdfReader

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(APP_DIR, "review_bot.db")
UPLOAD_DIR = os.path.join(APP_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "change-this-before-deployment")
ALLOWED = {"pdf", "docx", "txt"}

def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with db() as c:
        c.execute("""CREATE TABLE IF NOT EXISTS interns (
            id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, email TEXT, phone TEXT,
            project_title TEXT, project_brief TEXT, stage TEXT DEFAULT 'Not started',
            review_date TEXT, review_time TEXT, review_link TEXT, last_feedback TEXT,
            completed INTEGER DEFAULT 0, created_at TEXT)""")
        c.execute("""CREATE TABLE IF NOT EXISTS submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT, intern_id INTEGER, filepath TEXT,
            feedback TEXT, score REAL, submitted_at TEXT,
            FOREIGN KEY(intern_id) REFERENCES interns(id))""")
init_db()

def send_email(to, subject, body):
    """Email sends only when SMTP environment variables are configured."""
    host = os.getenv("SMTP_HOST")
    user = os.getenv("SMTP_USER")
    password = os.getenv("SMTP_PASSWORD")
    port = int(os.getenv("SMTP_PORT", "587"))
    if not all([host, user, password, to]):
        return False, "Email not sent: configure SMTP_HOST, SMTP_USER, SMTP_PASSWORD."
    msg = EmailMessage()
    msg["From"] = os.getenv("FROM_EMAIL", user)
    msg["To"] = to
    msg["Subject"] = subject
    msg.set_content(body)
    try:
        with smtplib.SMTP(host, port, timeout=20) as server:
            server.starttls()
            server.login(user, password)
            server.send_message(msg)
        return True, "Email sent."
    except Exception as e:
        return False, f"Email error: {e}"

def extract_text(path):
    ext = path.rsplit(".", 1)[-1].lower()
    try:
        if ext == "txt":
            return open(path, "r", encoding="utf-8", errors="ignore").read()
        if ext == "docx":
            doc = Document(path)
            return "\n".join(p.text for p in doc.paragraphs)
        if ext == "pdf":
            reader = PdfReader(path)
            return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as e:
        return f"Could not extract text: {e}"
    return ""

def review_submission(brief, submission):
    """Free local heuristic evaluator. It uses no paid API or external AI service.

    Scores are indicators for a mentor, not a definitive grade. It checks requirement
    keyword coverage, document structure, evidence indicators, and basic readability.
    """
    import re
    from collections import Counter

    stop_words = set("""a an and are as at be been being by can could did do does for from had has have
        he her hers him his how i if in into is it its may me might more most must my no not of on
        or our ours she should so some such than that the their theirs them then there these they
        this those through to too under up us very was we were what when where which while who why
        will with would you your yours project task intern requirement requirements expected include
        including use using used make made build built develop developed create created work working
        system application platform tool tools document report submission result results provide
        provided show shows explain explanation implement implementation demonstrate""".split())

    def tokens(text):
        return re.findall(r"[a-zA-Z][a-zA-Z0-9+#.-]*", (text or "").lower())

    brief_tokens = [w for w in tokens(brief) if (len(w) >= 3 or w in {"ai", "ml", "ui", "ux", "db", "api"}) and w not in stop_words]
    submission_tokens = tokens(submission)
    submission_set = set(submission_tokens)
    expected = set(brief_tokens)
    if not (brief or "").strip():
        report = {
            "score": 0, "coverage": 0, "structure": 0, "evidence": 0, "clarity": 0,
            "present": [], "missing": [], "strengths": [],
            "improvements": ["Add a clear project brief in the Excel Project Brief column before evaluating."],
            "word_count": len(submission_tokens), "status": "Brief required"
        }
        return 0, format_review(report), report
    if not (submission or "").strip():
        report = {
            "score": 0, "coverage": 0, "structure": 0, "evidence": 0, "clarity": 0,
            "present": [], "missing": sorted(expected)[:35], "strengths": [],
            "improvements": ["No readable text was extracted. Check that the PDF contains selectable text, or upload a DOCX/TXT file."],
            "word_count": 0, "status": "No readable text"
        }
        return 0, format_review(report), report

    present = sorted(expected & submission_set)
    missing = sorted(expected - submission_set)
    coverage = round(100 * len(present) / max(1, len(expected)))

    # Detect common report sections, including headings and common variants.
    lower_text = (submission or "").lower()
    section_groups = {
        "Introduction / objective": ["introduction", "objective", "problem statement", "overview"],
        "Method / implementation": ["methodology", "method", "implementation", "workflow", "algorithm"],
        "Tools / technology": ["technology", "technologies", "tools used", "software used", "tech stack"],
        "Results / testing": ["result", "results", "testing", "test cases", "evaluation", "output"],
        "Conclusion / next steps": ["conclusion", "future work", "future enhancement", "recommendation"]
    }
    found_sections = [name for name, variants in section_groups.items() if any(v in lower_text for v in variants)]
    structure = round(100 * len(found_sections) / len(section_groups))

    evidence_groups = {
        "test or validation evidence": ["test case", "tested", "testing", "validation", "passed", "failed", "unit test"],
        "output or result evidence": ["screenshot", "output", "result", "accuracy", "measurement", "log"],
        "technical implementation evidence": ["code", "function", "api", "database", "architecture", "module", "configuration"],
        "examples or data evidence": ["example", "sample", "dataset", "input", "expected", "actual"]
    }
    found_evidence = [name for name, variants in evidence_groups.items() if any(v in lower_text for v in variants)]
    evidence = round(100 * len(found_evidence) / len(evidence_groups))

    word_count = len(submission_tokens)
    paragraphs = [x.strip() for x in re.split(r"\n\s*\n", submission or "") if x.strip()]
    if word_count >= 250 and len(paragraphs) >= 3:
        clarity = 100
    elif word_count >= 120:
        clarity = 75
    elif word_count >= 50:
        clarity = 50
    else:
        clarity = 25

    score = round(0.65 * coverage + 0.15 * structure + 0.15 * evidence + 0.05 * clarity)
    strengths = []
    if coverage >= 70: strengths.append("Many of the brief's key terms appear in the submission.")
    if len(found_sections) >= 3: strengths.append("The document includes several useful report sections.")
    if len(found_evidence) >= 2: strengths.append("Some technical or testing evidence indicators were found.")
    if not strengths: strengths.append("The submission was received and text was extracted for review.")

    improvements = []
    if missing:
        improvements.append("Review the brief topics not detected: " + ", ".join(missing[:18]) + ("…" if len(missing) > 18 else "."))
    if len(found_sections) < 4:
        improvements.append("Add clear sections for introduction/objectives, methodology, tools, results/testing, and conclusion.")
    if len(found_evidence) < 3:
        improvements.append("Include evidence such as screenshots, test cases, sample inputs/outputs, logs, or implementation details.")
    if word_count < 120:
        improvements.append("The extracted submission is short. Add enough explanation for the mentor to verify the work.")
    improvements.append("Ask the mentor to verify correctness, completeness, originality, and whether the project runs as expected.")

    if score >= 80: status = "Strong preliminary coverage"
    elif score >= 60: status = "Moderate preliminary coverage"
    else: status = "Needs further review"
    report = {
        "score": score, "coverage": coverage, "structure": structure, "evidence": evidence,
        "clarity": clarity, "present": present[:35], "missing": missing[:35],
        "strengths": strengths, "improvements": improvements,
        "found_sections": found_sections, "found_evidence": found_evidence,
        "word_count": word_count, "status": status
    }
    return score, format_review(report), report


def format_review(report):
    """Create a plain-text report that can also be sent by email."""
    return "\n".join([
        "AUTOMATED PRELIMINARY REVIEW — LOCAL PYTHON HEURISTICS",
        "This report is an indicator only. A mentor must verify the submission.",
        "",
        f"Overall preliminary score: {report['score']} / 100",
        f"Status: {report['status']}",
        f"Requirement keyword coverage: {report['coverage']}%",
        f"Document structure: {report['structure']}%",
        f"Evidence indicators: {report['evidence']}%",
        f"Basic readability indicator: {report['clarity']}%",
        f"Extracted word count: {report['word_count']}",
        "",
        "Strengths:", *[f"- {x}" for x in report['strengths']],
        "",
        "Potentially missing brief topics:",
        (", ".join(report['missing']) if report['missing'] else "No missing brief keywords detected."),
        "",
        "Recommended improvements:", *[f"- {x}" for x in report['improvements']],
        "",
        "Scoring guide: 65% brief keyword coverage, 15% document structure, 15% evidence indicators, 5% basic readability.",
        "Limitations: this does not prove project quality, working code, plagiarism, or actual understanding. Keyword matches can miss synonyms and can be present without correct implementation."
    ])

def review_link_for(intern_id):
    return request.url_root.rstrip("/") + url_for("submit_project", intern_id=intern_id)

@app.route("/")
def index():
    with db() as c:
        interns = c.execute("SELECT * FROM interns ORDER BY id DESC").fetchall()
        count = c.execute("SELECT COUNT(*) n FROM interns").fetchone()["n"]
        done = c.execute("SELECT COUNT(*) n FROM interns WHERE completed=1").fetchone()["n"]
    return render_template("index.html", interns=interns, count=count, done=done)

@app.route("/upload_excel", methods=["POST"])
def upload_excel():
    f = request.files.get("excel")
    if not f or not f.filename.lower().endswith((".xlsx", ".xlsm")):
        flash("Upload an .xlsx Excel file.")
        return redirect(url_for("index"))
    try:
        wb = load_workbook(f, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            raise ValueError("The spreadsheet is empty.")
        headers = [str(x or "").strip().lower().replace(" ", "_") for x in rows[0]]
        aliases = {
            "name": ["name", "intern_name", "student_name"],
            "email": ["email", "email_address"],
            "phone": ["phone", "phone_number", "mobile"],
            "project_title": ["project_title", "project", "title"],
            "project_brief": ["project_brief", "brief", "original_project_document", "requirements"]
        }
        indexes = {}
        for key, options in aliases.items():
            for opt in options:
                if opt in headers:
                    indexes[key] = headers.index(opt); break
        if "name" not in indexes or "email" not in indexes:
            raise ValueError("Excel must contain at least Name and Email columns.")
        imported = 0
        with db() as c:
            for row in rows[1:]:
                if not row or not row[indexes["email"]] or not row[indexes["name"]]:
                    continue
                get = lambda key, default="": str(row[indexes[key]] or default).strip() if key in indexes and indexes[key] < len(row) else default
                c.execute("""INSERT INTO interns(name,email,phone,project_title,project_brief,created_at)
                    VALUES(?,?,?,?,?,?)""",
                    (get("name"), get("email"), get("phone"), get("project_title", "Intern project"),
                     get("project_brief"), datetime.now().isoformat(timespec="minutes")))
                imported += 1
        flash(f"Imported {imported} interns. Review their details before sending emails.")
    except Exception as e:
        flash(f"Excel import failed: {e}")
    return redirect(url_for("index"))

@app.route("/intern/<int:intern_id>/task", methods=["POST"])
def send_task(intern_id):
    with db() as c:
        i = c.execute("SELECT * FROM interns WHERE id=?", (intern_id,)).fetchone()
    if not i:
        flash("Intern not found."); return redirect(url_for("index"))
    link = review_link_for(intern_id)
    body = f"""Hello {i['name']},

Welcome to the project review process.

Project: {i['project_title']}

Project brief / expected requirements:
{i['project_brief'] or 'Your mentor will provide the project requirements separately.'}

Upload your project document at this review link:
{link}

Please submit a PDF, DOCX, or TXT file. The automated feedback is preliminary; your mentor will validate it.

Regards,
Internship Review Team"""
    ok, msg = send_email(i["email"], f"Project task: {i['project_title']}", body)
    with db() as c:
        c.execute("UPDATE interns SET stage='Task sent' WHERE id=?", (intern_id,))
    flash(msg)
    return redirect(url_for("index"))

@app.route("/intern/<int:intern_id>/schedule", methods=["POST"])
def schedule(intern_id):
    date = request.form.get("date", "").strip()
    time = request.form.get("time", "").strip()
    if not date or not time:
        flash("Choose a review date and time."); return redirect(url_for("index"))
    with db() as c:
        i = c.execute("SELECT * FROM interns WHERE id=?", (intern_id,)).fetchone()
        c.execute("UPDATE interns SET review_date=?, review_time=?, stage='Review scheduled' WHERE id=?",
                  (date, time, intern_id))
    link = review_link_for(intern_id)
    body = f"""Hello {i['name']},

Your project review is scheduled for:
Date: {date}
Time: {time}

Submit your latest project document using this link:
{link}

Please upload before or during your review slot. Keep your project brief and implementation evidence ready.

Regards,
Internship Review Team"""
    ok, msg = send_email(i["email"], "Your project review is scheduled", body)
    flash(msg if ok else f"Schedule saved. {msg}")
    return redirect(url_for("index"))

@app.route("/submit/<int:intern_id>", methods=["GET", "POST"])
def submit_project(intern_id):
    with db() as c:
        intern = c.execute("SELECT * FROM interns WHERE id=?", (intern_id,)).fetchone()
    if not intern:
        return "Invalid review link", 404
    if request.method == "POST":
        f = request.files.get("document")
        if not f or "." not in f.filename or f.filename.rsplit(".",1)[-1].lower() not in ALLOWED:
            flash("Please upload a PDF, DOCX, or TXT file.")
            return redirect(request.url)
        safe = secure_filename(f.filename)
        path = os.path.join(UPLOAD_DIR, f"intern{intern_id}_{datetime.now().strftime('%Y%m%d%H%M%S')}_{safe}")
        f.save(path)
        text = extract_text(path)
        score, feedback, report = review_submission(intern["project_brief"], text)
        with db() as c:
            c.execute("INSERT INTO submissions(intern_id,filepath,feedback,score,submitted_at) VALUES(?,?,?,?,?)",
                      (intern_id, path, feedback, score, datetime.now().isoformat(timespec="minutes")))
            c.execute("UPDATE interns SET stage='Feedback ready',last_feedback=? WHERE id=?", (feedback, intern_id))
        email_body = f"Hello {intern['name']},\n\nYour project submission has been reviewed by the local preliminary evaluation tool.\n\n{feedback}\n\nPlease discuss the report with your mentor. The score is an indicator, not a final grade.\n\nRegards,\nInternship Review Team"
        email_ok, email_msg = send_email(intern["email"], f"Preliminary project feedback: {intern['project_title']}", email_body)
        email_status = email_msg if email_ok else f"Report saved, but email was not sent. {email_msg}"
        return render_template("feedback.html", intern=intern, feedback=feedback, score=score, report=report, email_status=email_status)
    return render_template("submit.html", intern=intern)

@app.route("/intern/<int:intern_id>/feedback_email", methods=["POST"])
def feedback_email(intern_id):
    with db() as c:
        i = c.execute("SELECT * FROM interns WHERE id=?", (intern_id,)).fetchone()
    if not i: flash("Intern not found."); return redirect(url_for("index"))
    body = f"""Hello {i['name']},

Review feedback:
{i['last_feedback'] or 'No submission feedback is available yet.'}

Next review date/time: {i['review_date'] or 'Not scheduled'} {i['review_time'] or ''}
Submit your corrected document here: {review_link_for(intern_id)}

Please address the feedback and upload the next version. The mentor should confirm completion before the process is stopped.

Regards,
Internship Review Team"""
    ok, msg = send_email(i["email"], "Project review feedback and next steps", body)
    flash(msg)
    return redirect(url_for("index"))

@app.route("/intern/<int:intern_id>/complete", methods=["POST"])
def complete(intern_id):
    with db() as c:
        c.execute("UPDATE interns SET completed=1,stage='Completed' WHERE id=?", (intern_id,))
    flash("Marked complete. Automated review cycle will stop for this intern.")
    return redirect(url_for("index"))

@app.route("/intern/<int:intern_id>/reopen", methods=["POST"])
def reopen(intern_id):
    with db() as c:
        c.execute("UPDATE interns SET completed=0,stage='In progress' WHERE id=?", (intern_id,))
    flash("Review cycle reopened.")
    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(debug=True)
