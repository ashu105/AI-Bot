# AI Intern Review Bot

## What it does
- Imports an Excel `.xlsx` roster for many interns.
- Stores name, email, phone, project title and original project brief.
- Creates an individual submission URL for each intern.
- Sends task emails and review schedule emails when SMTP is configured.
- Checks a configured Gmail inbox every minute for unread intern emails with project attachments.
- Accepts PDF, DOCX or TXT project documents.
- Compares extracted text with the original brief using transparent keyword overlap.
- Shows preliminary feedback and automatically emails it, including the currently scheduled next review date and submission link, when SMTP is configured.
- Includes a local, rule-based workflow assistant for questions about importing, scheduling, submissions, and completion.
- Lets a mentor mark an intern complete; completed interns cannot submit further documents or be scheduled for another review.

## Important limitations
This is a local workflow prototype, not an LLM-powered chatbot or production-grade autonomous evaluator. The assistant uses fixed rule-based replies, and the reviewer uses keyword matching; neither validates code execution, plagiarism, security, or actual feature correctness. Review links are based on sequential database IDs and submissions are not restricted to a time window. A human mentor must verify feedback and approve completion. Add authentication, signed/expiring links, time-window rules, file-size limits, and other production safeguards before deployment. Do not upload sensitive student data to an unapproved hosting provider.

## Excel columns
Required: `Name`, `Email`
Optional: `Phone`, `Project Title`, `Project Brief`
Use one row per intern. Put the original requirements in Project Brief, or provide a project-specific brief through an approved workflow.

## Run locally
1. Install Python 3.10+.
2. Open a terminal in this folder.
3. Run:
   ```bash
   python -m venv .venv
   # Windows:
   .venv\\Scripts\\activate
   # macOS/Linux:
   source .venv/bin/activate
   pip install -r requirements.txt
   python app.py
   ```
4. Open `http://127.0.0.1:5000`.

## Enable email
Set `SMTP_HOST`, `SMTP_PORT` (defaults to `587`), `SMTP_USER`, `SMTP_PASSWORD`, and optionally `FROM_EMAIL` as environment variables. To monitor Gmail, enable IMAP in Gmail settings and set `IMAP_USER` and `IMAP_PASSWORD`; the app falls back to the SMTP username and password when these are omitted. Gmail can use the same newly generated App Password for SMTP and IMAP. Set `PUBLIC_BASE_URL` to the externally reachable root URL where this app is deployed, so the feedback email contains a working next-review link. Without SMTP, feedback is reviewed and saved but cannot be emailed. Without IMAP credentials or the public URL, inbox monitoring is disabled and the app reports the configuration issue. Never paste a password into chat or commit it into project files; if a password has been shared, revoke it and generate a replacement. In production, use secure secrets, authentication, rate limits, audit logs, and bounce handling.

The app checks unread inbox messages once per minute and also provides a manual **Check inbox now** action. It matches the sender address to the imported intern email, processes PDF/DOCX/TXT attachments, marks handled messages as read, and prevents processing the same message twice. Unsupported attachments and messages from unknown senders are skipped. Feedback uses keyword matching only and must be checked by a mentor.

## Suggested production AI review
Replace `review_submission()` with a controlled LLM workflow only after selecting and approving a provider. It should:
1. Receives only the assigned intern's original requirements and submission.
2. Produces structured feedback: requirement, evidence found, issue, severity, recommended correction, and mentor-confirmation flag.
3. Quotes page/section evidence from the submitted document.
4. Never marks a project complete automatically; a mentor approves completion.
5. Records model/version, timestamp, prompt version, and reviewer override.
