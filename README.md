# MCQ Exam System

A Flask + SQLite browser application for sequential MCQ testing.

## Features

- Student name and roll number captured before start.
- Previous test history can be searched by student name or roll number.
- Admin can view one overall report containing all students sorted by score, with links to individual reports.
- Multiple ordered test sections.
- Multiple saved tests can be created and managed independently by the admin.
- Students choose which saved test to take before entering the exam.
- Admin can set an optional overall test duration and edit each section's duration.
- Students can choose whether an attempt is kept in history; the choice is shown when the test starts.
- Each section has its own countdown timer.
- Every question is compulsory.
- One question at a time; after submission, an answer cannot be changed.
- Images can be uploaded for the question and/or any answer option.
- Passing mark is 60% per section.
- A student proceeds to the next section only after passing the current section.
- Failure in a section ends the test and shows the result report.
- Before starting, a student can choose to complete all sections or stop when a section is failed.
- Expiry of a section timer ends the test and shows the report.
- SQLite database stores students, attempts, responses, sections and questions.
- Password-protected admin panel can add/delete sections and questions.
- One normal importer can create questions from one or more structured `.docx` Word and/or text-based `.pdf` files in the same test.
- Claude AI import can extract unstructured MCQs, answer keys, and embedded Word images.

## Run in a browser

### Windows

```bat
cd mcq_exam_app
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python seed.py
python app.py
```

### macOS / Linux

```bash
cd mcq_exam_app
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python seed.py
python app.py
```

Then open:

- Student test: http://127.0.0.1:5000/
- Admin: http://127.0.0.1:5000/admin

Development admin password: `admin123`

For safer local/shared use, set environment variables before starting:

Alternatively, for local development, copy/edit `.env.example` as `.env`. The app loads this file automatically, so no terminal environment-variable commands are needed. Never commit or share `.env` because it contains credentials.

Windows PowerShell:

```powershell
$env:ADMIN_PASSWORD="choose-a-strong-password"
$env:SECRET_KEY="choose-a-long-random-secret"
$env:CLAUDE_API_KEY="your-claude-api-key"
# Optional, if you want to select a different installed Claude model:
$env:CLAUDE_MODEL="claude-sonnet-4-5-20250929"
python app.py
```

macOS/Linux:

```bash
export ADMIN_PASSWORD="choose-a-strong-password"
export SECRET_KEY="choose-a-long-random-secret"
export CLAUDE_API_KEY="your-claude-api-key"
python app.py
```

## Word/PDF import format

Create a `.docx` file where each line/paragraph follows this format:

```text
SECTION: General Knowledge
TIME: 10
Q: Capital of Pakistan?
A: Lahore
B: Karachi
C: Islamabad
D: Quetta
ANSWER: C

Q: 2 + 2 equals?
A: 3
B: 4
C: 5
D: 6
ANSWER: B

SECTION: Science
TIME: 15
Q: Water freezes at?
A: 0 C
B: 10 C
C: 50 C
D: 100 C
ANSWER: A
```

`TIME` is in minutes. `SECTION` and `TIME` can be repeated when a new section begins.

The normal importer extracts text questions, answers, and embedded PNG/JPG/GIF/WEBP images from Word and text-based PDFs. Place an image in the same paragraph as its `Q:`, `A:`, `B:`, `C:`, or `D:` line, or immediately beneath that line; it will be attached to that question or option automatically. A scanned/image-only PDF requires OCR and cannot be read by normal import.

## AI Word import

Set `CLAUDE_API_KEY` before starting the app, then use **AI import Word (.docx)** in the admin panel. It accepts normal Word exam papers rather than the strict format above, sends its text and up to 20 embedded images to Claude, and creates sections, questions, options, answer keys, and recognised question/option images. Check the imported answer key before publishing a high-stakes exam.

## Notes for production

This is a complete working starter application, but for deployment on a school network or the public internet you should add CSRF protection, proper admin accounts, HTTPS, backups, stronger access controls, and a production WSGI server.
