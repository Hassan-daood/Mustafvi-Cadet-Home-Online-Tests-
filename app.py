import os
import re
import hmac
import uuid
import json
import mimetypes
import sqlite3
import zipfile
from math import ceil
from io import BytesIO
import pymupdf as fitz
from datetime import datetime, timezone, timedelta
from functools import wraps
from pathlib import Path

from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from anthropic import Anthropic
from flask import (
    Flask, flash, redirect, render_template, request, send_file, session, url_for
)
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import UniqueConstraint, text
from sqlalchemy.exc import OperationalError
from werkzeug.utils import secure_filename
from werkzeug.security import check_password_hash, generate_password_hash

BASE_DIR = Path(__file__).resolve().parent


def load_local_env():
    """Load local development settings without ever committing credentials to source."""
    env_file = BASE_DIR / ".env"
    if not env_file.exists():
        return
    for raw_line in env_file.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key:
            os.environ.setdefault(key, value)


load_local_env()
UPLOAD_DIR = BASE_DIR / "static" / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
ADMIN_PASSWORD_FILE = BASE_DIR / "instance" / "admin_password.hash"

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "change-this-secret-key")
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///exam.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024
app.config["UPLOAD_FOLDER"] = str(UPLOAD_DIR)

db = SQLAlchemy(app)
PASS_PERCENTAGE = 60.0
TEST_TITLE = os.environ.get("TEST_NAME", "Mustafvi Cadet Home MCQ Examination")
ALLOWED_IMAGE_EXTENSIONS = {"png", "jpg", "jpeg", "gif", "webp"}
ALLOWED_DOC_EXTENSIONS = {"docx", "pdf"}
CLAUDE_MODEL = os.environ.get("CLAUDE_MODEL", "claude-sonnet-4-5-20250929")

PKT = timezone(timedelta(hours=5))


def utcnow():
    return datetime.utcnow()


def to_pkdt(utc_dt):
    if utc_dt is None:
        return None
    if utc_dt.tzinfo is not None:
        return utc_dt.astimezone(PKT)
    return utc_dt.replace(tzinfo=timezone.utc).astimezone(PKT)


app.jinja_env.filters["to_pkdt"] = to_pkdt
app.jinja_env.filters["strftime"] = lambda dt, fmt: dt.strftime(fmt) if dt else ""


class Test(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False, unique=True)
    duration_minutes = db.Column(db.Integer, nullable=False, default=0)
    duration_seconds = db.Column(db.Integer, nullable=False, default=0)
    expected_total_marks = db.Column(db.Integer, nullable=False, default=1)
    pass_percentage = db.Column(db.Float, nullable=False, default=60.0)
    created_at = db.Column(db.DateTime, default=utcnow, nullable=False)
    sections = db.relationship(
        "Section", backref="test", cascade="all, delete-orphan", lazy=True
    )


class Section(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    test_id = db.Column(db.Integer, db.ForeignKey("test.id"), nullable=True)
    title = db.Column(db.String(200), nullable=False)
    duration_minutes = db.Column(db.Integer, nullable=False, default=10)
    duration_seconds = db.Column(db.Integer, nullable=False, default=0)
    pass_percentage = db.Column(db.Float, nullable=False, default=60.0)
    order_no = db.Column(db.Integer, nullable=False, default=1)
    questions = db.relationship(
        "Question", backref="section", cascade="all, delete-orphan", lazy=True
    )


class Question(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    section_id = db.Column(db.Integer, db.ForeignKey("section.id"), nullable=False)
    text = db.Column(db.Text, nullable=False)
    image = db.Column(db.String(255), nullable=True)
    option_a = db.Column(db.Text, nullable=False)
    option_b = db.Column(db.Text, nullable=False)
    option_c = db.Column(db.Text, nullable=False)
    option_d = db.Column(db.Text, nullable=False)
    option_a_image = db.Column(db.String(255), nullable=True)
    option_b_image = db.Column(db.String(255), nullable=True)
    option_c_image = db.Column(db.String(255), nullable=True)
    option_d_image = db.Column(db.String(255), nullable=True)
    correct_option = db.Column(db.String(1), nullable=False)


class Student(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    roll_no = db.Column(db.String(100), nullable=False)
    class_name = db.Column(db.String(100), nullable=False, default="")
    created_at = db.Column(db.DateTime, default=utcnow, nullable=False)


class Attempt(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    test_id = db.Column(db.Integer, db.ForeignKey("test.id"), nullable=True)
    student_id = db.Column(db.Integer, db.ForeignKey("student.id"), nullable=False)
    started_at = db.Column(db.DateTime, default=utcnow, nullable=False)
    finished_at = db.Column(db.DateTime, nullable=True)
    status = db.Column(db.String(30), default="in_progress", nullable=False)
    stopped_reason = db.Column(db.String(100), nullable=True)
    save_history = db.Column(db.Boolean, nullable=False, default=True)
    student = db.relationship("Student", backref="attempts")
    test = db.relationship("Test", backref="attempts")


class SectionAttempt(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    attempt_id = db.Column(db.Integer, db.ForeignKey("attempt.id"), nullable=False)
    section_id = db.Column(db.Integer, db.ForeignKey("section.id"), nullable=False)
    started_at = db.Column(db.DateTime, default=utcnow, nullable=False)
    finished_at = db.Column(db.DateTime, nullable=True)
    percentage = db.Column(db.Float, nullable=True)
    passed = db.Column(db.Boolean, nullable=True)
    timed_out = db.Column(db.Boolean, default=False, nullable=False)
    attempt = db.relationship("Attempt", backref="section_attempts")
    section = db.relationship("Section")
    __table_args__ = (UniqueConstraint("attempt_id", "section_id", name="uq_attempt_section"),)


class Response(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    attempt_id = db.Column(db.Integer, db.ForeignKey("attempt.id"), nullable=False)
    question_id = db.Column(db.Integer, db.ForeignKey("question.id"), nullable=False)
    selected_option = db.Column(db.String(1), nullable=False)
    is_correct = db.Column(db.Boolean, nullable=False)
    answered_at = db.Column(db.DateTime, default=utcnow, nullable=False)
    question = db.relationship("Question")
    __table_args__ = (UniqueConstraint("attempt_id", "question_id", name="uq_attempt_question"),)


with app.app_context():
    db.create_all()
    if "pass_percentage" not in {row[1] for row in db.session.execute(text("PRAGMA table_info(test)")).all()}:
        db.session.execute(text("ALTER TABLE test ADD COLUMN pass_percentage FLOAT NOT NULL DEFAULT 60.0"))
        db.session.commit()
    inspector = db.inspect(db.engine)
    section_columns = {column["name"] for column in inspector.get_columns("section")}
    attempt_columns = {column["name"] for column in inspector.get_columns("attempt")}
    test_columns = {column["name"] for column in inspector.get_columns("test")}
    student_columns = {column["name"] for column in inspector.get_columns("student")}
    if "test_id" not in section_columns:
        with db.engine.begin() as connection:
            connection.execute(text("ALTER TABLE section ADD COLUMN test_id INTEGER"))
    if "test_id" not in attempt_columns:
        with db.engine.begin() as connection:
            connection.execute(text("ALTER TABLE attempt ADD COLUMN test_id INTEGER"))
    if "duration_minutes" not in test_columns:
        with db.engine.begin() as connection:
            connection.execute(text("ALTER TABLE test ADD COLUMN duration_minutes INTEGER NOT NULL DEFAULT 0"))
    if "duration_seconds" not in test_columns:
        with db.engine.begin() as connection:
            connection.execute(text("ALTER TABLE test ADD COLUMN duration_seconds INTEGER NOT NULL DEFAULT 0"))
    if "duration_seconds" not in section_columns:
        with db.engine.begin() as connection:
            connection.execute(text("ALTER TABLE section ADD COLUMN duration_seconds INTEGER NOT NULL DEFAULT 0"))
    if "pass_percentage" not in section_columns:
        with db.engine.begin() as connection:
            connection.execute(text("ALTER TABLE section ADD COLUMN pass_percentage FLOAT NOT NULL DEFAULT 60"))
    if "save_history" not in attempt_columns:
        with db.engine.begin() as connection:
            connection.execute(text("ALTER TABLE attempt ADD COLUMN save_history BOOLEAN NOT NULL DEFAULT 1"))
    if "class_name" not in student_columns:
        try:
            with db.engine.begin() as connection:
                connection.execute(text("ALTER TABLE student ADD COLUMN class_name VARCHAR(100) NOT NULL DEFAULT ''"))
        except OperationalError:
            current_student_columns = {
                row[1] for row in db.session.execute(text("PRAGMA table_info(student)")).all()
            }
            if "class_name" not in current_student_columns:
                raise
    if "expected_total_marks" not in test_columns:
        try:
            with db.engine.begin() as connection:
                connection.execute(text("ALTER TABLE test ADD COLUMN expected_total_marks INTEGER NOT NULL DEFAULT 1"))
        except OperationalError:
            current_test_columns = {
                column["name"] for column in db.inspect(db.engine).get_columns("test")
            }
            if "expected_total_marks" not in current_test_columns:
                raise
        for existing_test in Test.query.all():
            question_count = (
                db.session.query(db.func.count(Question.id))
                .join(Section, Question.section_id == Section.id)
                .filter(Section.test_id == existing_test.id)
                .scalar()
            )
            existing_test.expected_total_marks = max(question_count or 0, 1)
        db.session.commit()
    legacy_test = Test.query.order_by(Test.id.asc()).first()
    if not legacy_test and Section.query.count():
        legacy_test = Test(title="Legacy Test")
        db.session.add(legacy_test)
        db.session.flush()
    if legacy_test:
        Section.query.filter(Section.test_id.is_(None)).update(
            {Section.test_id: legacy_test.id}, synchronize_session=False
        )
        for attempt in Attempt.query.filter(Attempt.test_id.is_(None)).all():
            section_attempt = SectionAttempt.query.filter_by(attempt_id=attempt.id).first()
            attempt.test_id = section_attempt.section.test_id if section_attempt and section_attempt.section else legacy_test.id
        db.session.commit()


def allowed_file(filename, allowed):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in allowed


def save_image(file_storage):
    if not file_storage or not file_storage.filename:
        return None
    if not allowed_file(file_storage.filename, ALLOWED_IMAGE_EXTENSIONS):
        raise ValueError("Image must be PNG, JPG, JPEG, GIF or WEBP.")
    ext = secure_filename(file_storage.filename).rsplit(".", 1)[1].lower()
    filename = f"{uuid.uuid4().hex}.{ext}"
    file_storage.save(UPLOAD_DIR / filename)
    return filename


def save_embedded_docx_image(image):
    """Save a supported image extracted from a .docx file."""
    if not image:
        return None
    blob, media_type = image
    extension = {
        "image/png": "png", "image/jpeg": "jpg", "image/gif": "gif", "image/webp": "webp",
    }.get(media_type)
    if not extension:
        return None
    filename = f"{uuid.uuid4().hex}.{extension}"
    (UPLOAD_DIR / filename).write_bytes(blob)
    return filename


def get_attempt():
    attempt_id = session.get("attempt_id")
    if not attempt_id:
        return None
    return db.session.get(Attempt, attempt_id)


def ordered_sections(test_id=None):
    query = Section.query
    if test_id is not None:
        query = query.filter_by(test_id=test_id)
    return query.order_by(Section.order_no.asc(), Section.id.asc()).all()


def student_roll_sort_key(student):
    roll_no = student.roll_no.strip()
    roll_key = (0, int(roll_no)) if roll_no.isdigit() else (1, roll_no.casefold())
    return roll_key, student.class_name.casefold(), student.id


def total_test_seconds(test):
    return sum(
        section.duration_minutes * 60 + section.duration_seconds
        for section in test.sections
    ) if test else 0


def section_total_marks(section):
    return len(section.questions)


def get_section_passing_marks(section):
    return ceil(section_total_marks(section) * section.pass_percentage / 100 - 1e-7)


def test_marks_summary(test):
    if not test:
        return {"total_marks": 0, "passing_marks": 0, "pass_percentage": 0.0}
    sections = ordered_sections(test.id)
    total_marks = sum(section_total_marks(section) for section in sections)
    passing_marks = sum(get_section_passing_marks(section) for section in sections)
    pass_percentage = (
        sum(section.pass_percentage * section_total_marks(section) for section in sections) / total_marks
        if total_marks else test.pass_percentage
    )
    return {
        "total_marks": total_marks,
        "passing_marks": passing_marks,
        "pass_percentage": pass_percentage,
    }


def report_test_details(test):
    if not test:
        return None
    summary = test_marks_summary(test)
    return {
        "title": test.title,
        "total_marks": summary["total_marks"],
        "passing_marks": summary["passing_marks"],
        "pass_percentage": summary["pass_percentage"],
        "sections": [{
            "title": section.title,
            "total_marks": section_total_marks(section),
            "passing_marks": get_section_passing_marks(section),
            "pass_percentage": section.pass_percentage,
        } for section in ordered_sections(test.id)],
    }


def unique_attempt_tests(rows):
    tests = []
    seen_ids = set()
    for row in rows:
        test = row["attempt"].test
        if test and test.id not in seen_ids:
            seen_ids.add(test.id)
            tests.append(test)
    return tests


def test_score(attempt):
    total = Question.query.join(Section).filter(Section.test_id == attempt.test_id).count()
    correct = (
        db.session.query(Response)
        .join(Question, Response.question_id == Question.id)
        .join(Section, Question.section_id == Section.id)
        .filter(Response.attempt_id == attempt.id, Section.test_id == attempt.test_id, Response.is_correct.is_(True))
        .count()
    )
    return total, correct, round(correct / total * 100, 2) if total else 0


def attempt_overall_passed(attempt, score, pass_percentage):
    """Match the final result page: an attempt stopped by a failed section fails overall."""
    return (
        attempt.status != "in_progress"
        and score >= pass_percentage
        and attempt.stopped_reason == "All sections completed"
    )


def get_section_attempt(attempt_id, section_id):
    return SectionAttempt.query.filter_by(attempt_id=attempt_id, section_id=section_id).first()


def remaining_seconds(section_attempt):
    duration = section_attempt.section.duration_minutes * 60 + section_attempt.section.duration_seconds
    elapsed = int((utcnow() - section_attempt.started_at).total_seconds())
    return max(0, duration - elapsed)
def overall_remaining_seconds(attempt):
    duration = total_test_seconds(attempt.test)
    if duration <= 0:
        return None
    elapsed = int((utcnow() - attempt.started_at).total_seconds())
    return max(0, duration - elapsed)


def effective_remaining_seconds(attempt, section_attempt):
    section_remaining = remaining_seconds(section_attempt)
    overall_remaining = overall_remaining_seconds(attempt)
    return min(section_remaining, overall_remaining) if overall_remaining is not None else section_remaining


def start_section(attempt, section):
    sa = get_section_attempt(attempt.id, section.id)
    if not sa:
        sa = SectionAttempt(attempt_id=attempt.id, section_id=section.id, started_at=utcnow())
        db.session.add(sa)
        db.session.commit()
    return sa


def finish_exam(attempt, reason):
    attempt.status = "completed"
    attempt.finished_at = utcnow()
    attempt.stopped_reason = reason
    db.session.commit()


def finalise_section(attempt, section, timed_out=False):
    sa = get_section_attempt(attempt.id, section.id)
    if not sa:
        sa = start_section(attempt, section)

    if sa.finished_at:
        return sa

    total = Question.query.filter_by(section_id=section.id).count()
    correct = (
        db.session.query(Response)
        .join(Question, Response.question_id == Question.id)
        .filter(Response.attempt_id == attempt.id, Question.section_id == section.id, Response.is_correct.is_(True))
        .count()
    )
    percentage = (correct / total * 100.0) if total else 0.0
    sa.percentage = round(percentage, 2)
    required_marks = get_section_passing_marks(section)
    sa.passed = correct >= required_marks
    sa.finished_at = utcnow()
    sa.timed_out = timed_out
    db.session.commit()
    return sa


def next_section(attempt, section):
    sections = ordered_sections(attempt.test_id)
    current_index = next(i for i, item in enumerate(sections) if item.id == section.id)
    if current_index + 1 >= len(sections):
        return None
    following = sections[current_index + 1]
    start_section(attempt, following)
    return following


def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("admin_logged_in"):
            return redirect(url_for("admin_login"))
        return view(*args, **kwargs)
    return wrapped


# The standard class list. It is shown on the student dashboard and is kept in the
# database; deleting test history never removes these students.
CLASS_ROSTER = [
    ("1", "Fahad Majeed", "8th"),
    ("2", "M Irfan", "8th"),
    ("3", "Ali Hassan", "8th"),
    ("4", "Sultan Nabeel", "8th"),
    ("5", "M Ali Ahsan", "8th"),
    ("6", "Hassan Rizwan", "8th"),
    ("7", "Ali Ahmad", "8th"),
    ("8", "Abdul Qadir", "8th"),
    ("9", "M Shahzaib Asad", "8th"),
    ("10", "M Uabid Moeen", "8th"),
    ("11", "Muazam Ali", "8th"),
    ("12", "Shaheryar Ghuman", "8th"),
    ("13", "M Ayaz", "8th"),
    ("14", "Eissa Bin Rashid", "8th"),
    ("15", "Murad Abbas", "8th"),
    ("16", "Abdul Raheem", "8th"),
    ("17", "Abdul Hadi (Tariq)", "8th"),
    ("18", "M Asad", "8th"),
    ("19", "M Huzaifa Fareed", "8th"),
    ("20", "Zain Raza Atta", "7th"),
    ("21", "Noor Aalam", "7th"),
    ("22", "M Mahad Pahat", "7th"),
    ("23", "Abdullah Sabir", "7th"),
    ("24", "Abdul Muqeet", "7th"),
    ("25", "Ibrahim Asif", "7th"),
    ("26", "M Musa", "7th"),
    ("27", "M. Ali Ayaz (new)", "7th"),
    ("28", "Naeem Akram", "7th"),
    ("29", "Junaid Ali", "7th"),
    ("30", "M Fahad Ch", "7th"),
    ("31", "Taha Ali", "7th"),
    ("32", "Abubakar", "7th"),
    ("33", "Usman", "7th"),
    ("34", "M Zeeshan Mughal", "7th"),
]


def restore_class_roster():
    """Add class-list students that are missing. Never edits or removes anyone."""
    existing_names = set()
    existing_rolls = set()
    for student in Student.query.all():
        existing_names.add(student.name.strip().casefold())
        existing_rolls.add(student.roll_no.strip())
    added = []
    skipped = []
    for roll_no, name, class_name in CLASS_ROSTER:
        if name.casefold() in existing_names:
            continue
        if roll_no in existing_rolls:
            skipped.append((roll_no, name))
            continue
        db.session.add(Student(name=name, roll_no=roll_no, class_name=class_name))
        added.append((roll_no, name))
    if added:
        db.session.commit()
    return added, skipped


with app.app_context():
    # If the student table is empty (for example after the old "Delete all history"
    # removed everyone) the standard class list is loaded again automatically.
    if Student.query.count() == 0:
        restore_class_roster()


STUDENT_UNLOCK_MINUTES = int(os.environ.get("STUDENT_UNLOCK_MINUTES", "30"))


def valid_student_password(password):
    """STUDENT_PASSWORD (if set) opens the student pages; otherwise the admin password does."""
    student_password = os.environ.get("STUDENT_PASSWORD", "")
    if student_password:
        return hmac.compare_digest(password.encode("utf-8"), student_password.encode("utf-8"))
    return valid_admin_password(password)


def lock_student_access():
    session.pop("student_access_unlocked", None)
    session.pop("student_access_unlocked_at", None)


def student_access_ok():
    """True only while the student password was entered recently and not yet used to start a test."""
    if not session.get("student_access_unlocked"):
        return False
    unlocked_at = session.get("student_access_unlocked_at")
    now = datetime.now(timezone.utc).timestamp()
    if unlocked_at is None or now - unlocked_at > STUDENT_UNLOCK_MINUTES * 60:
        lock_student_access()
        return False
    return True


def valid_admin_password(password):
    if ADMIN_PASSWORD_FILE.exists():
        return check_password_hash(ADMIN_PASSWORD_FILE.read_text(encoding="utf-8").strip(), password)
    return password == os.environ.get("ADMIN_PASSWORD", "admin123")


@app.get("/")
def index():
    if not student_access_ok():
        return redirect(url_for("student_login"))
    tests = Test.query.order_by(Test.created_at.asc(), Test.id.asc()).all()
    selected_test_id = request.args.get("test_id", type=int)
    selected_test = db.session.get(Test, selected_test_id) if selected_test_id else (tests[0] if tests else None)
    return render_template(
        "student_landing.html",
        tests=tests,
        selected_test=selected_test,
    )


@app.route("/student/login", methods=["GET", "POST"])
def student_login():
    if request.method == "POST":
        if not valid_student_password(request.form.get("password", "")):
            flash("Incorrect password.", "error")
            return redirect(url_for("student_login"))
        session["student_access_unlocked"] = True
        session["student_access_unlocked_at"] = datetime.now(timezone.utc).timestamp()
        return redirect(url_for("index"))

    lock_student_access()
    return render_template("student_landing.html", password_gate=True)


@app.post("/student/register")
@admin_required
def register_student():
    name = request.form.get("name", "").strip()
    roll_no = request.form.get("roll_no", "").strip()
    class_name = request.form.get("class_name", "").strip()
    registration_url = url_for(
        "admin_student_registrations",
        class_filter=request.form.get("class_filter", "").strip(),
        q=request.form.get("q", "").strip(),
    )
    if not name or not class_name:
        flash("Student name and class are required.", "error")
        return redirect(registration_url)

    existing_student = Student.query.filter_by(name=name).order_by(Student.id.asc()).first()
    if existing_student:
        flash(f"Student is already registered with roll number {existing_student.roll_no}.", "success")
        return redirect(registration_url)

    if not roll_no:
        numeric_rolls = [
            int(existing.roll_no)
            for existing in Student.query.all()
            if existing.roll_no.strip().isdigit()
        ]
        roll_no = str(max(numeric_rolls, default=0) + 1)
    elif Student.query.filter_by(roll_no=roll_no).first():
        flash("That roll number is already assigned to another student.", "error")
        return redirect(registration_url)

    student = Student(name=name, roll_no=roll_no, class_name=class_name)
    db.session.add(student)
    db.session.commit()
    flash(f"Registration complete. Your roll number is {student.roll_no}.", "success")
    return redirect(registration_url)


@app.get("/admin/students")
@admin_required
def admin_student_registrations():
    class_filter = request.args.get("class_filter", "").strip()
    query = request.args.get("q", "").strip()
    sort_by = request.args.get("sort_by", "roll_asc")
    students_query = Student.query
    renumber_preview = sort_by == "renumber"
    if renumber_preview:
        class_filter = ""
        query = ""
    if class_filter:
        students_query = students_query.filter(Student.class_name == class_filter)
    if query:
        search = f"%{query}%"
        students_query = students_query.filter(
            db.or_(Student.name.ilike(search), Student.roll_no.ilike(search), Student.class_name.ilike(search))
        )
    students = students_query.order_by(Student.id.asc()).all()
    if sort_by == "class_roll":
        students.sort(key=lambda student: (
            student.class_name.casefold(),
            (0, int(student.roll_no)) if student.roll_no.strip().isdigit() else (1, student.roll_no.casefold()),
        ))
    elif sort_by == "renumber":
        students.sort(key=student_roll_sort_key)
    else:
        sort_by = "roll_asc"
        students.sort(key=lambda student: (
            (0, int(student.roll_no)) if student.roll_no.strip().isdigit() else (1, student.roll_no.casefold()),
            student.class_name.casefold(),
        ))
    classes = [
        class_name
        for (class_name,) in db.session.query(Student.class_name)
        .filter(Student.class_name != "")
        .distinct()
        .order_by(Student.class_name.asc())
        .all()
    ]
    return render_template(
        "student_registrations.html",
        students=students,
        classes=classes,
        class_filter=class_filter,
        query=query,
        sort_by=sort_by,
        renumber_preview=renumber_preview,
    )


@app.post("/admin/students/renumber")
@admin_required
def renumber_students():
    database_path = Path(db.engine.url.database)
    backup_path = database_path.with_name(
        f"exam.db.before-renumber-{datetime.now().strftime('%Y%m%d-%H%M%S-%f')}.bak"
    )
    with sqlite3.connect(str(database_path)) as source, sqlite3.connect(str(backup_path)) as backup:
        source.backup(backup)

    students = Student.query.order_by(Student.id.asc()).all()
    students.sort(key=student_roll_sort_key)
    for index, student in enumerate(students, start=1):
        student.roll_no = str(index)
    db.session.commit()
    flash(f"Roll numbers were reassigned sequentially from 1 to {len(students)}. Backup saved as {backup_path.name}.", "success")
    return redirect(url_for("admin_student_registrations", sort_by="roll_asc"))


@app.post("/admin/students/<int:student_id>/update")
@admin_required
def update_student(student_id):
    student = db.session.get(Student, student_id)
    if not student:
        flash("Student not found.", "error")
        return redirect(url_for("admin_student_registrations"))

    name = request.form.get("name", "").strip()
    roll_no = request.form.get("roll_no", "").strip()
    class_name = request.form.get("class_name", "").strip()
    registration_url = url_for(
        "admin_student_registrations",
        class_filter=request.form.get("class_filter", "").strip(),
        q=request.form.get("q", "").strip(),
    )
    if not name or not roll_no or not class_name:
        flash("Name, roll number, and class are required.", "error")
        return redirect(registration_url)
    duplicate_roll = Student.query.filter(Student.roll_no == roll_no, Student.id != student.id).first()
    if duplicate_roll:
        flash("That roll number is already assigned to another student.", "error")
        return redirect(registration_url)

    student.name = name
    student.roll_no = roll_no
    student.class_name = class_name
    db.session.commit()
    flash("Student details updated.", "success")
    return redirect(registration_url)


@app.post("/admin/students/<int:student_id>/delete")
@admin_required
def delete_student(student_id):
    student = db.session.get(Student, student_id)
    if not student:
        flash("Student not found.", "error")
        return redirect(url_for("admin_student_registrations"))

    attempt_ids = [attempt.id for attempt in student.attempts]
    if attempt_ids:
        Response.query.filter(Response.attempt_id.in_(attempt_ids)).delete(synchronize_session=False)
        SectionAttempt.query.filter(SectionAttempt.attempt_id.in_(attempt_ids)).delete(synchronize_session=False)
        Attempt.query.filter(Attempt.id.in_(attempt_ids)).delete(synchronize_session=False)
    db.session.delete(student)
    db.session.commit()
    flash("Student and their saved test history were deleted.", "success")
    return redirect(url_for("admin_student_registrations"))


@app.post("/student/setup")
def prepare_student_dashboard():
    if not student_access_ok():
        return redirect(url_for("student_login"))
    test_id = request.form.get("test_id", type=int)
    failure_mode = request.form.get("section_failure_mode", "stop")
    selected_test = db.session.get(Test, test_id)
    if not selected_test or failure_mode not in {"all", "stop"}:
        flash("Please choose a valid test and failure mode.", "error")
        return redirect(url_for("index"))
    session["student_setup_test_id"] = selected_test.id
    session["student_setup_failure_mode"] = failure_mode
    return redirect(url_for("student_dashboard", test_id=selected_test.id))


@app.get("/student/dashboard")
def student_dashboard():
    if not student_access_ok():
        return redirect(url_for("student_login"))
    tests = Test.query.order_by(Test.created_at.asc()).all()
    selected_test_id = request.args.get("test_id", type=int) or session.get("student_setup_test_id")
    selected_test = db.session.get(Test, selected_test_id) if selected_test_id else (tests[0] if tests else None)
    sections = ordered_sections(selected_test.id) if selected_test else []
    marks_summary = test_marks_summary(selected_test)
    total_marks = marks_summary["total_marks"]
    pass_percentage = marks_summary["pass_percentage"] if total_marks else PASS_PERCENTAGE
    passing_marks = marks_summary["passing_marks"]
    section_passing_marks = {
        section.id: get_section_passing_marks(section)
        for section in sections
    }
    section_pass_percentages = {
        section.id: section.pass_percentage
        for section in sections
    }
    students = sorted(
        Student.query.order_by(Student.id.asc()).all(),
        key=lambda student: (0, int(student.roll_no)) if student.roll_no.strip().isdigit() else (1, student.roll_no.casefold()),
    )
    return render_template(
        "index.html",
        tests=tests,
        selected_test=selected_test,
        sections=sections,
        selected_test_total_seconds=total_test_seconds(selected_test),
        total_marks=total_marks,
        passing_marks=passing_marks,
        section_passing_marks=section_passing_marks,
        section_pass_percentages=section_pass_percentages,
        pass_percentage=pass_percentage,
        section_failure_mode=session.get("student_setup_failure_mode", "stop"),
        students=students,
        form_name=session.pop("student_form_name", ""),
        form_roll_no=session.pop("student_form_roll_no", ""),
    )


CHART_HEIGHT = 216
DIST_CHART_HEIGHT = 360


def chart_percentage_ticks(plot_top, plot_height):
    """Y-axis ticks every 10% (0, 10, 20 ... 100) for every dashboard graph."""
    return [
        {"value": tick, "y": plot_top + (100 - tick) * plot_height / 100, "major": tick % 50 == 0}
        for tick in range(0, 101, 10)
    ]


def chart_tooltip_lines(lines, anchor_x, top_y, chart_width):
    """Instant tooltip with several text rows (an empty row leaves a blank gap)."""
    row_height = 15
    width = round(max(len(line) for line in lines) * 6.8 + 22, 1)
    x = min(max(anchor_x - width / 2, 4), chart_width - width - 4)
    y = max(top_y, 2)
    return {
        "x": round(x, 1),
        "y": round(y, 1),
        "width": width,
        "height": len(lines) * row_height + 10,
        "text_x": round(x + width / 2, 1),
        "lines": [
            {"text": line, "y": round(y + 2 + row_height * (index + 1), 1)}
            for index, line in enumerate(lines)
        ],
    }


def chart_tooltip(text, anchor_x, top_y, chart_width):
    """Geometry for an instant (CSS-hover) tooltip box, kept inside the chart."""
    width = round(len(text) * 6.8 + 18, 1)
    x = min(max(anchor_x - width / 2, 4), chart_width - width - 4)
    y = max(top_y, 2)
    return {
        "text": text,
        "x": round(x, 1),
        "y": round(y, 1),
        "width": width,
        "height": 22,
        "text_x": round(x + width / 2, 1),
        "text_y": round(y + 15, 1),
    }


def build_student_progress_analytics(attempt_rows):
    analytics = {
        "progress_student": None,
        "progress_points": [],
        "progress_line": "",
        "chart_ticks": [],
        "chart_width": 760,
        "chart_height": CHART_HEIGHT,
        "chart_mid_y": CHART_HEIGHT // 2,
        "fail_bands": [],
        "average_percentage": 0,
        "passed_attempts": 0,
        "passed_percentage": 0,
        "progress_total_tests": 0,
    }
    if not attempt_rows or len({row["attempt"].student_id for row in attempt_rows}) != 1:
        return analytics

    ordered_rows = sorted(attempt_rows, key=lambda row: row["attempt"].started_at)
    # Fixed width (not growing with attempt count) so this chart always fits
    # its card without needing a horizontal scroll, unlike the "all tests" chart.
    chart_width = 420
    plot_left = 60
    plot_right = 18
    plot_top = 18
    plot_bottom = 48
    plot_width = chart_width - plot_left - plot_right
    plot_height = CHART_HEIGHT - plot_top - plot_bottom
    progress_points = []
    fail_bands = []
    for index, row in enumerate(ordered_rows):
        x = plot_left + (plot_width / 2 if len(ordered_rows) == 1 else index * plot_width / (len(ordered_rows) - 1))
        score = max(0, min(100, row["score"]))
        pass_percentage = max(0, min(100, row["pass_percentage"]))
        threshold_y = plot_top + (100 - pass_percentage) * plot_height / 100
        band_left = plot_left if index == 0 else (progress_points[index - 1]["x"] + x) / 2
        if index == len(ordered_rows) - 1:
            band_right = plot_left + plot_width
        else:
            next_x = plot_left + (index + 1) * plot_width / (len(ordered_rows) - 1)
            band_right = (x + next_x) / 2
        fail_bands.append({
            "x": round(band_left, 2),
            "y": round(threshold_y, 2),
            "width": round(band_right - band_left, 2),
            "height": round(plot_top + plot_height - threshold_y, 2),
        })
        attempt_date = to_pkdt(row["attempt"].started_at)
        progress_points.append({
            "x": round(x, 2),
            "y": round(plot_top + (100 - score) * plot_height / 100, 2),
            "score": row["score"],
            "failed": row["attempt"].status != "in_progress" and row["score"] < row["pass_percentage"],
            "date": attempt_date.strftime("%d %b %Y"),
            "date_label": attempt_date.strftime("%d %b"),
            "attempt_id": row["attempt"].id,
            "tip": chart_tooltip_lines(
                (
                    [f"{row['total_correct']} / {row['total_questions']}", f"{row['score']}%"]
                    if row.get("total_questions") is not None else [f"{row['score']}%"]
                ),
                x, plot_top + (100 - score) * plot_height / 100 - (52 if row.get("total_questions") is not None else 32), chart_width,
            ),
        })
    passed_attempts = sum(
        row["score"] >= row["pass_percentage"]
        and row["attempt"].status != "in_progress"
        for row in ordered_rows
    )
    analytics.update({
        "progress_student": ordered_rows[0]["attempt"].student,
        "progress_points": progress_points,
        "fail_bands": fail_bands,
        "progress_line": " ".join(f"{point['x']},{point['y']}" for point in progress_points),
        "chart_ticks": chart_percentage_ticks(plot_top, plot_height),
        "chart_width": chart_width,
        "chart_height": CHART_HEIGHT,
        "chart_mid_y": CHART_HEIGHT // 2,
        "average_percentage": round(sum(row["score"] for row in ordered_rows) / len(ordered_rows), 1),
        "passed_attempts": passed_attempts,
        "passed_percentage": round(passed_attempts / len(ordered_rows) * 100, 1),
        "progress_total_tests": len(ordered_rows),
    })
    return analytics


def build_overall_progress_analytics(student_rows, selected_test, all_tests):
    # Every attempt row shown in the results table below counts here too,
    # so the average/passing-rate figures always match what the table shows.
    progress_rows = list(student_rows)
    progress_rows.sort(key=lambda row: (
        row["attempt"].test.title.casefold() if all_tests and row["attempt"].test else "",
        (0, int(row["attempt"].student.roll_no))
        if row["attempt"].student.roll_no.strip().isdigit()
        else (1, row["attempt"].student.roll_no.casefold()),
        row["attempt"].student.name.casefold(),
    ))
    total = len(progress_rows)
    top_score_row = max(progress_rows, key=lambda row: row["score"], default=None)
    low_score_row = min(progress_rows, key=lambda row: row["score"], default=None)
    average_percentage = round(sum(row["score"] for row in progress_rows) / total, 1) if total else 0
    passing_percentage = (
        test_marks_summary(selected_test)["pass_percentage"]
        if selected_test else
        round(sum(
            test_marks_summary(row["attempt"].test)["pass_percentage"] if row["attempt"].test else PASS_PERCENTAGE
            for row in progress_rows
        ) / total, 1) if total else 0
    )
    passed_count = sum(
        row["score"] >= row["pass_percentage"]
        and row["attempt"].status != "in_progress"
        for row in progress_rows
    )

    # "All tests" donut: the average of each individual test's own passing
    # rate (not every attempt pooled together), so tests with more attempts
    # don't skew the figure.
    tests_average_passing_rate = None
    if all_tests:
        rows_by_test = {}
        for row in progress_rows:
            rows_by_test.setdefault(row["attempt"].test_id, []).append(row)
        per_test_rates = []
        for test_rows in rows_by_test.values():
            test_total = len(test_rows)
            test_passed = sum(
                row["score"] >= row["pass_percentage"]
                and row["attempt"].status != "in_progress"
                for row in test_rows
            )
            per_test_rates.append(test_passed / test_total * 100 if test_total else 0)
        tests_average_passing_rate = round(sum(per_test_rates) / len(per_test_rates), 1) if per_test_rates else 0

    # Individual-test graph: one numbered X position and one actual-score dot per student.
    dist_plot_left, dist_plot_right, dist_plot_top, dist_plot_bottom = 60, 18, 18, 92
    dist_plot_height = DIST_CHART_HEIGHT - dist_plot_top - dist_plot_bottom
    dist_baseline = dist_plot_top + dist_plot_height
    dist_pass_percentage = test_marks_summary(selected_test)["pass_percentage"] if selected_test else PASS_PERCENTAGE
    pass_midpoint = dist_pass_percentage + (100 - dist_pass_percentage) / 2
    dist_chart_width = max(700, dist_plot_left + dist_plot_right + max(total, 1) * 48)
    dist_plot_width = dist_chart_width - dist_plot_left - dist_plot_right
    dist_chart_ticks = chart_percentage_ticks(dist_plot_top, dist_plot_height)
    score_distribution = []
    score_bands = []
    bands = [(0, dist_pass_percentage), (dist_pass_percentage, pass_midpoint), (pass_midpoint, 100)]
    band_colors = ["#e11d48", "#84cc16", "#047857"]
    for index, ((lo, hi), color) in enumerate(zip(bands, band_colors)):
        count = sum(1 for row in progress_rows if lo <= row["score"] < hi or (hi == 100 and row["score"] == 100))
        y_top = dist_plot_top + (100 - hi) * dist_plot_height / 100
        y_bottom = dist_plot_top + (100 - lo) * dist_plot_height / 100
        score_bands.append({
            "x": dist_plot_left, "y": round(y_top, 2), "width": dist_plot_width,
            "height": round(y_bottom - y_top, 2), "color": color, "count": count,
            "label_y": round((y_top + y_bottom) / 2 + 4, 2),
            "label": ["Fail", "Pass", "High score"][index], "min": lo, "max": hi,
        })
        score_distribution.append(score_bands[-1])
    for index, row in enumerate(progress_rows):
        score = max(0, min(100, row["score"]))
        x = dist_plot_left + (dist_plot_width / 2 if total == 1 else index * dist_plot_width / (total - 1))
        y = dist_plot_top + (100 - score) * dist_plot_height / 100
        score_distribution.append({
            "dot": True, "cx": round(x, 1), "cy": round(y, 1),
            "attempt_id": row["attempt"].id, "number": index + 1,
            "color": band_colors[0] if score < dist_pass_percentage else band_colors[1] if score < pass_midpoint else band_colors[2],
            "tip": chart_tooltip_lines(
                [row["attempt"].student.name, "", f"{row['correct']} / {row['total']}", "", f"{row['score']}%"],
                x, y - 100, dist_chart_width,
            ),
        })

    # Website chart is sized to its card so all students and the in-chart labels
    # fit without horizontal scrolling. Keep the document chart dimensions above.
    web_dist_chart_width = 760
    web_dist_plot_right = 135
    web_dist_plot_width = web_dist_chart_width - dist_plot_left - web_dist_plot_right
    web_score_distribution = []
    for item in score_distribution:
        if not item.get("dot"):
            web_score_distribution.append({
                **item,
                "width": web_dist_plot_width,
                "label_x": web_dist_chart_width - 14,
            })
            continue
        index = item["number"] - 1
        x = dist_plot_left + (web_dist_plot_width / 2 if total == 1 else index * web_dist_plot_width / (total - 1))
        y = item["cy"]
        web_score_distribution.append({
            **item,
            "cx": round(x, 1),
            "tip": chart_tooltip_lines(
                [progress_rows[index]["attempt"].student.name, "", f"{progress_rows[index]['correct']} / {progress_rows[index]['total']}", "", f"{progress_rows[index]['score']}%"],
                x, y - 100, web_dist_chart_width,
            ),
        })

    # Fixed width so this chart also always fits its card without scrolling,
    # regardless of how many students/tests are plotted.
    chart_width = 420
    plot_left, plot_right, plot_top, plot_bottom = 60, 18, 18, 48
    plot_width = chart_width - plot_left - plot_right
    plot_height = CHART_HEIGHT - plot_top - plot_bottom
    progress_points = []
    if all_tests:
        # One point per test: the average percentage across all students who
        # took that test (not one point per individual student attempt).
        tests_in_order = []
        rows_by_test_id = {}
        for row in progress_rows:
            test_id = row["attempt"].test_id
            if test_id not in rows_by_test_id:
                rows_by_test_id[test_id] = []
                tests_in_order.append(test_id)
            rows_by_test_id[test_id].append(row)
        test_count = len(tests_in_order)
        for index, test_id in enumerate(tests_in_order):
            test_rows = rows_by_test_id[test_id]
            test_obj = test_rows[0]["attempt"].test
            test_title = test_obj.title if test_obj else "Test"
            avg_score = round(sum(r["score"] for r in test_rows) / len(test_rows), 1)
            x = plot_left + (plot_width / 2 if test_count == 1 else index * plot_width / (test_count - 1))
            score = max(0, min(100, avg_score))
            progress_points.append({
                "x": round(x, 2),
                "y": round(plot_top + (100 - score) * plot_height / 100, 2),
                "score": avg_score,
                "label": str(index + 1),
                "name": test_title,
                "test_id": test_id,
                "tip": chart_tooltip(
                    f"{avg_score}%",
                    x, plot_top + (100 - score) * plot_height / 100 - 32, chart_width,
                ),
                "name_tip": chart_tooltip(test_title, x, CHART_HEIGHT - 74, chart_width),
                "detail": f"{test_title}: {avg_score}% average ({len(test_rows)} student{'' if len(test_rows) == 1 else 's'})",
            })
    else:
        for index, row in enumerate(progress_rows):
            attempt = row["attempt"]
            student = attempt.student
            x = plot_left + (plot_width / 2 if total == 1 else index * plot_width / (total - 1))
            score = max(0, min(100, row["score"]))
            test_title = attempt.test.title if attempt.test else "Test"
            progress_points.append({
                "x": round(x, 2),
                "y": round(plot_top + (100 - score) * plot_height / 100, 2),
                "score": row["score"],
                "label": student.roll_no,
                "detail": f"{student.name} · Roll {student.roll_no} · {test_title}: {row['score']}%",
            })

    return {
        "progress_rows": progress_rows,
        "progress_points": progress_points,
        "progress_line": " ".join(f"{point['x']},{point['y']}" for point in progress_points),
        "chart_ticks": chart_percentage_ticks(plot_top, plot_height),
        "chart_width": chart_width,
        "chart_height": CHART_HEIGHT,
        "chart_mid_y": CHART_HEIGHT // 2,
        "average_percentage": average_percentage,
        "overall_passing_percentage": round(passing_percentage, 1),
        "passed_count": passed_count,
        "progress_total_students": total,
        "top_score_student": ({
            "name": top_score_row["attempt"].student.name,
            "marks": top_score_row["correct"],
            "total_marks": top_score_row["total"],
            "percentage": top_score_row["score"],
        } if top_score_row else None),
        "low_score_student": ({
            "name": low_score_row["attempt"].student.name,
            "marks": low_score_row["correct"],
            "total_marks": low_score_row["total"],
            "percentage": low_score_row["score"],
        } if low_score_row else None),
        "passed_percentage": round(passed_count / total * 100, 1) if total else 0,
        "tests_average_passing_rate": tests_average_passing_rate,
        "score_distribution": score_distribution,
        "score_bands": score_bands,
        "web_score_distribution": web_score_distribution,
        "web_dist_chart_width": web_dist_chart_width,
        "dist_chart_width": dist_chart_width,
        "dist_chart_height": DIST_CHART_HEIGHT,
        "dist_baseline": dist_baseline,
        "dist_mid_y": dist_plot_top + dist_plot_height // 2,
        "dist_chart_ticks": dist_chart_ticks,
    }


@app.route("/history", methods=["GET", "POST"])
def history():
    name = request.values.get("name", "").strip()
    roll_no = request.values.get("roll_no", "").strip()
    tests = Test.query.order_by(Test.created_at.asc(), Test.id.asc()).all()
    selected_test_id = request.values.get("test_id", type=int)
    selected_test = db.session.get(Test, selected_test_id) if selected_test_id else None
    attempts = []
    start_date = request.values.get("start_date", "").strip()
    end_date = request.values.get("end_date", "").strip()
    start_boundary = None
    end_boundary = None
    if start_date:
        try:
            local_start = datetime.strptime(start_date, "%Y-%m-%d").replace(tzinfo=PKT)
            start_boundary = local_start.astimezone(timezone.utc).replace(tzinfo=None)
        except ValueError:
            start_date = ""
    if end_date:
        try:
            local_end = datetime.strptime(end_date, "%Y-%m-%d") + timedelta(days=1)
            end_boundary = local_end.replace(tzinfo=PKT).astimezone(timezone.utc).replace(tzinfo=None)
        except ValueError:
            end_date = ""
    searched = bool(name or roll_no or start_date or end_date or selected_test_id)

    if searched:
        query = (
            Attempt.query
            .join(Student)
            .filter(Attempt.save_history.is_(True))
        )
        if selected_test:
            query = query.filter(Attempt.test_id == selected_test.id)
        if name:
            query = query.filter(Student.name.ilike(f"%{name}%"))
        if roll_no:
            query = query.filter(Student.roll_no == roll_no)
        if start_boundary:
            query = query.filter(Attempt.started_at >= start_boundary)
            effective_end = end_boundary or start_boundary + timedelta(days=1)
            query = query.filter(Attempt.started_at < effective_end)
        elif end_boundary:
            query = query.filter(Attempt.started_at < end_boundary)
        attempts = query.order_by(Attempt.started_at.desc()).all()

    attempt_rows = []
    for attempt in attempts:
        overall_total, overall_correct, overall_score = test_score(attempt)
        total = 0
        correct = 0
        section_details = []
        for section_attempt in attempt.section_attempts:
            if not section_attempt.section:
                continue
            sec = section_attempt.section
            section_questions = Question.query.filter_by(section_id=section_attempt.section_id).count()
            section_correct = (
                db.session.query(Response)
                .join(Question, Response.question_id == Question.id)
                .filter(
                    Response.attempt_id == attempt.id,
                    Question.section_id == section_attempt.section_id,
                    Response.is_correct.is_(True),
                )
                .count()
            )
            total += section_questions
            correct += section_correct
            sec_percentage = round(section_correct / section_questions * 100, 2) if section_questions else 0
            section_details.append({
                "title": sec.title,
                "correct": section_correct,
                "total": section_questions,
                "percentage": sec_percentage,
                "passed": section_attempt.passed,
            })
        attempt_rows.append({
            "attempt": attempt,
            "score": overall_score,
            "sections": len([sa for sa in attempt.section_attempts if sa.section]),
            "total_questions": overall_total,
            "total_correct": overall_correct,
            "pass_percentage": test_marks_summary(attempt.test)["pass_percentage"] if attempt.test else PASS_PERCENTAGE,
            "section_details": section_details,
        })

    students_by_roll = {}
    for student in (
        Student.query.join(Attempt)
        .filter(Attempt.save_history.is_(True))
        .order_by(Student.id.asc())
        .all()
    ):
        students_by_roll.setdefault(student.roll_no, student)
    students = sorted(
        students_by_roll.values(),
        key=lambda student: (0, int(student.roll_no)) if student.roll_no.strip().isdigit() else (1, student.roll_no.casefold()),
    )

    progress_analytics = build_student_progress_analytics(attempt_rows)
    history_tests = unique_attempt_tests(attempt_rows)
    history_test_details = [report_test_details(test) for test in history_tests]

    return render_template(
        "history.html",
        name=name,
        roll_no=roll_no,
        tests=tests,
        selected_test=selected_test,
        attempt_rows=attempt_rows,
        report_test_details=history_test_details,
        searched=searched,
        start_date=start_date,
        end_date=end_date,
        students=students,
        **progress_analytics,
    )


# =============================================================================
# Word downloads
# The .docx files copy the website's look (colours, tables, pills, charts) and
# contain only the report itself: no buttons, filters or filler text.
# =============================================================================
WORD_FONT = "Arial"
WORD_INK = "1F2937"
WORD_MUTED = "64748B"
WORD_DARK = "062C24"
WORD_BRAND = "0E7C56"
WORD_GREEN = "16A34A"
WORD_RED = "DC2626"
WORD_BLUE = "2563EB"
WORD_AMBER = "B45309"
WORD_SLATE = "475569"
WORD_BORDER = "E2E8F0"
WORD_PASS_BG = "F0FDF4"
WORD_FAIL_BG = "FEF2F2"
WORD_FAIL_ROW = "FEE2E2"
WORD_FAIL_TEXT = "991B1B"
WORD_SOFT = "F8FAFC"
SCORE_BAR_COLOURS = ["EF4444", "FACC15", "38BDF8", "84CC16", "16A34A"]

_CHART_FONT_CACHE = {}


def chart_font(size_px, bold=False):
    from PIL import ImageFont
    size_px = max(int(round(size_px)), 6)
    key = (size_px, bold)
    if key in _CHART_FONT_CACHE:
        return _CHART_FONT_CACHE[key]
    names = (
        ["DejaVuSans-Bold.ttf", "arialbd.ttf", "LiberationSans-Bold.ttf", "Arial Bold.ttf", "Helvetica-Bold.ttf"]
        if bold else
        ["DejaVuSans.ttf", "arial.ttf", "LiberationSans-Regular.ttf", "Arial.ttf", "Helvetica.ttf"]
    )
    font = None
    for name in names:
        try:
            font = ImageFont.truetype(name, size_px)
            break
        except OSError:
            continue
    if font is None:
        try:
            font = ImageFont.load_default(size=size_px)
        except TypeError:
            font = ImageFont.load_default()
    _CHART_FONT_CACHE[key] = font
    return font


class ChartCanvas:
    """Tiny drawing helper working in the same units as the website's SVG charts."""

    SUPERSAMPLE = 2

    def __init__(self, width, height, scale=3):
        from PIL import Image, ImageDraw
        self.scale = scale
        self.s = scale * self.SUPERSAMPLE
        self.width = width
        self.height = height
        self.image = Image.new("RGBA", (int(width * self.s), int(height * self.s)), (255, 255, 255, 0))
        self.draw = ImageDraw.Draw(self.image)

    def rect(self, x0, y0, x1, y1, fill=None, outline=None, width=0, radius=0):
        s = self.s
        box = [x0 * s, y0 * s, x1 * s, y1 * s]
        fill = f"#{fill}" if fill else None
        outline = f"#{outline}" if outline else None
        line_width = max(int(width * s), 1) if outline else 0
        if radius:
            self.draw.rounded_rectangle(box, radius=radius * s, fill=fill, outline=outline, width=line_width)
        else:
            self.draw.rectangle(box, fill=fill, outline=outline, width=line_width)

    def line(self, points, colour, width=1):
        s = self.s
        self.draw.line([(x * s, y * s) for x, y in points], fill=f"#{colour}", width=max(int(width * s), 1), joint="curve")

    def circle(self, cx, cy, r, fill=None, outline=None, width=0):
        s = self.s
        self.draw.ellipse(
            [(cx - r) * s, (cy - r) * s, (cx + r) * s, (cy + r) * s],
            fill=f"#{fill}" if fill else None,
            outline=f"#{outline}" if outline else None,
            width=max(int(width * s), 1) if outline else 0,
        )

    def arc(self, cx, cy, r, start, end, colour, thickness):
        s = self.s
        self.draw.arc(
            [(cx - r) * s, (cy - r) * s, (cx + r) * s, (cy + r) * s],
            start=start, end=end, fill=f"#{colour}", width=max(int(thickness * s), 1),
        )

    def text(self, x, y, value, size, colour, bold=False, anchor="ls"):
        self.draw.text((x * self.s, y * self.s), str(value), font=chart_font(size * self.s, bold), fill=f"#{colour}", anchor=anchor)

    def vertical_text(self, cx, cy, value, size, colour, bold=True):
        from PIL import Image, ImageDraw
        font = chart_font(size * self.s, bold)
        probe = ImageDraw.Draw(Image.new("RGBA", (1, 1)))
        left, top, right, bottom = probe.textbbox((0, 0), str(value), font=font)
        label = Image.new("RGBA", (right - left + 8, bottom - top + 8), (255, 255, 255, 0))
        ImageDraw.Draw(label).text((4 - left, 4 - top), str(value), font=font, fill=f"#{colour}")
        label = label.rotate(90, expand=True)
        self.image.alpha_composite(label, (int(cx * self.s - label.width / 2), int(cy * self.s - label.height / 2)))

    def text_width(self, value, size, bold=False):
        return chart_font(size * self.s, bold).getlength(str(value)) / self.s

    def fit_text(self, value, size, max_width, bold=False):
        value = str(value)
        if self.text_width(value, size, bold) <= max_width:
            return value
        while len(value) > 1 and self.text_width(value + "…", size, bold) > max_width:
            value = value[:-1]
        return value.rstrip() + "…"

    def gradient_card(self, x0, y0, x1, y1, start, end, border, bar, radius=8):
        """Rounded card with a soft left-to-right gradient, a thin border and a coloured left edge."""
        from PIL import Image, ImageDraw
        s = self.s
        width, height = int((x1 - x0) * s), int((y1 - y0) * s)
        mask = Image.linear_gradient("L").rotate(90).resize((width, height))
        body = Image.composite(
            Image.new("RGBA", (width, height), f"#{end}"),
            Image.new("RGBA", (width, height), f"#{start}"),
            mask,
        )
        shape = Image.new("L", (width, height), 0)
        ImageDraw.Draw(shape).rounded_rectangle([0, 0, width - 1, height - 1], radius=radius * s, fill=255)
        self.image.paste(body, (int(x0 * s), int(y0 * s)), shape)
        self.draw.rounded_rectangle(
            [x0 * s, y0 * s, x1 * s, y1 * s], radius=radius * s, outline=f"#{border}", width=max(int(s), 1),
        )
        self.draw.rounded_rectangle(
            [x0 * s, y0 * s, (x0 + 4) * s, y1 * s], radius=radius * s, fill=f"#{bar}",
            corners=(True, False, False, True),
        )

    def to_png(self):
        from PIL import Image
        final = self.image.resize((int(self.width * self.scale), int(self.height * self.scale)), Image.LANCZOS)
        buffer = BytesIO()
        final.save(buffer, format="PNG")
        buffer.seek(0)
        return buffer


def draw_percentage_axes(canvas, ox, oy, ticks, width, y_title, y_title_mid):
    for tick in ticks:
        colour = "64748B" if tick["major"] else "B6C2D0"
        canvas.line([(ox + 60, oy + tick["y"]), (ox + width - 18, oy + tick["y"])], colour, 1)
        canvas.text(ox + 53, oy + tick["y"] + 3.5, f"{tick['value']}%", 10, WORD_MUTED, anchor="rs")
    canvas.vertical_text(ox + 15, oy + y_title_mid, y_title, 12, "4B5563")


def draw_line_chart(canvas, ox, oy, analytics, label_key, x_title, y_title):
    width = analytics["chart_width"]
    height = analytics["chart_height"]
    for band in analytics.get("fail_bands", []):
        canvas.rect(
            ox + band["x"], oy + band["y"],
            ox + band["x"] + band["width"], oy + band["y"] + band["height"],
            fill=WORD_FAIL_BG,
        )
    draw_percentage_axes(canvas, ox, oy, analytics["chart_ticks"], width, y_title, analytics["chart_mid_y"])
    points = analytics["progress_points"]
    if len(points) > 1:
        canvas.line([(ox + p["x"], oy + p["y"]) for p in points], WORD_BRAND, 2.5)
    for point in points:
        failed = point.get("failed", False)
        canvas.circle(
            ox + point["x"], oy + point["y"], 4,
            fill=WORD_RED if failed else "FFFFFF",
            outline=WORD_RED if failed else WORD_BRAND,
            width=2.5,
        )
        canvas.text(ox + point["x"], oy + height - 30, point[label_key], 12, WORD_MUTED, anchor="ms")
    canvas.text(ox + width / 2, oy + height - 5, x_title, 12, "4B5563", bold=True, anchor="ms")


def draw_distribution_chart(canvas, ox, oy, analytics):
    width = analytics["dist_chart_width"]
    baseline = analytics["dist_baseline"]
    canvas.vertical_text(ox + 15, oy + analytics["dist_mid_y"], "Student percentage", 12, "4B5563")
    for tick in analytics["dist_chart_ticks"]:
        canvas.text(ox + 53, oy + tick["y"] + 3.5, f"{tick['value']}%", 10, WORD_MUTED, anchor="rs")
    band_names = {"Fail": "Failed", "Pass": "Average", "High score": "Good"}
    for band in (item for item in analytics["score_distribution"] if not item.get("dot")):
        fill = "FFF1F2" if band["label"] == "Fail" else "F7FEE7" if band["label"] == "Pass" else "ECFDF5"
        canvas.rect(ox + band["x"], oy + band["y"], ox + band["x"] + band["width"], oy + band["y"] + band["height"], fill=fill)
        canvas.text(ox + band["x"] + band["width"] - 12, oy + band["label_y"], f"{band_names[band['label']]} · {band['count']} students", 15, band["color"][1:], bold=True, anchor="rs")
    for point in (item for item in analytics["score_distribution"] if item.get("dot")):
        canvas.circle(ox + point["cx"], oy + point["cy"], 6, fill=point["color"][1:], outline="FFFFFF", width=2)
        canvas.text(ox + point["cx"], oy + baseline + 20, point["number"], 11, WORD_MUTED, bold=True, anchor="ms")
    canvas.text(ox + width / 2, oy + baseline + 40, "Student number", 12, "4B5563", bold=True, anchor="ms")
    canvas.text(
        ox + width - 18, oy + baseline + 72,
        f"Total students: {analytics['progress_total_students']}", 12, "4B5563", bold=True, anchor="rs",
    )


def draw_donut(canvas, ox, oy, value, title, caption):
    centre_x = ox + 75
    centre_y = oy + 80
    canvas.text(centre_x, oy + 14, title, 12, "111827", bold=True, anchor="ms")
    canvas.arc(centre_x, centre_y, 50 - 7, 0, 360, "DBE3DF", 14)
    value = max(0, min(100, float(value)))
    if value >= 99.95:
        canvas.arc(centre_x, centre_y, 50 - 7, 0, 360, WORD_BRAND, 14)
    elif value > 0:
        canvas.arc(centre_x, centre_y, 50 - 7, -90, -90 + 3.6 * value, WORD_BRAND, 14)
    canvas.text(centre_x, centre_y + 4.5, f"{value:g}%" if value != int(value) else f"{value:.1f}%", 13, "111827", bold=True, anchor="ms")
    canvas.text(centre_x, oy + 152, caption, 10, WORD_MUTED, anchor="ms")


SCORE_HIGHLIGHT_HEIGHT = 66


def draw_score_highlights(canvas, ox, oy, width, highest, lowest):
    """Green "highest" and red "lowest" cards side by side (same look as the website)."""
    gap = 10
    card_width = (width - gap) / 2
    cards = (
        ("Highest score", highest, "F0FDF4", "DCFCE7", "A7F3D0", "059669", "0F766E", "047857"),
        ("Lowest score", lowest, "FFF7F7", "FEE2E2", "FECACA", "DC2626", "B91C1C", "B91C1C"),
    )
    for index, (label, student, start, end, border, bar, label_colour, percent_colour) in enumerate(cards):
        if not student:
            continue
        left = ox + index * (card_width + gap)
        canvas.gradient_card(left, oy, left + card_width, oy + SCORE_HIGHLIGHT_HEIGHT, start, end, border, bar)
        text_left = left + 16
        text_width = card_width - 28
        canvas.text(text_left, oy + 18, label.upper(), 9.5, label_colour, bold=True)
        canvas.text(text_left, oy + 35, canvas.fit_text(student["name"], 12.5, text_width, bold=True), 12.5, WORD_INK, bold=True)
        marks_text = f"{student['marks']} / {student['total_marks']} marks  ·  "
        canvas.text(text_left, oy + 55, marks_text, 10, "4B5563")
        canvas.text(text_left + canvas.text_width(marks_text, 10), oy + 56, f"{student['percentage']}%", 15, percent_colour, bold=True)


def draw_pass_box(canvas, ox, oy, count_text, label):
    canvas.rect(ox, oy, ox + 150, oy + 46, fill="ECFDF5")
    canvas.rect(ox, oy, ox + 3, oy + 46, fill=WORD_GREEN)
    canvas.text(ox + 12, oy + 21, count_text, 17, "047857", bold=True)
    canvas.text(ox + 12, oy + 37, label, 10, "4B5563", bold=True)


def render_analytics_panel(chart_title, chart_width, chart_height, draw_chart, donuts, pass_text, pass_label, highlights=None):
    """One picture holding the chart, the two circular diagrams and the passed/total box.

    highlights = (highest, lowest) adds the website's green/red score cards under the chart."""
    pad, gap, donut_width, title_height = 14, 18, 150, 24
    cards_height = (SCORE_HIGHLIGHT_HEIGHT + 12) if highlights and any(highlights) else 0
    width = pad + chart_width + gap + donut_width + gap + donut_width + pad
    height = max(pad + title_height + chart_height + cards_height + pad, pad + 170 + 8 + 46 + pad)
    canvas = ChartCanvas(width, height)
    canvas.rect(0.5, 0.5, width - 0.5, height - 0.5, fill=WORD_SOFT, outline=WORD_BORDER, width=1, radius=10)
    canvas.text(pad, pad + 12, chart_title, 12, "111827", bold=True)
    draw_chart(canvas, pad, pad + title_height)
    if cards_height:
        cards_width = min(chart_width - 20, 520)
        draw_score_highlights(
            canvas, pad + (chart_width - cards_width) / 2, pad + title_height + chart_height + 12,
            cards_width, *highlights,
        )
    first_donut_x = pad + chart_width + gap
    second_donut_x = first_donut_x + donut_width + gap
    draw_donut(canvas, first_donut_x, pad, *donuts[0])
    draw_donut(canvas, second_donut_x, pad, *donuts[1])
    draw_pass_box(canvas, second_donut_x, pad + 178, pass_text, pass_label)
    return canvas.to_png()


# ---------- python-docx helpers (kept schema-ordered so Word opens the files cleanly) ----------

def _insert_before(parent, element, successors):
    for child in parent:
        if child.tag in {qn(name) for name in successors}:
            child.addprevious(element)
            return element
    parent.append(element)
    return element


def word_cell_fill(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shade = OxmlElement("w:shd")
    shade.set(qn("w:val"), "clear")
    shade.set(qn("w:color"), "auto")
    shade.set(qn("w:fill"), fill)
    _insert_before(tc_pr, shade, ["w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark"])


def word_cell_borders(cell, **sides):
    """sides: top/left/bottom/right = (size_in_eighths_of_a_point, hex_colour)."""
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        if side in sides:
            size, colour = sides[side]
            edge = OxmlElement(f"w:{side}")
            edge.set(qn("w:val"), "single")
            edge.set(qn("w:sz"), str(size))
            edge.set(qn("w:space"), "0")
            edge.set(qn("w:color"), colour)
            borders.append(edge)
    _insert_before(tc_pr, borders, ["w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark"])


def word_run(paragraph, text, size=9, bold=False, colour=WORD_INK, fill=None, italic=False):
    run = paragraph.add_run(text)
    run.font.name = WORD_FONT
    run._r.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), WORD_FONT)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor.from_string(colour)
    if fill:
        shade = OxmlElement("w:shd")
        shade.set(qn("w:val"), "clear")
        shade.set(qn("w:color"), "auto")
        shade.set(qn("w:fill"), fill)
        _insert_before(run._r.get_or_add_rPr(), shade, ["w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath"])
    return run


def word_pill(paragraph, text, fill, size=9):
    return word_run(paragraph, f"\u00a0\u00a0{text}\u00a0\u00a0", size=size, bold=True, colour="FFFFFF", fill=fill)


def word_paragraph(container, align=None, before=0, after=0):
    paragraph = container.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    if align == "center":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    return paragraph


def word_cell_text(cell, text="", size=9, bold=False, colour=WORD_INK, align=None, fill=None):
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    if align == "center":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if text != "":
        word_run(paragraph, str(text), size=size, bold=bold, colour=colour)
    if fill:
        word_cell_fill(cell, fill)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    return paragraph


def word_box(container, text, fill, size=9, width=None, align="center", gap=1):
    """A padded coloured box (the website's badges) placed inside a cell or the page body."""
    if width is None:
        width = max(0.7, len(str(text)) * 0.078 * size / 9 + 0.4)
    if hasattr(container, "_tc") and not container.tables and len(container.paragraphs) == 1 and not container.paragraphs[0].text:
        container._tc.remove(container.paragraphs[0]._p)
    box = container.add_table(rows=1, cols=1)
    box.alignment = WD_TABLE_ALIGNMENT.CENTER if align == "center" else WD_TABLE_ALIGNMENT.LEFT
    box.autofit = False
    box.columns[0].width = Inches(width)
    tbl_pr = box._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    _insert_before(tbl_pr, layout, ["w:tblCellMar", "w:tblLook"])
    margins = OxmlElement("w:tblCellMar")
    for side, value in (("top", 45), ("left", 60), ("bottom", 45), ("right", 60)):
        edge = OxmlElement(f"w:{side}")
        edge.set(qn("w:w"), str(value))
        edge.set(qn("w:type"), "dxa")
        margins.append(edge)
    _insert_before(tbl_pr, margins, ["w:tblLook"])
    cell = box.rows[0].cells[0]
    cell.width = Inches(width)
    word_cell_text(cell, text, size=size, bold=True, colour="FFFFFF", align="center", fill=fill)
    # python-docx leaves one empty paragraph after a nested table; keep it invisible.
    trailing = container.paragraphs[-1]
    trailing.paragraph_format.space_before = Pt(0)
    trailing.paragraph_format.space_after = Pt(0)
    trailing.paragraph_format.line_spacing = Pt(gap)
    word_run(trailing, "", size=1)
    return box


def word_cell_pill(cell, text, pill_fill, background=None, size=9):
    if background:
        word_cell_fill(cell, background)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    return word_box(cell, text, pill_fill, size=size)


def word_new_document(landscape=False):
    document = Document()
    section = document.sections[0]
    if landscape:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = Cm(29.7), Cm(21.0)
    else:
        section.page_width, section.page_height = Cm(21.0), Cm(29.7)
    section.left_margin = section.right_margin = Inches(0.55)
    section.top_margin = section.bottom_margin = Inches(0.55)
    normal = document.styles["Normal"]
    normal.font.name = WORD_FONT
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), WORD_FONT)
    normal.font.size = Pt(10)
    normal.font.color.rgb = RGBColor.from_string(WORD_INK)
    return document


def word_usable_width(document):
    section = document.sections[0]
    return (section.page_width - section.left_margin - section.right_margin) / 914400


def word_style_table(table, widths, border_colour=WORD_BORDER):
    """Thin light borders, fixed layout, comfortable padding and real column widths."""
    table.autofit = False
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        edge = OxmlElement(f"w:{side}")
        edge.set(qn("w:val"), "single")
        edge.set(qn("w:sz"), "4")
        edge.set(qn("w:space"), "0")
        edge.set(qn("w:color"), border_colour)
        borders.append(edge)
    _insert_before(tbl_pr, borders, ["w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook"])
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    _insert_before(tbl_pr, layout, ["w:tblCellMar", "w:tblLook"])
    margins = OxmlElement("w:tblCellMar")
    for side, value in (("top", 70), ("left", 90), ("bottom", 70), ("right", 90)):
        edge = OxmlElement(f"w:{side}")
        edge.set(qn("w:w"), str(value))
        edge.set(qn("w:type"), "dxa")
        margins.append(edge)
    _insert_before(tbl_pr, margins, ["w:tblLook"])
    for index, width in enumerate(widths):
        table.columns[index].width = Inches(width)


def word_table(document, headers, widths, header_fill=WORD_DARK, border_colour=WORD_BORDER):
    """Table with the website's dark green header row."""
    table = document.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    word_style_table(table, widths, border_colour)
    header_row = table.rows[0]
    header_row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
    for index, (cell, label) in enumerate(zip(header_row.cells, headers)):
        cell.width = Inches(widths[index])
        word_cell_text(cell, label.upper(), size=7.5, bold=True, colour="FFFFFF", fill=header_fill)
    return table


def word_add_row(table, widths, keep_together=True):
    row = table.add_row()
    if keep_together:
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    for index, cell in enumerate(row.cells):
        cell.width = Inches(widths[index])
    return row.cells


def scale_widths(widths, total):
    factor = total / sum(widths)
    return [round(value * factor, 3) for value in widths]


def word_picture(document, buffer, width_inches):
    paragraph = word_paragraph(document, align="center", before=4, after=8)
    paragraph.add_run().add_picture(buffer, width=Inches(width_inches))


def word_heading(document, text, size=15, before=6, after=4, colour="111827"):
    paragraph = word_paragraph(document, before=before, after=after)
    paragraph.paragraph_format.keep_with_next = True
    word_run(paragraph, text, size=size, bold=True, colour=colour)
    return paragraph


def word_test_details(document, tests):
    """Add the test and section mark requirements used by non-result report exports."""
    seen = set()
    for test in tests:
        if not test or test.id in seen:
            continue
        seen.add(test.id)
        summary = test_marks_summary(test)
        word_heading(document, test.title, size=13, before=8, after=4)
        usable = word_usable_width(document)
        widths = scale_widths([2.2, 0.8, 1.0, 1.1, 1.0], usable)
        table = word_table(document, ["Section", "Total marks", "Passing marks", "Passing %", "Status"], widths)
        for section in ordered_sections(test.id):
            cells = word_add_row(table, widths)
            total_marks = section_total_marks(section)
            passing_marks = get_section_passing_marks(section)
            values = [section.title, total_marks, passing_marks, f"{section.pass_percentage:g}%", "Section pass"]
            for index, value in enumerate(values):
                word_cell_text(cells[index], value, size=9)
        cells = word_add_row(table, widths)
        values = ["TEST TOTAL", summary["total_marks"], summary["passing_marks"], f"{summary['pass_percentage']:.2f}%", "Overall pass"]
        for index, value in enumerate(values):
            word_cell_text(cells[index], value, size=9, bold=True, colour=WORD_GREEN, fill=WORD_PASS_BG)


def word_response(buffer, filename):
    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


def word_status_pill(cell, attempt, score, pass_percentage, background=True):
    if attempt.status == "in_progress":
        word_cell_pill(cell, "IN PROGRESS", WORD_SLATE, WORD_SOFT if background else None, size=8)
    elif attempt_overall_passed(attempt, score, pass_percentage):
        word_cell_pill(cell, "PASS", WORD_GREEN, WORD_PASS_BG if background else None, size=8)
    else:
        word_cell_pill(cell, "FAIL", "B91C1C", WORD_FAIL_BG if background else None, size=8)


def history_attempt_rows(attempts):
    """Same figures the History page shows for each attempt."""
    attempt_rows = []
    for attempt in attempts:
        overall_total, overall_correct, overall_score = test_score(attempt)
        section_details = []
        for section_attempt in attempt.section_attempts:
            if not section_attempt.section:
                continue
            section_questions = Question.query.filter_by(section_id=section_attempt.section_id).count()
            section_correct = (
                db.session.query(Response)
                .join(Question, Response.question_id == Question.id)
                .filter(
                    Response.attempt_id == attempt.id,
                    Question.section_id == section_attempt.section_id,
                    Response.is_correct.is_(True),
                )
                .count()
            )
            section_details.append({
                "title": section_attempt.section.title,
                "correct": section_correct,
                "total": section_questions,
                "percentage": round(section_correct / section_questions * 100, 2) if section_questions else 0,
                "passed": section_attempt.passed,
            })
        attempt_rows.append({
            "attempt": attempt,
            "score": overall_score,
            "sections": len(section_details),
            "total_questions": overall_total,
            "total_correct": overall_correct,
            "pass_percentage": test_marks_summary(attempt.test)["pass_percentage"] if attempt.test else PASS_PERCENTAGE,
            "section_details": section_details,
        })
    return attempt_rows


def attempt_section_reports(attempt):
    """Same section / answer-review data the Result Report page shows."""
    section_reports = []
    responses = {response.question_id: response for response in Response.query.filter_by(attempt_id=attempt.id).all()}
    for section_attempt in sorted(attempt.section_attempts, key=lambda item: (item.section.order_no, item.section.id)):
        questions = Question.query.filter_by(section_id=section_attempt.section_id).order_by(Question.id.asc()).all()
        answered = sum(1 for question in questions if question.id in responses)
        correct = sum(1 for question in questions if responses.get(question.id) and responses[question.id].is_correct)
        review = []
        for question in questions:
            response = responses.get(question.id)
            selected = response.selected_option if response else None
            review.append({
                "text": question.text,
                "selected_option": selected,
                "selected_text": getattr(question, f"option_{selected.lower()}") if selected else "Not answered",
                "correct_option": question.correct_option,
                "correct_text": getattr(question, f"option_{question.correct_option.lower()}"),
                "is_correct": bool(response and response.is_correct),
            })
        section_reports.append({
            "title": section_attempt.section.title,
            "pass_percentage": section_attempt.section.pass_percentage if section_attempt.section else PASS_PERCENTAGE,
            "total": len(questions),
            "passing_marks": get_section_passing_marks(section_attempt.section) if section_attempt.section else 0,
            "answered": answered,
            "correct": correct,
            "percentage": section_attempt.percentage if section_attempt.percentage is not None else 0,
            "passed": bool(section_attempt.passed),
            "timed_out": section_attempt.timed_out,
            "review": review,
        })
    return section_reports


@app.get("/history/download-word")
def download_history_word():
    name = request.args.get("name", "").strip()
    roll_no = request.args.get("roll_no", "").strip()
    selected_test_id = request.args.get("test_id", type=int)
    selected_test = db.session.get(Test, selected_test_id) if selected_test_id else None
    start_date, end_date, start_boundary, end_boundary = parse_report_dates(
        request.args.get("start_date"), request.args.get("end_date")
    )
    if not (name or roll_no or start_date or end_date or selected_test_id):
        flash("Search for test history before downloading the report.", "error")
        return redirect(url_for("history"))

    attempts_query = Attempt.query.join(Student).filter(Attempt.save_history.is_(True))
    if selected_test:
        attempts_query = attempts_query.filter(Attempt.test_id == selected_test.id)
    if name:
        attempts_query = attempts_query.filter(Student.name.ilike(f"%{name}%"))
    if roll_no:
        attempts_query = attempts_query.filter(Student.roll_no == roll_no)
    if start_boundary:
        attempts_query = attempts_query.filter(Attempt.started_at >= start_boundary)
        attempts_query = attempts_query.filter(Attempt.started_at < (end_boundary or start_boundary + timedelta(days=1)))
    elif end_boundary:
        attempts_query = attempts_query.filter(Attempt.started_at < end_boundary)
    attempts = attempts_query.order_by(Attempt.started_at.desc()).all()
    if not attempts:
        flash("No test history matched this search.", "error")
        return redirect(url_for(
            "history", name=name, roll_no=roll_no,
            test_id=selected_test.id if selected_test else "",
            start_date=start_date, end_date=end_date,
        ))

    attempt_rows = history_attempt_rows(attempts)
    analytics = build_student_progress_analytics(attempt_rows)
    word_passed_attempts = sum(
        attempt_overall_passed(row["attempt"], row["score"], row["pass_percentage"])
        for row in attempt_rows
    )
    analytics["passed_attempts"] = word_passed_attempts
    analytics["passed_percentage"] = round(word_passed_attempts / len(attempt_rows) * 100, 1) if attempt_rows else 0
    document = word_new_document(landscape=True)
    usable = word_usable_width(document)
    word_test_details(document, [row["attempt"].test for row in attempt_rows])

    if analytics["progress_student"]:
        panel = render_analytics_panel(
            f"{analytics['progress_student'].name} · Test progress",
            analytics["chart_width"], analytics["chart_height"],
            lambda canvas, ox, oy: draw_line_chart(canvas, ox, oy, analytics, "date_label", "Test date", "Percentage"),
            [
                (analytics["average_percentage"], "Average score",
                 f"Across {analytics['progress_total_tests']} test{'' if analytics['progress_total_tests'] == 1 else 's'}"),
                (analytics["passed_percentage"], "Passing rate", "Tests passed"),
            ],
            f"{analytics['passed_attempts']} / {analytics['progress_total_tests']}",
            "Tests passed / total",
        )
        word_picture(document, panel, usable)

    word_heading(document, f"{len(attempt_rows)} previous attempt{'' if len(attempt_rows) == 1 else 's'}", size=15, before=2, after=6)

    summary_widths = scale_widths([1.9, 0.8, 2.0, 1.9, 1.0, 2.0, 0.9], usable)
    section_widths = scale_widths([4.0, 2.2, 2.0, 1.6], usable)
    for position, row in enumerate(attempt_rows):
        attempt = row["attempt"]
        card = document.add_table(rows=2, cols=7)
        card.alignment = WD_TABLE_ALIGNMENT.CENTER
        word_style_table(card, summary_widths)
        labels = [
            "Student", "Roll #", f"Attempt #{len(attempt_rows) - position} · Test",
            "Overall score", "Result", "Date", "Sections",
        ]
        for index, label in enumerate(labels):
            top_cell, bottom_cell = card.rows[0].cells[index], card.rows[1].cells[index]
            top_cell.width = bottom_cell.width = Inches(summary_widths[index])
            word_cell_text(top_cell, label.upper(), size=7, colour=WORD_MUTED, fill=WORD_SOFT)
            word_cell_fill(bottom_cell, WORD_SOFT)
        row_cells = card.rows[1].cells
        word_cell_text(row_cells[0], attempt.student.name, size=9.5, bold=True)
        word_cell_text(row_cells[1], attempt.student.roll_no, size=9.5, bold=True)
        word_cell_text(row_cells[2], attempt.test.title if attempt.test else "Test", size=9.5, bold=True)
        word_cell_pill(row_cells[3], f"{row['score']}% | {row['total_correct']} / {row['total_questions']}", WORD_BLUE, WORD_SOFT, size=9)
        if attempt_overall_passed(attempt, row["score"], row["pass_percentage"]):
            word_cell_pill(row_cells[4], "PASS", WORD_GREEN, WORD_SOFT, size=8.5)
        else:
            word_cell_pill(row_cells[4], "FAIL", "B91C1C", WORD_SOFT, size=8.5)
        word_cell_text(row_cells[5], to_pkdt(attempt.started_at).strftime("%d %b %Y, %I:%M %p"), size=9, bold=True)
        word_cell_text(row_cells[6], row["sections"], size=9.5, bold=True)
        for cell in card.rows[0].cells + card.rows[1].cells:
            word_cell_borders(cell, top=(4, WORD_BORDER), bottom=(4, WORD_BORDER), right=(4, WORD_BORDER), left=(4, WORD_BORDER))
        for cell in (card.rows[0].cells[0], card.rows[1].cells[0]):
            word_cell_borders(cell, top=(4, WORD_BORDER), bottom=(4, WORD_BORDER), right=(4, WORD_BORDER), left=(36, "047857"))
        for card_row in card.rows:
            card_row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        for cell in card.rows[0].cells:
            cell.paragraphs[0].paragraph_format.keep_with_next = True

        if len(row["section_details"]) > 1:
            heading = word_paragraph(document, before=6, after=3)
            heading.paragraph_format.keep_with_next = True
            word_run(heading, "Section-wise Results", size=10.5, bold=True, colour="111827")
            table = word_table(document, ["Section", "Correct / Total", "Percentage", "Passed"], section_widths)
            for section in row["section_details"]:
                cells = word_add_row(table, section_widths)
                failed = not section["passed"]
                fill = WORD_FAIL_ROW if failed else WORD_PASS_BG
                text_colour = WORD_FAIL_TEXT if failed else WORD_INK
                word_cell_text(cells[0], section["title"], size=9.5, bold=failed, colour=text_colour, fill=fill)
                word_cell_text(cells[1], f"{section['correct']} / {section['total']}", size=9.5, bold=failed, colour=text_colour, fill=fill)
                word_cell_text(cells[2], f"{section['percentage']}%", size=9.5, bold=failed, colour=text_colour, fill=fill)
                if section["passed"]:
                    word_cell_pill(cells[3], "PASS", WORD_GREEN, fill, size=8.5)
                else:
                    word_cell_pill(cells[3], "FAIL", "B91C1C", fill, size=8.5)
                if failed:
                    for fail_cell in cells:
                        word_cell_borders(fail_cell, top=(12, WORD_RED), bottom=(12, WORD_RED))
                    word_cell_borders(cells[0], left=(48, WORD_RED), top=(12, WORD_RED), bottom=(12, WORD_RED))
        word_paragraph(document, after=10)

    buffer = BytesIO()
    document.save(buffer)
    return word_response(buffer, "test_history_report.docx")


def parse_report_dates(start_date, end_date):
    """Turn 'YYYY-MM-DD' strings (Pakistan time) into UTC boundaries; bad values are ignored."""
    start_date = (start_date or "").strip()
    end_date = (end_date or "").strip()
    start_boundary = None
    end_boundary = None
    if start_date:
        try:
            start_boundary = datetime.strptime(start_date, "%Y-%m-%d").replace(tzinfo=PKT).astimezone(timezone.utc).replace(tzinfo=None)
        except ValueError:
            start_date = ""
    if end_date:
        try:
            end_boundary = (datetime.strptime(end_date, "%Y-%m-%d") + timedelta(days=1)).replace(tzinfo=PKT).astimezone(timezone.utc).replace(tzinfo=None)
        except ValueError:
            end_date = ""
    return start_date, end_date, start_boundary, end_boundary


@app.get("/admin/overall-results")
@admin_required
def overall_results():
    tests = Test.query.order_by(Test.created_at.asc(), Test.id.asc()).all()
    test_selection = request.args.get("test_id", "")
    all_tests = test_selection == "all"
    selected_test_id = request.args.get("test_id", type=int)
    order_by = request.args.get("order_by", "")
    valid_orderings = {
        "date_asc": ("date", "asc"),
        "date_desc": ("date", "desc"),
        "roll_asc": ("roll", "asc"),
        "roll_desc": ("roll", "desc"),
        "marks_asc": ("marks", "asc"),
        "marks_desc": ("marks", "desc"),
    }
    sort_by, sort_order = valid_orderings.get(
        order_by,
        (request.args.get("sort_by", "marks"), request.args.get("sort_order", "desc")),
    )
    if sort_by not in {"date", "roll", "marks"}:
        sort_by = "marks"
    if sort_order not in {"asc", "desc"}:
        sort_order = "desc"
    order_by = f"{sort_by}_{sort_order}"
    selected_test = None if all_tests or not selected_test_id else db.session.get(Test, selected_test_id)
    report_selected = all_tests or selected_test is not None
    attempts_query = Attempt.query.filter(Attempt.save_history.is_(True))
    if selected_test:
        attempts_query = attempts_query.filter_by(test_id=selected_test.id)
    elif not all_tests or not report_selected:
        attempts_query = attempts_query.filter_by(test_id=None)
    start_date, end_date, start_boundary, end_boundary = parse_report_dates(
        request.args.get("start_date"), request.args.get("end_date")
    )
    if start_boundary:
        attempts_query = attempts_query.filter(Attempt.started_at >= start_boundary)
    if end_boundary:
        attempts_query = attempts_query.filter(Attempt.started_at < end_boundary)
    attempts = attempts_query.order_by(Attempt.started_at.desc()).all() if report_selected else []
    student_rows = []
    for attempt in attempts:
        total, correct, score = test_score(attempt)
        student_rows.append({
            "attempt": attempt,
            "total": total,
            "correct": correct,
            "score": score,
            "pass_percentage": test_marks_summary(attempt.test)["pass_percentage"] if attempt.test else PASS_PERCENTAGE,
        })

    if sort_by == "date":
        sort_key = lambda row: row["attempt"].started_at
    elif sort_by == "roll":
        def sort_key(row):
            roll_no = (row["attempt"].student.roll_no or "").strip()
            return (0, int(roll_no)) if roll_no.isdigit() else (1, roll_no.casefold())
    else:
        sort_key = lambda row: (row["score"], row["correct"])
    student_rows.sort(key=sort_key, reverse=sort_order == "desc")
    students = (
        Student.query.join(Attempt)
        .filter(Attempt.save_history.is_(True))
        .distinct()
        .order_by(Student.name.asc(), Student.id.asc())
        .all()
    )

    selected_student = None
    overall_progress = build_overall_progress_analytics(student_rows, selected_test, all_tests) if report_selected else None
    overall_tests = unique_attempt_tests(student_rows)
    if selected_test and selected_test not in overall_tests:
        overall_tests = [selected_test]
    overall_test_details = [report_test_details(test) for test in overall_tests]
    progress_analytics = build_student_progress_analytics(student_rows) if selected_student else build_student_progress_analytics([])
    return render_template(
        "overall_results.html",
        tests=tests,
        students=students,
        selected_test=selected_test,
        selected_student=selected_student,
        all_tests=all_tests,
        test_title="All Tests" if all_tests else (selected_test.title if selected_test else "Overall Results"),
        student_rows=student_rows,
        pass_percentage=test_marks_summary(selected_test)["pass_percentage"] if selected_test else PASS_PERCENTAGE,
        report_selected=report_selected,
        start_date=start_date,
        end_date=end_date,
        sort_by=sort_by,
        sort_order=sort_order,
        order_by=order_by,
        overall_progress=overall_progress,
        report_test_details=overall_test_details,
        **progress_analytics,
    )


@app.get("/admin/overall-results/download-word")
@admin_required
def download_overall_results_word():
    test_selection = request.args.get("test_id", "")
    all_tests = test_selection == "all"
    selected_test_id = request.args.get("test_id", type=int)
    if not all_tests and not selected_test_id:
        return redirect(url_for("overall_results"))
    selected_test = None if all_tests else db.session.get(Test, selected_test_id)
    attempts_query = Attempt.query.filter(Attempt.save_history.is_(True))
    if selected_test:
        attempts_query = attempts_query.filter_by(test_id=selected_test.id)
    elif not all_tests:
        attempts_query = attempts_query.filter_by(test_id=None)
    _start_date, _end_date, start_boundary, end_boundary = parse_report_dates(
        request.args.get("start_date"), request.args.get("end_date")
    )
    if start_boundary:
        attempts_query = attempts_query.filter(Attempt.started_at >= start_boundary)
    if end_boundary:
        attempts_query = attempts_query.filter(Attempt.started_at < end_boundary)
    attempts = attempts_query.order_by(Attempt.started_at.desc()).all()
    student_rows = []
    for attempt in attempts:
        # Use the full test denominator, including sections not reached after a failure.
        total, correct, score = test_score(attempt)
        student_rows.append({
            "attempt": attempt,
            "total": total,
            "correct": correct,
            "score": score,
            "pass_percentage": test_marks_summary(attempt.test)["pass_percentage"] if attempt.test else PASS_PERCENTAGE,
        })
    student_rows.sort(key=lambda row: (row["score"], row["correct"], row["attempt"].started_at), reverse=True)

    document = word_new_document(landscape=True)
    usable = word_usable_width(document)
    eyebrow = word_paragraph(document, after=2)
    word_run(eyebrow, "Overall Result Report", size=9, colour=WORD_INK)
    test_title = "All Tests" if all_tests else (selected_test.title if selected_test else "Overall Results")
    word_heading(document, test_title, size=22, before=2, after=8)
    word_test_details(
        document,
        [row["attempt"].test for row in student_rows] if all_tests else [selected_test],
    )

    if student_rows:
        analytics = build_overall_progress_analytics(student_rows, selected_test, all_tests)
        word_passed_count = sum(
            attempt_overall_passed(row["attempt"], row["score"], row["pass_percentage"])
            for row in student_rows
        )
        analytics["passed_count"] = word_passed_count
        analytics["passed_percentage"] = round(word_passed_count / len(student_rows) * 100, 1)
        if all_tests:
            word_rows_by_test = {}
            for row in student_rows:
                word_rows_by_test.setdefault(row["attempt"].test_id, []).append(row)
            test_passing_rates = [
                sum(attempt_overall_passed(row["attempt"], row["score"], row["pass_percentage"]) for row in rows)
                / len(rows) * 100
                for rows in word_rows_by_test.values() if rows
            ]
            analytics["tests_average_passing_rate"] = round(sum(test_passing_rates) / len(test_passing_rates), 1) if test_passing_rates else 0
        if all_tests:
            panel = render_analytics_panel(
                "Overall student score distribution",
                analytics["chart_width"], analytics["chart_height"],
                lambda canvas, ox, oy: draw_line_chart(canvas, ox, oy, analytics, "label", "Test number", "Percentage"),
                [
                    (analytics["average_percentage"], "Overall average percentage", "All selected tests"),
                    (analytics["tests_average_passing_rate"] or 0, "Passing rate", "Average across all tests"),
                ],
                f"{analytics['passed_count']} / {analytics['progress_total_students']}",
                "Students passed / total",
                highlights=(analytics.get("top_score_student"), analytics.get("low_score_student")),
            )
        else:
            panel = render_analytics_panel(
                "Student score distribution",
                analytics["dist_chart_width"], analytics["dist_chart_height"],
                lambda canvas, ox, oy: draw_distribution_chart(canvas, ox, oy, analytics),
                [
                    (analytics["average_percentage"], "Overall average percentage", "Selected test"),
                    (analytics["passed_percentage"], "Passing rate", "Students passed"),
                ],
                f"{analytics['passed_count']} / {analytics['progress_total_students']}",
                "Students passed / total",
                highlights=(analytics.get("top_score_student"), analytics.get("low_score_student")),
            )
        word_picture(document, panel, usable)

        headers = ["Rank", "Student", "Roll number", "Obtained / Total marks", "Percentage", "Date and time", "Status"]
        widths = [0.6, 2.3, 1.2, 1.8, 1.4, 2.3, 1.1]
        if all_tests:
            headers.insert(1, "Test")
            widths.insert(1, 1.6)
        widths = scale_widths(widths, usable)
        table = word_table(document, headers, widths)
        for rank, row_data in enumerate(student_rows, start=1):
            attempt = row_data["attempt"]
            cells = word_add_row(table, widths)
            index = 0
            word_cell_text(cells[index], rank, size=9.5, bold=True)
            index += 1
            if all_tests:
                word_cell_text(cells[index], attempt.test.title if attempt.test else "Test", size=9.5)
                index += 1
            word_cell_text(cells[index], attempt.student.name, size=9.5)
            word_cell_text(cells[index + 1], attempt.student.roll_no, size=9.5)
            word_cell_pill(cells[index + 2], f"{row_data['correct']} / {row_data['total']}", WORD_BLUE, size=9.5)
            word_cell_pill(cells[index + 3], f"{row_data['score']}%", WORD_BLUE, size=9.5)
            word_cell_text(cells[index + 4], to_pkdt(attempt.started_at).strftime("%d %b %Y, %I:%M:%S %p"), size=9)
            word_status_pill(cells[index + 5], attempt, row_data["score"], row_data["pass_percentage"])

    buffer = BytesIO()
    document.save(buffer)
    return word_response(buffer, "overall_result_report.docx")


@app.get("/admin/overall-results/download-excel")
@admin_required
def download_overall_results_excel():
    selected_value = request.args.get("test_id", "")
    selected_test_id = request.args.get("test_id", type=int)
    all_tests = selected_value == "all" or not selected_value
    attempts = (
        Attempt.query.filter(Attempt.save_history.is_(True))
        .order_by(Attempt.started_at.desc())
        .all()
    )
    if selected_test_id:
        attempts = [attempt for attempt in attempts if attempt.test_id == selected_test_id]
        test_ids = [selected_test_id]
    elif all_tests:
        test_ids = list(dict.fromkeys(attempt.test_id for attempt in attempts if attempt.test_id))
    else:
        test_ids = []
    tests = [db.session.get(Test, test_id) for test_id in test_ids]
    tests = [test for test in tests if test]
    conducted_ids = {attempt.test_id for attempt in attempts}
    if all_tests:
        tests = [test for test in tests if test.id in conducted_ids]

    student_results = {}
    for attempt in attempts:
        student_key = (attempt.student.name, attempt.student.roll_no)
        results_for_student = student_results.setdefault(student_key, {})
        if attempt.test_id in results_for_student:
            continue
        total, correct, score = test_score(attempt)
        results_for_student[attempt.test_id] = {
            "correct": correct, "total": total, "score": score,
            "passed": score >= test_marks_summary(attempt.test)["pass_percentage"],
        }

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Test Marks"
    header_fill = PatternFill("solid", fgColor="123B35")
    meta_fill = PatternFill("solid", fgColor="E8F3EF")
    subheader_fill = PatternFill("solid", fgColor="DCE9E5")
    pass_fill = PatternFill("solid", fgColor="DCFCE7")
    fail_fill = PatternFill("solid", fgColor="FEE2E2")
    white_bold = Font(bold=True, color="FFFFFF")
    sheet.merge_cells("A1:B2")
    sheet["A1"] = "Student"
    for cell in (sheet["A1"], sheet["B1"], sheet["A2"], sheet["B2"]):
        cell.fill = header_fill
        cell.font = white_bold
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for column in (1, 2):
        sheet.cell(row=3, column=column, value="Name" if column == 1 else "Roll number")
        sheet.cell(row=3, column=column).fill = subheader_fill
        sheet.cell(row=3, column=column).font = Font(bold=True, color="123B35")
        sheet.cell(row=3, column=column).alignment = Alignment(horizontal="center")

    test_columns = {}
    for test_index, test in enumerate(tests):
        first_col = 3 + test_index * 3
        last_col = first_col + 2
        test_columns[test.id] = first_col
        sheet.merge_cells(start_row=1, start_column=first_col, end_row=1, end_column=last_col)
        title_cell = sheet.cell(row=1, column=first_col, value=test.title)
        title_cell.font = white_bold
        title_cell.fill = header_fill
        title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        summary = test_marks_summary(test)
        detail = f"{len(test.sections)} sections  ·  Total {summary['total_marks']}  ·  Passing {summary['passing_marks']} ({summary['pass_percentage']:.2f}%)"
        sheet.merge_cells(start_row=2, start_column=first_col, end_row=2, end_column=last_col)
        detail_cell = sheet.cell(row=2, column=first_col, value=detail)
        detail_cell.fill = meta_fill
        detail_cell.font = Font(bold=True, color="123B35", size=9)
        detail_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for offset, label in enumerate(("Marks", "Percentage", "Status")):
            cell = sheet.cell(row=3, column=first_col + offset, value=label)
            cell.fill = subheader_fill
            cell.font = Font(bold=True, color="123B35")
            cell.alignment = Alignment(horizontal="center")

    ordered_students = sorted(student_results.items(), key=lambda item: (
        (0, int(item[0][1])) if item[0][1].strip().isdigit() else (1, item[0][1].casefold()),
        item[0][0].casefold(),
    ))
    for row_index, (student, results) in enumerate(ordered_students, start=4):
        sheet.cell(row=row_index, column=1, value=student[0])
        sheet.cell(row=row_index, column=2, value=student[1])
        for test in tests:
            first_col = test_columns[test.id]
            result = results.get(test.id)
            if not result:
                continue
            values = [f"{result['correct']} / {result['total']}", result["score"], "PASS" if result["passed"] else "FAIL"]
            for offset, value in enumerate(values):
                cell = sheet.cell(row=row_index, column=first_col + offset, value=value)
                cell.fill = pass_fill if result["passed"] else fail_fill
                cell.font = Font(color="166534" if result["passed"] else "B91C1C", bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            sheet.cell(row=row_index, column=first_col + 1).number_format = '0.00"%"'

    sheet.freeze_panes = "C4"
    sheet.auto_filter.ref = f"A3:{get_column_letter(max(2, 2 + len(tests) * 3))}{max(3, sheet.max_row)}"
    sheet.column_dimensions["A"].width = 28
    sheet.column_dimensions["B"].width = 16
    for column_index in range(3, 3 + len(tests) * 3):
        sheet.column_dimensions[get_column_letter(column_index)].width = 15
    sheet.row_dimensions[1].height = 32
    sheet.row_dimensions[2].height = 30
    sheet.row_dimensions[3].height = 24

    buffer = BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=(f"{tests[0].title[:48]}_marks.xlsx" if selected_test_id and tests else "all_conducted_tests_marks.xlsx"),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/history/<int:attempt_id>/delete")
@admin_required
def delete_history_attempt(attempt_id):
    attempt = db.session.get(Attempt, attempt_id)
    if attempt:
        student = attempt.student
        Response.query.filter_by(attempt_id=attempt.id).delete(synchronize_session=False)
        SectionAttempt.query.filter_by(attempt_id=attempt.id).delete(synchronize_session=False)
        db.session.delete(attempt)
        db.session.flush()
        if student and not Attempt.query.filter_by(student_id=student.id).first():
            db.session.delete(student)
        db.session.commit()
        flash("Previous test record deleted.", "success")
    if request.form.get("from") == "overall":
        test_selection = request.form.get("test_id", "")
        return redirect(url_for("overall_results", test_id="all" if test_selection == "all" else request.form.get("test_id", type=int)))
    return redirect(url_for("history"))


@app.post("/history/delete-all")
@admin_required
def delete_all_history():
    Response.query.delete(synchronize_session=False)
    SectionAttempt.query.delete(synchronize_session=False)
    Attempt.query.delete(synchronize_session=False)
    # Students are NOT deleted: the class list on the student dashboard must stay.
    db.session.commit()
    session.pop("attempt_id", None)
    session.pop("student_id", None)
    flash("All previous test history was deleted.", "success")
    return redirect(url_for("history"))


@app.post("/start")
def start_exam():
    name = request.form.get("name", "")
    roll_no = request.form.get("roll_no", "")
    test_id = request.form.get("test_id", type=int)
    section_failure_mode = request.form.get(
        "section_failure_mode",
        session.get("student_setup_failure_mode", "stop"),
    )
    selected_test = db.session.get(Test, test_id)
    sections = ordered_sections(test_id) if selected_test else []
    if not student_access_ok():
        flash("Enter the student password before starting a test.", "error")
        return redirect(url_for("student_login"))
    student = Student.query.filter_by(roll_no=roll_no.strip()).first()
    if not student or student.name.casefold() != name.strip().casefold():
        session["student_form_name"] = name
        session["student_form_roll_no"] = roll_no
        flash("Wrong name or roll number. Please correct the details and try again.", "error")
        return redirect(url_for("student_dashboard"))
    if not selected_test:
        flash("Please choose a valid test.", "error")
        return redirect(url_for("index"))
    if section_failure_mode not in {"all", "stop"}:
        flash("Choose how the test should handle a failed section.", "error")
        return redirect(url_for("index"))
    if not sections:
        flash("No test sections are available yet. Please contact the administrator.", "error")
        return redirect(url_for("index"))
    if any(len(s.questions) == 0 for s in sections):
        flash("Every section must contain at least one question before the test can start.", "error")
        return redirect(url_for("index"))

    attempt = Attempt(student_id=student.id, test_id=selected_test.id, save_history=True)
    db.session.add(attempt)
    db.session.commit()
    session["attempt_id"] = attempt.id
    session["student_id"] = student.id
    session["section_failure_mode"] = "stop"
    start_section(attempt, sections[0])
    lock_student_access()  # the next student must enter the password again
    return redirect(url_for("exam", section_id=sections[0].id))


@app.post("/attempt/history-preference")
def update_history_preference():
    attempt = get_attempt()
    section_id = request.form.get("section_id", type=int)
    if not attempt or attempt.status != "in_progress":
        return redirect(url_for("index"))
    attempt.save_history = True
    db.session.commit()
    return redirect(url_for("exam", section_id=section_id))


@app.route("/exam/<int:section_id>")
def exam(section_id):
    attempt = get_attempt()
    if not attempt or attempt.status != "in_progress":
        return redirect(url_for("index"))

    section = db.session.get(Section, section_id)
    if not section or section.test_id != attempt.test_id:
        return redirect(url_for("index"))

    sa = get_section_attempt(attempt.id, section.id)
    if not sa:
        return redirect(url_for("index"))

    if effective_remaining_seconds(attempt, sa) <= 0:
        result = finalise_section(attempt, section, timed_out=True)
        following = (
            next_section(attempt, section)
            if overall_remaining_seconds(attempt) != 0 and session.get("section_failure_mode", "stop") == "all"
            else None
        )
        if following:
            flash(f"{section.title} failed because time expired. Continuing with {following.title}.", "error")
            return redirect(url_for("exam", section_id=following.id))
        reason = f"Overall test time expired" if overall_remaining_seconds(attempt) == 0 else f"Time expired in section: {section.title}"
        finish_exam(attempt, reason)
        return redirect(url_for("results"))

    answered_ids = {
        r.question_id for r in Response.query.filter_by(attempt_id=attempt.id).all()
    }
    questions = Question.query.filter_by(section_id=section.id).order_by(Question.id.asc()).all()
    current = next((q for q in questions if q.id not in answered_ids), None)

    if current is None:
        result = finalise_section(attempt, section)
        if not result.passed and session.get("section_failure_mode", "stop") == "stop":
            finish_exam(attempt, f"Failed section: {section.title}")
            return redirect(url_for("results"))

        following = next_section(attempt, section)
        if following:
            if result.passed:
                flash(f"{section.title} passed. Moving to {following.title}.", "success")
            else:
                flash(f"{section.title} failed. Continuing with {following.title}.", "error")
            return redirect(url_for("exam", section_id=following.id))

        finish_exam(attempt, "All sections completed")
        return redirect(url_for("results"))

    question_no = questions.index(current) + 1
    section_scores = {}
    for exam_section in ordered_sections(attempt.test_id):
        section_questions = Question.query.filter_by(section_id=exam_section.id).count()
        section_correct = (
            db.session.query(Response)
            .join(Question, Response.question_id == Question.id)
            .filter(
                Response.attempt_id == attempt.id,
                Question.section_id == exam_section.id,
                Response.is_correct.is_(True),
            )
            .count()
        )
        section_scores[exam_section.id] = {
            "correct": section_correct,
            "total": section_questions,
        }
    return render_template(
        "exam.html",
        attempt=attempt,
        test=attempt.test,
        exam_sections=ordered_sections(attempt.test_id),
        test_total_marks=Question.query.join(Section).filter(Section.test_id == attempt.test_id).count(),
        test_total_seconds=total_test_seconds(attempt.test),
        section=section,
        question=current,
        question_no=question_no,
        question_count=len(questions),
        remaining_seconds=effective_remaining_seconds(attempt, sa),
        section_scores=section_scores,
    )


@app.post("/answer/<int:question_id>")
def answer(question_id):
    attempt = get_attempt()
    if not attempt or attempt.status != "in_progress":
        return redirect(url_for("index"))

    question = db.session.get(Question, question_id)
    if not question or question.section.test_id != attempt.test_id:
        return redirect(url_for("index"))

    sa = get_section_attempt(attempt.id, question.section_id)
    if not sa:
        return redirect(url_for("index"))

    if effective_remaining_seconds(attempt, sa) <= 0:
        result = finalise_section(attempt, question.section, timed_out=True)
        following = (
            next_section(attempt, question.section)
            if overall_remaining_seconds(attempt) != 0 and session.get("section_failure_mode", "stop") == "all"
            else None
        )
        if following:
            flash(f"{question.section.title} failed because time expired. Continuing with {following.title}.", "error")
            return redirect(url_for("exam", section_id=following.id))
        reason = "Overall test time expired" if overall_remaining_seconds(attempt) == 0 else f"Time expired in section: {question.section.title}"
        finish_exam(attempt, reason)
        return redirect(url_for("results"))

    existing = Response.query.filter_by(attempt_id=attempt.id, question_id=question.id).first()
    if existing:
        flash("That question was already answered and cannot be changed.", "error")
        return redirect(url_for("exam", section_id=question.section_id))

    selected = request.form.get("answer", "").strip().upper()
    if selected not in {"A", "B", "C", "D"}:
        flash("You must select one answer before continuing.", "error")
        return redirect(url_for("exam", section_id=question.section_id))

    response = Response(
        attempt_id=attempt.id,
        question_id=question.id,
        selected_option=selected,
        is_correct=(selected == question.correct_option.upper()),
    )
    db.session.add(response)
    db.session.commit()
    return redirect(url_for("exam", section_id=question.section_id))


@app.route("/timeout/<int:section_id>")
def timeout(section_id):
    attempt = get_attempt()
    section = db.session.get(Section, section_id)
    if attempt and section and section.test_id == attempt.test_id and attempt.status == "in_progress":
        result = finalise_section(attempt, section, timed_out=True)
        following = (
            next_section(attempt, section)
            if overall_remaining_seconds(attempt) != 0 and session.get("section_failure_mode", "stop") == "all"
            else None
        )
        if following:
            flash(f"{section.title} failed because time expired. Continuing with {following.title}.", "error")
            return redirect(url_for("exam", section_id=following.id))
        reason = "Overall test time expired" if overall_remaining_seconds(attempt) == 0 else f"Time expired in section: {section.title}"
        finish_exam(attempt, reason)
    return redirect(url_for("results"))


@app.route("/results")
def results():
    attempt = get_attempt()
    if not attempt:
        return redirect(url_for("index"))

    section_reports = []
    for sa in sorted(attempt.section_attempts, key=lambda x: (x.section.order_no, x.section.id)):
        questions = Question.query.filter_by(section_id=sa.section_id).order_by(Question.id.asc()).all()
        total = len(questions)
        responses = {
            response.question_id: response
            for response in Response.query.filter_by(attempt_id=attempt.id).all()
        }
        answered = (
            db.session.query(Response)
            .join(Question, Response.question_id == Question.id)
            .filter(Response.attempt_id == attempt.id, Question.section_id == sa.section_id)
            .count()
        )
        correct = (
            db.session.query(Response)
            .join(Question, Response.question_id == Question.id)
            .filter(Response.attempt_id == attempt.id, Question.section_id == sa.section_id, Response.is_correct.is_(True))
            .count()
        )
        review = []
        for question in questions:
            response = responses.get(question.id)
            selected = response.selected_option if response else None
            review.append({
                "text": question.text,
                "selected_option": selected,
                "selected_text": getattr(question, f"option_{selected.lower()}") if selected else "Not answered",
                "correct_option": question.correct_option,
                "correct_text": getattr(question, f"option_{question.correct_option.lower()}"),
                "is_correct": bool(response and response.is_correct),
            })
        section_reports.append({
            "title": sa.section.title,
            "pass_percentage": sa.section.pass_percentage if sa.section else PASS_PERCENTAGE,
            "total": total,
            "answered": answered,
            "correct": correct,
            "percentage": sa.percentage if sa.percentage is not None else 0,
            "passed": bool(sa.passed),
            "timed_out": sa.timed_out,
            "review": review,
        })

    attempted_total = sum(r["total"] for r in section_reports)
    correct_total = sum(r["correct"] for r in section_reports)
    overall_total, overall_correct, overall = test_score(attempt)
    pass_percentage = test_marks_summary(attempt.test)["pass_percentage"] if attempt.test else PASS_PERCENTAGE
    overall_pass = overall >= pass_percentage and attempt.stopped_reason == "All sections completed"

    return render_template(
        "results.html",
        attempt=attempt,
        section_reports=section_reports,
        overall=overall,
        overall_correct=overall_correct,
        overall_total=overall_total,
        overall_pass=overall_pass,
        pass_percentage=pass_percentage,
    )


@app.route("/history/<int:attempt_id>/results")
def history_view_results(attempt_id):
    roll_no = request.args.get("roll_no", "").strip()
    name = request.args.get("name", "").strip()
    from_overall_results = request.args.get("from") == "overall"
    test_selection = request.args.get("test_id", "")
    overall_test_id = request.args.get("test_id", type=int)
    back_url = (
        url_for("overall_results", test_id="all")
        if from_overall_results and test_selection == "all"
        else
        url_for("overall_results", test_id=overall_test_id)
        if from_overall_results and overall_test_id
        else url_for("overall_results") if from_overall_results else url_for("history")
    )
    if roll_no or name:
        from urllib.parse import urlencode
        params = urlencode({"roll_no": roll_no, "name": name})
        if not from_overall_results:
            back_url = f"{back_url}?{params}"

    attempt = db.session.get(Attempt, attempt_id)
    if not attempt:
        flash("Attempt not found.", "error")
        return redirect(url_for("history"))

    section_reports = []
    for sa in sorted(attempt.section_attempts, key=lambda x: (x.section.order_no, x.section.id)):
        questions = Question.query.filter_by(section_id=sa.section_id).order_by(Question.id.asc()).all()
        total = len(questions)
        responses = {
            response.question_id: response
            for response in Response.query.filter_by(attempt_id=attempt.id).all()
        }
        answered = (
            db.session.query(Response)
            .join(Question, Response.question_id == Question.id)
            .filter(Response.attempt_id == attempt.id, Question.section_id == sa.section_id)
            .count()
        )
        correct = (
            db.session.query(Response)
            .join(Question, Response.question_id == Question.id)
            .filter(Response.attempt_id == attempt.id, Question.section_id == sa.section_id, Response.is_correct.is_(True))
            .count()
        )
        review = []
        for question in questions:
            response = responses.get(question.id)
            selected = response.selected_option if response else None
            review.append({
                "text": question.text,
                "selected_option": selected,
                "selected_text": getattr(question, f"option_{selected.lower()}") if selected else "Not answered",
                "correct_option": question.correct_option,
                "correct_text": getattr(question, f"option_{question.correct_option.lower()}"),
                "is_correct": bool(response and response.is_correct),
            })
        section_reports.append({
            "title": sa.section.title,
            "pass_percentage": sa.section.pass_percentage if sa.section else PASS_PERCENTAGE,
            "total": total,
            "answered": answered,
            "correct": correct,
            "percentage": sa.percentage if sa.percentage is not None else 0,
            "passed": bool(sa.passed),
            "timed_out": sa.timed_out,
            "review": review,
        })

    attempted_total = sum(r["total"] for r in section_reports)
    correct_total = sum(r["correct"] for r in section_reports)
    overall_total, overall_correct, overall = test_score(attempt)
    pass_percentage = test_marks_summary(attempt.test)["pass_percentage"] if attempt.test else PASS_PERCENTAGE
    overall_pass = overall >= pass_percentage and attempt.stopped_reason == "All sections completed"

    return render_template(
        "results.html",
        attempt=attempt,
        section_reports=section_reports,
        overall=overall,
        overall_correct=overall_correct,
        overall_total=overall_total,
        overall_pass=overall_pass,
        pass_percentage=pass_percentage,
        is_history=True,
        back_url=back_url,
        back_label="Back to overall result" if from_overall_results else "Back to History",
    )


@app.get("/history/<int:attempt_id>/download-word")
def download_attempt_word(attempt_id):
    attempt = db.session.get(Attempt, attempt_id)
    if not attempt:
        flash("Attempt not found.", "error")
        return redirect(url_for("history"))

    section_reports = attempt_section_reports(attempt)
    overall_total, overall_correct, overall = test_score(attempt)
    pass_percentage = test_marks_summary(attempt.test)["pass_percentage"] if attempt.test else PASS_PERCENTAGE
    overall_pass = overall >= pass_percentage and attempt.stopped_reason == "All sections completed"
    accent = WORD_GREEN if overall_pass else WORD_RED
    tint = "DCFCE7" if overall_pass else "FEE2E2"

    document = word_new_document(landscape=False)
    usable = word_usable_width(document)
    word_heading(document, "Result Report", size=22, before=0, after=8)

    verdict = document.add_table(rows=1, cols=1)
    verdict.alignment = WD_TABLE_ALIGNMENT.CENTER
    verdict.autofit = False
    cell = verdict.rows[0].cells[0]
    cell.width = Inches(usable)
    word_cell_fill(cell, tint)
    word_cell_borders(cell, top=(24, accent), bottom=(24, accent), left=(24, accent), right=(24, accent))
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    cell.paragraphs[0].paragraph_format.space_after = Pt(0)
    failed_sections = [report["title"] for report in section_reports if not report["passed"]]
    if not overall_pass:
        label = "Failed sections:" if len(failed_sections) > 1 else "Failed section:"
        word_box(cell, f"{label} {', '.join(failed_sections)}", WORD_RED, size=11, width=usable - 0.5, gap=6)
    word_box(cell, f"Marks: {overall_correct} / {overall_total}", WORD_BLUE, size=13, width=2.4, gap=4)
    verdict_paragraph = word_paragraph(cell, align="center", before=2, after=2)
    word_run(verdict_paragraph, "PASS" if overall_pass else "FAIL", size=40, bold=True, colour="166534" if overall_pass else WORD_FAIL_TEXT)
    word_box(cell, f"Percentage: {overall}%", WORD_GREEN if overall_pass else WORD_RED, size=11, width=2.6)

    spacer = word_paragraph(document, after=4)
    details = [
        ("Name", attempt.student.name),
        ("Roll number", attempt.student.roll_no),
        ("Test", attempt.test.title if attempt.test else "Test"),
        ("Date and time", to_pkdt(attempt.started_at).strftime("%d %b %Y, %I:%M:%S %p")),
        ("History", "Saved" if attempt.save_history else "Not saved"),
        ("Pass requirement", f"The complete test uses {pass_percentage}% overall."),
        ("Reason", attempt.stopped_reason or "Completed"),
    ]
    for label, value in details:
        line = word_paragraph(document, before=2, after=2)
        word_run(line, f"{label}: ", size=10, bold=True)
        word_run(line, str(value), size=10)
    word_paragraph(document, after=4)

    section_widths = scale_widths([2.1, 0.8, 0.8, 0.7, 0.9, 1.0, 0.9, 1.2], usable)
    table = word_table(document, ["Section", "Answered", "Correct", "Total", "Percentage", "Passing marks", "Pass mark", "Status"], section_widths)
    for report in section_reports:
        cells = word_add_row(table, section_widths)
        failed = not report["passed"]
        fill = WORD_FAIL_ROW if failed else WORD_PASS_BG
        text_colour = WORD_FAIL_TEXT if failed else WORD_INK
        values = [report["title"], report["answered"], report["correct"], report["total"], f"{report['percentage']}%", report["passing_marks"], f"{report['pass_percentage']}%"]
        for index, value in enumerate(values):
            word_cell_text(cells[index], value, size=9.5, bold=failed, colour=text_colour, fill=fill)
        if report["timed_out"]:
            word_cell_pill(cells[6], "TIME EXPIRED", WORD_AMBER, fill, size=8)
        elif report["passed"]:
            word_cell_pill(cells[6], "PASS", WORD_GREEN, fill, size=8)
        else:
            word_cell_pill(cells[6], "FAIL", "B91C1C", fill, size=8)
        if failed:
            for fail_cell in cells:
                word_cell_borders(fail_cell, top=(12, WORD_RED), bottom=(12, WORD_RED))
            word_cell_borders(cells[0], left=(48, WORD_RED), top=(12, WORD_RED), bottom=(12, WORD_RED))

    word_heading(document, "Answer review", size=15, before=14, after=4)
    review_widths = scale_widths([0.4, 2.7, 1.6, 1.6, 0.95], usable)
    for report in section_reports:
        heading = word_paragraph(document, before=8, after=4)
        heading.paragraph_format.keep_with_next = True
        word_run(heading, report["title"], size=12, bold=True, colour="111827")
        if not report["passed"]:
            word_run(heading, "  ")
            word_pill(heading, "FAILED SECTION", WORD_RED, size=8)
        review_table = word_table(document, ["#", "Question", "Your answer", "Correct answer", "Result"], review_widths)
        for number, item in enumerate(report["review"], start=1):
            cells = word_add_row(review_table, review_widths)
            fill = WORD_PASS_BG if item["is_correct"] else WORD_FAIL_BG
            word_cell_text(cells[0], number, size=9, fill=fill)
            word_cell_text(cells[1], item["text"], size=9, fill=fill)
            your_answer = f"{item['selected_option']}. {item['selected_text']}" if item["selected_option"] else "Not answered"
            word_cell_text(cells[2], your_answer, size=9, fill=fill)
            word_cell_text(cells[3], f"{item['correct_option']}. {item['correct_text']}", size=9, fill=fill)
            if item["is_correct"]:
                word_cell_pill(cells[4], "CORRECT", WORD_GREEN, fill, size=8)
            else:
                word_cell_pill(cells[4], "WRONG", "B91C1C", fill, size=8)

    buffer = BytesIO()
    document.save(buffer)
    filename = f"{attempt.student.name.replace(' ', '_')}_result.docx"
    return word_response(buffer, filename)


@app.route("/restart")
def restart():
    session.pop("attempt_id", None)
    session.pop("student_id", None)
    lock_student_access()
    return redirect(url_for("student_login"))


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        password = request.form.get("password", "")
        if valid_admin_password(password):
            session["admin_logged_in"] = True
            session["admin_entry_granted"] = True
            return redirect(url_for("admin_dashboard"))
        flash("Incorrect admin password.", "error")
    return render_template("admin_login.html")


@app.route("/admin/logout")
def admin_logout():
    session.pop("admin_logged_in", None)
    return redirect(url_for("index"))


@app.route("/admin")
@admin_required
def admin_dashboard():
    tests = Test.query.order_by(Test.created_at.asc(), Test.id.asc()).all()
    selected_test_id = request.args.get("test_id", type=int)
    selected_test = db.session.get(Test, selected_test_id) if selected_test_id else None
    sections = ordered_sections(selected_test.id) if selected_test else []
    test_totals = {test.id: total_test_seconds(test) for test in tests}
    test_mark_summaries = {test.id: test_marks_summary(test) for test in tests}
    section_mark_summaries = {
        section.id: {
            "total_marks": section_total_marks(section),
            "passing_marks": get_section_passing_marks(section),
            "pass_percentage": section.pass_percentage,
        }
        for section in sections
    }
    students = sorted(
        Student.query.order_by(Student.id.asc()).all(),
        key=lambda student: (0, int(student.roll_no)) if student.roll_no.strip().isdigit() else (1, student.roll_no.casefold()),
    )
    return render_template(
        "admin.html",
        tests=tests,
        selected_test=selected_test,
        sections=sections,
        selected_test_total_seconds=total_test_seconds(selected_test),
        test_totals=test_totals,
        test_mark_summaries=test_mark_summaries,
        section_mark_summaries=section_mark_summaries,
        students=students,
    )


@app.get("/admin/password-change")
@admin_required
def admin_password_change_page():
    return render_template("admin_password_change.html")


@app.post("/admin/password-change")
@admin_required
def change_admin_password():
    current_password = request.form.get("current_password", "")
    new_password = request.form.get("new_password", "")
    confirm_password = request.form.get("confirm_password", "")
    if ADMIN_PASSWORD_FILE.exists():
        current_valid = check_password_hash(ADMIN_PASSWORD_FILE.read_text(encoding="utf-8").strip(), current_password)
    else:
        current_valid = current_password == os.environ.get("ADMIN_PASSWORD", "admin123")
    if not current_valid:
        flash("Current password is incorrect.", "error")
    elif len(new_password) < 6:
        flash("New password must be at least 6 characters.", "error")
    elif new_password != confirm_password:
        flash("New passwords do not match.", "error")
    else:
        ADMIN_PASSWORD_FILE.parent.mkdir(parents=True, exist_ok=True)
        ADMIN_PASSWORD_FILE.write_text(generate_password_hash(new_password), encoding="utf-8")
        flash("Admin password changed successfully.", "success")
    return redirect(url_for("admin_password_change_page"))


@app.post("/admin/test/add")
@admin_required
def add_test():
    title = request.form.get("title", "").strip()
    if not title:
        flash("Test name is required.", "error")
        return redirect(url_for("admin_dashboard"))
    if Test.query.filter_by(title=title).first():
        flash("A test with this name already exists.", "error")
        return redirect(url_for("admin_dashboard"))
    test = Test(title=title)
    db.session.add(test)
    db.session.commit()
    flash("New test created. You can now import questions into it.", "success")
    return redirect(url_for("admin_dashboard", test_id=test.id))


@app.post("/admin/test/<int:test_id>/delete")
@admin_required
def delete_test(test_id):
    test = db.session.get(Test, test_id)
    if not test:
        return redirect(url_for("admin_dashboard"))
    attempt_ids = [attempt.id for attempt in Attempt.query.filter_by(test_id=test.id).all()]
    if attempt_ids:
        Response.query.filter(Response.attempt_id.in_(attempt_ids)).delete(synchronize_session=False)
        SectionAttempt.query.filter(SectionAttempt.attempt_id.in_(attempt_ids)).delete(synchronize_session=False)
        Attempt.query.filter(Attempt.id.in_(attempt_ids)).delete(synchronize_session=False)
    db.session.delete(test)
    db.session.commit()
    flash("Test and its sections were deleted.", "success")
    return redirect(url_for("admin_dashboard"))


@app.post("/admin/test/<int:test_id>/update")
@admin_required
def update_test(test_id):
    test = db.session.get(Test, test_id)
    title = request.form.get("title", "").strip()
    pass_percentage = request.form.get("pass_percentage", type=float)
    duplicate = Test.query.filter(Test.title == title, Test.id != test_id).first() if title else None
    if not test or not title or duplicate or pass_percentage is None or not 0 <= pass_percentage <= 100:
        flash("Enter a unique test name and an overall passing percentage from 0 to 100.", "error")
        return redirect(url_for("admin_dashboard", test_id=test_id))
    test.title = title
    test.pass_percentage = pass_percentage
    for section in test.sections:
        section.pass_percentage = pass_percentage
    db.session.commit()
    flash("Test settings saved and applied to all sections.", "success")
    return redirect(url_for("admin_dashboard", test_id=test.id))


@app.post("/admin/section/add")
@admin_required
def add_section():
    test_id = request.form.get("test_id", type=int)
    test = db.session.get(Test, test_id)
    title = request.form.get("title", "").strip()
    duration = request.form.get("duration_minutes", type=int)
    seconds = request.form.get("duration_seconds", type=int)
    pass_percentage = request.form.get("pass_percentage", type=float)
    order_no = request.form.get("order_no", type=int)
    if (not test or not title or not duration or duration < 1 or seconds is None or not 0 <= seconds <= 59
            or (pass_percentage is not None and not 0 <= pass_percentage <= 100)):
        flash("Section title and a positive duration are required.", "error")
        return redirect(url_for("admin_dashboard", test_id=test_id) if test_id else url_for("admin_dashboard"))
    if order_no is None:
        order_no = (db.session.query(db.func.max(Section.order_no)).filter_by(test_id=test.id).scalar() or 0) + 1
    current_test_summary = test_marks_summary(test)
    default_pass_percentage = (
        current_test_summary["pass_percentage"]
        if current_test_summary["total_marks"] else test.pass_percentage
    )
    db.session.add(Section(
        test_id=test.id, title=title, duration_minutes=duration, duration_seconds=seconds,
        pass_percentage=pass_percentage if pass_percentage is not None else default_pass_percentage, order_no=order_no,
    ))
    db.session.commit()
    flash("Section added.", "success")
    return redirect(url_for("admin_dashboard", test_id=test.id))


@app.post("/admin/section/<int:section_id>/update")
@admin_required
def update_section(section_id):
    section = db.session.get(Section, section_id)
    duration = request.form.get("duration_minutes", type=int)
    seconds = request.form.get("duration_seconds", type=int)
    title = request.form.get("title", "").strip()
    pass_percentage = request.form.get("pass_percentage", type=float)
    pass_marks = request.form.get("pass_marks", type=int)
    threshold_basis = request.form.get("threshold_basis", "percentage")
    if (not section or not title or duration is None or duration < 1 or seconds is None or not 0 <= seconds <= 59
            or (pass_percentage is not None and not 0 <= pass_percentage <= 100)):
        flash("Enter a section name and valid timing.", "error")
        return redirect(url_for("admin_dashboard", test_id=section.test_id if section else None))
    total_marks = section_total_marks(section)
    if threshold_basis == "marks":
        if pass_marks is None or not 0 <= pass_marks <= total_marks:
            flash("Passing marks must be between zero and the section total.", "error")
            return redirect(url_for("admin_dashboard", test_id=section.test_id) + f"#admin-section-{section.id}")
        pass_percentage = pass_marks / total_marks * 100 if total_marks else 0
    else:
        if pass_percentage is None:
            pass_percentage = section.pass_percentage
    section.title = title
    section.duration_minutes = duration
    section.duration_seconds = seconds
    section.pass_percentage = pass_percentage
    summary = test_marks_summary(section.test)
    section.test.expected_total_marks = summary["total_marks"]
    section.test.pass_percentage = summary["pass_percentage"]
    db.session.commit()
    flash("Section settings updated.", "success")
    return redirect(url_for("admin_dashboard", test_id=section.test_id) + f"#admin-section-{section.id}")


@app.post("/admin/section/<int:section_id>/delete")
@admin_required
def delete_section(section_id):
    section = db.session.get(Section, section_id)
    if section:
        test_id = section.test_id
        question_ids = [q.id for q in Question.query.with_entities(Question.id).filter_by(section_id=section.id).all()]
        if question_ids:
            Response.query.filter(Response.question_id.in_(question_ids)).delete(synchronize_session=False)
        SectionAttempt.query.filter_by(section_id=section.id).delete(synchronize_session=False)
        db.session.delete(section)
        db.session.commit()
        flash("Section and its questions were deleted.", "success")
    return redirect(url_for("admin_dashboard", test_id=test_id if section else None))


@app.post("/admin/sections/delete-all")
@admin_required
def delete_all_sections():
    test_id = request.form.get("test_id", type=int)
    section_ids = [section.id for section in Section.query.filter_by(test_id=test_id).all()]
    question_ids = [question.id for question in Question.query.filter(Question.section_id.in_(section_ids)).all()] if section_ids else []
    if question_ids:
        Response.query.filter(Response.question_id.in_(question_ids)).delete(synchronize_session=False)
    if section_ids:
        SectionAttempt.query.filter(SectionAttempt.section_id.in_(section_ids)).delete(synchronize_session=False)
        Section.query.filter(Section.id.in_(section_ids)).delete(synchronize_session=False)
        Question.query.filter(Question.section_id.in_(section_ids)).delete(synchronize_session=False)
    db.session.commit()
    flash("All sections and questions were deleted.", "success")
    return redirect(url_for("admin_dashboard", test_id=test_id) if test_id else url_for("admin_dashboard"))


@app.post("/admin/question/add")
@admin_required
def add_question():
    section_id = request.form.get("section_id", type=int)
    section = db.session.get(Section, section_id)
    if not section:
        flash("Choose a valid section.", "error")
        return redirect(url_for("admin_dashboard"))

    text = request.form.get("text", "").strip()
    options = {letter: request.form.get(f"option_{letter.lower()}", "").strip() for letter in "ABCD"}
    correct = request.form.get("correct_option", "").upper()
    if not text or any(not v for v in options.values()) or correct not in {"A", "B", "C", "D"}:
        flash("Question, all four option texts, and the correct option are required.", "error")
        return redirect(url_for("admin_dashboard"))

    try:
        q = Question(
            section_id=section.id,
            text=text,
            option_a=options["A"],
            option_b=options["B"],
            option_c=options["C"],
            option_d=options["D"],
            correct_option=correct,
            image=save_image(request.files.get("image")),
            option_a_image=save_image(request.files.get("option_a_image")),
            option_b_image=save_image(request.files.get("option_b_image")),
            option_c_image=save_image(request.files.get("option_c_image")),
            option_d_image=save_image(request.files.get("option_d_image")),
        )
    except ValueError as e:
        flash(str(e), "error")
        return redirect(url_for("admin_dashboard"))

    db.session.add(q)
    db.session.commit()
    flash("Question added.", "success")
    return redirect(url_for("admin_dashboard", test_id=section.test_id))


@app.post("/admin/question/<int:question_id>/update")
@admin_required
def update_question(question_id):
    question = db.session.get(Question, question_id)
    if not question:
        flash("Question not found.", "error")
        return redirect(url_for("admin_dashboard"))

    text = request.form.get("text", "").strip()
    options = {letter: request.form.get(f"option_{letter.lower()}", "").strip() for letter in "ABCD"}
    correct = request.form.get("correct_option", "").upper()
    if not text or any(not v for v in options.values()) or correct not in {"A", "B", "C", "D"}:
        flash("Question, all four option texts, and the correct option are required.", "error")
        return redirect(url_for("admin_dashboard"))

    question.text = text
    question.option_a = options["A"]
    question.option_b = options["B"]
    question.option_c = options["C"]
    question.option_d = options["D"]
    question.correct_option = correct
    db.session.commit()
    flash("Question updated.", "success")
    return redirect(url_for("admin_dashboard", test_id=question.section.test_id))


@app.post("/admin/question/<int:question_id>/delete")
@admin_required
def delete_question(question_id):
    question = db.session.get(Question, question_id)
    if question:
        test_id = question.section.test_id
        Response.query.filter_by(question_id=question.id).delete(synchronize_session=False)
        db.session.delete(question)
        db.session.commit()
        flash("Question deleted.", "success")
        return redirect(url_for("admin_dashboard", test_id=test_id))
    return redirect(url_for("admin_dashboard"))


def parse_import_entries(entries):
    """
    Expected paragraph format:
      SECTION: General Knowledge
      TIME: 10
      Q: Capital of Pakistan?
      A: Lahore
      B: Karachi
      C: Islamabad
      D: Quetta
      ANSWER: C

    Repeat Q/A/B/C/D/ANSWER blocks. SECTION and TIME may appear before each group.
    """
    parsed = []
    current_section = "Imported Section"
    current_time = 10
    block = {}
    last_field = None

    def flush_block():
        nonlocal block, last_field
        required = {"q", "a", "b", "c", "d", "answer"}
        if required.issubset(block):
            answer = block["answer"].strip().upper()[:1]
            if answer in {"A", "B", "C", "D"}:
                parsed.append({
                    "section": current_section,
                    "time": current_time,
                    "q": block["q"],
                    "a": block["a"],
                    "b": block["b"],
                    "c": block["c"],
                    "d": block["d"],
                    "answer": answer,
                    "q_image": block.get("q_image"),
                    "a_image": block.get("a_image"),
                    "b_image": block.get("b_image"),
                    "c_image": block.get("c_image"),
                    "d_image": block.get("d_image"),
                })
        block = {}
        last_field = None

    for line, images in entries:
        if not line:
            if images and last_field and last_field not in block:
                block[last_field] = images[0]
            continue
        upper = line.upper()
        if upper.startswith("SECTION:"):
            flush_block()
            current_section = line.split(":", 1)[1].strip() or "Imported Section"
        elif upper.startswith("TIME:"):
            value = line.split(":", 1)[1].strip()
            match = re.search(r"\d+", value)
            if match:
                current_time = max(1, int(match.group()))
        elif re.match(r"^Q\s*:", line, flags=re.I):
            flush_block()
            block["q"] = line.split(":", 1)[1].strip()
            last_field = "q_image"
        elif re.match(r"^[ABCD]\s*:", line, flags=re.I):
            key = line[0].lower()
            block[key] = line.split(":", 1)[1].strip()
            last_field = f"{key}_image"
        elif upper.startswith("ANSWER:"):
            block["answer"] = line.split(":", 1)[1].strip()
            flush_block()
        if images and last_field and last_field not in block:
            block[last_field] = images[0]
    flush_block()
    return parsed


def parse_docx_questions(path, assign_images_in_order=False):
    doc = Document(path)

    def paragraph_images(paragraph):
        images = []
        for element in paragraph._p.iter():
            if element.tag.endswith("}blip"):
                relationship_id = element.get(qn("r:embed"))
                part = doc.part.related_parts.get(relationship_id)
                if part and part.content_type in {"image/png", "image/jpeg", "image/gif", "image/webp"}:
                    images.append((part.blob, part.content_type))
        return images

    entries = [(p.text.strip(), paragraph_images(p)) for p in doc.paragraphs]
    parsed = parse_import_entries(entries)
    if assign_images_in_order:
        all_images = [image for _, images in entries for image in images]
        used_images = {
            id(item[key]) for item in parsed for key in ("q_image", "a_image", "b_image", "c_image", "d_image")
            if item.get(key)
        }
        remaining_images = [image for image in all_images if id(image) not in used_images]
        for item in parsed:
            if not item.get("q_image") and remaining_images:
                item["q_image"] = remaining_images.pop(0)
    return parsed


def parse_pdf_questions(path):
    """Extract labelled MCQs and displayed images from a text-based PDF."""
    image_types = {"png": "image/png", "jpeg": "image/jpeg", "jpg": "image/jpeg", "gif": "image/gif", "webp": "image/webp"}
    entries = []
    with fitz.open(path) as pdf:
        for page in pdf:
            blocks = sorted(page.get_text("dict")["blocks"], key=lambda block: (block["bbox"][1], block["bbox"][0]))
            for block in blocks:
                if block["type"] == 0:
                    text = "\n".join(
                        "".join(span["text"] for span in line["spans"])
                        for line in block.get("lines", [])
                    )
                    for line in text.splitlines():
                        if line.strip():
                            entries.append((line.strip(), []))
                elif block["type"] == 1:
                    media_type = image_types.get(block.get("ext", "").lower())
                    if media_type and block.get("image"):
                        entries.append(("", [(block["image"], media_type)]))
    return parse_import_entries(entries)


def parse_uploaded_questions(path, assign_images_in_order=False):
    if path.suffix.lower() == ".pdf":
        return parse_pdf_questions(path)
    return parse_docx_questions(path, assign_images_in_order=assign_images_in_order)


def extract_docx_for_ai(path):
    """Return readable text and embedded image files from a Word document."""
    doc = Document(path)
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            if any(cells):
                lines.append(" | ".join(cells))

    images = []
    with zipfile.ZipFile(path) as archive:
        files = sorted(name for name in archive.namelist() if name.startswith("word/media/"))
        for name in files[:20]:
            media_type = mimetypes.guess_type(name)[0] or "image/png"
            if media_type in {"image/png", "image/jpeg", "image/gif", "image/webp"}:
                images.append({"data": archive.read(name), "media_type": media_type})
    return "\n".join(lines), images


def claude_extract_questions(text, images):
    api_key = os.environ.get("CLAUDE_API_KEY")
    if not api_key:
        raise RuntimeError("CLAUDE_API_KEY is not configured on the server.")
    prompt = """You extract exam papers. Read this Word document text and the attached images.
Extract only complete MCQs: a question, exactly four choices, and a determinable correct answer.
Images are numbered in attachment order starting at 1; only associate them where clearly relevant.
Return ONLY valid JSON with this exact schema:
{"sections":[{"title":"string","duration_minutes":10,"questions":[{"text":"string","options":{"A":"string","B":"string","C":"string","D":"string"},"correct_option":"A","image_index":null,"option_image_indices":{"A":null,"B":null,"C":null,"D":null}}]}]}
Use Imported Section and 10 minutes where absent. Never invent questions, answers, or image associations.

DOCUMENT TEXT:
""" + text
    content = [{"type": "text", "text": prompt}]
    import base64
    for image in images:
        content.append({"type": "image", "source": {
            "type": "base64", "media_type": image["media_type"],
            "data": base64.b64encode(image["data"]).decode("ascii"),
        }})
    response = Anthropic(api_key=api_key).messages.create(
        model=CLAUDE_MODEL, max_tokens=16000,
        messages=[{"role": "user", "content": content}],
    )
    answer = "".join(block.text for block in response.content if block.type == "text").strip()
    answer = re.sub(r"^```(?:json)?\s*|\s*```$", "", answer, flags=re.I).strip()
    try:
        return json.loads(answer)
    except json.JSONDecodeError as exc:
        raise RuntimeError("Claude returned an invalid import response. Please try again.") from exc


def save_imported_image(images, index):
    if not isinstance(index, int) or not 1 <= index <= len(images):
        return None
    image = images[index - 1]
    extension = {"image/png": "png", "image/jpeg": "jpg", "image/gif": "gif", "image/webp": "webp"}[image["media_type"]]
    filename = f"{uuid.uuid4().hex}.{extension}"
    (UPLOAD_DIR / filename).write_bytes(image["data"])
    return filename


def save_ai_import(payload, images, test_id):
    sections_data = payload.get("sections", []) if isinstance(payload, dict) else []
    if not isinstance(sections_data, list):
        return 0
    next_order = (db.session.query(db.func.max(Section.order_no)).filter_by(test_id=test_id).scalar() or 0) + 1
    imported = 0
    for raw_section in sections_data:
        if not isinstance(raw_section, dict) or not isinstance(raw_section.get("questions"), list):
            continue
        title = str(raw_section.get("title", "Imported Section")).strip()[:200] or "Imported Section"
        duration = raw_section.get("duration_minutes", 10)
        duration = duration if isinstance(duration, int) and duration > 0 else 10
        section = Section.query.filter_by(test_id=test_id, title=title).first()
        if not section:
            section = Section(test_id=test_id, title=title, duration_minutes=duration, order_no=next_order)
            next_order += 1
            db.session.add(section)
            db.session.flush()
        for item in raw_section["questions"]:
            options = item.get("options", {}) if isinstance(item, dict) else {}
            correct = str(item.get("correct_option", "")).upper() if isinstance(item, dict) else ""
            question_text = str(item.get("text", "")).strip() if isinstance(item, dict) else ""
            if (not question_text or not isinstance(options, dict) or correct not in {"A", "B", "C", "D"}
                    or any(not str(options.get(letter, "")).strip() for letter in "ABCD")):
                continue
            option_images = item.get("option_image_indices", {})
            option_images = option_images if isinstance(option_images, dict) else {}
            db.session.add(Question(
                section_id=section.id, text=question_text, correct_option=correct,
                option_a=str(options["A"]).strip(), option_b=str(options["B"]).strip(),
                option_c=str(options["C"]).strip(), option_d=str(options["D"]).strip(),
                image=save_imported_image(images, item.get("image_index")),
                option_a_image=save_imported_image(images, option_images.get("A")),
                option_b_image=save_imported_image(images, option_images.get("B")),
                option_c_image=save_imported_image(images, option_images.get("C")),
                option_d_image=save_imported_image(images, option_images.get("D")),
            ))
            imported += 1
    return imported


@app.post("/admin/import-docx")
@admin_required
def import_docx():
    test_id = request.form.get("test_id", type=int)
    test = db.session.get(Test, test_id)
    if not test:
        flash("Choose a test before importing questions.", "error")
        return redirect(url_for("admin_dashboard"))
    uploads = [file for file in request.files.getlist("document_files") if file and file.filename]
    if not uploads:
        flash("Choose at least one .docx or .pdf file.", "error")
        return redirect(url_for("admin_dashboard", test_id=test.id))
    if any(not allowed_file(uploaded.filename, ALLOWED_DOC_EXTENSIONS) for uploaded in uploads):
        flash("Only .docx and .pdf files are supported.", "error")
        return redirect(url_for("admin_dashboard", test_id=test.id))

    temp_paths = []
    parsed = []
    assign_images_in_order = request.form.get("assign_images_in_order") == "yes"
    try:
        for uploaded in uploads:
            extension = Path(uploaded.filename).suffix.lower()
            temp_path = BASE_DIR / f"import_{uuid.uuid4().hex}{extension}"
            temp_paths.append(temp_path)
            uploaded.save(temp_path)
            file_questions = parse_uploaded_questions(
                temp_path, assign_images_in_order=assign_images_in_order
            )
            parsed.extend(file_questions)
    except Exception as exc:
        flash(f"Could not read the selected file: {exc}", "error")
        return redirect(url_for("admin_dashboard", test_id=test.id))
    finally:
        for temp_path in temp_paths:
            if temp_path.exists():
                temp_path.unlink()

    if not parsed:
        flash("No questions were found. Check the required import format shown on the admin page.", "error")
        return redirect(url_for("admin_dashboard"))

    if request.form.get("combine_into_one") == "yes":
        combined_title = request.form.get("combined_section_title", "").strip() or "Combined Import"
        combined_duration = request.form.get("combined_duration_minutes", type=int) or 60
        combined_duration = max(1, combined_duration)
        for item in parsed:
            item["section"] = combined_title
            item["time"] = combined_duration

    section_cache = {}
    next_order = (db.session.query(db.func.max(Section.order_no)).filter_by(test_id=test.id).scalar() or 0) + 1
    imported_count = 0
    for item in parsed:
        key = item["section"].strip()
        if key not in section_cache:
            section = Section.query.filter_by(test_id=test.id, title=key).first()
            if not section:
                section = Section(test_id=test.id, title=key, duration_minutes=item["time"], order_no=next_order)
                next_order += 1
                db.session.add(section)
                db.session.flush()
            section_cache[key] = section
        section = section_cache[key]
        db.session.add(Question(
            section_id=section.id,
            text=item["q"],
            option_a=item["a"],
            option_b=item["b"],
            option_c=item["c"],
            option_d=item["d"],
            correct_option=item["answer"],
            image=save_embedded_docx_image(item.get("q_image")),
            option_a_image=save_embedded_docx_image(item.get("a_image")),
            option_b_image=save_embedded_docx_image(item.get("b_image")),
            option_c_image=save_embedded_docx_image(item.get("c_image")),
            option_d_image=save_embedded_docx_image(item.get("d_image")),
        ))
        imported_count += 1

    db.session.commit()
    flash(f"Imported {imported_count} questions from {len(uploads)} file(s).", "success")
    return redirect(url_for("admin_dashboard", test_id=test.id))


@app.post("/admin/import-docx-ai")
@admin_required
def import_docx_ai():
    test_id = request.form.get("test_id", type=int)
    test = db.session.get(Test, test_id)
    if not test:
        flash("Choose a test before importing questions.", "error")
        return redirect(url_for("admin_dashboard"))
    uploads = [f for f in request.files.getlist("word_files") if f and f.filename]
    if not uploads:
        flash("Choose at least one valid .docx file.", "error")
        return redirect(url_for("admin_dashboard", test_id=test.id))
    if any(not allowed_file(f.filename, {"docx"}) for f in uploads):
        flash("Only .docx files are supported for AI import.", "error")
        return redirect(url_for("admin_dashboard", test_id=test.id))

    combine_into_one = request.form.get("combine_into_one") == "yes"
    combined_title = request.form.get("combined_section_title", "").strip() or "Combined Import"
    combined_duration = max(1, request.form.get("combined_duration_minutes", type=int) or 60)

    temp_paths = []
    all_text_parts = []
    all_images = []
    try:
        for uploaded in uploads:
            temp_path = BASE_DIR / f"ai_import_{uuid.uuid4().hex}.docx"
            temp_paths.append(temp_path)
            uploaded.save(temp_path)
            text, images = extract_docx_for_ai(temp_path)
            if text:
                all_text_parts.append(text)
            all_images.extend(images)
    except Exception as exc:
        flash(f"Could not read selected files: {exc}", "error")
        return redirect(url_for("admin_dashboard", test_id=test.id))
    finally:
        for temp_path in temp_paths:
            if temp_path.exists():
                temp_path.unlink()

    if not all_text_parts and not all_images:
        flash("The selected Word files are empty.", "error")
        return redirect(url_for("admin_dashboard"))

    try:
        combined_text = "\n\n--- PAGE BREAK ---\n\n".join(all_text_parts)
        payload = claude_extract_questions(combined_text, all_images)
        if combine_into_one:
            sections_data = payload.get("sections", []) if isinstance(payload, dict) else []
            if sections_data:
                for raw_section in sections_data:
                    if isinstance(raw_section, dict):
                        raw_section["title"] = combined_title
                        raw_section["duration_minutes"] = combined_duration
        imported_count = save_ai_import(payload, all_images, test.id)
        if not imported_count:
            raise RuntimeError("No complete MCQs were found. The documents need questions, four options, and answers.")
        db.session.commit()
    except Exception as exc:
        db.session.rollback()
        flash(f"AI import failed: {exc}", "error")
        return redirect(url_for("admin_dashboard", test_id=test.id))

    flash(f"AI imported {imported_count} questions from {len(uploads)} file(s).", "success")
    return redirect(url_for("admin_dashboard", test_id=test.id))


if __name__ == "__main__":
    app.run(debug=True)
