from flask import Flask, render_template, request, redirect, url_for, session, send_from_directory, jsonify
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import openpyxl
import os

app = Flask(__name__)
app.secret_key = "mps_lms_secret_2025"

# ─────────────────────────────────────────────
# DATABASE — Supabase PostgreSQL (persistent, free, never expires)
# ─────────────────────────────────────────────
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get(
    "DATABASE_URL",
    "postgresql://postgres:Kratika%4012345@db.fxyyepynocnjegevvsbq.supabase.co:5432/postgres"
)
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {"pool_pre_ping": True}
app.config["UPLOAD_FOLDER"] = "uploads"
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

db = SQLAlchemy(app)

# ─────────────────────────────────────────────
# MODELS
# ─────────────────────────────────────────────

class Admin(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True)
    password = db.Column(db.String(50))


class Teacher(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True)
    password = db.Column(db.String(50))
    name = db.Column(db.String(100))
    subject = db.Column(db.String(100), default="Computer Science")


class Student(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    admission_no = db.Column(db.String(50), unique=True)
    student_name = db.Column(db.String(100))
    dob = db.Column(db.String(20))
    student_class = db.Column(db.String(10))
    section = db.Column(db.String(10))
    roll_no = db.Column(db.String(10))
    parent_mobile = db.Column(db.String(15))
    status = db.Column(db.String(20), default="Active")


class Note(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(200))
    original_name = db.Column(db.String(200))
    chapter = db.Column(db.String(100))
    chapter_no = db.Column(db.Integer, default=0)
    section = db.Column(db.String(20))
    description = db.Column(db.String(300))
    file_type = db.Column(db.String(20), default="notes")
    uploaded_at = db.Column(db.String(30), default="")


class Assignment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200))
    chapter = db.Column(db.String(100))
    chapter_no = db.Column(db.Integer, default=0)
    section = db.Column(db.String(20))
    description = db.Column(db.Text)
    due_date = db.Column(db.String(30))
    created_at = db.Column(db.String(30), default="")
    filename = db.Column(db.String(200), default="")


class ChapterProgress(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey("student.id"))
    chapter_no = db.Column(db.Integer)
    status = db.Column(db.String(20), default="pending")
    updated_at = db.Column(db.String(30), default="")


class AssignmentSubmission(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey("student.id"))
    assignment_id = db.Column(db.Integer, db.ForeignKey("assignment.id"))
    status = db.Column(db.String(20), default="pending")
    submitted_at = db.Column(db.String(30), default="")
    remarks = db.Column(db.String(300), default="")


# ─────────────────────────────────────────────
# CHAPTERS & SECTIONS
# ─────────────────────────────────────────────

CHAPTERS = [
    {"no": 1, "name": "Internet Services"},
    {"no": 2, "name": "More on Excel"},
    {"no": 3, "name": "Formulas & Functions"},
    {"no": 4, "name": "Canva"},
    {"no": 5, "name": "Introduction to GIMP"},
    {"no": 6, "name": "Present Yourself Online"},
    {"no": 7, "name": "Introduction to Python"},
    {"no": 8, "name": "Introduction to Data Science"},
    {"no": 9, "name": "Robots & Sensors"},
]

SECTIONS = ["6A", "6B", "6C", "6D", "6E", "6F", "6G"]

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def now_str():
    return datetime.now().strftime("%d %b %Y, %I:%M %p")

def safe_filename(filename):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_")
    return ts + filename.replace(" ", "_")

# ─────────────────────────────────────────────
# HOME
# ─────────────────────────────────────────────

@app.route("/")
def home():
    return render_template("home.html")

# ─────────────────────────────────────────────
# ADMIN LOGIN
# ─────────────────────────────────────────────

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]
        admin = Admin.query.filter_by(username=username, password=password).first()
        if admin:
            session["admin"] = username
            return redirect("/admin")
        return render_template("login.html", error="Invalid username or password!")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

# ─────────────────────────────────────────────
# TEACHER LOGIN
# ─────────────────────────────────────────────

@app.route("/teacher_login", methods=["GET", "POST"])
def teacher_login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]
        teacher = Teacher.query.filter_by(username=username, password=password).first()
        if teacher:
            session["teacher"] = username
            session["teacher_name"] = teacher.name
            return redirect("/teacher")
        return render_template("teacher_login.html", error="Invalid credentials!")
    return render_template("teacher_login.html")

@app.route("/teacher_logout")
def teacher_logout():
    session.pop("teacher", None)
    session.pop("teacher_name", None)
    return redirect("/")

# ─────────────────────────────────────────────
# ADMIN DASHBOARD
# ─────────────────────────────────────────────

@app.route("/admin")
def admin():
    if "admin" not in session:
        return redirect("/login")
    total_students = Student.query.count()
    total_notes = Note.query.filter_by(file_type="notes").count()
    total_books = Note.query.filter_by(file_type="book").count()
    total_quizzes = Note.query.filter(Note.file_type.in_(["quiz", "test"])).count()
    total_assignments = Assignment.query.count()
    recent_notes = Note.query.order_by(Note.id.desc()).limit(5).all()
    recent_assignments = Assignment.query.order_by(Assignment.id.desc()).limit(5).all()
    section_counts = {}
    for sec in SECTIONS:
        section_counts[sec] = Student.query.filter_by(section=sec).count()
    return render_template("admin.html",
        total_students=total_students,
        total_notes=total_notes,
        total_books=total_books,
        total_quizzes=total_quizzes,
        total_assignments=total_assignments,
        recent_notes=recent_notes,
        recent_assignments=recent_assignments,
        section_counts=section_counts,
        sections=SECTIONS
    )

# ─────────────────────────────────────────────
# STUDENT MANAGEMENT
# ─────────────────────────────────────────────

@app.route("/students")
def students():
    if "admin" not in session:
        return redirect("/login")
    search = request.args.get("search", "")
    section = request.args.get("section", "")
    data = Student.query
    if search:
        data = data.filter(
            (Student.student_name.contains(search)) |
            (Student.admission_no.contains(search))
        )
    if section:
        data = data.filter_by(section=section)
    students_list = data.all()
    total = Student.query.count()
    return render_template("students.html", students=students_list, total=total, sections=SECTIONS, search=search, section=section)

@app.route("/delete/<int:id>")
def delete_student(id):
    if "admin" not in session:
        return redirect("/login")
    student = db.session.get(Student, id)
    if student:
        db.session.delete(student)
        db.session.commit()
    return redirect("/students")

@app.route("/upload_students", methods=["GET", "POST"])
def upload_students():
    if "admin" not in session:
        return redirect("/login")
    if request.method == "POST":
        file = request.files["file"]
        os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
        path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
        file.save(path)
        added = 0
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        headers = [cell.value for cell in ws[1]]
        for row in ws.iter_rows(min_row=2, values_only=True):
            data = dict(zip(headers, row))
            if not Student.query.filter_by(admission_no=str(data["Admission_No"])).first():
                student = Student(
                    admission_no=str(data["Admission_No"]),
                    student_name=data["Student_Name"],
                    dob=str(data["DOB"]),
                    student_class=str(data["Class"]),
                    section=str(data["Section"]),
                    roll_no=str(data["Roll_No"]),
                    parent_mobile=str(data["Parent_Mobile"])
                )
                db.session.add(student)
                added += 1
        db.session.commit()
        return render_template("success.html", total=added, filename=file.filename)
    return render_template("upload.html")

# ─────────────────────────────────────────────
# TEACHER PORTAL
# ─────────────────────────────────────────────

@app.route("/teacher")
def teacher():
    if "teacher" not in session and "admin" not in session:
        return redirect("/teacher_login")
    notes = Note.query.filter_by(file_type="notes").order_by(Note.id.desc()).all()
    books = Note.query.filter_by(file_type="book").order_by(Note.id.desc()).all()
    quizzes = Note.query.filter(Note.file_type.in_(["quiz", "test"])).order_by(Note.id.desc()).all()
    assignments = Assignment.query.order_by(Assignment.id.desc()).all()
    total_students = Student.query.count()
    teacher_name = session.get("teacher_name", "Teacher")
    return render_template("teacher.html",
        notes=notes,
        books=books,
        quizzes=quizzes,
        assignments=assignments,
        chapters=CHAPTERS,
        sections=SECTIONS,
        total_students=total_students,
        teacher_name=teacher_name
    )

@app.route("/upload_notes", methods=["POST"])
def upload_notes():
    if "teacher" not in session and "admin" not in session:
        return redirect("/teacher_login")
    file = request.files.get("file")
    chapter = request.form["chapter"]
    chapter_no = int(request.form.get("chapter_no", 0))
    section = request.form["section"]
    description = request.form.get("description", "")
    file_type = request.form.get("file_type", "notes")

    if file and file.filename:
        os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
        safe_name = safe_filename(file.filename)
        path = os.path.join(app.config["UPLOAD_FOLDER"], safe_name)
        file.save(path)
        note = Note(
            filename=safe_name,
            original_name=file.filename,
            chapter=chapter,
            chapter_no=chapter_no,
            section=section,
            description=description,
            file_type=file_type,
            uploaded_at=now_str()
        )
        db.session.add(note)
        db.session.commit()
    return redirect("/teacher")

@app.route("/delete_note/<int:id>")
def delete_note(id):
    if "teacher" not in session and "admin" not in session:
        return redirect("/teacher_login")
    note = db.session.get(Note, id)
    if note:
        fpath = os.path.join(app.config["UPLOAD_FOLDER"], note.filename)
        if os.path.exists(fpath):
            os.remove(fpath)
        db.session.delete(note)
        db.session.commit()
    return redirect("/teacher")

@app.route("/create_assignment", methods=["POST"])
def create_assignment():
    if "teacher" not in session and "admin" not in session:
        return redirect("/teacher_login")
    title = request.form["title"]
    chapter = request.form["chapter"]
    chapter_no = int(request.form.get("chapter_no", 0))
    section = request.form["section"]
    description = request.form.get("description", "")
    due_date = request.form.get("due_date", "")

    filename = ""
    file = request.files.get("file")
    if file and file.filename:
        os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
        safe_name = safe_filename(file.filename)
        file.save(os.path.join(app.config["UPLOAD_FOLDER"], safe_name))
        filename = safe_name

    assignment = Assignment(
        title=title,
        chapter=chapter,
        chapter_no=chapter_no,
        section=section,
        description=description,
        due_date=due_date,
        filename=filename,
        created_at=now_str()
    )
    db.session.add(assignment)
    db.session.commit()
    return redirect("/teacher")

@app.route("/delete_assignment/<int:id>")
def delete_assignment(id):
    if "teacher" not in session and "admin" not in session:
        return redirect("/teacher_login")
    a = db.session.get(Assignment, id)
    if a:
        db.session.delete(a)
        db.session.commit()
    return redirect("/teacher")

# ─────────────────────────────────────────────
# STUDENT PORTAL
# ─────────────────────────────────────────────

@app.route("/student")
def student_portal():
    student_data = None
    notes = []
    books = []
    quizzes = []
    assignments = []
    progress_map = {}
    submitted_ids = set()
    chapter_stats = {}

    if "student_id" in session:
        student_data = db.session.get(Student, session["student_id"])
        if student_data:
            sec = student_data.section

            def get_materials(ftype):
                return Note.query.filter(
                    Note.file_type == ftype,
                    (Note.section == sec) | (Note.section == "All Sections")
                ).order_by(Note.chapter_no, Note.id).all()

            notes = get_materials("notes")
            books = get_materials("book")
            quizzes = Note.query.filter(
                Note.file_type.in_(["quiz", "test"]),
                (Note.section == sec) | (Note.section == "All Sections")
            ).order_by(Note.chapter_no, Note.id).all()

            assignments = Assignment.query.filter(
                (Assignment.section == sec) | (Assignment.section == "") | (Assignment.section == "All Sections")
            ).order_by(Assignment.chapter_no, Assignment.id).all()

            progs = ChapterProgress.query.filter_by(student_id=student_data.id).all()
            progress_map = {p.chapter_no: p.status for p in progs}

            subs = AssignmentSubmission.query.filter_by(student_id=student_data.id).all()
            submitted_ids = {s.assignment_id: s.status for s in subs}

            for ch in CHAPTERS:
                cno = ch["no"]
                chapter_stats[cno] = {
                    "notes": sum(1 for n in notes if n.chapter_no == cno),
                    "books": sum(1 for b in books if b.chapter_no == cno),
                    "quizzes": sum(1 for q in quizzes if q.chapter_no == cno),
                    "assignments": sum(1 for a in assignments if a.chapter_no == cno),
                    "status": progress_map.get(cno, "pending"),
                }

    completed = sum(1 for s in progress_map.values() if s == "completed")
    in_progress = sum(1 for s in progress_map.values() if s == "in_progress")

    return render_template("student.html",
        student=student_data,
        notes=notes,
        books=books,
        quizzes=quizzes,
        assignments=assignments,
        chapters=CHAPTERS,
        progress_map=progress_map,
        submitted_ids=submitted_ids,
        chapter_stats=chapter_stats,
        completed=completed,
        in_progress=in_progress,
        total_chapters=len(CHAPTERS)
    )

@app.route("/student_login", methods=["POST"])
def student_login():
    admission_no = request.form["admission_no"]
    dob = request.form.get("dob", "").strip()
    student = Student.query.filter_by(admission_no=admission_no).first()
    if student and (not dob or student.dob == dob):
        session["student_id"] = student.id
        return redirect("/student")
    return render_template("student.html",
        student=None, error="Admission No not found. Please contact your teacher.",
        chapters=CHAPTERS, notes=[], books=[], quizzes=[], assignments=[],
        progress_map={}, submitted_ids={}, chapter_stats={},
        completed=0, in_progress=0, total_chapters=len(CHAPTERS)
    )

@app.route("/student_logout")
def student_logout():
    session.pop("student_id", None)
    return redirect("/student")

@app.route("/update_progress", methods=["POST"])
def update_progress():
    if "student_id" not in session:
        return jsonify({"ok": False})
    student_id = session["student_id"]
    chapter_no = int(request.form["chapter_no"])
    status = request.form["status"]
    existing = ChapterProgress.query.filter_by(student_id=student_id, chapter_no=chapter_no).first()
    if existing:
        existing.status = status
        existing.updated_at = now_str()
    else:
        prog = ChapterProgress(student_id=student_id, chapter_no=chapter_no, status=status, updated_at=now_str())
        db.session.add(prog)
    db.session.commit()
    return redirect("/student")

@app.route("/submit_assignment/<int:assignment_id>")
def submit_assignment(assignment_id):
    if "student_id" not in session:
        return redirect("/student")
    student_id = session["student_id"]
    existing = AssignmentSubmission.query.filter_by(student_id=student_id, assignment_id=assignment_id).first()
    if not existing:
        sub = AssignmentSubmission(
            student_id=student_id,
            assignment_id=assignment_id,
            status="submitted",
            submitted_at=now_str()
        )
        db.session.add(sub)
        db.session.commit()
    return redirect("/student")

# ─────────────────────────────────────────────
# FILE SERVE
# ─────────────────────────────────────────────

@app.route("/uploads/<filename>")
def serve_upload(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)

# ─────────────────────────────────────────────
# DB INIT — runs on every startup
# ─────────────────────────────────────────────

with app.app_context():
    os.makedirs("uploads", exist_ok=True)
    db.create_all()
    if Admin.query.count() == 0:
        db.session.add(Admin(username="admin", password="1234"))
        db.session.commit()
    if Teacher.query.count() == 0:
        db.session.add(Teacher(username="teacher", password="1234", name="Computer Teacher"))
        db.session.commit()

if __name__ == "__main__":
    app.run(debug=True)
