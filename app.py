from flask import Flask, render_template, request, redirect, session, send_from_directory, jsonify
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import openpyxl
import os

app = Flask(__name__)
app.secret_key = "mps_lms_secret_2025"

app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get(
    "DATABASE_URL",
    "postgresql://postgres.fxyyepynocnjegevvsbq:Kratika%4012345@aws-1-ap-northeast-1.pooler.supabase.com:5432/postgres"
)
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {"pool_pre_ping": True}
app.config["UPLOAD_FOLDER"] = "uploads"
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

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
    video_done = db.Column(db.Boolean, default=False)
    notes_done = db.Column(db.Boolean, default=False)
    book_done = db.Column(db.Boolean, default=False)
    quiz_done = db.Column(db.Boolean, default=False)
    game_done = db.Column(db.Boolean, default=False)
    feedback_done = db.Column(db.Boolean, default=False)
    updated_at = db.Column(db.String(30), default="")

class AssignmentSubmission(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey("student.id"))
    assignment_id = db.Column(db.Integer, db.ForeignKey("assignment.id"))
    status = db.Column(db.String(20), default="pending")
    submitted_at = db.Column(db.String(30), default="")
    remarks = db.Column(db.String(300), default="")

class Quiz(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200))
    chapter = db.Column(db.String(100))
    chapter_no = db.Column(db.Integer, default=0)
    section = db.Column(db.String(20))
    description = db.Column(db.String(300), default="")
    quiz_type = db.Column(db.String(20), default="mcq")  # mcq / link
    external_link = db.Column(db.String(500), default="")
    created_at = db.Column(db.String(30), default="")
    questions = db.relationship("QuizQuestion", backref="quiz", cascade="all, delete-orphan")
    attempts = db.relationship("QuizAttempt", backref="quiz", cascade="all, delete-orphan")

class QuizQuestion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    quiz_id = db.Column(db.Integer, db.ForeignKey("quiz.id"))
    question_text = db.Column(db.Text)
    option_a = db.Column(db.String(300))
    option_b = db.Column(db.String(300))
    option_c = db.Column(db.String(300))
    option_d = db.Column(db.String(300))
    correct_option = db.Column(db.String(1))

class QuizAttempt(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey("student.id"))
    quiz_id = db.Column(db.Integer, db.ForeignKey("quiz.id"))
    score = db.Column(db.Integer, default=0)
    total = db.Column(db.Integer, default=0)
    attempted_at = db.Column(db.String(30), default="")
    student = db.relationship("Student", backref="quiz_attempts")

class Game(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200))
    chapter = db.Column(db.String(100))
    chapter_no = db.Column(db.Integer, default=0)
    section = db.Column(db.String(20))
    created_at = db.Column(db.String(30), default="")
    pairs = db.relationship("GamePair", backref="game", cascade="all, delete-orphan")
    scores = db.relationship("GameScore", backref="game", cascade="all, delete-orphan")

class GamePair(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    game_id = db.Column(db.Integer, db.ForeignKey("game.id"))
    term = db.Column(db.String(200))
    definition = db.Column(db.String(200))

class GameScore(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey("student.id"))
    game_id = db.Column(db.Integer, db.ForeignKey("game.id"))
    score = db.Column(db.Integer, default=0)
    time_seconds = db.Column(db.Integer, default=999)
    played_at = db.Column(db.String(30), default="")
    student = db.relationship("Student", backref="game_scores")

class ChapterVideo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    chapter_no = db.Column(db.Integer)
    chapter = db.Column(db.String(100))
    section = db.Column(db.String(20))
    youtube_url = db.Column(db.String(500))
    title = db.Column(db.String(200))
    added_at = db.Column(db.String(30), default="")

class ChapterConfig(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    chapter_no = db.Column(db.Integer, unique=True)
    is_active = db.Column(db.Boolean, default=False)
    active_month = db.Column(db.String(20), default="")
    has_video = db.Column(db.Boolean, default=False)
    has_notes = db.Column(db.Boolean, default=False)
    has_book = db.Column(db.Boolean, default=False)
    has_quiz = db.Column(db.Boolean, default=False)
    has_game = db.Column(db.Boolean, default=False)
    has_feedback = db.Column(db.Boolean, default=True)

class Feedback(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey("student.id"))
    chapter_no = db.Column(db.Integer)
    rating = db.Column(db.Integer, default=5)
    comment = db.Column(db.Text, default="")
    submitted_at = db.Column(db.String(30), default="")
    student = db.relationship("Student", backref="feedbacks")

# ─────────────────────────────────────────────
# CONSTANTS
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
MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"]

def now_str():
    return datetime.now().strftime("%d %b %Y, %I:%M %p")

def safe_filename(filename):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_")
    return ts + filename.replace(" ", "_")

def get_youtube_embed(url):
    if not url:
        return ""
    if "embed" in url:
        return url
    if "youtu.be/" in url:
        vid = url.split("youtu.be/")[-1].split("?")[0]
    elif "v=" in url:
        vid = url.split("v=")[-1].split("&")[0]
    else:
        return url
    return f"https://www.youtube.com/embed/{vid}"

def get_chapter_progress_pct(prog, config):
    if not prog or not config:
        return 0
    total = 0
    done = 0
    if config.has_video:
        total += 1
        if prog.video_done: done += 1
    if config.has_notes:
        total += 1
        if prog.notes_done: done += 1
    if config.has_book:
        total += 1
        if prog.book_done: done += 1
    if config.has_quiz:
        total += 1
        if prog.quiz_done: done += 1
    if config.has_game:
        total += 1
        if prog.game_done: done += 1
    if config.has_feedback:
        total += 1
        if prog.feedback_done: done += 1
    if total == 0:
        return 0
    return int((done / total) * 100)

# ─────────────────────────────────────────────
# HOME
# ─────────────────────────────────────────────

@app.route("/")
def home():
    return render_template("home.html")

# ─────────────────────────────────────────────
# AUTH
# ─────────────────────────────────────────────

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        admin = Admin.query.filter_by(username=request.form["username"], password=request.form["password"]).first()
        if admin:
            session["admin"] = admin.username
            return redirect("/admin")
        return render_template("login.html", error="Invalid credentials!")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

@app.route("/teacher_login", methods=["GET","POST"])
def teacher_login():
    if request.method == "POST":
        teacher = Teacher.query.filter_by(username=request.form["username"], password=request.form["password"]).first()
        if teacher:
            session["teacher"] = teacher.username
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
# ADMIN
# ─────────────────────────────────────────────

@app.route("/admin")
def admin():
    if "admin" not in session:
        return redirect("/login")
    total_students = Student.query.count()
    total_notes = Note.query.filter_by(file_type="notes").count()
    total_books = Note.query.filter_by(file_type="book").count()
    total_quizzes = Quiz.query.count()
    total_assignments = Assignment.query.count()
    recent_notes = Note.query.order_by(Note.id.desc()).limit(5).all()
    recent_assignments = Assignment.query.order_by(Assignment.id.desc()).limit(5).all()
    section_counts = {sec: Student.query.filter_by(section=sec).count() for sec in SECTIONS}
    return render_template("admin.html",
        total_students=total_students, total_notes=total_notes,
        total_books=total_books, total_quizzes=total_quizzes,
        total_assignments=total_assignments, recent_notes=recent_notes,
        recent_assignments=recent_assignments, section_counts=section_counts, sections=SECTIONS)

@app.route("/students")
def students():
    if "admin" not in session:
        return redirect("/login")
    search = request.args.get("search","")
    section = request.args.get("section","")
    data = Student.query
    if search:
        data = data.filter((Student.student_name.contains(search))|(Student.admission_no.contains(search)))
    if section:
        data = data.filter_by(section=section)
    return render_template("students.html", students=data.all(), total=Student.query.count(), sections=SECTIONS, search=search, section=section)

@app.route("/delete/<int:id>")
def delete_student(id):
    if "admin" not in session: return redirect("/login")
    s = db.session.get(Student, id)
    if s: db.session.delete(s); db.session.commit()
    return redirect("/students")

@app.route("/upload_students", methods=["GET","POST"])
def upload_students():
    if "admin" not in session: return redirect("/login")
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
                db.session.add(Student(
                    admission_no=str(data["Admission_No"]), student_name=data["Student_Name"],
                    dob=str(data["DOB"]), student_class=str(data["Class"]),
                    section=str(data["Section"]), roll_no=str(data["Roll_No"]),
                    parent_mobile=str(data["Parent_Mobile"])))
                added += 1
        db.session.commit()
        return render_template("success.html", total=added, filename=file.filename)
    return render_template("upload.html")

# ─────────────────────────────────────────────
# TEACHER
# ─────────────────────────────────────────────

@app.route("/teacher")
def teacher():
    if "teacher" not in session and "admin" not in session:
        return redirect("/teacher_login")
    notes = Note.query.filter_by(file_type="notes").order_by(Note.id.desc()).all()
    books = Note.query.filter_by(file_type="book").order_by(Note.id.desc()).all()
    assignments = Assignment.query.order_by(Assignment.id.desc()).all()
    quizzes = Quiz.query.order_by(Quiz.id.desc()).all()
    games = Game.query.order_by(Game.id.desc()).all()
    videos = ChapterVideo.query.order_by(ChapterVideo.id.desc()).all()
    configs = {c.chapter_no: c for c in ChapterConfig.query.all()}
    feedbacks = Feedback.query.order_by(Feedback.id.desc()).limit(20).all()
    return render_template("teacher.html",
        notes=notes, books=books, assignments=assignments,
        quizzes=quizzes, games=games, videos=videos,
        configs=configs, chapters=CHAPTERS, sections=SECTIONS,
        months=MONTHS, teacher_name=session.get("teacher_name","Teacher"),
        total_students=Student.query.count(), feedbacks=feedbacks)

# Upload notes/books
@app.route("/upload_notes", methods=["POST"])
def upload_notes():
    if "teacher" not in session and "admin" not in session: return redirect("/teacher_login")
    file = request.files.get("file")
    chapter = request.form["chapter"]
    chapter_no = int(request.form.get("chapter_no", 0))
    section = request.form["section"]
    description = request.form.get("description","")
    file_type = request.form.get("file_type","notes")
    if file and file.filename:
        os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
        safe_name = safe_filename(file.filename)
        file.save(os.path.join(app.config["UPLOAD_FOLDER"], safe_name))
        db.session.add(Note(filename=safe_name, original_name=file.filename, chapter=chapter,
            chapter_no=chapter_no, section=section, description=description,
            file_type=file_type, uploaded_at=now_str()))
        db.session.commit()
        # update config
        cfg = ChapterConfig.query.filter_by(chapter_no=chapter_no).first()
        if cfg:
            if file_type == "notes": cfg.has_notes = True
            if file_type == "book": cfg.has_book = True
            db.session.commit()
    return redirect("/teacher")

@app.route("/delete_note/<int:id>")
def delete_note(id):
    if "teacher" not in session and "admin" not in session: return redirect("/teacher_login")
    note = db.session.get(Note, id)
    if note:
        fpath = os.path.join(app.config["UPLOAD_FOLDER"], note.filename)
        if os.path.exists(fpath): os.remove(fpath)
        db.session.delete(note); db.session.commit()
    return redirect("/teacher")

# Video
@app.route("/add_video", methods=["POST"])
def add_video():
    if "teacher" not in session and "admin" not in session: return redirect("/teacher_login")
    chapter_no = int(request.form["chapter_no"])
    chapter = request.form["chapter"]
    section = request.form["section"]
    youtube_url = request.form["youtube_url"]
    title = request.form.get("title","Chapter Video")
    existing = ChapterVideo.query.filter_by(chapter_no=chapter_no, section=section).first()
    if existing:
        existing.youtube_url = youtube_url; existing.title = title; existing.added_at = now_str()
    else:
        db.session.add(ChapterVideo(chapter_no=chapter_no, chapter=chapter, section=section,
            youtube_url=youtube_url, title=title, added_at=now_str()))
    cfg = ChapterConfig.query.filter_by(chapter_no=chapter_no).first()
    if cfg: cfg.has_video = True
    db.session.commit()
    return redirect("/teacher")

@app.route("/delete_video/<int:id>")
def delete_video(id):
    if "teacher" not in session and "admin" not in session: return redirect("/teacher_login")
    v = db.session.get(ChapterVideo, id)
    if v: db.session.delete(v); db.session.commit()
    return redirect("/teacher")

# Chapter config (activate/deactivate)
@app.route("/save_chapter_config", methods=["POST"])
def save_chapter_config():
    if "teacher" not in session and "admin" not in session: return redirect("/teacher_login")
    chapter_no = int(request.form["chapter_no"])
    cfg = ChapterConfig.query.filter_by(chapter_no=chapter_no).first()
    if not cfg:
        cfg = ChapterConfig(chapter_no=chapter_no)
        db.session.add(cfg)
    cfg.is_active = "is_active" in request.form
    cfg.active_month = request.form.get("active_month","")
    cfg.has_video = "has_video" in request.form
    cfg.has_notes = "has_notes" in request.form
    cfg.has_book = "has_book" in request.form
    cfg.has_quiz = "has_quiz" in request.form
    cfg.has_game = "has_game" in request.form
    cfg.has_feedback = "has_feedback" in request.form
    db.session.commit()
    return redirect("/teacher")

# Assignment
@app.route("/create_assignment", methods=["POST"])
def create_assignment():
    if "teacher" not in session and "admin" not in session: return redirect("/teacher_login")
    filename = ""
    file = request.files.get("file")
    if file and file.filename:
        os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
        safe_name = safe_filename(file.filename)
        file.save(os.path.join(app.config["UPLOAD_FOLDER"], safe_name))
        filename = safe_name
    db.session.add(Assignment(
        title=request.form["title"], chapter=request.form["chapter"],
        chapter_no=int(request.form.get("chapter_no",0)), section=request.form["section"],
        description=request.form.get("description",""), due_date=request.form.get("due_date",""),
        filename=filename, created_at=now_str()))
    db.session.commit()
    return redirect("/teacher")

@app.route("/delete_assignment/<int:id>")
def delete_assignment(id):
    if "teacher" not in session and "admin" not in session: return redirect("/teacher_login")
    a = db.session.get(Assignment, id)
    if a: db.session.delete(a); db.session.commit()
    return redirect("/teacher")

# Quiz
@app.route("/create_quiz", methods=["POST"])
def create_quiz():
    if "teacher" not in session and "admin" not in session: return redirect("/teacher_login")
    chapter_no = int(request.form.get("chapter_no",0))
    quiz_type = request.form.get("quiz_type","mcq")
    quiz = Quiz(title=request.form["title"], chapter=request.form["chapter"],
        chapter_no=chapter_no, section=request.form["section"],
        description=request.form.get("description",""), quiz_type=quiz_type,
        external_link=request.form.get("external_link",""), created_at=now_str())
    db.session.add(quiz); db.session.flush()
    if quiz_type == "mcq":
        questions = request.form.getlist("question_text")
        for i, qtext in enumerate(questions):
            if qtext.strip():
                db.session.add(QuizQuestion(
                    quiz_id=quiz.id, question_text=qtext.strip(),
                    option_a=request.form.getlist("option_a")[i],
                    option_b=request.form.getlist("option_b")[i],
                    option_c=request.form.getlist("option_c")[i],
                    option_d=request.form.getlist("option_d")[i],
                    correct_option=request.form.getlist("correct_option")[i]))
    cfg = ChapterConfig.query.filter_by(chapter_no=chapter_no).first()
    if cfg: cfg.has_quiz = True
    db.session.commit()
    return redirect("/teacher")

@app.route("/delete_quiz/<int:id>")
def delete_quiz(id):
    if "teacher" not in session and "admin" not in session: return redirect("/teacher_login")
    q = db.session.get(Quiz, id)
    if q: db.session.delete(q); db.session.commit()
    return redirect("/teacher")

@app.route("/quiz_results/<int:quiz_id>")
def quiz_results(quiz_id):
    if "teacher" not in session and "admin" not in session: return redirect("/teacher_login")
    quiz = db.session.get(Quiz, quiz_id)
    attempts = QuizAttempt.query.filter_by(quiz_id=quiz_id).order_by(QuizAttempt.score.desc()).all()
    return render_template("quiz_results.html", quiz=quiz, attempts=attempts)

# Game
@app.route("/create_game", methods=["POST"])
def create_game():
    if "teacher" not in session and "admin" not in session: return redirect("/teacher_login")
    chapter_no = int(request.form.get("chapter_no",0))
    game = Game(title=request.form["title"], chapter=request.form["chapter"],
        chapter_no=chapter_no, section=request.form["section"], created_at=now_str())
    db.session.add(game); db.session.flush()
    terms = request.form.getlist("term")
    definitions = request.form.getlist("definition")
    for t, d in zip(terms, definitions):
        if t.strip() and d.strip():
            db.session.add(GamePair(game_id=game.id, term=t.strip(), definition=d.strip()))
    cfg = ChapterConfig.query.filter_by(chapter_no=chapter_no).first()
    if cfg: cfg.has_game = True
    db.session.commit()
    return redirect("/teacher")

@app.route("/delete_game/<int:id>")
def delete_game(id):
    if "teacher" not in session and "admin" not in session: return redirect("/teacher_login")
    g = db.session.get(Game, id)
    if g: db.session.delete(g); db.session.commit()
    return redirect("/teacher")

@app.route("/game_leaderboard/<int:game_id>")
def game_leaderboard(game_id):
    if "teacher" not in session and "admin" not in session: return redirect("/teacher_login")
    game = db.session.get(Game, game_id)
    scores = GameScore.query.filter_by(game_id=game_id).order_by(GameScore.score.desc(), GameScore.time_seconds).all()
    return render_template("game_leaderboard.html", game=game, scores=scores)

# Feedbacks view
@app.route("/view_feedbacks")
def view_feedbacks():
    if "teacher" not in session and "admin" not in session: return redirect("/teacher_login")
    chapter_no = request.args.get("chapter_no", 0, type=int)
    query = Feedback.query
    if chapter_no:
        query = query.filter_by(chapter_no=chapter_no)
    feedbacks = query.order_by(Feedback.id.desc()).all()
    return render_template("view_feedbacks.html", feedbacks=feedbacks, chapters=CHAPTERS, selected_chapter=chapter_no)

# ─────────────────────────────────────────────
# STUDENT PORTAL
# ─────────────────────────────────────────────

@app.route("/student")
def student_portal():
    if "student_id" not in session:
        return render_template("student.html", student=None, chapters=CHAPTERS,
            notes=[], books=[], quizzes=[], assignments=[], games=[], videos=[],
            progress_map={}, submitted_ids={}, attempted_quiz_ids={},
            game_scores={}, feedbacks_done=set(), chapter_configs={},
            chapter_pcts={}, completed=0, in_progress=0, total_chapters=len(CHAPTERS))

    student = db.session.get(Student, session["student_id"])
    if not student:
        session.pop("student_id", None)
        return redirect("/student")

    sec = student.section
    configs = {c.chapter_no: c for c in ChapterConfig.query.all()}

    def mat(ftype):
        return Note.query.filter(Note.file_type==ftype,
            (Note.section==sec)|(Note.section=="All Sections")).order_by(Note.chapter_no, Note.id).all()

    notes = mat("notes")
    books = mat("book")
    quizzes = Quiz.query.filter((Quiz.section==sec)|(Quiz.section=="All Sections")).order_by(Quiz.chapter_no).all()
    assignments = Assignment.query.filter((Assignment.section==sec)|(Assignment.section=="")|(Assignment.section=="All Sections")).order_by(Assignment.chapter_no).all()
    games = Game.query.filter((Game.section==sec)|(Game.section=="All Sections")).order_by(Game.chapter_no).all()
    videos = ChapterVideo.query.filter((ChapterVideo.section==sec)|(ChapterVideo.section=="All Sections")).order_by(ChapterVideo.chapter_no).all()

    progs = {p.chapter_no: p for p in ChapterProgress.query.filter_by(student_id=student.id).all()}
    submitted_ids = {s.assignment_id: s.status for s in AssignmentSubmission.query.filter_by(student_id=student.id).all()}
    attempted_quiz_ids = {a.quiz_id: a for a in QuizAttempt.query.filter_by(student_id=student.id).all()}
    game_scores = {s.game_id: s for s in GameScore.query.filter_by(student_id=student.id).all()}
    feedbacks_done = {f.chapter_no for f in Feedback.query.filter_by(student_id=student.id).all()}

    chapter_pcts = {}
    for ch in CHAPTERS:
        cno = ch["no"]
        cfg = configs.get(cno)
        prog = progs.get(cno)
        chapter_pcts[cno] = get_chapter_progress_pct(prog, cfg)

    completed = sum(1 for p in chapter_pcts.values() if p == 100)
    in_progress = sum(1 for p in chapter_pcts.values() if 0 < p < 100)

    return render_template("student.html",
        student=student, chapters=CHAPTERS, notes=notes, books=books,
        quizzes=quizzes, assignments=assignments, games=games, videos=videos,
        progress_map=progs, submitted_ids=submitted_ids,
        attempted_quiz_ids=attempted_quiz_ids, game_scores=game_scores,
        feedbacks_done=feedbacks_done, chapter_configs=configs,
        chapter_pcts=chapter_pcts, completed=completed,
        in_progress=in_progress, total_chapters=len(CHAPTERS))

@app.route("/student_login", methods=["POST"])
def student_login():
    admission_no = request.form["admission_no"]
    dob = request.form.get("dob","").strip()
    student = Student.query.filter_by(admission_no=admission_no).first()
    if student and (not dob or student.dob == dob):
        session["student_id"] = student.id
        return redirect("/student")
    return render_template("student.html", student=None, error="Admission No not found.",
        chapters=CHAPTERS, notes=[], books=[], quizzes=[], assignments=[], games=[], videos=[],
        progress_map={}, submitted_ids={}, attempted_quiz_ids={}, game_scores={},
        feedbacks_done=set(), chapter_configs={}, chapter_pcts={},
        completed=0, in_progress=0, total_chapters=len(CHAPTERS))

@app.route("/student_logout")
def student_logout():
    session.pop("student_id", None)
    return redirect("/student")

# Mark video watched
@app.route("/mark_video/<int:chapter_no>")
def mark_video(chapter_no):
    if "student_id" not in session: return redirect("/student")
    prog = ChapterProgress.query.filter_by(student_id=session["student_id"], chapter_no=chapter_no).first()
    if not prog:
        prog = ChapterProgress(student_id=session["student_id"], chapter_no=chapter_no)
        db.session.add(prog)
    prog.video_done = True; prog.updated_at = now_str()
    db.session.commit()
    return redirect("/student")

# Mark notes read
@app.route("/mark_notes/<int:chapter_no>")
def mark_notes(chapter_no):
    if "student_id" not in session: return redirect("/student")
    prog = ChapterProgress.query.filter_by(student_id=session["student_id"], chapter_no=chapter_no).first()
    if not prog:
        prog = ChapterProgress(student_id=session["student_id"], chapter_no=chapter_no)
        db.session.add(prog)
    prog.notes_done = True; prog.updated_at = now_str()
    db.session.commit()
    return redirect("/student")

# Mark book read
@app.route("/mark_book/<int:chapter_no>")
def mark_book(chapter_no):
    if "student_id" not in session: return redirect("/student")
    prog = ChapterProgress.query.filter_by(student_id=session["student_id"], chapter_no=chapter_no).first()
    if not prog:
        prog = ChapterProgress(student_id=session["student_id"], chapter_no=chapter_no)
        db.session.add(prog)
    prog.book_done = True; prog.updated_at = now_str()
    db.session.commit()
    return redirect("/student")

# Submit assignment
@app.route("/submit_assignment/<int:assignment_id>")
def submit_assignment(assignment_id):
    if "student_id" not in session: return redirect("/student")
    if not AssignmentSubmission.query.filter_by(student_id=session["student_id"], assignment_id=assignment_id).first():
        db.session.add(AssignmentSubmission(student_id=session["student_id"],
            assignment_id=assignment_id, status="submitted", submitted_at=now_str()))
        db.session.commit()
    return redirect("/student")

# Quiz attempt
@app.route("/attempt_quiz/<int:quiz_id>", methods=["GET","POST"])
def attempt_quiz(quiz_id):
    if "student_id" not in session: return redirect("/student")
    quiz = db.session.get(Quiz, quiz_id)
    if not quiz: return redirect("/student")
    if quiz.quiz_type == "link":
        return redirect(quiz.external_link)
    existing = QuizAttempt.query.filter_by(student_id=session["student_id"], quiz_id=quiz_id).first()
    if existing: return redirect("/student")
    if request.method == "POST":
        score = sum(1 for q in quiz.questions if request.form.get(f"q_{q.id}","").upper() == q.correct_option.upper())
        total = len(quiz.questions)
        db.session.add(QuizAttempt(student_id=session["student_id"], quiz_id=quiz_id,
            score=score, total=total, attempted_at=now_str()))
        # update progress
        prog = ChapterProgress.query.filter_by(student_id=session["student_id"], chapter_no=quiz.chapter_no).first()
        if not prog:
            prog = ChapterProgress(student_id=session["student_id"], chapter_no=quiz.chapter_no)
            db.session.add(prog)
        prog.quiz_done = True; prog.updated_at = now_str()
        db.session.commit()
        return render_template("quiz_score.html", score=score, total=total, quiz=quiz)
    return render_template("attempt_quiz.html", quiz=quiz)

# Game play
@app.route("/play_game/<int:game_id>")
def play_game(game_id):
    if "student_id" not in session: return redirect("/student")
    game = db.session.get(Game, game_id)
    if not game: return redirect("/student")
    existing = GameScore.query.filter_by(student_id=session["student_id"], game_id=game_id).first()
    leaderboard = GameScore.query.filter_by(game_id=game_id).order_by(GameScore.score.desc(), GameScore.time_seconds).limit(10).all()
    return render_template("play_game.html", game=game, existing=existing, leaderboard=leaderboard)

@app.route("/save_game_score", methods=["POST"])
def save_game_score():
    if "student_id" not in session: return jsonify({"ok": False})
    data = request.get_json()
    game_id = data.get("game_id")
    score = data.get("score", 0)
    time_seconds = data.get("time_seconds", 999)
    game = db.session.get(Game, game_id)
    if not game: return jsonify({"ok": False})
    existing = GameScore.query.filter_by(student_id=session["student_id"], game_id=game_id).first()
    if existing:
        if score > existing.score or (score == existing.score and time_seconds < existing.time_seconds):
            existing.score = score; existing.time_seconds = time_seconds; existing.played_at = now_str()
    else:
        db.session.add(GameScore(student_id=session["student_id"], game_id=game_id,
            score=score, time_seconds=time_seconds, played_at=now_str()))
    # update progress
    prog = ChapterProgress.query.filter_by(student_id=session["student_id"], chapter_no=game.chapter_no).first()
    if not prog:
        prog = ChapterProgress(student_id=session["student_id"], chapter_no=game.chapter_no)
        db.session.add(prog)
    prog.game_done = True; prog.updated_at = now_str()
    db.session.commit()
    leaderboard = []
    for s in GameScore.query.filter_by(game_id=game_id).order_by(GameScore.score.desc(), GameScore.time_seconds).limit(10).all():
        leaderboard.append({"name": s.student.student_name, "score": s.score, "time": s.time_seconds})
    return jsonify({"ok": True, "leaderboard": leaderboard})

# Feedback
@app.route("/submit_feedback", methods=["POST"])
def submit_feedback():
    if "student_id" not in session: return redirect("/student")
    chapter_no = int(request.form["chapter_no"])
    existing = Feedback.query.filter_by(student_id=session["student_id"], chapter_no=chapter_no).first()
    if existing: return redirect("/student")
    db.session.add(Feedback(student_id=session["student_id"], chapter_no=chapter_no,
        rating=int(request.form.get("rating",5)), comment=request.form.get("comment",""),
        submitted_at=now_str()))
    prog = ChapterProgress.query.filter_by(student_id=session["student_id"], chapter_no=chapter_no).first()
    if not prog:
        prog = ChapterProgress(student_id=session["student_id"], chapter_no=chapter_no)
        db.session.add(prog)
    prog.feedback_done = True; prog.updated_at = now_str()
    db.session.commit()
    return redirect("/student")

@app.route("/uploads/<filename>")
def serve_upload(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)

# ─────────────────────────────────────────────
# DB INIT
# ─────────────────────────────────────────────
with app.app_context():
    os.makedirs("uploads", exist_ok=True)
    db.create_all()
    if Admin.query.count() == 0:
        db.session.add(Admin(username="admin", password="1234")); db.session.commit()
    if Teacher.query.count() == 0:
        db.session.add(Teacher(username="teacher", password="1234", name="Computer Teacher")); db.session.commit()
    for ch in CHAPTERS:
        if not ChapterConfig.query.filter_by(chapter_no=ch["no"]).first():
            db.session.add(ChapterConfig(chapter_no=ch["no"]))
    db.session.commit()

if __name__ == "__main__":
    app.run(debug=True)
