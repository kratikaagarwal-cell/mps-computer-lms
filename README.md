# MPS Computer LMS — Enhanced Edition
**MPS International School | Computer Department | Class 6**

---

## 🚀 Quick Setup

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the App
```bash
python app.py
```

### 3. Open in Browser
```
http://localhost:5000
```

---

## 🔐 Default Login Credentials

| Role    | Username  | Password |
|---------|-----------|----------|
| Admin   | `admin`   | `1234`   |
| Teacher | `teacher` | `1234`   |

**Student login:** Admission Number (no password required by default)

---

## 📌 What's New (vs Original)

### Teacher Portal
- ✅ Upload **Notes**, **Book PDFs**, **Quizzes**, and **Test Papers** separately
- ✅ Each file tagged to a specific **Chapter** + **Section**
- ✅ Attach files directly to **Assignments**
- ✅ Delete notes and assignments
- ✅ Tabbed UI: Upload / Notes / Books / Quizzes / Assignments / Chapter View
- ✅ Separate **Teacher Login** (username: `teacher`, password: `1234`)

### Student Portal
- ✅ Separate tabs for **Notes / Books / Quizzes / Assignments**
- ✅ **Chapter Progress Tracker** — mark each chapter as Pending / In Progress / Completed
- ✅ Visual progress bar showing % of syllabus completed
- ✅ **Mark Assignment as Submitted** button
- ✅ Chapter-wise counts showing how many materials are available per chapter

### Admin Dashboard
- ✅ Sidebar navigation
- ✅ Stats for Notes, Books, Quizzes, Assignments separately
- ✅ **Section-wise student bar chart**
- ✅ Recent uploads & assignments on dashboard

### Technical
- ✅ Timestamped filenames to prevent upload collisions
- ✅ Original filename stored separately for display
- ✅ Duplicate student protection on Excel import
- ✅ File deletion removes file from disk too

---

## 📚 Chapters (Class 6)
1. Internet Services
2. More on Excel
3. Formulas & Functions
4. Canva
5. Introduction to GIMP
6. Present Yourself Online
7. Introduction to Python
8. Introduction to Data Science
9. Robots & Sensors

## 🏫 Sections Supported
6A, 6B, 6C, 6D, 6E, 6F, 6G + All Sections

---

## 📁 Project Structure
```
MPS_LMS_Enhanced/
├── app.py                  ← Main Flask application
├── requirements.txt
├── uploads/                ← Uploaded files stored here
├── instance/
│   └── mps_lms.db          ← SQLite database (auto-created)
└── templates/
    ├── home.html
    ├── login.html           ← Admin login
    ├── teacher_login.html   ← Teacher login
    ├── admin.html
    ├── teacher.html
    ├── student.html
    ├── students.html
    ├── upload.html
    └── success.html
```

---

## 🗄️ Database Tables
- **Admin** — admin accounts
- **Teacher** — teacher accounts
- **Student** — student records
- **Note** — uploaded files (notes/books/quizzes/tests)
- **Assignment** — assignments created by teachers
- **ChapterProgress** — per-student chapter status tracking
- **AssignmentSubmission** — tracks which students submitted which assignments
