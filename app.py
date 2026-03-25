import datetime
import os
import re
from functools import wraps

from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

from extract_engine import compute_net_backlogs, detect_and_parse, parse_class_report

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
MIN_PASSWORD_LENGTH = 8
USERS_PER_PAGE = 10

app = Flask(__name__)
app.secret_key = os.environ.get("AURAS_SECRET_KEY", "change-this-secret-key")
app.config["UPLOAD_FOLDER"] = os.path.join(BASE_DIR, "uploads")
app.config["DOWNLOAD_FOLDER"] = os.path.join(BASE_DIR, "downloads")
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["DOWNLOAD_FOLDER"], exist_ok=True)

db_url = os.environ.get("DATABASE_URL", "sqlite:///" + os.path.join(BASE_DIR, "results.db"))
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)
app.config["SQLALCHEMY_DATABASE_URI"] = db_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)


class User(db.Model):
    __tablename__ = "users"

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default="faculty")
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.datetime.utcnow)


class Student(db.Model):
    __tablename__ = "students"

    register_no = db.Column(db.String(20), primary_key=True)
    name = db.Column(db.String(100))
    batch_year = db.Column(db.String(4))
    department = db.Column(db.String(10))


class SemesterResult(db.Model):
    __tablename__ = "semester_results"

    id = db.Column(db.Integer, primary_key=True)
    register_no = db.Column(db.String(20), db.ForeignKey("students.register_no"))
    semester = db.Column(db.String(5))
    sgpa = db.Column(db.Float)
    cgpa = db.Column(db.Float, default=0.0)
    backlogs = db.Column(db.Integer)


class CourseGrade(db.Model):
    __tablename__ = "course_grades"

    id = db.Column(db.Integer, primary_key=True)
    register_no = db.Column(db.String(20), db.ForeignKey("students.register_no"))
    semester = db.Column(db.String(5))
    course_code = db.Column(db.String(20))
    course_name = db.Column(db.String(150))
    grade = db.Column(db.String(5))


class Course(db.Model):
    __tablename__ = "courses"

    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(20))
    name = db.Column(db.String(150))
    semester = db.Column(db.String(5))
    batch_year = db.Column(db.String(4))
    department = db.Column(db.String(10))


def normalize_username(value):
    return (value or "").strip().lower()


def is_valid_role(value):
    return value in {"admin", "faculty"}


def password_error(password):
    if len(password or "") < MIN_PASSWORD_LENGTH:
        return f"Password must be at least {MIN_PASSWORD_LENGTH} characters long."
    return None


def admin_count():
    return User.query.filter_by(role="admin").count()


def is_last_admin(user):
    return user.role == "admin" and admin_count() == 1


def bootstrap_admin():
    if User.query.count() > 0:
        return

    username = normalize_username(os.environ.get("AURAS_BOOTSTRAP_ADMIN_USERNAME", "admin")) or "admin"
    password = os.environ.get("AURAS_BOOTSTRAP_ADMIN_PASSWORD", "admin12345")
    if password_error(password):
        password = "admin12345"

    db.session.add(
        User(
            username=username,
            password_hash=generate_password_hash(password),
            role="admin",
        )
    )
    db.session.commit()

    if username == "admin" and password == "admin12345":
        app.logger.warning(
            "Bootstrapped default admin credentials. Set AURAS_BOOTSTRAP_ADMIN_USERNAME "
            "and AURAS_BOOTSTRAP_ADMIN_PASSWORD for production use."
        )


with app.app_context():
    auto_init_db = os.environ.get("AURAS_AUTO_INIT_DB", "true").strip().lower() in {"1", "true", "yes", "on"}
    if auto_init_db:
        db.create_all()
        bootstrap_admin()


SEMS = [f"S{i}" for i in range(1, 9)]
REG_NO_CLASS_RE = re.compile(r"^[A-Z]+(?P<yy>\d{2})(?P<dept>[A-Z]{2,4})\d{3}$")


def semester_sort_key(value):
    """Sort semesters safely, pushing malformed values to the end."""
    text = (value or "").strip().upper()
    match = re.match(r"^S(\d+)$", text)
    if match:
        return (0, int(match.group(1)), text)
    return (1, 999, text)


def current_year():
    return session.get("year")


def current_dept():
    return session.get("dept")


def current_username():
    return session.get("username", "")


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if "user_id" not in session:
            flash("Please log in to continue.", "warning")
            return redirect(url_for("login"))
        return fn(*args, **kwargs)

    return wrapper


def require_class(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if "user_id" not in session:
            flash("Please log in to continue.", "warning")
            return redirect(url_for("login"))
        if "year" not in session or "dept" not in session:
            flash("Select a class to continue.", "warning")
            return redirect(url_for("select_class"))
        return fn(*args, **kwargs)

    return wrapper


def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if "user_id" not in session:
            flash("Please log in to continue.", "warning")
            return redirect(url_for("login"))
        if session.get("role") != "admin":
            flash("Admin access is required for that page.", "danger")
            return redirect(url_for("index") if "year" in session and "dept" in session else url_for("select_class"))
        return fn(*args, **kwargs)

    return wrapper


def class_name():
    year = current_year()
    dept = current_dept()
    if not year or not dept:
        return ""
    return f"{dept}{str(year)[-2:]}"


def matches_selected_class(register_no, parsed_dept=None):
    """Return True when a register number belongs to the active class (year+dept)."""
    match = REG_NO_CLASS_RE.match((register_no or "").upper())
    if not match:
        return False

    selected_year = str(current_year() or "")[-2:]
    selected_dept = (current_dept() or "").upper()

    if not selected_year or not selected_dept:
        return False

    reg_year = match.group("yy")
    reg_dept = match.group("dept")

    if reg_year != selected_year:
        return False

    if parsed_dept:
        return parsed_dept.upper() == selected_dept and reg_dept == selected_dept

    return reg_dept == selected_dept


def build_sgpa_matrix():
    results = (
        db.session.query(SemesterResult)
        .join(Student)
        .filter(Student.batch_year == current_year(), Student.department == current_dept())
        .all()
    )

    reg_nos = sorted({r.register_no for r in results})
    sems_used = sorted({r.semester for r in results if r.semester}, key=semester_sort_key)

    sgpa_map = {}
    for result in results:
        sgpa_map.setdefault(result.register_no, {})[result.semester] = result.sgpa

    students = {
        student.register_no: student.name
        for student in Student.query.filter_by(batch_year=current_year(), department=current_dept()).all()
    }

    matrix = []
    for register_no in reg_nos:
        row = {"register_no": register_no, "name": students.get(register_no, ""), "sems": {}}
        for sem in sems_used:
            row["sems"][sem] = sgpa_map.get(register_no, {}).get(sem, "-")
        matrix.append(row)

    return matrix, sems_used


def build_backlog_matrix():
    grades = (
        db.session.query(CourseGrade)
        .join(Student)
        .filter(Student.batch_year == current_year(), Student.department == current_dept())
        .all()
    )

    grade_rows = [
        {
            "register_no": grade.register_no,
            "semester": grade.semester,
            "course_code": grade.course_code,
            "grade": grade.grade,
        }
        for grade in grades
    ]

    net = compute_net_backlogs(grade_rows)
    sems_used = sorted({grade.semester for grade in grades if grade.semester}, key=semester_sort_key)

    students = {
        student.register_no: student.name
        for student in Student.query.filter_by(batch_year=current_year(), department=current_dept()).all()
    }

    matrix = []
    for register_no in sorted(net.keys()):
        row = {"register_no": register_no, "name": students.get(register_no, ""), "sems": {}}
        for sem in sems_used:
            row["sems"][sem] = net.get(register_no, {}).get(sem, "-")
        matrix.append(row)

    return matrix, sems_used


def get_courses_for_class():
    courses = Course.query.filter_by(batch_year=current_year(), department=current_dept()).all()
    return sorted(courses, key=lambda course: (semester_sort_key(course.semester), course.code or ""))


def get_user_or_none(user_id):
    return db.session.get(User, user_id)


def parse_positive_int(value, default=1):
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        return default
    return parsed if parsed > 0 else default


def manage_users_redirect():
    search_query = (request.form.get("q") or "").strip()
    page = parse_positive_int(request.form.get("page"), default=1)
    return redirect(url_for("manage_users", q=search_query or None, page=page))


@app.route("/login", methods=["GET", "POST"])
def login():
    if "user_id" in session:
        return redirect(url_for("select_class"))

    if request.method == "POST":
        username = normalize_username(request.form.get("username"))
        password = request.form.get("password", "")
        user = User.query.filter_by(username=username).first()

        if user and check_password_hash(user.password_hash, password):
            session.clear()
            session["user_id"] = user.id
            session["username"] = user.username
            session["role"] = user.role
            flash(f"Welcome back, {user.username}.", "success")
            return redirect(url_for("select_class"))

        flash("Invalid username or password.", "danger")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("You have been logged out.", "info")
    return redirect(url_for("login"))


@app.route("/select_class", methods=["GET", "POST"])
@login_required
def select_class():
    if request.method == "POST":
        session["year"] = request.form["year"]
        session["dept"] = request.form["dept"]
        return redirect(url_for("index"))

    years = list(range(datetime.date.today().year, 2019, -1))
    return render_template(
        "selection.html",
        years=years,
        depts=["CS", "EC", "ME", "CE", "EEE"],
        role=session.get("role"),
        username=current_username(),
    )


@app.route("/")
@require_class
def index():
    sgpa_matrix, sems = build_sgpa_matrix()
    backlog_matrix, _ = build_backlog_matrix()
    courses = get_courses_for_class()
    return render_template(
        "index.html",
        class_name=class_name(),
        role=session.get("role"),
        username=current_username(),
        sgpa_matrix=sgpa_matrix,
        backlog_matrix=backlog_matrix,
        sems=sems,
        courses=courses,
    )


@app.route("/users")
@admin_required
def manage_users():
    search_query = (request.args.get("q") or "").strip()
    page = parse_positive_int(request.args.get("page"), default=1)

    base_query = User.query
    if search_query:
        base_query = base_query.filter(User.username.ilike(f"%{search_query}%"))

    total_users = User.query.count()
    filtered_count = base_query.count()
    total_pages = max(1, (filtered_count + USERS_PER_PAGE - 1) // USERS_PER_PAGE)
    page = min(page, total_pages)

    users = (
        base_query.order_by(User.username.asc())
        .offset((page - 1) * USERS_PER_PAGE)
        .limit(USERS_PER_PAGE)
        .all()
    )

    start_index = (page - 1) * USERS_PER_PAGE + 1 if filtered_count else 0
    end_index = min(page * USERS_PER_PAGE, filtered_count)

    return render_template(
        "users.html",
        users=users,
        username=current_username(),
        role=session.get("role"),
        search_query=search_query,
        page=page,
        total_pages=total_pages,
        total_users=total_users,
        filtered_count=filtered_count,
        start_index=start_index,
        end_index=end_index,
    )


@app.route("/users/create", methods=["POST"])
@admin_required
def create_user():
    username = normalize_username(request.form.get("username"))
    password = request.form.get("password", "")
    role = request.form.get("role", "faculty")

    if not username:
        flash("Username is required.", "danger")
        return manage_users_redirect()
    if not is_valid_role(role):
        flash("Invalid role selected.", "danger")
        return manage_users_redirect()

    error = password_error(password)
    if error:
        flash(error, "danger")
        return manage_users_redirect()

    if User.query.filter_by(username=username).first():
        flash("That username already exists.", "danger")
        return manage_users_redirect()

    db.session.add(User(username=username, password_hash=generate_password_hash(password), role=role))
    db.session.commit()
    flash(f"User '{username}' created successfully.", "success")
    return manage_users_redirect()


@app.route("/users/<int:user_id>/update", methods=["POST"])
@admin_required
def update_user(user_id):
    user = get_user_or_none(user_id)
    if not user:
        flash("User not found.", "danger")
        return manage_users_redirect()

    username = normalize_username(request.form.get("username"))
    role = request.form.get("role", user.role)

    if not username:
        flash("Username is required.", "danger")
        return manage_users_redirect()
    if not is_valid_role(role):
        flash("Invalid role selected.", "danger")
        return manage_users_redirect()

    duplicate = User.query.filter(User.username == username, User.id != user.id).first()
    if duplicate:
        flash("That username already exists.", "danger")
        return manage_users_redirect()

    if user.role == "admin" and role != "admin" and is_last_admin(user):
        flash("The last remaining admin cannot be changed to faculty.", "danger")
        return manage_users_redirect()

    user.username = username
    user.role = role
    db.session.commit()

    if session.get("user_id") == user.id:
        session["username"] = user.username
        session["role"] = user.role

    flash(f"User '{username}' updated successfully.", "success")
    return manage_users_redirect()


@app.route("/users/<int:user_id>/reset_password", methods=["POST"])
@admin_required
def reset_user_password(user_id):
    user = get_user_or_none(user_id)
    if not user:
        flash("User not found.", "danger")
        return manage_users_redirect()

    password = request.form.get("password", "")
    error = password_error(password)
    if error:
        flash(error, "danger")
        return manage_users_redirect()

    user.password_hash = generate_password_hash(password)
    db.session.commit()
    flash(f"Password reset for '{user.username}'.", "success")
    return manage_users_redirect()


@app.route("/users/<int:user_id>/delete", methods=["POST"])
@admin_required
def delete_user(user_id):
    user = get_user_or_none(user_id)
    if not user:
        flash("User not found.", "danger")
        return manage_users_redirect()

    if session.get("user_id") == user.id:
        flash("You cannot delete the account currently logged in.", "danger")
        return manage_users_redirect()

    if is_last_admin(user):
        flash("The last remaining admin cannot be deleted.", "danger")
        return manage_users_redirect()

    username = user.username
    db.session.delete(user)
    db.session.commit()
    flash(f"User '{username}' deleted successfully.", "success")
    return manage_users_redirect()


@app.route("/upload", methods=["POST"])
@require_class
def upload_file():
    uploaded = 0
    errors = []
    files = request.files.getlist("pdf_file")

    if not files or all(not upload.filename for upload in files):
        flash("Please select at least one PDF file to upload.", "warning")
        return redirect(url_for("index"))

    for upload in files:
        if not upload or not upload.filename:
            continue
        if not upload.filename.lower().endswith(".pdf"):
            errors.append(f"{upload.filename}: Only PDF files are allowed.")
            continue

        path = os.path.join(app.config["UPLOAD_FOLDER"], secure_filename(upload.filename))
        upload.save(path)

        try:
            semester, courses_dict, students = detect_and_parse(path)
        except Exception as exc:
            errors.append(f"{upload.filename}: {exc}")
            continue

        if not semester or not students:
            errors.append(f"{upload.filename}: Could not extract data.")
            continue

        if semester_sort_key(semester)[0] != 0:
            errors.append(f"{upload.filename}: Invalid semester value '{semester}'.")
            continue

        try:
            for code, name in courses_dict.items():
                existing = Course.query.filter_by(
                    code=code,
                    semester=semester,
                    batch_year=current_year(),
                    department=current_dept(),
                ).first()
                if not existing:
                    db.session.add(
                        Course(
                            code=code,
                            name=name,
                            semester=semester,
                            batch_year=current_year(),
                            department=current_dept(),
                        )
                    )

            for student_data in students:
                register_no = student_data.get("register_no")
                if not register_no:
                    continue

                parsed_dept = student_data.get("_department")
                if not matches_selected_class(register_no, parsed_dept=parsed_dept):
                    continue

                student = db.session.get(Student, register_no)

                if not student:
                    db.session.add(
                        Student(
                            register_no=register_no,
                            name=student_data.get("name", ""),
                            batch_year=current_year(),
                            department=current_dept(),
                        )
                    )
                elif not student.name and student_data.get("name"):
                    student.name = student_data["name"]

                SemesterResult.query.filter_by(register_no=register_no, semester=semester).delete()
                CourseGrade.query.filter_by(register_no=register_no, semester=semester).delete()

                db.session.add(
                    SemesterResult(
                        register_no=register_no,
                        semester=semester,
                        sgpa=student_data.get("sgpa", 0.0),
                        cgpa=student_data.get("cgpa", 0.0),
                        backlogs=student_data.get("backlogs", 0),
                    )
                )

                for code, grade in student_data.get("grades", {}).items():
                    db.session.add(
                        CourseGrade(
                            register_no=register_no,
                            semester=semester,
                            course_code=code,
                            course_name=courses_dict.get(code, ""),
                            grade=grade,
                        )
                    )

            db.session.commit()
            uploaded += 1
        except Exception as exc:
            db.session.rollback()
            errors.append(f"{upload.filename}: Failed to save parsed data ({exc}).")
            continue

    if uploaded:
        flash(f"Successfully imported {uploaded} file(s).", "success")
    if errors:
        flash("Upload issues: " + "; ".join(errors), "warning")

    return redirect(url_for("index"))


@app.route("/api/subject_analysis")
@require_class
def subject_analysis():
    code = request.args.get("code", "")
    sem = request.args.get("sem", "")

    query = (
        CourseGrade.query.join(Student)
        .filter(
            Student.batch_year == current_year(),
            Student.department == current_dept(),
            CourseGrade.course_code == code,
        )
    )
    if sem:
        query = query.filter(CourseGrade.semester == sem)

    rows = query.all()
    if not rows:
        return jsonify({"code": code, "name": "", "semester": sem, "distribution": [], "total": 0})

    from collections import Counter

    counts = Counter(row.grade for row in rows)
    grade_order = ["S", "A+", "A", "B+", "B", "C+", "C", "D", "P", "F", "FE", "LP", "I"]
    fail_grades = {"F", "FE", "LP", "I"}
    grade_points = {"S": 10, "A+": 9, "A": 8.5, "B+": 8, "B": 7, "C+": 6, "C": 5, "D": 4, "P": 3, "F": 0, "FE": 0, "LP": 0, "I": 0}
    quality_grades = {"S", "A+", "A", "B+"}

    distribution = []
    for grade in grade_order:
        if grade in counts:
            distribution.append({"grade": grade, "count": counts[grade], "fail": grade in fail_grades})
    for grade, count in counts.items():
        if grade not in grade_order:
            distribution.append({"grade": grade, "count": count, "fail": False})

    total = len(rows)
    grade_points_sum = sum(grade_points.get(row.grade, 0) for row in rows)
    avg_grade_point = round(grade_points_sum / total, 2) if total else 0

    topper_grade = ""
    topper_count = 0
    for grade in grade_order:
        if grade in counts:
            topper_grade = grade
            topper_count = counts[grade]
            break

    quality_count = sum(counts.get(grade, 0) for grade in quality_grades)
    quality_index = round(quality_count / total * 100, 1) if total else 0

    return jsonify(
        {
            "code": code,
            "name": rows[0].course_name,
            "semester": sem,
            "distribution": distribution,
            "total": total,
            "avg_grade_point": avg_grade_point,
            "topper_grade": topper_grade,
            "topper_count": topper_count,
            "quality_index": quality_index,
        }
    )


@app.route("/download_report")
@require_class
def download_report():
    # Lazy import to keep app cold-start lighter on Render.
    from report_generator import generate_excel_report

    sem = request.args.get("sem", "")
    code = request.args.get("code", "")

    results = (
        db.session.query(SemesterResult)
        .join(Student)
        .filter(Student.batch_year == current_year(), Student.department == current_dept())
        .all()
    )
    grades = (
        db.session.query(CourseGrade)
        .join(Student)
        .filter(Student.batch_year == current_year(), Student.department == current_dept())
        .all()
    )
    students = Student.query.filter_by(batch_year=current_year(), department=current_dept()).all()

    results_data = [
        {"register_no": result.register_no, "semester": result.semester, "sgpa": result.sgpa, "backlogs": result.backlogs}
        for result in results
    ]
    grades_data = [
        {
            "register_no": grade.register_no,
            "semester": grade.semester,
            "course_code": grade.course_code,
            "course_name": grade.course_name,
            "grade": grade.grade,
        }
        for grade in grades
    ]
    students_data = [{"register_no": student.register_no, "name": student.name} for student in students]

    net_backlogs = compute_net_backlogs(grades_data)

    if code and sem:
        filename = f"AURAS_Report_{class_name()}_{code}_{sem}.xlsx"
    else:
        filename = f"AURAS_Report_{class_name()}.xlsx"
    path = os.path.join(app.config["DOWNLOAD_FOLDER"], filename)

    generate_excel_report(results_data, grades_data, students_data, net_backlogs, path, selected_course_code=code if code else None)
    return send_file(path, as_attachment=True)


@app.route("/api/semester_all_subjects")
@require_class
def semester_all_subjects():
    """Return analysis data for ALL subjects in a given semester."""
    sem = request.args.get("sem", "")
    if not sem:
        return jsonify({"error": "sem parameter required"}), 400

    # Get all courses for this semester
    courses = Course.query.filter_by(
        semester=sem,
        batch_year=current_year(),
        department=current_dept(),
    ).all()

    if not courses:
        return jsonify({"semester": sem, "subjects": []})

    from collections import Counter

    fail_grades = {"F", "FE", "LP", "I", "Absent"}
    grade_points = {"S": 10, "A+": 9, "A": 8.5, "B+": 8, "B": 7,
                    "C+": 6, "C": 5, "D": 4, "P": 3,
                    "F": 0, "FE": 0, "LP": 0, "I": 0, "Absent": 0}
    quality_grades = {"S", "A+", "A", "B+"}
    grade_order = ["S", "A+", "A", "B+", "B", "C+", "C", "D", "P",
                   "F", "FE", "LP", "I", "Absent"]

    subjects = []
    for course in courses:
        rows = (
            CourseGrade.query.join(Student)
            .filter(
                Student.batch_year == current_year(),
                Student.department == current_dept(),
                CourseGrade.course_code == course.code,
                CourseGrade.semester == sem,
            )
            .all()
        )
        if not rows:
            continue

        counts = Counter(row.grade for row in rows)
        total = len(rows)

        distribution = []
        for grade in grade_order:
            if grade in counts:
                distribution.append({
                    "grade": grade,
                    "count": counts[grade],
                    "fail": grade in fail_grades,
                })

        pass_cnt = sum(v for k, v in counts.items() if k not in fail_grades)
        fail_cnt = total - pass_cnt
        pass_pct = round(pass_cnt / total * 100, 1) if total else 0

        gp_sum = sum(grade_points.get(row.grade, 0) for row in rows)
        avg_gp = round(gp_sum / total, 2) if total else 0

        quality_cnt = sum(counts.get(g, 0) for g in quality_grades)
        quality_idx = round(quality_cnt / total * 100, 1) if total else 0

        subjects.append({
            "code": course.code,
            "name": course.name,
            "total": total,
            "pass_count": pass_cnt,
            "fail_count": fail_cnt,
            "pass_pct": pass_pct,
            "avg_grade_point": avg_gp,
            "quality_index": quality_idx,
            "distribution": distribution,
        })

    # Sort by code
    subjects.sort(key=lambda s: s["code"])
    return jsonify({"semester": sem, "subjects": subjects})


@app.route("/download_semester_report")
@require_class
def download_semester_report():
    """Download Excel report with analysis for ALL subjects in a semester."""
    from report_generator import generate_all_subjects_excel_report

    sem = request.args.get("sem", "")
    if not sem:
        return redirect(url_for("index"))

    grades = (
        db.session.query(CourseGrade)
        .join(Student)
        .filter(
            Student.batch_year == current_year(),
            Student.department == current_dept(),
            CourseGrade.semester == sem,
        )
        .all()
    )
    all_grades = (
        db.session.query(CourseGrade)
        .join(Student)
        .filter(
            Student.batch_year == current_year(),
            Student.department == current_dept(),
        )
        .all()
    )
    results = (
        db.session.query(SemesterResult)
        .join(Student)
        .filter(
            Student.batch_year == current_year(),
            Student.department == current_dept(),
        )
        .all()
    )
    students_list = Student.query.filter_by(
        batch_year=current_year(), department=current_dept()
    ).all()
    courses = Course.query.filter_by(
        semester=sem,
        batch_year=current_year(),
        department=current_dept(),
    ).all()

    grades_data = [
        {
            "register_no": g.register_no,
            "semester": g.semester,
            "course_code": g.course_code,
            "course_name": g.course_name,
            "grade": g.grade,
        }
        for g in grades
    ]
    all_grades_data = [
        {
            "register_no": g.register_no,
            "semester": g.semester,
            "course_code": g.course_code,
            "course_name": g.course_name,
            "grade": g.grade,
        }
        for g in all_grades
    ]
    results_data = [
        {
            "register_no": r.register_no,
            "semester": r.semester,
            "sgpa": r.sgpa,
            "backlogs": r.backlogs,
        }
        for r in results
    ]
    students_data = [
        {"register_no": s.register_no, "name": s.name}
        for s in students_list
    ]
    courses_data = [
        {"code": c.code, "name": c.name}
        for c in courses
    ]

    filename = f"AURAS_AllSubjects_{class_name()}_{sem}.xlsx"
    path = os.path.join(app.config["DOWNLOAD_FOLDER"], filename)
    net_backlogs = compute_net_backlogs(all_grades_data)

    generate_all_subjects_excel_report(
        grades_data,
        students_data,
        courses_data,
        sem,
        path,
        results_data=results_data,
        net_backlogs=net_backlogs,
    )
    return send_file(path, as_attachment=True)


@app.route("/clear_database", methods=["POST"])
@require_class
@admin_required
def clear_database():
    scope = request.form.get("scope", "class")
    if scope == "all":
        CourseGrade.query.delete()
        SemesterResult.query.delete()
        Student.query.delete()
        Course.query.delete()
    else:
        register_nos = [
            student.register_no
            for student in Student.query.filter_by(batch_year=current_year(), department=current_dept()).all()
        ]
        if register_nos:
            CourseGrade.query.filter(CourseGrade.register_no.in_(register_nos)).delete(synchronize_session=False)
            SemesterResult.query.filter(SemesterResult.register_no.in_(register_nos)).delete(synchronize_session=False)
        Student.query.filter_by(batch_year=current_year(), department=current_dept()).delete()
        Course.query.filter_by(batch_year=current_year(), department=current_dept()).delete()

    db.session.commit()
    flash("Database records cleared successfully.", "success")
    return redirect(url_for("index"))


@app.route("/health")
def health_check():
    """A lightweight endpoint for Render's free tier health checks and pings."""
    return "OK", 200


if __name__ == "__main__":
    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    os.makedirs(app.config["DOWNLOAD_FOLDER"], exist_ok=True)
    debug_mode = os.environ.get("FLASK_DEBUG", "0").strip() in {"1", "true", "True"}
    app.run(debug=debug_mode)
