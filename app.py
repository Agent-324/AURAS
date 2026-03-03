import os
import datetime
from flask import Flask, render_template, request, redirect, url_for, session, send_file, jsonify
from flask_sqlalchemy import SQLAlchemy
from werkzeug.utils import secure_filename
from extract_engine import parse_class_report, compute_net_backlogs
from report_generator import generate_excel_report, generate_subject_excel_report

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

app = Flask(__name__)
app.secret_key = 'auras_secure_key'
app.config['UPLOAD_FOLDER']   = os.path.join(BASE_DIR, 'uploads')
app.config['DOWNLOAD_FOLDER'] = os.path.join(BASE_DIR, 'downloads')
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(BASE_DIR, 'results.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# ─── Models ───────────────────────────────────────────────────────────────────

class Student(db.Model):
    __tablename__ = 'students'
    register_no  = db.Column(db.String(20), primary_key=True)
    name         = db.Column(db.String(100))
    batch_year   = db.Column(db.String(4))
    department   = db.Column(db.String(10))

class SemesterResult(db.Model):
    __tablename__ = 'semester_results'
    id           = db.Column(db.Integer, primary_key=True)
    register_no  = db.Column(db.String(20), db.ForeignKey('students.register_no'))
    semester     = db.Column(db.String(5))
    sgpa         = db.Column(db.Float)
    cgpa         = db.Column(db.Float, default=0.0)
    backlogs     = db.Column(db.Integer)          # raw backlogs this semester

class CourseGrade(db.Model):
    __tablename__ = 'course_grades'
    id           = db.Column(db.Integer, primary_key=True)
    register_no  = db.Column(db.String(20), db.ForeignKey('students.register_no'))
    semester     = db.Column(db.String(5))
    course_code  = db.Column(db.String(20))
    course_name  = db.Column(db.String(150))
    grade        = db.Column(db.String(5))

class Course(db.Model):
    __tablename__ = 'courses'
    id           = db.Column(db.Integer, primary_key=True)
    code         = db.Column(db.String(20))
    name         = db.Column(db.String(150))
    semester     = db.Column(db.String(5))
    batch_year   = db.Column(db.String(4))
    department   = db.Column(db.String(10))

with app.app_context():
    db.create_all()

# ─── Helpers ──────────────────────────────────────────────────────────────────

SEMS = [f'S{i}' for i in range(1, 9)]

def current_year():  return session.get('year')
def current_dept():  return session.get('dept')

def require_class(fn):
    from functools import wraps
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if 'year' not in session or 'dept' not in session:
            return redirect(url_for('select_class'))
        return fn(*args, **kwargs)
    return wrapper


def class_name():
    return f"STM{current_dept()}{str(current_year())[-2:]}"


def build_sgpa_matrix():
    """Return (students_list, semesters_with_data, matrix)
       matrix: {register_no: {semester: sgpa_or_'-'}}
    """
    results = (db.session.query(SemesterResult)
               .join(Student)
               .filter(Student.batch_year == current_year(),
                       Student.department == current_dept())
               .all())

    reg_nos   = sorted({r.register_no for r in results})
    sems_used = sorted({r.semester for r in results},
                       key=lambda s: int(s[1:]))

    # sgpa lookup
    sgpa_map = {}
    for r in results:
        sgpa_map.setdefault(r.register_no, {})[r.semester] = r.sgpa

    # student names
    students = {s.register_no: s.name
                for s in Student.query.filter_by(batch_year=current_year(),
                                                 department=current_dept()).all()}

    matrix = []
    for rn in reg_nos:
        row = {'register_no': rn, 'name': students.get(rn, ''), 'sems': {}}
        for sem in sems_used:
            row['sems'][sem] = sgpa_map.get(rn, {}).get(sem, '-')
        matrix.append(row)

    return matrix, sems_used


def build_backlog_matrix():
    """Return (matrix, sems_used)
       matrix: [{register_no, name, sems: {sem: net_count_or_'-'}}]
    """
    grades = (db.session.query(CourseGrade)
              .join(Student)
              .filter(Student.batch_year == current_year(),
                      Student.department == current_dept())
              .all())

    grade_rows = [{'register_no': g.register_no, 'semester': g.semester,
                   'course_code': g.course_code, 'grade': g.grade}
                  for g in grades]

    net = compute_net_backlogs(grade_rows)

    sems_used = sorted({g.semester for g in grades}, key=lambda s: int(s[1:]))

    students = {s.register_no: s.name
                for s in Student.query.filter_by(batch_year=current_year(),
                                                 department=current_dept()).all()}

    reg_nos = sorted(net.keys())
    matrix  = []
    for rn in reg_nos:
        row = {'register_no': rn, 'name': students.get(rn, ''), 'sems': {}}
        for sem in sems_used:
            row['sems'][sem] = net.get(rn, {}).get(sem, '-')
        matrix.append(row)

    return matrix, sems_used


def get_courses_for_class():
    """Return list of {code, name, semester} for current class, ordered by semester."""
    courses = (Course.query
               .filter_by(batch_year=current_year(), department=current_dept())
               .all())
    return sorted(courses, key=lambda c: (int(c.semester[1:]), c.code))


# ─── Routes ───────────────────────────────────────────────────────────────────

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        session['user'] = request.form['username']
        session['role'] = 'admin' if request.form['username'] == 'admin' else 'faculty'
        return redirect(url_for('select_class'))
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/select_class', methods=['GET', 'POST'])
def select_class():
    if 'user' not in session:
        return redirect(url_for('login'))
    if request.method == 'POST':
        session['year'] = request.form['year']
        session['dept'] = request.form['dept']
        return redirect(url_for('index'))
    years = list(range(datetime.date.today().year, 2019, -1))
    return render_template('selection.html', years=years,
                           depts=['CS', 'EC', 'ME', 'CE', 'EEE'])


@app.route('/')
@require_class
def index():
    sgpa_matrix, sems = build_sgpa_matrix()
    backlog_matrix, _ = build_backlog_matrix()
    courses           = get_courses_for_class()
    return render_template('index.html',
                           class_name=class_name(),
                           role=session.get('role'),
                           sgpa_matrix=sgpa_matrix,
                           backlog_matrix=backlog_matrix,
                           sems=sems,
                           courses=courses)


@app.route('/upload', methods=['POST'])
@require_class
def upload_file():
    uploaded = 0
    errors   = []

    for file in request.files.getlist('pdf_file'):
        if not (file and file.filename.endswith('.pdf')):
            continue

        path = os.path.join(app.config['UPLOAD_FOLDER'],
                            secure_filename(file.filename))
        file.save(path)

        try:
            semester, courses_dict, students = parse_class_report(path)
        except Exception as e:
            errors.append(f"{file.filename}: {e}")
            continue

        if not semester or not students:
            errors.append(f"{file.filename}: Could not extract data.")
            continue

        # Upsert courses
        for code, name in courses_dict.items():
            existing = Course.query.filter_by(
                code=code, semester=semester,
                batch_year=current_year(), department=current_dept()).first()
            if not existing:
                db.session.add(Course(code=code, name=name, semester=semester,
                                      batch_year=current_year(),
                                      department=current_dept()))

        for s in students:
            reg_no = s['register_no']

            # Upsert student
            student = Student.query.get(reg_no)
            if not student:
                db.session.add(Student(register_no=reg_no, name=s['name'],
                                       batch_year=current_year(),
                                       department=current_dept()))
            else:
                # Update name if empty
                if not student.name and s['name']:
                    student.name = s['name']

            # Remove old result for same semester (re-upload scenario)
            SemesterResult.query.filter_by(
                register_no=reg_no, semester=semester).delete()
            CourseGrade.query.filter_by(
                register_no=reg_no, semester=semester).delete()

            db.session.add(SemesterResult(
                register_no=reg_no, semester=semester,
                sgpa=s['sgpa'], cgpa=s['cgpa'],
                backlogs=s['backlogs']))

            for code, grade in s['grades'].items():
                db.session.add(CourseGrade(
                    register_no=reg_no, semester=semester,
                    course_code=code,
                    course_name=courses_dict.get(code, ''),
                    grade=grade))

        db.session.commit()
        uploaded += 1

    session['upload_msg'] = (f"Successfully imported {uploaded} file(s)." +
                             (' Errors: ' + '; '.join(errors) if errors else ''))
    return redirect(url_for('index'))


# ─── API: subject-wise grade distribution ─────────────────────────────────────

@app.route('/api/subject_analysis')
@require_class
def subject_analysis():
    code = request.args.get('code', '')
    sem  = request.args.get('sem', '')

    q = (CourseGrade.query
         .join(Student)
         .filter(Student.batch_year == current_year(),
                 Student.department == current_dept(),
                 CourseGrade.course_code == code))
    if sem:
        q = q.filter(CourseGrade.semester == sem)

    rows = q.all()
    if not rows:
        return jsonify({'code': code, 'name': '', 'semester': sem, 'distribution': [], 'total': 0})

    from collections import Counter
    counts = Counter(r.grade for r in rows)
    GRADE_ORDER = ['S', 'A+', 'A', 'B+', 'B', 'C+', 'C', 'D', 'P', 'F', 'FE', 'LP', 'I']
    FAIL_GRADES_SET = {'F', 'FE', 'LP', 'I'}
    GRADE_POINTS = {'S': 10, 'A+': 9, 'A': 8.5, 'B+': 8, 'B': 7, 'C+': 6, 'C': 5, 'D': 4, 'P': 3, 'F': 0, 'FE': 0, 'LP': 0, 'I': 0}
    QUALITY_GRADES = {'S', 'A+', 'A', 'B+'}

    distribution = []
    for g in GRADE_ORDER:
        if g in counts:
            distribution.append({
                'grade': g,
                'count': counts[g],
                'fail':  g in FAIL_GRADES_SET
            })
    # Any grades not in GRADE_ORDER
    for g, cnt in counts.items():
        if g not in GRADE_ORDER:
            distribution.append({'grade': g, 'count': cnt, 'fail': False})

    total = len(rows)
    name = rows[0].course_name if rows else code

    # Compute analytics
    grade_points_sum = sum(GRADE_POINTS.get(r.grade, 0) for r in rows)
    avg_grade_point = round(grade_points_sum / total, 2) if total else 0

    # Topper grade: first grade in GRADE_ORDER that has count > 0
    topper_grade = ''
    topper_count = 0
    for g in GRADE_ORDER:
        if g in counts:
            topper_grade = g
            topper_count = counts[g]
            break

    # Quality index: % of students with B+ or above
    quality_count = sum(counts.get(g, 0) for g in QUALITY_GRADES)
    quality_index = round(quality_count / total * 100, 1) if total else 0

    return jsonify({
        'code': code,
        'name': name,
        'semester': sem,
        'distribution': distribution,
        'total': total,
        'avg_grade_point': avg_grade_point,
        'topper_grade': topper_grade,
        'topper_count': topper_count,
        'quality_index': quality_index
    })


# ─── Download report ──────────────────────────────────────────────────────────

@app.route('/download_report')
@require_class
def download_report():
    results = (db.session.query(SemesterResult)
               .join(Student)
               .filter(Student.batch_year == current_year(),
                       Student.department == current_dept())
               .all())

    grades = (db.session.query(CourseGrade)
              .join(Student)
              .filter(Student.batch_year == current_year(),
                      Student.department == current_dept())
              .all())

    students = (Student.query
                .filter_by(batch_year=current_year(), department=current_dept())
                .all())

    results_data = [{'register_no': r.register_no, 'semester': r.semester,
                     'sgpa': r.sgpa, 'backlogs': r.backlogs} for r in results]
    grades_data  = [{'register_no': g.register_no, 'semester': g.semester,
                     'course_code': g.course_code, 'course_name': g.course_name,
                     'grade': g.grade} for g in grades]
    students_data = [{'register_no': s.register_no, 'name': s.name}
                     for s in students]

    net_backlogs = compute_net_backlogs(grades_data)

    short_year = str(current_year())[-2:]
    fname = f"AURAS_Report_{class_name()}.xlsx"
    path  = os.path.join(app.config['DOWNLOAD_FOLDER'], fname)

    generate_excel_report(results_data, grades_data, students_data,
                          net_backlogs, path)
    return send_file(path, as_attachment=True)


# ─── Download single-subject report ──────────────────────────────────────────

@app.route('/download_subject_report')
@require_class
def download_subject_report():
    sem  = request.args.get('sem', '')
    code = request.args.get('code', '')

    if not sem or not code:
        return redirect(url_for('index'))

    grades = (db.session.query(CourseGrade)
              .join(Student)
              .filter(Student.batch_year == current_year(),
                      Student.department == current_dept())
              .all())

    students = (Student.query
                .filter_by(batch_year=current_year(), department=current_dept())
                .all())

    grades_data  = [{'register_no': g.register_no, 'semester': g.semester,
                     'course_code': g.course_code, 'course_name': g.course_name,
                     'grade': g.grade} for g in grades]
    students_data = [{'register_no': s.register_no, 'name': s.name}
                     for s in students]

    fname = f"AURAS_{class_name()}_{code}_{sem}.xlsx"
    path  = os.path.join(app.config['DOWNLOAD_FOLDER'], fname)

    generate_subject_excel_report(grades_data, students_data, path, code, sem)
    return send_file(path, as_attachment=True)


# ─── Clear database ───────────────────────────────────────────────────────────

@app.route('/clear_database', methods=['POST'])
@require_class
def clear_database():
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
    scope = request.form.get('scope', 'class')
    if scope == 'all':
        CourseGrade.query.delete()
        SemesterResult.query.delete()
        Student.query.delete()
        Course.query.delete()
    else:
        # Only current class
        reg_nos = [s.register_no for s in
                   Student.query.filter_by(batch_year=current_year(),
                                           department=current_dept()).all()]
        CourseGrade.query.filter(CourseGrade.register_no.in_(reg_nos)).delete(synchronize_session=False)
        SemesterResult.query.filter(SemesterResult.register_no.in_(reg_nos)).delete(synchronize_session=False)
        Student.query.filter_by(batch_year=current_year(), department=current_dept()).delete()
        Course.query.filter_by(batch_year=current_year(), department=current_dept()).delete()
    db.session.commit()
    return redirect(url_for('index'))


if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'],   exist_ok=True)
    os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)
    app.run(debug=True)
