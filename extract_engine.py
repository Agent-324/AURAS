import re

FAIL_GRADES = {'F', 'FE', 'I', 'LP', 'Absent'}
GRADE_ORDER = ['S', 'A+', 'A', 'B+', 'B', 'C+', 'C', 'D', 'P', 'F', 'FE', 'LP', 'I', 'Absent']

GRADE_POINTS = {
    'S': 10, 'A+': 9, 'A': 8.5, 'B+': 8, 'B': 7,
    'C+': 6, 'C': 5, 'D': 4, 'P': 3,
    'F': 0, 'FE': 0, 'LP': 0, 'I': 0, 'Absent': 0,
}

# KTU 2019 scheme – known course-credit overrides.
# Use exact course-code mappings where available, then fall back to a
# prefix-based heuristic for anything we do not explicitly know yet.
KTU_CREDITS = {
    # ----- CS S4 -----
    'MAT206': 4, 'CST202': 4, 'CST204': 4, 'CST206': 4,
    'HUT200': 2, 'MCN202': 0, 'CSL202': 2, 'CSL204': 2,
    # ----- CS S5 -----
    'CST301': 4, 'CST303': 4, 'CST305': 4, 'CST307': 4, 'CST309': 3,
    'CSL331': 2, 'CSL333': 2,
    # ----- CE S5 -----
    'CET301': 4, 'CET303': 4, 'CET305': 4, 'CET307': 4, 'CET309': 4,
    'CEL331': 2, 'CEL333': 2,
    # ----- ME S5 -----
    'MET301': 4, 'MET303': 4, 'MET305': 4, 'MET307': 4,
    'MEL331': 2, 'MEL333': 2,
    # ----- EC S5 -----
    'ECT301': 4, 'ECT303': 4, 'ECT305': 4, 'ECT307': 4, 'ECT309': 4,
    'ECL331': 2, 'ECL333': 2,
    # ----- EEE S5 -----
    'EET301': 4, 'EET303': 4, 'EET305': 4, 'EET307': 4,
    'EEL331': 2, 'EEL333': 2,
    # ----- CD (Computer Science and Design) S5 -----
    'CDT305': 4, 'CDT307': 4,
    'CDL331': 2,
    # ----- Common -----
    'MCN301': 0, 'MCN201': 0, 'MCN202': 0,   # no-credit
    'HUT300': 3, 'HUT310': 3,
}


def _guess_credits(code):
    """Heuristic credit guess based on KTU naming conventions."""
    if code in KTU_CREDITS:
        return KTU_CREDITS[code]
    if code.startswith('MCN'):
        return 0
    if 'L' in code[2:4]:        # Lab courses (e.g. CSL, CEL, MEL)
        return 2
    if code.startswith('HUT'):
        return 3
    return 4                     # Default theory course


def _calculate_sgpa(grades_dict):
    """
    Calculate SGPA from {course_code: grade} dict.
    SGPA = Σ(Credit × GradePoint) / Σ(Credit)   [only credit-bearing courses]
    """
    total_credits = 0
    total_points = 0
    for code, grade in grades_dict.items():
        cr = _guess_credits(code)
        if cr == 0:
            continue   # skip non-credit courses
        gp = GRADE_POINTS.get(grade, 0)
        total_credits += cr
        total_points += cr * gp
    return round(total_points / total_credits, 2) if total_credits else 0.0


def _clean(cell):
    return str(cell).replace('\n', ' ').strip() if cell else ''


# ══════════════════════════════════════════════════════════════════════════════
# Parser 1 – Tabular "Semester Grade Card Report" PDFs (original format)
# ══════════════════════════════════════════════════════════════════════════════

def parse_class_report(pdf_path):
    """
    Parse a KTU class-wise semester grade card report PDF.

    Returns:
        semester (str)          – e.g. 'S4'
        courses  (dict)         – {code: name}
        students (list of dict) – keys: register_no, name, sgpa, cgpa,
                                         earned_credits, grades {code: grade},
                                         backlogs
    """
    semester = None
    courses = {}        # code -> full name
    students = []
    course_col_map = {} # column-index -> course_code  (set from first table that has header)
    sgpa_idx = cgpa_idx = earned_idx = None

    try:
        import pdfplumber
    except ModuleNotFoundError as exc:
        raise RuntimeError("pdfplumber is required to parse uploaded PDF reports.") from exc

    with pdfplumber.open(pdf_path) as pdf:
        full_text = '\n'.join(p.extract_text() or '' for p in pdf.pages)

        # ── semester ──────────────────────────────────────────────────────
        m = re.search(r'Semester\s*[:\-]\s*(S\d)', full_text, re.IGNORECASE)
        if m:
            semester = m.group(1)

        # ── collect all tables from every page ────────────────────────────
        all_tables = []
        for page in pdf.pages:
            all_tables.extend(page.extract_tables() or [])

        for table in all_tables:
            if not table:
                continue

            # ── course-info table: cells look like "MAT206 - GRAPH THEORY" ──
            for row in table:
                for cell in row:
                    c = _clean(cell)
                    m = re.match(r'([A-Z]{2,4}\d{3,4})\s*[-–]\s*(.+)', c)
                    if m:
                        courses[m.group(1)] = m.group(2).strip()

            # ── detect student-grades table by looking for 'Student' header ─
            header = [_clean(c) for c in table[0]]
            is_header_row = 'Student' in header and 'SGPA' in header

            if is_header_row:
                # Build column mapping from this header
                for i, h in enumerate(header):
                    # course codes are split across two lines, e.g. 'MAT20\n6' → 'MAT206'
                    code_candidate = re.sub(r'\s+', '', h)   # remove all whitespace
                    if re.match(r'^[A-Z]{2,4}\d{3,4}$', code_candidate):
                        course_col_map[i] = code_candidate
                    elif 'SGPA' in h:
                        sgpa_idx = i
                    elif 'CGPA' in h:
                        cgpa_idx = i
                    elif 'Earned' in h or 'Earn' in h:
                        earned_idx = i
                data_rows = table[1:]
            else:
                # Pages 2+ – no header, use same column map established above
                data_rows = table

            if not course_col_map:
                continue   # haven't found a header yet, skip

            for row in data_rows:
                if not row or not row[0]:
                    continue
                student_cell = _clean(row[0])

                # Register numbers: STM23CS046, LSTM23CS055, MLT23CS011 …
                rm = re.match(r'([A-Z]+\d{2}[A-Z]+\d{3})', student_cell)
                if not rm:
                    continue

                reg_no = rm.group(1)
                name   = student_cell[len(reg_no):].lstrip('- ').strip()

                grades = {}
                for col_i, code in course_col_map.items():
                    if col_i < len(row) and row[col_i]:
                        g = _clean(row[col_i])
                        if g and g not in ('None', '-'):
                            grades[code] = g

                def _float(idx):
                    try:
                        return float(_clean(row[idx])) if idx is not None and idx < len(row) and row[idx] else 0.0
                    except ValueError:
                        return 0.0

                students.append({
                    'register_no':   reg_no,
                    'name':          name,
                    'sgpa':          _float(sgpa_idx),
                    'cgpa':          _float(cgpa_idx),
                    'earned_credits': _float(earned_idx),
                    'grades':        grades,
                    'backlogs':      sum(1 for g in grades.values() if g in FAIL_GRADES),
                })

    # De-duplicate (same reg_no may appear on multiple pages due to table split)
    seen = {}
    for s in students:
        seen[s['register_no']] = s
    students = list(seen.values())

    return semester, courses, students


# ══════════════════════════════════════════════════════════════════════════════
# Parser 2 – KTU Exam Result PDFs  (inline "CourseCode(Grade)" format)
# ══════════════════════════════════════════════════════════════════════════════

# Regex patterns for the exam-result format
_DEPT_HEADER_RE = re.compile(
    r'([A-Z][A-Z &]+(?:ENGINEERING|TECHNOLOGY|SCIENCE))\s*\[',
    re.IGNORECASE,
)
_COURSE_LINE_RE = re.compile(
    r'([A-Z]{2,4}\d{3,4})\s+(.+)',
)
_GRADE_PAIR_RE = re.compile(
    r'([A-Z]{2,4}\d{3,4})\(([^)]+)\)',
)
_REG_NO_RE = re.compile(
    r'^((?:L?STM|ECE|MLT)\d{2}[A-Z]{2}\d{3})',
)

_DEPT_ABBREV = {
    'CIVIL': 'CE',
    'MECHANICAL': 'ME',
    'COMPUTER SCIENCE': 'CS',
    'COMPUTER SCIENCE AND DESIGN': 'CD',
    'COMPUTER SCIENCE AND ENGINEERING': 'CS',
    'ELECTRONICS AND COMMUNICATION': 'EC',
    'ELECTRONICS': 'EC',
    'ELECTRICAL AND ELECTRONICS': 'EEE',
    'ELECTRICAL': 'EEE',
    'INFORMATION TECHNOLOGY': 'IT',
}

def _detect_dept_abbrev(dept_full_name):
    """Return short department code from the full department name."""
    up = dept_full_name.upper().strip()
    for key, abbrev in _DEPT_ABBREV.items():
        if key in up:
            return abbrev
    return up[:3]


def parse_exam_result(pdf_path):
    """
    Parse a KTU exam result PDF that uses inline CourseCode(Grade) format.

    Returns:
        semester (str)          – e.g. 'S5'
        courses  (dict)         – {code: name}
        students (list of dict) – same shape as parse_class_report
    """
    try:
        import pdfplumber
    except ModuleNotFoundError as exc:
        raise RuntimeError("pdfplumber is required.") from exc

    semester = None
    courses = {}                # code -> full name
    students = []
    current_dept = None

    with pdfplumber.open(pdf_path) as pdf:
        full_text = '\n'.join(p.extract_text() or '' for p in pdf.pages)

        # ── Detect semester from title ────────────────────────────────────
        # e.g. "B.Tech S5 (R, S) Exam Nov 2025"
        m = re.search(r'B\.?Tech\s+(S\d)', full_text, re.IGNORECASE)
        if m:
            semester = m.group(1)
        else:
            m = re.search(r'\(S(\d)\s+Result\)', full_text, re.IGNORECASE)
            if m:
                semester = f'S{m.group(1)}'

        # ── Process all pages ─────────────────────────────────────────────
        for page in pdf.pages:
            text = page.extract_text() or ''
            lines = text.split('\n')

            in_course_list = False   # True between "Course Code   Course" and "Register No"

            # First pass: join continuation lines.
            # A continuation line starts with a course code (e.g. "CET309(B)")
            # rather than a register number or header.
            merged_lines = []
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                # Is this a continuation of the previous student line?
                # (starts with CourseCode( but NOT a register number)
                if (merged_lines
                        and _GRADE_PAIR_RE.match(line)
                        and not _REG_NO_RE.match(line)
                        and not _DEPT_HEADER_RE.search(line)):
                    merged_lines[-1] += ' ' + line
                else:
                    merged_lines.append(line)

            for line in merged_lines:
                # Detect department header
                dm = _DEPT_HEADER_RE.search(line)
                if dm:
                    current_dept = _detect_dept_abbrev(dm.group(1))
                    in_course_list = False
                    continue

                # Detect course list header
                if re.match(r'Course\s+Code\s+Course', line, re.IGNORECASE):
                    in_course_list = True
                    continue

                # Detect student data header (ends course list)
                if re.match(r'Register\s+No\s+Course\s+Code', line, re.IGNORECASE):
                    in_course_list = False
                    continue

                # Parse course definitions
                if in_course_list:
                    cm = _COURSE_LINE_RE.match(line)
                    if cm:
                        courses[cm.group(1)] = cm.group(2).strip()
                    continue

                # Parse student grade lines
                rm = _REG_NO_RE.match(line)
                if rm:
                    reg_no = rm.group(1)
                    # Extract all CourseCode(Grade) pairs from the merged line
                    grade_pairs = _GRADE_PAIR_RE.findall(line)
                    grades = {code: grade for code, grade in grade_pairs}

                    sgpa = _calculate_sgpa(grades)
                    backlogs = sum(1 for g in grades.values() if g in FAIL_GRADES)

                    students.append({
                        'register_no': reg_no,
                        'name': '',       # This format doesn't include student names
                        'sgpa': sgpa,
                        'cgpa': 0.0,      # Not available in this format
                        'earned_credits': 0.0,
                        'grades': grades,
                        'backlogs': backlogs,
                        '_department': current_dept,  # Extra field for filtering
                    })

            # Also check tables (pdfplumber sometimes captures these better)
            tables = page.extract_tables() or []
            for table in tables:
                for row in table:
                    if not row or not row[0]:
                        continue
                    cell0 = _clean(row[0])
                    rm = _REG_NO_RE.match(cell0)
                    if not rm:
                        continue
                    reg_no = rm.group(1)
                    # Already parsed from text? Skip.
                    if any(s['register_no'] == reg_no for s in students):
                        continue
                    # Combine all cells and extract grade pairs
                    combined = ' '.join(_clean(c) for c in row if c)
                    grade_pairs = _GRADE_PAIR_RE.findall(combined)
                    grades = {code: grade for code, grade in grade_pairs}
                    if not grades:
                        continue
                    sgpa = _calculate_sgpa(grades)
                    backlogs = sum(1 for g in grades.values() if g in FAIL_GRADES)
                    students.append({
                        'register_no': reg_no,
                        'name': '',
                        'sgpa': sgpa,
                        'cgpa': 0.0,
                        'earned_credits': 0.0,
                        'grades': grades,
                        'backlogs': backlogs,
                        '_department': current_dept,
                    })

    # De-duplicate: keep the entry with the most grades
    seen = {}
    for s in students:
        key = s['register_no']
        if key not in seen or len(s['grades']) > len(seen[key]['grades']):
            seen[key] = s
    students = list(seen.values())

    return semester, courses, students


# ══════════════════════════════════════════════════════════════════════════════
# Auto-detect & parse
# ══════════════════════════════════════════════════════════════════════════════

def detect_and_parse(pdf_path):
    """
    Auto-detect PDF format and use the appropriate parser.
    Returns the same tuple as parse_class_report / parse_exam_result.
    """
    try:
        import pdfplumber
    except ModuleNotFoundError as exc:
        raise RuntimeError("pdfplumber is required.") from exc

    # Peek at the first page to decide
    with pdfplumber.open(pdf_path) as pdf:
        first_text = (pdf.pages[0].extract_text() or '') if pdf.pages else ''

    # The tabular "Semester Grade Card Report" format has the string
    # "Semester Grade Card Report" in the title, and tables with Student / SGPA
    # columns.  The exam-result format has "Exam" in the title and inline grades.
    is_grade_card = 'Semester Grade Card Report' in first_text
    is_exam_result = bool(re.search(r'B\.?Tech\s+S\d.*Exam', first_text, re.IGNORECASE))

    if is_grade_card:
        return parse_class_report(pdf_path)
    elif is_exam_result:
        return parse_exam_result(pdf_path)
    else:
        # Try the tabular parser first; fall back to exam-result
        semester, courses, students = parse_class_report(pdf_path)
        if students:
            return semester, courses, students
        return parse_exam_result(pdf_path)


def compute_net_backlogs(grade_rows):
    """
    Given a list of dicts {register_no, semester, course_code, grade},
    return {register_no: {semester: net_backlog_count}}

    Logic: accumulate failed courses in an outstanding set.
            Clear a course from outstanding if it appears with a passing grade
            in any subsequent semester.
    """
    SEMS = [f'S{i}' for i in range(1, 9)]

    # Build {reg_no: {semester: {course_code: grade}}}
    data = {}
    for r in grade_rows:
        data.setdefault(r['register_no'], {}) \
            .setdefault(r['semester'], {})[r['course_code']] = r['grade']

    result = {}
    for reg_no, sem_data in data.items():
        outstanding = set()   # course codes with outstanding fail
        result[reg_no] = {}
        for sem in SEMS:
            if sem not in sem_data:
                continue
            grades_this_sem = sem_data[sem]
            # Add new failures
            for code, grade in grades_this_sem.items():
                if grade in FAIL_GRADES:
                    outstanding.add(code)
            # Clear backlogs where student passed this semester
            cleared = {c for c in outstanding
                       if grades_this_sem.get(c) and grades_this_sem[c] not in FAIL_GRADES}
            outstanding -= cleared
            result[reg_no][sem] = len(outstanding)

    return result
