import re

FAIL_GRADES = {'F', 'FE', 'I', 'LP'}
GRADE_ORDER = ['S', 'A+', 'A', 'B+', 'B', 'C+', 'C', 'D', 'P', 'F', 'FE', 'LP', 'I']


def _clean(cell):
    return str(cell).replace('\n', ' ').strip() if cell else ''


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
