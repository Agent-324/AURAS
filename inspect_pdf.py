import sys
sys.path.insert(0, '.')
from extract_engine import detect_and_parse

pdf_path = r'../input/result_STM (7) (3).pdf'
semester, courses, students = detect_and_parse(pdf_path)

# Check CE students for CET309
ce_students = [s for s in students if s.get('_department') == 'CE']
print(f'CE students: {len(ce_students)}')
# Check total grades for CE students with full set
for s in ce_students[:3]:
    print(f"  {s['register_no']}: {len(s['grades'])} grades -> {s['grades']}")

# Check the first LSTM student which should have CET309
lstm_19 = [s for s in students if s['register_no'] == 'LSTM23CE019']
if lstm_19:
    print(f"\nLSTM23CE019 grades: {lstm_19[0]['grades']}")

# Check students with missing courses
ce_full = [s for s in ce_students if len(s['grades']) >= 8]
ce_partial = [s for s in ce_students if len(s['grades']) < 8]
print(f"\nCE full (>=8 subjects): {len(ce_full)}")
print(f"CE partial (<8 subjects): {len(ce_partial)}")

# Show a full CE student
if ce_full:
    s = ce_full[0]
    print(f"  {s['register_no']}: SGPA={s['sgpa']} grades={s['grades']}")
