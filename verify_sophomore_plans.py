import sys
import re
import major_requirements as mr
import csv

######################################################################

def gather_courses(cur_term, student, subject):

    courses = []

    for s, course, term, cr, grade, _ in student.enrollments:
        if subject != s: continue
        if term == cur_term:
            course = course + '*'
        elif int(term) > int(cur_term):
            course = course + '?'
        courses.append(course)

    return ', '.join(courses)

######################################################################

def fix_no_e90(student):

    if student.gradyear == 2022:
        
        has_e90 = False
        
        grad_term = student.senior_term()
        grad_term_count = 0
        grad_term_engr_count = 0
            
        for e in student.enrollments:
            if (e.subject, e.course) == ('ENGR', '90'):
                assert e.term == grad_term
                has_e90 = True
            elif e.term == grad_term:
                grad_term_count += 1
                if e.subject == 'ENGR':
                    grad_term_engr_count += 1
                    
        if not has_e90:
            assert grad_term_count <= 3
            assert grad_term_engr_count <= 2
            print('Warning: Had to enroll {} in E90 along with {} other courses ({} ENGR courses)!'.format(student, grad_term_count, grad_term_engr_count), file=sys.stderr)
            student.enroll('ENGR', '90', grad_term, 1.0)

######################################################################    

def deal_with(cur_term, major_or_minor, students, requirements,
              subject_columns, do_emails=False, fix_student=None):

    total_pass = 0

    records = []
    email_records = None

    headers = ['ID', 'First', 'Last', 'Email', 'Class', 'Advisor', 'ENGR GPA', 'Eng/math/sci GPA', 'Honors', 'Abroad?', 'Accept?']

    email_headers = ['Email', 'cc', 'First', 'Last', 'Degree', 'Missing']

    for r in requirements:
        headers.append(r.name + '?')

    for subject in subject_columns:
        headers.append(subject)

    records.append(headers)

    emails_str = None

    for student in sorted(students.values(), key=mr.student_key):

        if not len(student.enrollments):
            print('*** NO ENROLLMENTS FOR {} ***'.format(student), file=sys.stderr)
        elif fix_student is not None:
            fix_student(student)
        
        all_ok, results = mr.verify_requirements(student, requirements)
        if all_ok: total_pass += 1
        
        row = []
        
        row.append(student.id)
        row.append(student.first)
        row.append(student.last)
        row.append(student.email_id)
        row.append(student.gradyear)
        row.append(student.advisor)
        
        if student.gpa('ENGR') is not None:
            row.append('{:.2f}'.format(student.gpa('ENGR')))
        else:
            row.append('')
            
        if student.gpa() is not None:
            row.append('{:.2f}'.format(student.gpa()))
        else:
            row.append('')

        row.append(student.honors)

        row.append(' and '.join(sorted(student.abroad)))

        decision = student.decision
        
        if not decision and all_ok:
            decision = 'yes'
        
        row.append(decision)

        for (ok, total, _), r in zip(results, requirements):
            if ok:
                okstr = 'y'
            elif abs(r.threshold) <= 1:
                okstr = 'n'
            else:
                okstr = '{}/{}'.format(mr.maybe_int(total), abs(mr.maybe_int(r.threshold)))
            row.append(okstr)

        for subject in subject_columns:
            row.append(gather_courses(cur_term, student, subject))

        records.append(row)

        if do_emails and not all_ok and not student.decision:
            
            student_email = student.email_id + '@swarthmore.edu'

            adv = mr.check_engr_advisor(student.advisor)
            if adv is not None:
                advisor_email = adv + '@swarthmore.edu'
            else:
                advisor_email = ''

            if not len(student.enrollments):
                missing = 'Please submit a sophomore plan on mySwarthmore ASAP!'
            else:
                missing, _ = mr.get_missing(requirements, results)
                missing = 'Remaining requirements:\n\n' + missing

            if email_records is None:
                email_records = [email_headers]

            email_records.append([student_email, advisor_email,
                                  student.first, student.last,
                                  major_or_minor, missing])

                
    #print('\n'.join(tabulate(records)))

    output_filename = 'outputs/{}_requirements.csv'.format(major_or_minor)

    ostr = open(output_filename, 'w')
    writer = csv.writer(ostr)
    writer.writerows(records)

    print('{} of {} prospective {}s satisfy requirements.\n'.format(
        total_pass, len(students), major_or_minor))

    print('wrote', output_filename, file=sys.stderr)

    if email_records is not None:
        filename = 'outputs/{}_sophomore_plans.csv'.format(major_or_minor)
        with open(filename, 'w') as ostr:
            writer = csv.writer(ostr)
            writer.writerows(email_records)
        print('wrote', filename, file=sys.stderr)

######################################################################

def deal_with_majors(cur_term, students, requirements):

    deal_with(cur_term, 'major', students, requirements, 
              ['ENGR', 'MATH', 'PHYS', 'ASTR', 'BIOL', 'CHEM'],
              True, fix_no_e90)
    
######################################################################

def deal_with_minors(cur_term, students, requirements):

    deal_with(cur_term, 'minor', students, requirements, 
              ['ENGR'])

######################################################################
# parse major/minor decisions already made

def parse_decisions(elookup, filename):

    print('reading ' + filename)

    lineno = 0
    
    with open(filename, 'r') as istr:
        for line in istr:
            lineno += 1
            line = line.rstrip()
            if not line:
                continue
            email_id, decision = line.split()
            if email_id in elookup:
                s = elookup[email_id]
                s.decision = decision

######################################################################
# scrape report PDF -> HTML -> txt

def scrape_report(elookup, filename):

    EMAIL_REGEX = re.compile(r'(\w+)@swarthmore.edu$')
    
    TERM_REGEX = re.compile(r'^(Spring|Fall) ([0-9][0-9][0-9][0-9])')
    
    PROJ_SUBJECT_REGEX = re.compile(r'^[^(]+ \(([A-Z][A-Z][A-Z][A-Z]) (\w+\.?\w*)\):')

    PROJ_CREDIT_REGEX = re.compile(r'\[CREDITS: +([0-9](\.[0-9]+)?)( .*credit.*)?\]')

    HIST_SUBJECT_REGEX = re.compile(r'^([A-Z][A-Z][A-Z][A-Z]) (\w+)( \w+)?$')
    HIST_CREDIT_REGEX = re.compile(r'^([0-9.]+)$')
    HIST_GRADE_REGEX = re.compile(r'^CR \((.*)\)$')

    istr = open(filename, 'r')
    print('reading ' + filename)
    debug = False

    email_id = None
    student = None
    
    term = None
    
    pending_enrollment = None
    
    state = 'inactive'
    lineno = 0

    peek = None

    while True:

        if peek:
            line = peek
            peek = None
        else:
            try:
                line = next(istr).rstrip()
            except StopIteration:
                break

        lineno += 1

        if debug: print('*** {:04d} *** '.format(lineno) + line)
        
        if not line:
            continue

        email_match = re.search(EMAIL_REGEX, line)

        if email_match:
            assert not pending_enrollment
            new_email_id = email_match.group(1)
            if new_email_id != email_id:
                term = None
                state = 'inactive'
                email_id = new_email_id
                if email_id not in elookup:
                    student = None
                else:
                    student = elookup[email_id]
        elif '@swarthmore.edu' in line:
            print('{}:{} unexpected email address'.format(filename, lineno))
            sys.exit(1)

        if student is None:
            continue

        term_match = re.search(TERM_REGEX, line)
        
        if term_match:
            
            assert not pending_enrollment 
            term = mr.parse_term(term_match.group(0))
            
        elif line == '--Sophomore Plan Course Projection--':
            
            assert email_id and term
            assert not pending_enrollment 
            state = 'projecting'
            
        elif line == '----------------Academic History--------------':
            
            assert email_id and term
            assert pending_enrollment is None
            state = 'history'
            
        elif state == 'projecting':
            
            dept_match = re.search(PROJ_SUBJECT_REGEX, line)

            if '[CREDITS: ] Abroad' in line:
                student = elookup[email_id]
                student.abroad.add(mr.unparse_term(term))
                assert not dept_match
                continue

            if dept_match:
                assert pending_enrollment is None
                pending_enrollment = dict()
                pending_enrollment['subject'] = dept_match.group(1)
                pending_enrollment['course'] = dept_match.group(2)
                pending_enrollment['term'] = term
                pending_enrollment['grade'] = None

            credit_match = re.search(PROJ_CREDIT_REGEX, line)

            if pending_enrollment and not credit_match and '[' in line and ']' not in line:
                next_line = next(istr).rstrip()
                lineno += 1
                if debug: print('### {:04d} ### '.format(lineno) + next_line)
                line += ' ' + next_line
                credit_match = re.search(PROJ_CREDIT_REGEX, line)
                assert credit_match

            if credit_match:

                if debug: print('dept_match:', dept_match)
                if debug: print('pending_enrollment:', pending_enrollment)
                assert pending_enrollment is not None
                assert 'credit' not in pending_enrollment
                pending_enrollment['credit'] = credit_match.group(1)

            elif '[CREDITS: ]' in line and pending_enrollment:

                pending_enrollment['credit'] = '0'

            elif pending_enrollment and not dept_match:
                
                print('shoot')
                sys.exit(1)

        elif state == 'history':

            if pending_enrollment is None:

                subject_match = re.match(HIST_SUBJECT_REGEX, line)
            
                if subject_match:
                    pending_enrollment = dict()
                    pending_enrollment['subject'] = subject_match.group(1)
                    pending_enrollment['course'] = subject_match.group(2)
                    pending_enrollment['grade'] = None
                    pending_enrollment['term'] = term

                    # next line is title, ignore it
                    line = next(istr).rstrip()
                    lineno += 1
                    if debug: print('~~~ {:04d} ~~~ '.format(lineno) + line)

            elif 'credit' not in pending_enrollment:

                credit_match = re.match(HIST_CREDIT_REGEX, line)

                if debug and not credit_match:
                    print('^^^ {:04d} ^^^ '.format(lineno) + line)

                assert credit_match
                
                pending_enrollment['credit'] = credit_match.group(1)

                peek = next(istr).rstrip()
                
                grade_match = re.match(HIST_GRADE_REGEX, peek)

                if peek in mr.ALL_GRADES or grade_match:
                    if peek in mr.ALL_GRADES:
                        pending_enrollment['grade'] = peek
                    else:
                        pending_enrollment['grade'] = grade_match.group(1)
                    line = peek
                    peek = None
                    lineno += 1
                    if debug: print('### {:04d} ### '.format(lineno) + line)

        if pending_enrollment and 'credit' in pending_enrollment:
            pending_enrollment['source'] = '{}:{}'.format(filename, lineno)
            student.enroll(**pending_enrollment)
            pending_enrollment = None

######################################################################

def main():

    #major_requirements = [ item for item in mr.MAJOR_REQUIREMENTS if item.name != '5 ENGR electives' ]

    report_file = 'inputs/majors_report.txt'

    majors_file = 'inputs/majors.csv'
    minors_file = 'inputs/minors.csv'

    majors = mr.parse_students(majors_file)
    minors = mr.parse_students(minors_file)

    students = dict()
    
    if len(sys.argv) > 1:

        lastnames = set([l.lower() for l in sys.argv[1:]])
        
        def named_student(item):
            id, s = item
            return s.last.lower() in lastnames

        majors = dict(filter(named_student, majors.items()))
        minors = dict(filter(named_student, minors.items()))
    
    students.update(majors)
    students.update(minors)

    elookup = dict([(s.email_id, s) for s in students.values()])
    
    scrape_report(elookup, report_file)
    mr.parse_etc(elookup, 'inputs/placements_ocs_etc.txt')

    cur_term = mr.get_cur_term()

    deal_with_majors(cur_term, majors, mr.MAJOR_REQUIREMENTS)
    deal_with_minors(cur_term, minors, mr.MINOR_REQUIREMENTS)

if __name__ == '__main__':
    main()

    
