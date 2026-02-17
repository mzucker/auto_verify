#from openpyxl import load_workbook, Workbook
import datetime
import itertools
import sys
import csv
import re
from collections import defaultdict
from collections import namedtuple

# make a named tuple for enrollments
Enrollment = namedtuple('Enrollment', 'subject, course, term, credit, grade, source')

# get all permutations of A,B,C,D (stars come from off-campus?)
def get_letter_grades():

    rval = dict()

    for grade, base in zip('ABCD', [4.0, 3.0, 2.0, 1.0]):
        for modifier, delta in zip(['', '*', '+', '-'], [0, 0, 0.33, -0.33]):
            rval[grade+modifier] = min(base+delta, 4.0) # A+ does not count for more than A according to Martin warner

    return rval

######################################################################
# global variables

# for advising
ENGR_ADVISORS = dict(
    zucker='mzucker1',
    delano='mdelano1',
    everbach='ceverba1',
    molter='lmolter1',
    moser='amoser2',
    piovoso='mpiovos1',
)

# get all parse-able grades
LETTER_GRADES = get_letter_grades()

# used to when parsing Excel files
PASSING_GRADES = set(list(LETTER_GRADES.keys()) + ['CR', '4', '5'])
BAD_GRADES = set(['NC', 'W', 'INC'])

ALL_GRADES = PASSING_GRADES | BAD_GRADES

# crosslists between ENGR and other departments
ENGR_XLISTS = dict([
#    (('MATH', '24'), ('ENGR', '19')), 
    (('CPSC', '45'), ('ENGR', '22')),
    (('CPSC', '75'), ('ENGR', '23')),
    (('CPSC', '52'), ('ENGR', '25')),
    (('CPSC', '40'), ('ENGR', '26')),
    (('CPSC', '72'), ('ENGR', '27')),
    (('CPSC', '82'), ('ENGR', '28'))
])

# ignore all subjects besides these
ENGR_RELATED_SUBJECT = set('ENGR MATH STAT BIOL CHEM PHYS ASTR CPSC'.split())

# courses that don't count for credit in the ENGR major
ENGR_NONMAJOR_COURSES = set([
    ('ENGR', '1'),
    ('ENGR', '3'),
    ('ENGR', '4'),
    ('ENGR', '4A'),
    ('ENGR', '7'),
    ('ENGR', '10'),
])

ENGR_DR = set([('ENGR', '93')])

# courses that unconditionally count for credit in the ENGR major
ENGR_CORE_COURSES = set([
    ('ENGR', '6'),
    ('ENGR', '11'),
    ('ENGR', '11A'),
    ('ENGR', '11B'),
    ('ENGR', '12'),
    ('ENGR', '14'),
    ('ENGR', '41'),
    ('ENGR', '90'),
])

# E15 or E19 can be either core or elective, confusingly
E15_OR_E19 = set([
    ('ENGR', '15'),
    ('ENGR', '15A'),
    ('ENGR', '15B'),
    ('ENGR', '19'),
])
    
# these courses don't count towards the 8 math/science credits for ENGR
MATH_SCI_EXCLUSIONS = set([
    ('MATH', '3'),
    ('MATH', '15SP'),
    ('MATH', '3X'), # used to designate placement
    ('MATH', '4X'), # used to designate placement
    ('BIOL', '1SP'),
    ('BIOL', '2SP'),
    ('ASTR', '1'),
    ('PHYS', '1C'),
    ('PHYS', '2E'),
    ('STAT', '1'),
    ('CHEM', '3A'),
    ('CHEM', '3B'),    
    ('CHEM', '3C'),
])

# these courses are taught outside of the Engineering department
TAUGHT_OUTSIDE_ENGR = set([
    ('CPSC', '45'),
    ('CPSC', '75'),
    ('CPSC', '40')
])

# you can have at most one math/science course without a lab (dumb)
NONLAB_COURSES = set([
    ('PHYS', '5'),
    ('BIOL', '9'),
    ('ASTR', '14'),
])

REPEATABLE_COURSES = set([
    ('ENGR', '93'), # can repeat directed readings
    ('CPSC', '35'), # somehow sophomore plans had duplicate CS courses?
    ('CPSC', '21'),
    ('CPSC', '31'),
    ('BIOL', '139')
])

NSE_SUBJECTS = set([
    'BIOL',
    'CHEM',
    'CPSC',
    'ENGR',
    'MATH', 'STAT',
    'PHYS', 'ASTR',
])

######################################################################
# normalize course number

def course_normalize(crse):

    assert isinstance(crse, str)
    
    while crse.startswith('0'):
        crse = crse[1:]

    return crse

######################################################################
# parse a term as Spring/Fall YEAR

SEASONS = dict(Spring='02', Fall='04', January='01', Winter='01', Summer='03')
SEASONS_INV = { '02': 'Spring', '04': 'Fall', '01': 'Winter', '03': 'Summer' }

def parse_term(term):

    m = re.match(r'(Spring|Fall|January|Summer|winter) ([0-9][0-9][0-9][0-9])', term)
    assert m is not None

    season = m.group(1)
    year = m.group(2)

    return year + SEASONS[season]

######################################################################
# convert from 202104 to Fall 2021 example

def unparse_term(term):

    assert len(term) == 6
    assert all(x.isdigit() for x in term)
    year = term[:4]
    season = term[4:]
    return SEASONS_INV[season] + ' ' + year

######################################################################
# get cur term (either fall or spring, splits on july)

def get_cur_term():
    
    date = datetime.date.today()
    year = date.year
    month = date.month

    cur_term = str(year) + ('02' if month < 7 else '04')

    return cur_term

######################################################################
# class to represent students

class Student(object):

    def __init__(self, id, first, last, email_id, gradyear, honors, advisor):
        self.id = id
        self.first = first
        self.last = last
        self.email_id = email_id
        self.gradyear = gradyear
        self.honors = honors
        self.advisor = advisor
        self.enrollments = []
        self.abroad = set()
        self.decision = ''
        self.exceptions = set()

    def __str__(self):
        return '{} {}'.format(self.first, self.last)

    def has_taken(self, subject, course):
        for eold in self.enrollments:
            if eold.subject == subject and eold.course == course:
                return True
        return False

    def enroll(self, subject, course, term, credit, grade=None, source=None):

        #if subject not in ENGR_RELATED_SUBJECT:
        #    return

        course = course_normalize(course)
        credit = float(credit)

        if credit == 0:
            return
                
        enew = Enrollment(subject, course, term, credit, grade, source)
        
        for eold in self.enrollments:
            
            if eold.subject == subject and eold.course == course and (subject, course) not in REPEATABLE_COURSES:
                
                bad = False
                
                if eold.term == term:
                    if eold.credit == credit and eold.grade == grade:
                        return
                    else:
                        bad = True
                elif course == 'XXX':
                    bad = False
                    
                if bad:
                    print('duplicate course:', self.first, self.last, subject, course, file=sys.stderr)
                    print('  old: ', *eold, file=sys.stderr)
                    print('  new: ', *enew, file=sys.stderr)
                    sys.exit(1)
                    
            
                    
        self.enrollments.append(enew)


    def gpa(self, filter_subject=None):

        total = 0.0
        weight = 0.0
        
        for (subject,crse, term, cr, grade, _) in self.enrollments:

            if isinstance(filter_subject, str):
                include = filter_subject == subject
            elif isinstance(filter_subject, list) or isinstance(filter_subject, set):
                include = subject in filter_subject
            elif callable(filter_subject):
                include = filter_subject(subject)
            elif filter_subject is None:
                include = True
            else:
                raise RuntimeError('bad filter!')

            if not include:
                continue

            if grade in LETTER_GRADES:
                total += cr*LETTER_GRADES[grade]
                weight += cr

        if weight:
            return total / weight
        else:
            return None

    def senior_term(self):
        return str(self.gradyear) + '02'

######################################################################
# for checking requirements

# base class
class BaseRequirement(object):

    def check(self, enrollments):

        total = 0.0
        clist = []
        
        for e in enrollments:

            does_count = self.meets_requirement(e)

            if not does_count:
                key = (e.subject, e.course)
                if key in ENGR_XLISTS:
                    xsubj, xcourse = ENGR_XLISTS[key]
                    xlist = Enrollment(xsubj, xcourse, e.term, e.credit,
                                       e.grade, 'crosslist of ' + e.course)
                    does_count = self.meets_requirement(xlist)
            
            if does_count:
                total += e.credit
                clist.append(e)

        if self.threshold >= 0:
            ok = (total >= self.threshold)
        else:
            ok = (total <= -self.threshold)

        return ok, total, clist

# derived class to check multiple courses within single subject
class SingleSubjectRequirement(BaseRequirement):

    def __init__(self, name, subject, courses, threshold):
        self.name = name
        self.subject = subject
        self.courses = set(courses.split())
        self.threshold = threshold

    def meets_requirement(self, e):
        return e.subject == self.subject and e.course in self.courses

    def __repr__(self):
        return 'SingleSubjectRequirement({}, {}, {}, {})'.format(
            repr(self.name),
            repr(self.subject),
            repr(' '.join(self.courses)),
            repr(self.threshold))


# derived class to check courses across many subjects
class MultiSubjectRequirement(BaseRequirement):

    def __init__(self, name, subjects, filter_func, threshold):
        self.name = name
        self.subjects = set(subjects.split())
        self.filter_func = filter_func
        self.threshold = threshold

    def meets_requirement(self, e):
        return e.subject in self.subjects

    def check(self, enrollments):

        if self.filter_func:
            enrollments = self.filter_func(enrollments)

        return BaseRequirement.check(self, enrollments)

    def __repr__(self):
        return 'MultiSubjectRequirement({}, {}, {}, {})'.format(
            repr(self.name),
            repr(' '.join(self.subjects)),
            repr(self.filter_func),
            repr(self.threshold))

######################################################################
# for sorting students

def student_key(student):
    return student.last, student.first, student.id

######################################################################
# for sorting enrollments

def enrollment_key(enrollment):
    subject, course, term, cr, grade, source = enrollment
    course_match = re.match(r'^([0-9]+)([^0-9]*)$', course)
    if course_match:
        num = int(course_match.group(1))
        extra = course_match.group(2)
    else:
        num = 999
        extra = course
    return subject, num, extra, term, cr, grade, source

######################################################################
# for parsing Excel or CSV files 

def foreach_row(filenames, callback):

    for filename in filenames:

        if filename.lower().endswith('.xlsx'):

            print('reading ' + filename, file=sys.stderr)
            
            wb = load_workbook(filename)

            for sheet in wb.worksheets:

                header = None

                for row in sheet.rows:

                    if header is None:
                        header = [str(col.value) for col in row]
                    else:
                        lookup = dict(zip(header, [str(col.value) for col in row]))
                        callback(lookup)

        elif filename.lower().endswith('.csv'):

            print('reading ' + filename, file=sys.stderr)

            with open(filename) as f:
                reader = csv.reader(f)

                header = None

                for row in reader:

                    if header is None:
                        header = row
                    else:
                        lookup = dict(zip(header, row))
                        callback(lookup)
            
######################################################################
# for checking advisors

def check_engr_advisor(x):

    if not x:
        return False

    xl = x.lower()

    for lastname, email in ENGR_ADVISORS.items():
        if lastname in xl:
            return email

    return None
                                    
######################################################################
# make a list of tuples into a list of Requirement objects

def make_requirements(rlist):

    requirements = []

    for rtuple in rlist:

        assert len(rtuple) == 4
        name, subjects, courses, threshold = rtuple[:4]

        if callable(courses) or courses is None:
            r = MultiSubjectRequirement(name, subjects, courses, threshold)
        else:
            assert isinstance(courses, str)
            assert isinstance(subjects, str) and len(subjects) == 4
            r = SingleSubjectRequirement(name, subjects, courses, threshold)

        requirements.append(r)

    return requirements

######################################################################
# truncate .0 for printing

def maybe_int(x):
    ix = int(x)
    if x == ix:
        return ix
    else:
        return x

######################################################################
# make a printable course enrollment string

def printable_enrollment(e):
    if e.subject == 'ENGR':
        rval = 'E'+e.course
    else:
        rval = e.subject+' '+e.course
    if e.credit != 1:
        rval += ' ({} CR)'.format(maybe_int(e.credit))
    return rval

######################################################################
# make an ascii table that prints nicely

def tabulate(lines, pad_final):

    nfields = len(lines[0])
    maxlens = [0] * nfields

    for line in lines:
        for i, field in enumerate(line):
            maxlens[i] = max(maxlens[i], len(field))

    outputs = []

    for line in lines:
        loutputs = []
        for i, field in enumerate(line):
            if i+1 == len(line) and not pad_final:
                loutputs.append(field)
            else:
                loutputs.append(field + (' ' * (maxlens[i] - len(field))))
        outputs.append(' '.join(loutputs))

    return outputs

######################################################################

def verify_requirements(student, requirements):

    totals = defaultdict(float)
    lists = defaultdict(list)

    lines = []
    totals = []
    results = []

    all_ok = True

    for r in requirements:

        if r.name in student.exceptions:
            
            print('applying exception for student {} to rule {}'.format(
                student.email_id, r.name), file=sys.stderr)
            
            status = 'EXEMPT'
            ok = True
            clist = []
            benchmark = ''
            total = 0

        else:
        
            ok, total, clist = r.check(student.enrollments)

            all_ok = all_ok and ok

            thresh = maybe_int(r.threshold)
            total = maybe_int(total)

            if thresh > 0:
                benchmark = '{:>4} >= {:<4}'.format(total, thresh)
            else:
                benchmark = '{:>4} <= {:<4}'.format(total, -thresh)
            
            status = 'pass' if ok else 'FAIL'
            
        clist = ', '.join(map(printable_enrollment, sorted(clist)))
            
        lines.append((r.name, benchmark, status, clist))
        results.append((ok, total, clist))

    status = 'pass' if all_ok else 'FAIL'
    print(str(student) + ': ' + status)
    if not all_ok and student.abroad:
        print('*** NOTE: planned study abroad in {} ***'.format(' and '.join(sorted(student.abroad))))
    print()
    print('  ' + '\n  '.join(tabulate(lines, pad_final=False)))
    print()

    return all_ok, results
                    


######################################################################

def get_missing(requirements, results):

    numbad = 0

    outputs = []
    
    for r, (ok, total, clist) in zip(requirements, results):
        if not ok:
            if len(outputs):
                outputs.append('\n\n')
            if r.threshold > 1 and not re.search(r'[0-9] or [0-9]', r.name):
                of = '{} of '.format(maybe_int(r.threshold - total))
            else:
                of = ''
            outputs.append('  - {}{}'.format(of, r.name))
            if r.threshold > 1:
                outputs.append('\n    You have {} of {}'.format(total, maybe_int(r.threshold)))
                if (len(clist)):
                    outputs.append(': ' + clist)
                else:
                    outputs.append('.')
            numbad += 1

    output = ''.join(outputs)
    return output, numbad

######################################################################    

def exclude(exclusions, enrollments):
    rval = []
    for e in enrollments:
        if (e.subject, e.course) not in exclusions:
            rval.append(e)
    return rval

def count_if(threshold, courses, enrollments):
    
    total = 0.0

    rval = []

    for e in enrollments:
        if (e.subject, e.course) in courses:
            total += e.credit
            if ((threshold > 0 and total <= threshold) or
                (threshold < 0 and total > -threshold)):
                rval.append(e)
        else:
            rval.append(e)

    return rval

def filter_max1_dr(enrollments):

    dr_credit = 0.0

    rval = []

    for e in enrollments:

        if e.subject == 'ENGR' and e.course == '93':

            if dr_credit + e.credit > 1.0:

                allowed_credit = 1.0 - dr_credit

                e = Enrollment('ENGR', f'93 [orig {e.credit}]', e.term, 
                               allowed_credit, e.grade, e.source)

            dr_credit += e.credit
                
        rval.append(e)

    return rval

def at_most(max_count, courses, enrollments):
    return count_if(max_count, courses, enrollments)

def excess(threshold, courses, enrollments):
    return count_if(-threshold, courses, enrollments)
    
def filter_engr(enrollments):
    enrollments = exclude(ENGR_NONMAJOR_COURSES, enrollments)
    enrollments = at_most(1, TAUGHT_OUTSIDE_ENGR, enrollments)
    return enrollments

def filter_engr_nodr(enrollments):
    enrollments = exclude(ENGR_DR, enrollments)
    return filter_engr(enrollments)
    

def filter_engr_electives(enrollments):
    enrollments = exclude(ENGR_NONMAJOR_COURSES, enrollments)
    enrollments = exclude(ENGR_CORE_COURSES, enrollments)
    enrollments = at_most(1, TAUGHT_OUTSIDE_ENGR, enrollments)
    enrollments = excess(1, E15_OR_E19, enrollments)
    enrollments = filter_max1_dr(enrollments)
    return enrollments

def filter_engr_electives_nodr(enrollments):
    enrollments = exclude(ENGR_DR, enrollments)
    return filter_engr_electives(enrollments)

def filter_math_sci(enrollments):
    enrollments = exclude(MATH_SCI_EXCLUSIONS, enrollments)
    enrollments = at_most(3, NONLAB_COURSES, enrollments)
    return enrollments

######################################################################

MAJOR_REQUIREMENTS = make_requirements([
    ('E6',  'ENGR', '6', 1.0),
    ('E11', 'ENGR', '11 11A 11B', 1.0),
    ('E12', 'ENGR', '12', 1.0),
    ('E14', 'ENGR', '14', 1.0),
    ('E41', 'ENGR', '41', 1.0),
    ('E90', 'ENGR', '90', 1.0),
    ('E15 or E19', 'ENGR', '15 15A 15B 19', 1.0),
    ('5 ENGR electives', 'ENGR', filter_engr_electives, 5.0),
    ('MATH 33, 34, or 35', 'MATH', '33 34 35 3X 43 44 4X', 0.1),
    ('MATH 43 or 44', 'MATH', '43 44 4X', 0.1),
    ('2 PHYS credits', 'PHYS', filter_math_sci, 2.0),
    ('1 BIOL or CHEM credit', 'BIOL CHEM', filter_math_sci, 1.0),
    ('8 math/science credits', 'MATH STAT PHYS ASTR BIOL CHEM', filter_math_sci, 8.0)
])

MINOR_REQUIREMENTS = make_requirements([
    ('2+ core courses',  'ENGR', '6 11 11A 11B 12 14 15 15A 15B 19 41', 2.0),
    ('2+ electives excl DR', 'ENGR', filter_engr_electives_nodr, 2.0),
    ('5 ENGR credits', 'ENGR', filter_engr_nodr, 5.0),
    ('No E90', 'ENGR', '90', -0.01)
])

######################################################################
# parse CSV file generated by Major/Minor portal

def parse_students(filename):
    
    students = dict()

    majkey = 'Honors Major  Indicator'
    minkey = 'Honors Minor Indicator'

    def handle_student(lookup):
        
        if majkey in lookup:
            hkey = majkey
        else:
            hkey = minkey
            
        id = lookup['ID']
        first = lookup['Student First']
        last = lookup['Student Last']
        gradyear = lookup['Grad Year']
        email = lookup['Student Email']
        
        honors_discuss = lookup['Would like to Discuss Honors']
        honors_applicant = lookup[hkey]
        
        if honors_applicant == 'Yes':
            honors = 'yes'
        elif honors_discuss == 'Yes':
            honors = 'maybe'
        else:
            honors = 'no'
            

        premaj = lookup['Pre-Major Advisors']
        requested = lookup['Requested Advisor Name']
        sp = lookup['Sophomore Plan Advisor Name']

        if check_engr_advisor(sp):
            advisor, src = sp, '(assigned S.P.)'
        elif check_engr_advisor(premaj):
            advisor, src = premaj, '(pre-major)'
        elif check_engr_advisor(requested):
            advisor, src = requested, '(requested)'
        elif sp:
            advisor, src = sp, '(assigned S.P.)'
        else:
            advisor, src = premaj, '(pre-major)'

        if advisor.endswith('{'):
            advisor = advisor[:-1]
            
        advisor = advisor.replace('{', ' and ')

        email_id = email.replace('@swarthmore.edu', '')

        students[id] = Student(id, first, last, email_id,
                               int(gradyear), honors, advisor + ' ' + src)

    foreach_row([filename], handle_student)

    print(f'read {len(students)} students from {filename}')

    return students

######################################################################
# parse additional things for placements, study abroad, etc.

def parse_etc(elookup, filename):

    print('reading ' + filename, file=sys.stderr)

    lineno = 0

    course_expr = re.compile(r'^(\w+)\s+(\w\w\w\w)\s+(\w+)\s+([0-9][0-9][0-9][0-9][0-9][0-9])\s+([0-9]+\.?[0-9]*)$')
    rule_expr = re.compile(r'^(\w+)\s+\((.+)\)$')
    name_expr = re.compile(r'^(\w+)\s+"([^"]+)"\s+"([^"]+)"$')
    
    with open(filename, 'r') as istr:
        
        for line in istr:
            
            lineno += 1

            hpos = line.find('#')
            if hpos >= 0:
                line = line[:hpos]
            
            line = line.rstrip()
            
            if not line:
                continue

            m = course_expr.match(line)

            if m is not None:
                
                email_id = m.group(1)
                subject = m.group(2)
                course = m.group(3)
                term = m.group(4)
                credit = float(m.group(5))

                if email_id in elookup:
                    s = elookup[email_id]
                    s.enroll(subject, course, term, credit, grade=None,
                             source='{}:{}'.format(filename, lineno))
                    #print('enrolling {} in {}, {}, {}, {}, {}'.format(email_id, subject, course, term, credit, None), file=sys.stderr)

            else:

                m = rule_expr.match(line)

                if m is not None:


                    email_id = m.group(1)
                    rule_name = m.group(2)

                    if email_id in elookup:
                        s = elookup[email_id]
                        s.exceptions.add(rule_name)
                        #print('adding exception for {} to rule {}'.format(email_id, rule_name), file=sys.stderr)

                else:

                    m = name_expr.match(line)

                    if m is not None:

                        email_id = m.group(1)
                        first = m.group(2)
                        last = m.group(3)

                        if email_id in elookup:

                            s = elookup[email_id]
                            s.first = first
                            s.last = last

                            #print(f'setting name for {email_id} to {first} {last}', file=sys.stderr)

                    else:

                        raise RuntimeError('{}:{} parse error'.format(filename, lineno))
                

            '''
            
            email_id, subject, course, term, credit = line.split()
            assert len(subject) == 4
            credit = float(credit)
            if email_id in elookup:
                s = elookup[email_id]
                s.enroll(subject, course, term, credit, grade=None,
                         source='{}:{}'.format(filename, lineno))
                print('enrolling {} in {}, {}, {}, {}, {}'.format(email_id, subject, course, term, credit, None))
                if subject == 'ENGR' and 'XX' in course:
                    s.abroad.add(unparse_term(term))
            '''
