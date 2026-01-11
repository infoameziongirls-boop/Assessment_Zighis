from db import db
import os
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
from template_updater import AssessmentTemplateUpdater

import io
import csv
import random
import time
from functools import wraps
from werkzeug.utils import secure_filename
from datetime import datetime

from flask import Flask, render_template, redirect, url_for, flash, request, send_file, abort, jsonify, session
from flask_login import LoginManager, login_user, login_required, logout_user, current_user
from flask_bcrypt import Bcrypt
from flask_wtf import FlaskForm
from flask_wtf.file import FileField, FileAllowed
from wtforms import StringField, PasswordField, FloatField, SelectField, SelectMultipleField, DateField, TextAreaField, BooleanField
from wtforms.validators import InputRequired, Length, Optional, NumberRange

from config import config
from models import User, Student, Assessment, Setting, ActivityLog, Question, QuestionAttempt, Quiz, QuizAttempt, init_db
from excel_utils import ExcelTemplateHandler, ExcelBulkImporter, StudentBulkImporter, QuestionBulkImporter, create_default_template, create_student_import_template, create_question_import_template

def calculate_short_answer_score(student_answer, question):
    """Calculate score for short answer questions based on keywords"""
    if not question.keywords or not isinstance(question.keywords, list):
        # Fall back to exact match if no keywords
        return question.marks if student_answer.strip().lower() == question.correct_answer.strip().lower() else 0.0
    
    student_answer_lower = student_answer.lower().strip()
    keyword_matches = 0
    
    for keyword in question.keywords:
        keyword = keyword.lower().strip()
        if keyword in student_answer_lower:
            keyword_matches += 1
    
    total_keywords = len(question.keywords)
    if total_keywords == 0:
        return 0.0
    
    match_percentage = keyword_matches / total_keywords
    
    # 85% or more keywords = full marks
    if match_percentage >= 0.85:
        return question.marks
    # 50% or more keywords = half marks
    elif match_percentage >= 0.50:
        return question.marks * 0.5
    # Less than 50% = 0 marks
    else:
        return 0.0

# -------------------------
# Application Factory
# -------------------------
app = Flask(__name__, static_folder='public')

# Load configuration
env = os.environ.get('FLASK_ENV', 'development')
app.config.from_object(config[env])

# File upload configuration
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
app.config['TEMPLATE_FOLDER'] = os.path.join(os.path.dirname(__file__), 'templates_excel')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Create necessary folders
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['TEMPLATE_FOLDER'], exist_ok=True)

# -------------------------
# Extensions
# -------------------------
bcrypt = Bcrypt(app)
login_manager = LoginManager(app)
login_manager.login_view = "login"

# Initialize database
init_db(app, bcrypt)

# -------------------------
# Activity Logging
# -------------------------
def log_activity(user, action, details=None):
    """Log user activity for auditing purposes"""
    if not user or not user.is_authenticated:
        return
    try:
        ip_address = request.remote_addr if request else None
        log_entry = ActivityLog(
            user_id=user.id,
            action=action,
            details=details,
            ip_address=ip_address
        )
        db.session.add(log_entry)
        db.session.commit()
    except Exception as e:
        # Log to console if database logging fails
        print(f"Failed to log activity: {e}")

# -------------------------
# Forms - FIXED: Remove duplicate definitions
# -------------------------
# Forms - FIXED: Remove duplicate definitions
# -------------------------

class StudentLoginForm(FlaskForm):
    first_name = StringField("First Name", validators=[InputRequired(), Length(min=1, max=120)])
    student_number = StringField("Student Number", validators=[InputRequired(), Length(min=1, max=50)])

class LoginForm(FlaskForm):
    username = StringField("Username", validators=[InputRequired(), Length(min=3, max=80)])
    password = PasswordField("Password", validators=[InputRequired(), Length(min=4)])

class UserForm(FlaskForm):
    username = StringField("Username", validators=[InputRequired(), Length(min=3)])
    password = PasswordField("Password", validators=[InputRequired(), Length(min=6)])
    role = SelectField("Role", choices=app.config['USER_ROLES'])
    subject = SelectField("Subject (for teachers)", choices=[("", "-- Not Applicable --")] + app.config['LEARNING_AREAS'], validators=[Optional()])
    classes = SelectMultipleField("Classes (for teachers)", choices=app.config['CLASS_LEVELS'], validators=[Optional()])

class EditUserForm(FlaskForm):
    role = SelectField("Role", choices=app.config['USER_ROLES'])
    subject = SelectField("Subject (for teachers)", choices=[("", "-- Not Applicable --")] + app.config['LEARNING_AREAS'], validators=[Optional()])
    classes = SelectMultipleField("Classes (for teachers)", choices=app.config['CLASS_LEVELS'], validators=[Optional()])

class PasswordResetForm(FlaskForm):
    password = PasswordField("New Password", validators=[InputRequired(), Length(min=6)])

# StudentForm - ONE DEFINITION ONLY
class StudentForm(FlaskForm):
    student_number = StringField("Student Number", validators=[InputRequired(), Length(min=1, max=50)])
    first_name = StringField("First name", validators=[InputRequired()])
    last_name = StringField("Last name", validators=[InputRequired()])
    middle_name = StringField("Middle name", validators=[Optional()])
    class_name = SelectField("Class", choices=[("", "-- Select Class --")] + app.config['CLASS_LEVELS'], validators=[Optional()])
    study_area = SelectField("Study/Learning Area", choices=[("", "-- Select Study Area --")] + app.config['STUDY_AREAS'], validators=[Optional()])
    # Removed duplicate learning_area field since it's not in your models

class AssessmentForm(FlaskForm):
    student_number = StringField("Student Number", validators=[Optional()])
    student_name = SelectField("Student Name", choices=[], validators=[InputRequired()])
    reference_number = StringField("Reference Number", validators=[Optional()])
    category = SelectField("Category", choices=app.config['ASSESSMENT_CATEGORIES'], validators=[InputRequired()])
    subject = SelectField("Subject", choices=[("", "-- Select Subject --")] + app.config['LEARNING_AREAS'], validators=[InputRequired()])
    class_name = SelectField("Class", choices=[("", "-- Select Class --")] + app.config['CLASS_LEVELS'], validators=[Optional()])
    score = FloatField("Score", validators=[InputRequired(), NumberRange(min=0)])
    max_score = SelectField("Max Score", choices=[(50, '50'), (100, '100')], validators=[InputRequired()], default=100)
    term = SelectField("Term", choices=app.config['TERMS'], validators=[InputRequired()])
    academic_year = StringField("Academic Year", validators=[Optional()])
    session = StringField("Session", validators=[Optional()])
    assessor = StringField("Assessor", validators=[Optional()])
    comments = TextAreaField("Comments", validators=[Optional()])

class TeacherAssignmentForm(FlaskForm):
    subject = SelectField("Subject", choices=[("", "-- Select Subject --")] + app.config['LEARNING_AREAS'], validators=[InputRequired()])
    classes = SelectMultipleField("Classes", choices=app.config['CLASS_LEVELS'], validators=[Optional()])

class AssessmentFilterForm(FlaskForm):
    subject = SelectField("Subject", choices=[("", "-- All Subjects --")] + app.config['LEARNING_AREAS'], validators=[Optional()])
    class_name = SelectField("Class", choices=[("", "-- All Classes --")] + app.config['CLASS_LEVELS'], validators=[Optional()])
    category = SelectField("Category", choices=[("", "-- All Categories --")] + app.config['ASSESSMENT_CATEGORIES'], validators=[Optional()])

class BulkImportForm(FlaskForm):
    excel_file = FileField("Excel File", validators=[
        InputRequired(),
        FileAllowed(['xlsx', 'xls'], 'Excel files only!')
    ])

class StudentBulkImportForm(FlaskForm):
    excel_file = FileField("Excel File", validators=[
        InputRequired(),
        FileAllowed(['xlsx', 'xls'], 'Excel files only!')
    ])

class QuestionBulkImportForm(FlaskForm):
    excel_file = FileField("Excel File", validators=[
        InputRequired(),
        FileAllowed(['xlsx', 'xls'], 'Excel files only!')
    ])

class SettingsForm(FlaskForm):
    current_term = SelectField("Current Term", choices=app.config['TERMS'], validators=[InputRequired()])
    current_academic_year = StringField("Current Academic Year", validators=[InputRequired()])
    current_session = StringField("Current Session", validators=[InputRequired()])
    assessment_active = BooleanField("Assessment Entry Active", default=True)


class QuestionForm(FlaskForm):
    question_text = TextAreaField("Question Text", validators=[InputRequired(), Length(min=10, max=1000)])
    question_type = SelectField("Question Type", choices=[
        ('mcq', 'Multiple Choice Question'),
        ('true_false', 'True/False'),
        ('short_answer', 'Short Answer')
    ], validators=[InputRequired()])
    options = TextAreaField("Options (for MCQ only)", validators=[Optional()], 
                          render_kw={"placeholder": "Enter options one per line (A, B, C, D)"})
    correct_answer = StringField("Correct Answer", validators=[InputRequired()], 
                               render_kw={"placeholder": "For MCQ: A, B, C, or D. For True/False: True or False. For Short Answer: the expected answer"})
    marks = FloatField("Marks", validators=[InputRequired(), NumberRange(min=0.1, max=100)], default=1.0)
    keywords = TextAreaField("Keywords (for Short Answer only)", validators=[Optional()], 
                           render_kw={"placeholder": "Enter keywords one per line for flexible marking"})
    difficulty = SelectField("Difficulty", choices=[
        ('easy', 'Easy'),
        ('medium', 'Medium'),
        ('hard', 'Hard')
    ], validators=[InputRequired()])
    explanation = TextAreaField("Explanation (Optional)", validators=[Optional(), Length(max=500)])


class QuizForm(FlaskForm):
    title = StringField("Quiz Title", validators=[InputRequired(), Length(min=3, max=200)])
    subject = SelectField("Subject", validators=[InputRequired()])
    description = TextAreaField("Description", validators=[Optional(), Length(max=500)])
    questions = SelectMultipleField("Questions", validators=[InputRequired()], 
                                   render_kw={"size": 10})
    time_limit = FloatField("Time Limit (minutes)", validators=[Optional(), NumberRange(min=1, max=180)])
    is_active = BooleanField("Active", default=True)


# -------------------------
# Login manager
# -------------------------
@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# -------------------------
# Decorators
# -------------------------
def admin_required(f):
    @wraps(f)
    def wrapped(*args, **kwargs):
        if not current_user.is_authenticated:
            return login_manager.unauthorized()
        if not current_user.is_admin():
            abort(403)
        return f(*args, **kwargs)
    return wrapped

def teacher_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not current_user.is_authenticated or not current_user.is_teacher():
            abort(403)
        return f(*args, **kwargs)
    return decorated_function

def student_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not current_user.is_authenticated or not current_user.is_student():
            abort(403)
        return f(*args, **kwargs)
    return decorated_function

def student_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not current_user.is_authenticated or not current_user.is_student():
            abort(403)
        return f(*args, **kwargs)
    return decorated_function

# -------------------------
# Authentication Routes
# -------------------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    
    form = LoginForm()
    if form.validate_on_submit():
        user = User.query.filter_by(username=form.username.data.strip()).first()
        if user and user.check_password(form.password.data, bcrypt):
            login_user(user)
            log_activity(user, "login", f"User {user.username} logged in")
            flash("Logged in successfully", "success")
            next_page = request.args.get("next")
            return redirect(next_page or url_for("dashboard"))
        flash("Invalid credentials", "danger")
    return render_template("login.html", form=form)

@app.route("/logout")
@login_required
def logout():
    logout_user()
    flash("Logged out successfully", "info")
    return redirect(url_for("login"))

@app.route("/student/login", methods=["GET", "POST"])
def student_login():
    """Student login using first name and student number"""
    
    form = StudentLoginForm()
    if form.validate_on_submit():
        first_name = form.first_name.data.strip()
        student_number = form.student_number.data.strip()
        
        # Find student by first name and student number
        student = Student.query.filter_by(first_name=first_name, student_number=student_number).first()
        
        if student:
            # Check if there's a user account for this student
            user = User.query.filter_by(username=student_number).first()
            if not user:
                # Create a student user account if it doesn't exist
                password = app.config['DEFAULT_STUDENT_PASSWORD']
                pw_hash = bcrypt.generate_password_hash(password).decode("utf-8")
                user = User(
                    username=student_number,
                    password_hash=pw_hash,
                    role="student"
                )
                db.session.add(user)
                db.session.commit()
            
            login_user(user)
            log_activity(user, "student_login", f"Student {student.full_name()} ({student.student_number}) logged in")
            flash("Student login successful", "success")
            return redirect(url_for("student_dashboard"))
        else:
            flash("Invalid first name or student number. Please check your details.", "danger")
    
    return render_template("student_login.html", form=form)

@app.route("/student/logout")
@login_required
def student_logout():
    """Logout student"""
    logout_user()
    flash("Logged out successfully", "info")
    return redirect(url_for("student_login"))

# -------------------------
# Dashboard Routes
# -------------------------
@app.route("/")
@login_required
def dashboard():
    if hasattr(current_user, 'is_student') and current_user.is_student():
        return redirect(url_for("student_dashboard"))
    
    # Teacher/Admin dashboard
    student_count = Student.query.count()
    assessment_count = Assessment.query.filter_by(archived=False).count()
    users_count = User.query.count()
    
    # For teachers, show only their assessments and student summaries for their subject
    teacher_student_summaries = None
    if hasattr(current_user, 'is_teacher') and current_user.is_teacher():
        recent = Assessment.query.filter_by(teacher_id=current_user.id, archived=False)\
            .order_by(Assessment.date_recorded.desc()).limit(8).all()
        
        # Get student summaries for teacher's subject
        if current_user.subject:
            # Get all students who have assessments in teacher's subject by this teacher
            students_with_assessments = db.session.query(Student).join(Assessment)\
                .filter(Assessment.teacher_id == current_user.id)\
                .filter(Assessment.subject == current_user.subject)\
                .filter(Assessment.archived == False)\
                .distinct().all()
            
            teacher_student_summaries = []
            for student in students_with_assessments:
                final_grade = student.calculate_final_grade(subject=current_user.subject, teacher_id=current_user.id)
                assessment_count = len([a for a in student.assessments 
                                      if a.teacher_id == current_user.id and a.subject == current_user.subject and not a.archived])
                
                # Calculate GPA and grade
                gpa = None
                grade = None
                if final_grade is not None:
                    if final_grade >= 80:
                        gpa = 4.0
                        grade = 'A1'
                    elif final_grade >= 70:
                        gpa = 3.5
                        grade = 'B2'
                    elif final_grade >= 65:
                        gpa = 3.0
                        grade = 'B3'
                    elif final_grade >= 60:
                        gpa = 2.5
                        grade = 'C4'
                    elif final_grade >= 55:
                        gpa = 2.0
                        grade = 'C5'
                    elif final_grade >= 50:
                        gpa = 1.5
                        grade = 'C6'
                    elif final_grade >= 45:
                        gpa = 1.0
                        grade = 'D7'
                    elif final_grade >= 40:
                        gpa = 0.5
                        grade = 'E8'
                    else:
                        gpa = 0.0
                        grade = 'F9'
                
                teacher_student_summaries.append({
                    'student': student,
                    'final_grade': final_grade,
                    'gpa': gpa,
                    'grade': grade,
                    'assessment_count': assessment_count
                })
            
            # Sort by final grade descending
            teacher_student_summaries.sort(key=lambda x: x['final_grade'] or 0, reverse=True)
    else:
        recent = Assessment.query.filter_by(archived=False)\
            .order_by(Assessment.date_recorded.desc()).limit(8).all()
    
    return render_template(
        "dashboard.html",
        student_count=student_count,
        assessment_count=assessment_count,
        users_count=users_count,
        recent=recent,
        teacher_student_summaries=teacher_student_summaries
    )

@app.route("/student/dashboard")
@login_required
@student_required
def student_dashboard():
    """Student dashboard showing their assessments"""
    # Get student info using student number (which is the username)
    student = Student.query.filter_by(student_number=current_user.username).first()
    if not student:
        flash("Student record not found", "danger")
        return redirect(url_for("student_logout"))
    
    # Get filter parameters
    subject = request.args.get("subject", "")
    class_filter = request.args.get("class", "")
    category = request.args.get("category", "")
    
    # Get assessments
    query = Assessment.query.filter_by(student_id=student.id, archived=False)
    
    if subject:
        query = query.filter_by(subject=subject)
    if class_filter:
        query = query.filter_by(class_name=class_filter)
    if category:
        query = query.filter_by(category=category)
    
    assessments = query.order_by(Assessment.date_recorded.desc()).all()
    
    # Get unique subjects and classes for filter dropdowns
    subjects = sorted(set([a.subject for a in student.assessments if a.subject]))
    classes = sorted(set([a.class_name for a in student.assessments if a.class_name]))
    categories = sorted(set([a.category for a in student.assessments if a.category]))
    
    # Get quiz attempts for this student
    quiz_attempts = QuizAttempt.query.filter_by(student_id=student.id).order_by(QuizAttempt.completed_at.desc()).all()
    
    # Get quiz details for each attempt
    quiz_details = {}
    for attempt in quiz_attempts:
        quiz = Quiz.query.get(attempt.quiz_id)
        if quiz:
            quiz_details[attempt.id] = quiz
    
    # Get results grouped by teacher/subject instead of aggregated totals
    # Respect the subject filter - only show filtered subjects or all if no filter
    assessments_for_results = student.assessments
    if subject:
        assessments_for_results = [a for a in assessments_for_results if a.subject == subject]
    
    teacher_subjects = {}
    
    # Group assessments by teacher and subject
    for assessment in assessments_for_results:
        if assessment.archived:
            continue
            
        teacher_id = assessment.teacher_id
        subject_name = assessment.subject
        
        if teacher_id not in teacher_subjects:
            teacher_subjects[teacher_id] = {}
        
        if subject_name not in teacher_subjects[teacher_id]:
            teacher_subjects[teacher_id][subject_name] = []
        
        teacher_subjects[teacher_id][subject_name].append(assessment)
    
    # Calculate results per teacher/subject
    teacher_results = {}
    for teacher_id, subjects_data in teacher_subjects.items():
        teacher = User.query.get(teacher_id)
        teacher_name = teacher.username if teacher else f"Teacher {teacher_id}"
        
        teacher_results[teacher_name] = {}
        
        for subject_name, assessments_list in subjects_data.items():
            # Calculate final grade for this teacher/subject combination
            final_percent = student.calculate_final_grade(subject=subject_name, teacher_id=teacher_id)
            
            # Calculate GPA and grade based on this final percentage
            gpa = None
            grade = None
            if final_percent is not None:
                if final_percent >= 80:
                    gpa = 4.0
                    grade = 'A1'
                elif final_percent >= 70:
                    gpa = 3.5
                    grade = 'B2'
                elif final_percent >= 65:
                    gpa = 3.0
                    grade = 'B3'
                elif final_percent >= 60:
                    gpa = 2.5
                    grade = 'C4'
                elif final_percent >= 55:
                    gpa = 2.0
                    grade = 'C5'
                elif final_percent >= 50:
                    gpa = 1.5
                    grade = 'C6'
                elif final_percent >= 45:
                    gpa = 1.0
                    grade = 'D7'
                elif final_percent >= 40:
                    gpa = 0.5
                    grade = 'E8'
                else:
                    gpa = 0.0
                    grade = 'F9'
            
            teacher_results[teacher_name][subject_name] = {
                'final_percent': final_percent,
                'gpa': gpa,
                'grade': grade,
                'assessments': assessments_list
            }
    
    # Overall summary (keeping for backward compatibility, but not displayed prominently)
    summary = student.get_assessment_summary()
    final_percent = student.calculate_final_grade()
    gpa_grade = student.get_gpa_and_grade()
    
    # Calculate filtered summary
    filtered_assessments = assessments  # assessments is already filtered
    if subject:
        # Specific subject selected - show final grade for that subject
        final_percent_filtered = student.calculate_final_grade(subject=subject, teacher_id=current_user.id if current_user.is_teacher() else None)
        average_score = final_percent_filtered if final_percent_filtered is not None else 0.0
        # GPA and grade based on this subject's final grade
        if final_percent_filtered is not None:
            if final_percent_filtered >= 80:
                filtered_gpa = 4.0
                filtered_grade = 'A1'
            elif final_percent_filtered >= 70:
                filtered_gpa = 3.5
                filtered_grade = 'B2'
            elif final_percent_filtered >= 65:
                filtered_gpa = 3.0
                filtered_grade = 'B3'
            elif final_percent_filtered >= 60:
                filtered_gpa = 2.5
                filtered_grade = 'C4'
            elif final_percent_filtered >= 55:
                filtered_gpa = 2.0
                filtered_grade = 'C5'
            elif final_percent_filtered >= 50:
                filtered_gpa = 1.5
                filtered_grade = 'C6'
            elif final_percent_filtered >= 45:
                filtered_gpa = 1.0
                filtered_grade = 'D7'
            elif final_percent_filtered >= 40:
                filtered_gpa = 0.5
                filtered_grade = 'E8'
            else:
                filtered_gpa = 0.0
                filtered_grade = 'F9'
        else:
            filtered_gpa = 0.0
            filtered_grade = 'N/A'
    else:
        # All subjects - calculate average score from all assessments
        if filtered_assessments:
            total_marks = sum(a.max_score for a in filtered_assessments if a.max_score)
            obtained_marks = sum(a.score for a in filtered_assessments if a.score)
            average_score = (obtained_marks / total_marks * 100) if total_marks > 0 else 0.0
            
            # GPA and grade based on average score
            if average_score >= 80:
                filtered_gpa = 4.0
                filtered_grade = 'A1'
            elif average_score >= 70:
                filtered_gpa = 3.5
                filtered_grade = 'B2'
            elif average_score >= 65:
                filtered_gpa = 3.0
                filtered_grade = 'B3'
            elif average_score >= 60:
                filtered_gpa = 2.5
                filtered_grade = 'C4'
            elif average_score >= 55:
                filtered_gpa = 2.0
                filtered_grade = 'C5'
            elif average_score >= 50:
                filtered_gpa = 1.5
                filtered_grade = 'C6'
            elif average_score >= 45:
                filtered_gpa = 1.0
                filtered_grade = 'D7'
            elif average_score >= 40:
                filtered_gpa = 0.5
                filtered_grade = 'E8'
            else:
                filtered_gpa = 0.0
                filtered_grade = 'F9'
        else:
            average_score = 0.0
            filtered_gpa = 0.0
            filtered_grade = 'N/A'
    
    # Calculate comment based on GPA
    def get_comment(gpa_str):
        try:
            gpa = float(gpa_str)
            if gpa == 4.0: return "Excellent"
            elif gpa == 3.5: return "Very Good"
            elif gpa == 3.0: return "Good"
            elif gpa == 2.5: return "Average"
            elif gpa == 2.0: return "Below Average"
            elif gpa == 1.5: return "Credit"
            elif gpa == 1.0: return "Satisfactory"
            elif gpa == 0.5: return "Pass"
            else: return "Fail"
        except (ValueError, TypeError):
            return None
    
    comment = get_comment(gpa_grade['gpa']) if gpa_grade['gpa'] != 'N/A' else None
    
    return render_template(
        "student_dashboard.html",
        student=student,
        assessments=assessments,
        teacher_results=teacher_results,
        summary=summary,
        final_percent=final_percent,
        gpa_grade=gpa_grade,
        comment=comment,
        subjects=subjects,
        classes=classes,
        categories=categories,
        selected_subject=subject,
        selected_class=class_filter,
        selected_category=category,
        average_score=average_score,
        filtered_gpa=filtered_gpa,
        filtered_grade=filtered_grade,
        category_labels=app.config['CATEGORY_LABELS'],
        quiz_attempts=quiz_attempts,
        quiz_details=quiz_details
    )

# -------------------------
# Student Management Routes
# -------------------------
@app.route("/students")
@login_required
def students():
    search = request.args.get("search", "").strip()
    
    if search:
        students = Student.query.filter(
            (Student.student_number.ilike(f"%{search}%")) |
            (Student.first_name.ilike(f"%{search}%")) |
            (Student.last_name.ilike(f"%{search}%")) |
            (Student.reference_number.ilike(f"%{search}%"))
        ).order_by(Student.last_name, Student.first_name).all()
    else:
        students = Student.query.order_by(Student.last_name, Student.first_name).all()
    
    return render_template("students.html", students=students)

@app.route("/students/new", methods=["GET", "POST"])
@login_required
def student_new():
    # Only teachers and admins can create students
    if not (current_user.is_teacher() or current_user.is_admin()):
        abort(403)
        
    form = StudentForm()
    if form.validate_on_submit():
        exists = Student.query.filter_by(student_number=form.student_number.data.strip()).first()
        if exists:
            flash("Student number already exists", "warning")
        else:
            # Generate reference number (STU + random 6 digits)
            reference_number = f"STU{random.randint(100000, 999999)}"
            
            student = Student(
                student_number=form.student_number.data.strip(),
                first_name=form.first_name.data.strip(),
                last_name=form.last_name.data.strip(),
                middle_name=form.middle_name.data.strip() if form.middle_name.data else None,
                class_name=form.class_name.data if form.class_name.data else None,
                study_area=form.study_area.data if form.study_area.data else None,
                reference_number=reference_number
            )
            db.session.add(student)
            db.session.commit()
            
            log_activity(current_user, "create_student", f"Created student {student.full_name()} ({student.student_number})")
            flash(f"Student {student.full_name()} added successfully. Reference Number: {reference_number}", "success")
            return redirect(url_for("students"))
    
    return render_template("student_form.html", form=form, student=None)

@app.route("/students/<int:student_id>/edit", methods=["GET", "POST"])
@login_required
def student_edit(student_id):
    # Only teachers and admins can edit students
    if not (current_user.is_teacher() or current_user.is_admin()):
        abort(403)
        
    student = Student.query.get_or_404(student_id)
    form = StudentForm(obj=student)
    
    if form.validate_on_submit():
        student.student_number = form.student_number.data.strip()
        student.first_name = form.first_name.data.strip()
        student.last_name = form.last_name.data.strip()
        student.middle_name = form.middle_name.data.strip() if form.middle_name.data else None
        student.class_name = form.class_name.data if form.class_name.data else None
        student.study_area = form.study_area.data if form.study_area.data else None
        db.session.commit()
        log_activity(current_user, "edit_student", f"Edited student {student.full_name()} ({student.student_number})")
        flash(f"Student {student.full_name()} updated successfully", "success")
        return redirect(url_for("students"))
    
    return render_template("student_form.html", form=form, student=student)

@app.route("/students/<int:student_id>/delete", methods=["POST"])
@login_required
@admin_required
def student_delete(student_id):
    student = Student.query.get_or_404(student_id)
    student_name = student.full_name()
    db.session.delete(student)
    db.session.commit()
    log_activity(current_user, "delete_student", f"Deleted student {student_name} ({student.student_number})")
    flash(f"Student {student_name} deleted successfully", "info")
    return redirect(url_for("students"))

@app.route("/students/<int:student_id>")
@login_required
def student_view(student_id):
    student = Student.query.get_or_404(student_id)
    
    subject = request.args.get('subject')
    
    # Filter assessments by subject if specified
    if subject:
        assessments = [a for a in student.assessments if a.subject == subject]
    else:
        # Filter assessments by subject/class if teacher, and only show their own assessments
        if hasattr(current_user, 'is_teacher') and current_user.is_teacher():
            # Teachers only see their own assessments for this student
            assessments = [a for a in student.assessments if a.teacher_id == current_user.id]
            # Further filter by teacher's subject if they have one assigned
            if current_user.subject:
                assessments = [a for a in assessments if a.subject == current_user.subject]
        else:
            # Admins see all assessments
            assessments = student.assessments
    
    # Get assessment summary and final grade - but only for current teacher's assessments
    summary = student.get_assessment_summary(subject, teacher_id=current_user.id if current_user.is_teacher() else None)
    final_percent = student.calculate_final_grade(subject=subject, teacher_id=current_user.id if current_user.is_teacher() else None)
    
    # Get all subjects for this student - filtered by current teacher for teachers
    if current_user.is_teacher():
        all_subjects = sorted(set(a.subject for a in student.assessments if a.teacher_id == current_user.id))
    else:
        all_subjects = sorted(set(a.subject for a in student.assessments))
    
    # Calculate letter grade and GPA
    def get_letter_grade(percent):
        if percent >= 80: return 'A1'
        elif percent >= 70: return 'B2'
        elif percent >= 65: return 'B3'
        elif percent >= 60: return 'C4'
        elif percent >= 55: return 'C5'
        elif percent >= 50: return 'C6'
        elif percent >= 45: return 'D7'
        elif percent >= 40: return 'E8'
        else: return 'F9'
    
    def get_gpa(percent):
        if percent >= 80: return 4.0
        elif percent >= 70: return 3.5
        elif percent >= 65: return 3.0
        elif percent >= 60: return 2.5
        elif percent >= 55: return 2.0
        elif percent >= 50: return 1.5
        elif percent >= 45: return 1.0
        elif percent >= 40: return 0.5
        else: return 0.0
    
    letter_grade = get_letter_grade(final_percent) if final_percent is not None else None
    gpa = get_gpa(final_percent) if final_percent is not None else None
    
    def get_comment(gpa):
        if gpa == 4.0: return "Excellent"
        elif gpa == 3.5: return "Very Good"
        elif gpa == 3.0: return "Good"
        elif gpa == 2.5: return "Average"
        elif gpa == 2.0: return "Below Average"
        elif gpa == 1.5: return "Credit"
        elif gpa == 1.0: return "Satisfactory"
        elif gpa == 0.5: return "Pass"
        else: return "Fail"
    
    comment = get_comment(gpa) if gpa is not None else None
    
    # For admins, also prepare results grouped by teacher/subject
    teacher_results = None
    if current_user.is_admin():
        teacher_subjects = {}
        
        # Group assessments by teacher and subject
        for assessment in assessments:
            teacher_id = assessment.teacher_id
            subject_name = assessment.subject
            
            if teacher_id not in teacher_subjects:
                teacher_subjects[teacher_id] = {}
            
            if subject_name not in teacher_subjects[teacher_id]:
                teacher_subjects[teacher_id][subject_name] = []
            
            teacher_subjects[teacher_id][subject_name].append(assessment)
        
        # Calculate results per teacher/subject
        teacher_results = {}
        for teacher_id, subjects_data in teacher_subjects.items():
            teacher = User.query.get(teacher_id)
            teacher_name = teacher.username if teacher else f"Teacher {teacher_id}"
            
            teacher_results[teacher_name] = {}
            
            for subject_name, assessments_list in subjects_data.items():
                # Calculate final grade for this teacher/subject combination
                final_percent_teacher = student.calculate_final_grade(subject=subject_name, teacher_id=teacher_id)
                gpa_teacher = get_gpa(final_percent_teacher) if final_percent_teacher is not None else None
                
                teacher_results[teacher_name][subject_name] = {
                    'final_percent': final_percent_teacher,
                    'gpa': gpa_teacher,
                    'grade': get_letter_grade(final_percent_teacher) if final_percent_teacher is not None else None,
                    'assessments': assessments_list
                }
    
    return render_template(
        "student_view.html",
        student=student,
        assessments=assessments,
        teacher_results=teacher_results,
        summary=summary,
        final_percent=final_percent,
        letter_grade=letter_grade,
        gpa=gpa,
        comment=comment,
        subject=subject,
        all_subjects=all_subjects,
        category_labels=app.config['CATEGORY_LABELS'],
        study_areas_dict=dict(app.config['STUDY_AREAS'])
    )

@app.route("/students/bulk-import", methods=["GET", "POST"])
@login_required
def student_bulk_import():
    """Bulk import students from Excel file"""
    # Only teachers and admins can bulk import students
    if not (current_user.is_teacher() or current_user.is_admin()):
        abort(403)
        
    form = StudentBulkImportForm()
    
    if form.validate_on_submit():
        file = form.excel_file.data
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        # Save uploaded file
        file.save(filepath)
        
        try:
            # Import students
            importer = StudentBulkImporter(filepath)
            students_data = importer.import_students()
            
            # Process and save students
            success_count = 0
            error_count = 0
            errors = []
            
            for data in students_data:
                try:
                    # Check if student already exists
                    exists = Student.query.filter_by(student_number=data['student_number']).first()
                    if exists:
                        errors.append(f"Student {data['student_number']} already exists")
                        error_count += 1
                        continue
                    
                    # Generate reference number
                    reference_number = f"STU{random.randint(100000, 999999)}"
                    
                    student = Student(
                        student_number=data['student_number'],
                        first_name=data['first_name'],
                        last_name=data['last_name'],
                        middle_name=data.get('middle_name'),
                        class_name=data.get('class_name'),
                        study_area=data.get('study_area'),
                        reference_number=reference_number
                    )
                    db.session.add(student)
                    success_count += 1
                    
                except Exception as e:
                    errors.append(f"Error importing {data.get('student_number', 'unknown')}: {str(e)}")
                    error_count += 1
            
            db.session.commit()
            
            # Clean up uploaded file
            os.remove(filepath)
            
            flash(f"Bulk import completed. {success_count} students imported successfully. {error_count} errors.", "success")
            if errors:
                flash("Errors: " + "; ".join(errors[:5]), "warning")  # Show first 5 errors
            
            return redirect(url_for("students"))
            
        except Exception as e:
            flash(f"Error importing file: {str(e)}", "danger")
            return redirect(url_for("student_bulk_import"))
    
    return render_template("student_bulk_import.html", form=form)


@app.route("/teacher/questions/bulk_import", methods=["GET", "POST"])
@login_required
@teacher_required
def bulk_import_questions():
    """Bulk import questions from Excel file"""
    form = QuestionBulkImportForm()
    
    if form.validate_on_submit():
        file = form.excel_file.data
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        # Save uploaded file
        file.save(filepath)
        time.sleep(0.1)  # Allow file handle to be released
        
        try:
            # Import questions
            importer = QuestionBulkImporter(filepath)
            questions_data = importer.import_questions()
            
            # Process and save questions
            success_count = 0
            error_count = 0
            errors = []
            
            for data in questions_data:
                try:
                    # Create question
                    question = Question(
                        subject=current_user.subject,
                        question_text=data['question_text'],
                        question_type=data['question_type'],
                        options=data['options'],
                        correct_answer=data['correct_answer'],
                        difficulty=data['difficulty'],
                        explanation=data['explanation'],
                        created_by=current_user.id
                    )
                    db.session.add(question)
                    success_count += 1
                    
                except Exception as e:
                    errors.append(f"Error importing question '{data.get('question_text', 'unknown')[:50]}...': {str(e)}")
                    error_count += 1
            
            db.session.commit()
            
            # Clean up uploaded file
            os.remove(filepath)
            
            flash(f"Bulk import completed. {success_count} questions imported successfully. {error_count} errors.", "success")
            if errors:
                flash("Errors: " + "; ".join(errors[:5]), "warning")  # Show first 5 errors
            
            return redirect(url_for("teacher_question_bank"))
            
        except Exception as e:
            if "[WinError 32]" in str(e):
                flash("uploaded successful", "success")
            else:
                flash(f"Error importing file: {str(e)}", "danger")
            return redirect(url_for("bulk_import_questions"))
    
    return render_template("question_bulk_import.html", form=form)


# -------------------------
# Assessment Routes
# -------------------------
@app.route("/assessments")
@login_required
def assessments_list():
    page = request.args.get("page", 1, type=int)
    subject = request.args.get("subject", "")
    class_name = request.args.get("class", "")
    category = request.args.get("category", "")
    
    per_page = app.config['ASSESSMENTS_PER_PAGE']
    
    # Build query based on user role and filters
    if hasattr(current_user, 'is_teacher') and current_user.is_teacher():
        query = Assessment.query.filter_by(teacher_id=current_user.id, archived=False)
    else:
        query = Assessment.query.filter_by(archived=False)
    
    if subject:
        query = query.filter_by(subject=subject)
    if class_name:
        query = query.filter_by(class_name=class_name)
    if category:
        query = query.filter_by(category=category)
    
    # Get all filtered assessments for student performance calculation
    all_filtered_assessments = query.all()
    
    # Calculate overall metrics for display
    total_assessments = len(all_filtered_assessments)
    avg_score = 0.0
    avg_gpa = 0.0
    unique_students_affected = len(set(a.student_id for a in all_filtered_assessments)) if all_filtered_assessments else 0
    
    if all_filtered_assessments:
        # Calculate average score
        total_score = sum(a.score for a in all_filtered_assessments)
        total_max = sum(a.max_score for a in all_filtered_assessments)
        if total_max > 0:
            avg_score = (total_score / total_max) * 100
        
        # Calculate average GPA based on individual assessment grades
        total_gpa = 0.0
        assessment_count = 0
        for assessment in all_filtered_assessments:
            if assessment.max_score > 0:
                percent = (assessment.score / assessment.max_score) * 100
                if percent >= 80:
                    gpa = 4.0
                elif percent >= 70:
                    gpa = 3.5
                elif percent >= 65:
                    gpa = 3.0
                elif percent >= 60:
                    gpa = 2.5
                elif percent >= 55:
                    gpa = 2.0
                elif percent >= 50:
                    gpa = 1.5
                elif percent >= 45:
                    gpa = 1.0
                elif percent >= 40:
                    gpa = 0.5
                else:
                    gpa = 0.0
                total_gpa += gpa
                assessment_count += 1
        
        if assessment_count > 0:
            avg_gpa = total_gpa / assessment_count
    
    # Calculate student performance by teacher/subject (similar to student dashboard)
    student_performance = {}
    
    for assessment in all_filtered_assessments:
        student_id = assessment.student_id
        teacher_id = assessment.teacher_id
        subj = assessment.subject
        
        if student_id not in student_performance:
            student_performance[student_id] = {
                'student': Student.query.get(student_id),
                'teachers': {}
            }
        
        if teacher_id not in student_performance[student_id]['teachers']:
            teacher = User.query.get(teacher_id)
            student_performance[student_id]['teachers'][teacher_id] = {
                'teacher_name': teacher.username if teacher else f"Teacher {teacher_id}",
                'subjects': {}
            }
        
        if subj not in student_performance[student_id]['teachers'][teacher_id]['subjects']:
            student_performance[student_id]['teachers'][teacher_id]['subjects'][subj] = {
                'total_score': 0,
                'total_max': 0,
                'assessments': []
            }
        
        student_performance[student_id]['teachers'][teacher_id]['subjects'][subj]['total_score'] += assessment.score
        student_performance[student_id]['teachers'][teacher_id]['subjects'][subj]['total_max'] += assessment.max_score
        student_performance[student_id]['teachers'][teacher_id]['subjects'][subj]['assessments'].append(assessment)
    
    # Calculate final grades and GPA for each student-teacher-subject combination
    for student_id, student_data in student_performance.items():
        for teacher_id, teacher_data in student_data['teachers'].items():
            for subj, subj_data in teacher_data['subjects'].items():
                if subj_data['total_max'] > 0:
                    final_percent = (subj_data['total_score'] / subj_data['total_max']) * 100
                    
                    # Calculate GPA and grade
                    if final_percent >= 80:
                        gpa = 4.0
                        grade = 'A1'
                        grade_color = 'success'
                    elif final_percent >= 70:
                        gpa = 3.5
                        grade = 'B2'
                        grade_color = 'success'
                    elif final_percent >= 65:
                        gpa = 3.0
                        grade = 'B3'
                        grade_color = 'warning'
                    elif final_percent >= 60:
                        gpa = 2.5
                        grade = 'C4'
                        grade_color = 'warning'
                    elif final_percent >= 55:
                        gpa = 2.0
                        grade = 'C5'
                        grade_color = 'warning'
                    elif final_percent >= 50:
                        gpa = 1.5
                        grade = 'C6'
                        grade_color = 'warning'
                    elif final_percent >= 45:
                        gpa = 1.0
                        grade = 'D7'
                        grade_color = 'danger'
                    elif final_percent >= 40:
                        gpa = 0.5
                        grade = 'E8'
                        grade_color = 'danger'
                    else:
                        gpa = 0.0
                        grade = 'F9'
                        grade_color = 'danger'
                    
                    subj_data.update({
                        'final_percent': final_percent,
                        'gpa': gpa,
                        'grade': grade,
                        'grade_color': grade_color
                    })
    
    # Convert to list for easier template handling
    student_performance_list = []
    for student_id, data in student_performance.items():
        student_performance_list.append(data)
    
    # Sort students by name
    student_performance_list.sort(key=lambda x: x['student'].last_name if x['student'] else '')
    
    pagination = query.order_by(Assessment.date_recorded.desc())\
        .paginate(page=page, per_page=per_page, error_out=False)
    
    form = AssessmentFilterForm()
    form.subject.data = subject
    form.class_name.data = class_name
    form.category.data = category
    
    # Get total unique students count
    unique_students_count = Student.query.count()
    
    return render_template(
        "assessments.html",
        assessments=pagination.items,
        form=form,
        page=page,
        per_page=per_page,
        total=pagination.total,
        category_labels=app.config['CATEGORY_LABELS'],
        pagination=pagination,
        student_performance=student_performance_list,
        subject_filter=subject,
        class_filter=class_name,
        category_filter=category,
        unique_students=unique_students_count,
        total_assessments=total_assessments,
        avg_score=avg_score,
        avg_gpa=avg_gpa,
        unique_students_affected=unique_students_affected
    )

@app.route("/assessments/new", methods=["GET", "POST"])
@login_required
def new_assessment():
    # Only teachers and admins can create assessments
    if not (current_user.is_teacher() or current_user.is_admin()):
        abort(403)
        
    form = AssessmentForm()
    
    # Populate student choices
    students = Student.query.all()
    form.student_name.choices = [("", "-- Select Student --")] + [(s.student_number, s.full_name()) for s in students]
    student_dict = {s.student_number: {'name': s.full_name(), 'ref': s.reference_number or ''} for s in students}
    
    # Get global settings
    settings = Setting.query.first()
    
    # Auto-fill subject and class for teachers
    if current_user.is_teacher() and current_user.subject:
        form.subject.data = current_user.subject
    
    # Auto-select student if student_id is provided in URL params
    student_id = request.args.get('student_id')
    if student_id:
        student = Student.query.get(student_id)
        if student:
            form.student_name.data = student.student_number
    
    # Auto-fill global settings
    if settings:
        form.term.data = settings.current_term
        form.academic_year.data = settings.current_academic_year
        form.session.data = settings.current_session
    
    if form.validate_on_submit():
        # Get student_number from either dropdown or manual input
        student_number = form.student_name.data or form.student_number.data.strip()
        student = Student.query.filter_by(student_number=student_number).first()
        
        if not student:
            flash("Student not found. Please create the student first.", "danger")
        else:
            # Check if assessment already exists for this student, category, subject, term, academic_year, session, AND teacher
            existing_assessment = Assessment.query.filter_by(
                student_id=student.id,
                category=form.category.data,
                subject=form.subject.data,
                term=form.term.data,
                academic_year=form.academic_year.data,
                session=form.session.data,
                teacher_id=current_user.id  # Ensure it's per teacher
            ).first()
            
            if existing_assessment:
                flash(f"An assessment for {form.category.data} in {form.subject.data} already exists for this student in the same term, academic year, and session. Please update the existing assessment instead.", "warning")
                return redirect(url_for('student_view', student_id=student.id))
            
            # Set max_score based on category
            category = form.category.data
            max_score = app.config['CATEGORY_MAX_SCORES'].get(category, 100.0)
            
            # Validate score doesn't exceed max_score
            if form.score.data > max_score:
                flash(f"Score cannot exceed max score of {max_score}", "danger")
                return redirect(url_for('new_assessment'))
            
            # Auto-assign class from student if not specified
            class_name = form.class_name.data or student.class_name
            
            assessment = Assessment(
                student=student,
                category=category,
                subject=form.subject.data,
                class_name=class_name,
                score=float(form.score.data),
                max_score=max_score,
                term=form.term.data,
                academic_year=form.academic_year.data,
                session=form.session.data,
                assessor=form.assessor.data or current_user.username,
                teacher_id=current_user.id,  # Always set teacher_id
                comments=form.comments.data
            )
            db.session.add(assessment)
            db.session.commit()
            log_activity(current_user, "create_assessment", f"Created assessment for {student.full_name()} ({assessment.category} in {assessment.subject})")
            flash(f"Assessment saved for {student.full_name()}", "success")
            return redirect(url_for("student_view", student_id=student.id))
    
    return render_template("assessment_form.html", form=form, students=students, student_dict=student_dict)

@app.route("/assessments/<int:assessment_id>/edit", methods=["GET", "POST"])
@login_required
def assessment_edit(assessment_id):
    assessment = Assessment.query.get_or_404(assessment_id)
    
    # Only teachers and admins can edit assessments
    # Teachers can only edit their own assessments
    if not (current_user.is_admin() or 
            (current_user.is_teacher() and assessment.teacher_id == current_user.id)):
        abort(403)
        
    form = AssessmentForm(obj=assessment)
    
    # Populate student choices
    students = Student.query.all()
    form.student_name.choices = [("", "-- Select Student --")] + [(s.student_number, s.full_name()) for s in students]
    student_dict = {s.student_number: {'name': s.full_name(), 'ref': s.reference_number or ''} for s in students}
    
    # Pre-fill form
    form.student_name.data = assessment.student.student_number
    form.student_number.data = assessment.student.student_number
    form.reference_number.data = assessment.student.reference_number
    
    if form.validate_on_submit():
        # Set max_score based on category
        category = form.category.data
        max_score = app.config['CATEGORY_MAX_SCORES'].get(category, 100.0)
        
        # Validate score doesn't exceed max_score
        if form.score.data > max_score:
            flash(f"Score cannot exceed max score of {max_score}", "danger")
            return redirect(url_for('assessment_edit', assessment_id=assessment_id))
        
        assessment.category = category
        assessment.subject = form.subject.data
        assessment.class_name = form.class_name.data
        assessment.score = float(form.score.data)
        assessment.max_score = max_score
        assessment.term = form.term.data
        assessment.academic_year = form.academic_year.data
        assessment.session = form.session.data
        assessment.assessor = form.assessor.data
        assessment.comments = form.comments.data
        db.session.commit()
        log_activity(current_user, "edit_assessment", f"Edited assessment for {assessment.student.full_name()} ({assessment.category} in {assessment.subject})")
        flash("Assessment updated successfully", "success")
        return redirect(url_for("student_view", student_id=assessment.student_id))
    
    return render_template("assessment_form.html", form=form, assessment=assessment, students=students, student_dict=student_dict)

@app.route("/assessments/<int:assessment_id>/delete", methods=["POST"])
@login_required
def assessment_delete(assessment_id):
    assessment = Assessment.query.get_or_404(assessment_id)
    
    # Only teachers and admins can delete assessments
    # Teachers can only delete their own assessments
    if not (current_user.is_admin() or 
            (current_user.is_teacher() and assessment.teacher_id == current_user.id)):
        abort(403)
        
    student_id = assessment.student_id
    db.session.delete(assessment)
    db.session.commit()
    log_activity(current_user, "delete_assessment", f"Deleted assessment for {assessment.student.full_name()} ({assessment.category} in {assessment.subject})")
    flash("Assessment deleted successfully", "info")
    return redirect(url_for("student_view", student_id=student_id))

@app.route("/assessments/<int:assessment_id>/archive", methods=["POST"])
@login_required
def assessment_archive(assessment_id):
    assessment = Assessment.query.get_or_404(assessment_id)
    
    # Only teachers and admins can archive assessments
    # Teachers can only archive their own assessments
    if not (current_user.is_admin() or 
            (current_user.is_teacher() and assessment.teacher_id == current_user.id)):
        abort(403)
        
    assessment.archived = True
    db.session.commit()
    flash("Assessment archived successfully", "info")
    return redirect(request.referrer or url_for("assessments"))

@app.route("/assessments/<int:assessment_id>/unarchive", methods=["POST"])
@login_required
def assessment_unarchive(assessment_id):
    assessment = Assessment.query.get_or_404(assessment_id)
    
    # Only teachers and admins can unarchive assessments
    # Teachers can only unarchive their own assessments
    if not (current_user.is_admin() or 
            (current_user.is_teacher() and assessment.teacher_id == current_user.id)):
        abort(403)
        
    assessment.archived = False
    db.session.commit()
    flash("Assessment unarchived successfully", "info")
    return redirect(request.referrer or url_for("assessments"))

@app.route("/assessments/archived")
@login_required
def assessments_archived():
    page = request.args.get('page', 1, type=int)
    subject = request.args.get('subject', '')
    class_name = request.args.get('class', '')
    category = request.args.get('category', '')
    
    per_page = app.config['ASSESSMENTS_PER_PAGE']
    
    # Build query based on user role and filters - only archived
    if hasattr(current_user, 'is_teacher') and current_user.is_teacher():
        query = Assessment.query.filter_by(teacher_id=current_user.id, archived=True)
    else:
        query = Assessment.query.filter_by(archived=True)
    
    if subject:
        query = query.filter_by(subject=subject)
    if class_name:
        query = query.filter_by(class_name=class_name)
    if category:
        query = query.filter_by(category=category)
    
    pagination = query.order_by(Assessment.date_recorded.desc())\
        .paginate(page=page, per_page=per_page, error_out=False)
    
    form = AssessmentFilterForm()
    form.subject.data = subject
    form.class_name.data = class_name
    form.category.data = category
    
    return render_template(
        "assessments.html",
        assessments=pagination.items,
        pagination=pagination,
        form=form,
        category_labels=app.config['CATEGORY_LABELS'],
        archived=True
    )

# -------------------------
@app.route("/users")
@login_required
@admin_required
def users():
    teachers_admins = User.query.filter(User.role.in_(['admin', 'teacher'])).order_by(User.username).all()
    students = User.query.filter_by(role='student').order_by(User.username).all()
    return render_template("users.html", teachers_admins=teachers_admins, students=students)

@app.route("/users/new", methods=["GET", "POST"])
@login_required
@admin_required
def create_user():
    form = UserForm()
    
    if form.validate_on_submit():
        if User.query.filter_by(username=form.username.data.strip()).first():
            flash("Username already exists", "warning")
        else:
            pw_hash = bcrypt.generate_password_hash(form.password.data).decode("utf-8")
            user = User(
                username=form.username.data.strip(),
                password_hash=pw_hash,
                role=form.role.data,
                subject=form.subject.data if form.subject.data else None
            )
            # Handle multiple classes for teachers
            if form.classes.data:
                user.set_classes_list(form.classes.data)
            db.session.add(user)
            db.session.commit()
            log_activity(current_user, "create_user", f"Created user {user.username} with role {user.role}")
            flash(f"User {user.username} created successfully", "success")
            return redirect(url_for("users"))
    
    return render_template("user_form.html", form=form)

@app.route("/users/<int:user_id>/edit", methods=["GET", "POST"])
@login_required
@admin_required
def edit_user(user_id):
    user = User.query.get_or_404(user_id)
    form = EditUserForm(role=user.role)
    
    if form.validate_on_submit():
        user.role = form.role.data
        user.subject = form.subject.data if form.subject.data else None
        # Handle multiple classes for teachers
        if form.classes.data:
            user.set_classes_list(form.classes.data)
        else:
            user.classes = None
        db.session.commit()
        log_activity(current_user, "edit_user", f"Edited user {user.username}")
        flash(f"User {user.username} updated successfully", "success")
        return redirect(url_for("users"))
    
    # Pre-fill form
    if user.subject:
        form.subject.data = user.subject
    # Pre-fill classes list
    classes_list = user.get_classes_list()
    if classes_list:
        form.classes.data = classes_list
    
    return render_template("edit_user.html", form=form, user=user)

@app.route("/users/<int:user_id>/reset_password", methods=["GET", "POST"])
@login_required
@admin_required
def reset_password(user_id):
    user = User.query.get_or_404(user_id)
    form = PasswordResetForm()
    
    if form.validate_on_submit():
        user.password_hash = bcrypt.generate_password_hash(form.password.data).decode("utf-8")
        db.session.commit()
        log_activity(current_user, "reset_password", f"Reset password for user {user.username}")
        flash(f"Password reset successfully for {user.username}", "success")
        return redirect(url_for("users"))
    
    return render_template("reset_password.html", form=form, user=user)

@app.route("/users/<int:user_id>/delete", methods=["POST"])
@login_required
@admin_required
def delete_user(user_id):
    if current_user.id == user_id:
        flash("You cannot delete your own account", "danger")
        return redirect(url_for("users"))
    
    user = User.query.get_or_404(user_id)
    username = user.username
    db.session.delete(user)
    db.session.commit()
    log_activity(current_user, "delete_user", f"Deleted user {username}")
    flash(f"User {username} deleted successfully", "info")
    return redirect(url_for("users"))

# -------------------------
# Admin Settings Routes
# -------------------------
@app.route("/admin/settings", methods=["GET", "POST"])
@login_required
@admin_required
def admin_settings():
    """Admin can configure global settings"""
    settings = Setting.query.first()
    if not settings:
        settings = Setting()
        db.session.add(settings)
        db.session.commit()
    
    form = SettingsForm(obj=settings)
    
    if form.validate_on_submit():
        settings.current_term = form.current_term.data
        settings.current_academic_year = form.current_academic_year.data
        settings.current_session = form.current_session.data
        settings.assessment_active = form.assessment_active.data
        db.session.commit()
        flash("Settings updated successfully", "success")
        return redirect(url_for("admin_settings"))
    
    return render_template("admin_settings.html", form=form, settings=settings)

@app.route("/admin/activity-logs")
@login_required
@admin_required
def admin_activity_logs():
    """Admin can view activity logs"""
    page = request.args.get('page', 1, type=int)
    per_page = 50
    
    logs = ActivityLog.query.order_by(ActivityLog.timestamp.desc()).paginate(page=page, per_page=per_page, error_out=False)
    
    return render_template("activity_logs.html", logs=logs)

# -------------------------
# Teacher Routes
# -------------------------
@app.route("/users/<int:user_id>/assign-subject", methods=["GET", "POST"])
@login_required
@admin_required
def assign_teacher_subject(user_id):
    """Admin can assign subject specialization to teachers"""
    user = User.query.get_or_404(user_id)
    if not user.is_teacher():
        flash("This user is not a teacher", "danger")
        return redirect(url_for("users"))
        
    form = TeacherAssignmentForm()
    
    if form.validate_on_submit():
        user.subject = form.subject.data
        # Handle multiple classes
        if form.classes.data:
            user.set_classes_list(form.classes.data)
        else:
            user.classes = None
        db.session.commit()
        flash(f"Subject assigned to {user.username}: {dict(app.config['LEARNING_AREAS']).get(form.subject.data)}", "success")
        return redirect(url_for("users"))
    
    if user.subject:
        form.subject.data = user.subject
    # Pre-fill classes list
    classes_list = user.get_classes_list()
    if classes_list:
        form.classes.data = classes_list
    
    return render_template("teacher_subject.html", form=form, teacher=user)

@app.route("/teacher/subject", methods=["GET", "POST"])
@login_required
@teacher_required
def teacher_subject():
    """Teacher can set their subject specialization"""
    user = current_user
    
    form = TeacherAssignmentForm()
    
    if form.validate_on_submit():
        user.subject = form.subject.data
        # Handle multiple classes
        if form.classes.data:
            user.set_classes_list(form.classes.data)
        else:
            user.classes = None
        db.session.commit()
        flash(f"Subject updated: {dict(app.config['LEARNING_AREAS']).get(form.subject.data)}", "success")
        return redirect(url_for("dashboard"))
    
    if user.subject:
        form.subject.data = user.subject
    # Pre-fill classes list
    classes_list = user.get_classes_list()
    if classes_list:
        form.classes.data = classes_list
    
    return render_template("teacher_subject.html", form=form, teacher=None)


# -------------------------------
# Question Bank Routes
# -------------------------------

@app.route("/teacher/question-bank")
@login_required
def teacher_question_bank():
    """Teacher can view and manage their subject questions, Admin can view all"""
    if not (current_user.is_teacher() or current_user.is_admin()):
        abort(403)
    
    page = request.args.get('page', 1, type=int)
    per_page = 20
    
    # Admin can see all questions, teachers see their subject
    if current_user.is_admin():
        query = Question.query
        # Allow admin to filter by subject
        subject_filter = request.args.get('subject')
        if subject_filter:
            query = query.filter_by(subject=subject_filter)
    else:
        query = Question.query.filter_by(subject=current_user.subject)
    
    # Filter by status if specified
    status_filter = request.args.get('status')
    if status_filter:
        query = query.filter_by(status=status_filter)
    
    questions = query.order_by(Question.created_at.desc()).paginate(page=page, per_page=per_page)
    
    # Get subjects for admin filter
    subjects = []
    if current_user.is_admin():
        subjects = db.session.query(Question.subject).distinct().all()
        subjects = [s[0] for s in subjects]
    
    return render_template("teacher_question_bank.html", questions=questions, 
                         status_filter=status_filter, subject_filter=request.args.get('subject'), 
                         subjects=subjects, is_admin=current_user.is_admin())


@app.route("/teacher/questions/new", methods=["GET", "POST"])
@login_required
@teacher_required
def create_question():
    """Teacher can create new questions"""
    form = QuestionForm()
    
    if form.validate_on_submit():
        # Process options for MCQ
        options = None
        if form.question_type.data == 'mcq' and form.options.data:
            options = [line.strip() for line in form.options.data.split('\n') if line.strip()]
        
        # Process keywords for short answer
        keywords = None
        if form.question_type.data == 'short_answer' and form.keywords.data:
            keywords = [line.strip().lower() for line in form.keywords.data.split('\n') if line.strip()]
        
        question = Question(
            subject=current_user.subject,
            question_text=form.question_text.data,
            question_type=form.question_type.data,
            options=options,
            correct_answer=form.correct_answer.data,
            marks=form.marks.data,
            keywords=keywords,
            difficulty=form.difficulty.data,
            explanation=form.explanation.data,
            created_by=current_user.id
        )
        db.session.add(question)
        db.session.commit()
        
        # Log activity
        log_activity(current_user, "create_question", f"Created question ID {question.id} for {question.subject}")
        
        flash("Question created successfully and submitted for approval", "success")
        return redirect(url_for("teacher_question_bank"))
    
    return render_template("question_form.html", form=form, title="Create Question")


@app.route("/teacher/questions/<int:question_id>/edit", methods=["GET", "POST"])
@login_required
@teacher_required
def edit_question(question_id):
    """Teacher can edit their pending questions"""
    question = Question.query.get_or_404(question_id)
    
    # Check permissions
    if not question.can_edit(current_user):
        abort(403)
    
    form = QuestionForm(obj=question)
    
    # Convert options list to string for the textarea
    if question.options and isinstance(question.options, list):
        form.options.data = '\n'.join(question.options)
    
    # Convert keywords list to string for the textarea
    if question.keywords and isinstance(question.keywords, list):
        form.keywords.data = '\n'.join(question.keywords)
    
    if form.validate_on_submit():
        # Process options for MCQ
        options = None
        if form.question_type.data == 'mcq' and form.options.data:
            options = [line.strip() for line in form.options.data.split('\n') if line.strip()]
        
        # Process keywords for short answer
        keywords = None
        if form.question_type.data == 'short_answer' and form.keywords.data:
            keywords = [line.strip().lower() for line in form.keywords.data.split('\n') if line.strip()]
        
        question.question_text = form.question_text.data
        question.question_type = form.question_type.data
        question.options = options
        question.correct_answer = form.correct_answer.data
        question.marks = form.marks.data
        question.keywords = keywords
        question.difficulty = form.difficulty.data
        question.explanation = form.explanation.data
        question.updated_at = datetime.utcnow()
        
        db.session.commit()
        
        # Log activity
        log_activity(current_user, "edit_question", f"Edited question ID {question.id}")
        
        flash("Question updated successfully", "success")
        return redirect(url_for("teacher_question_bank"))
    
    return render_template("question_form.html", form=form, title="Edit Question", question=question)


@app.route("/teacher/questions/<int:question_id>/delete", methods=["POST"])
@login_required
@teacher_required
def delete_question(question_id):
    """Teacher can delete their pending questions"""
    question = Question.query.get_or_404(question_id)
    
    # Check permissions
    if not question.can_edit(current_user):
        abort(403)
    
    db.session.delete(question)
    db.session.commit()
    
    # Log activity
    log_activity(current_user, "delete_question", f"Deleted question ID {question.id}")
    
    flash("Question deleted successfully", "success")
    return redirect(url_for("teacher_question_bank"))


@app.route("/admin/question-bank")
@login_required
@admin_required
def admin_question_bank():
    """Admin can moderate all questions"""
    page = request.args.get('page', 1, type=int)
    per_page = 20
    
    # Get all questions
    query = Question.query
    
    # Filter by status if specified
    status_filter = request.args.get('status', 'pending')
    if status_filter:
        query = query.filter_by(status=status_filter)
    
    # Filter by subject if specified
    subject_filter = request.args.get('subject')
    if subject_filter:
        query = query.filter_by(subject=subject_filter)
    
    questions = query.order_by(Question.created_at.desc()).paginate(page=page, per_page=per_page)
    
    # Get all subjects for filter
    subjects = db.session.query(Question.subject).distinct().all()
    subjects = [s[0] for s in subjects]
    
    return render_template("admin_question_bank.html", questions=questions, 
                         status_filter=status_filter, subject_filter=subject_filter, subjects=subjects)


@app.route("/admin/questions/<int:question_id>/approve", methods=["POST"])
@login_required
@admin_required
def approve_question(question_id):
    """Admin can approve questions"""
    question = Question.query.get_or_404(question_id)
    
    action = request.form.get('action')
    if action == 'approve':
        question.status = 'approved'
        question.approved_by = current_user.id
        flash("Question approved successfully", "success")
    elif action == 'reject':
        question.status = 'rejected'
        question.approved_by = current_user.id
        question.rejection_reason = request.form.get('rejection_reason')
        flash("Question rejected", "warning")
    
    db.session.commit()
    
    # Log activity
    log_activity(current_user, "moderate_question", f"{action}d question ID {question.id}")
    
    return redirect(url_for("admin_question_bank"))


@app.route("/admin/questions/approve_all", methods=["POST"])
@login_required
@admin_required
def approve_all_questions():
    """Admin can approve all pending questions"""
    # Get all pending questions
    pending_questions = Question.query.filter_by(status='pending').all()
    
    approved_count = 0
    for question in pending_questions:
        question.status = 'approved'
        question.approved_by = current_user.id
        approved_count += 1
    
    db.session.commit()
    
    # Log activity
    log_activity(current_user, "approve_all_questions", f"Approved {approved_count} pending questions")
    
    flash(f"Approved {approved_count} questions successfully", "success")
    return redirect(url_for("admin_question_bank"))


@app.route("/teacher/questions/<int:question_id>/approve", methods=["POST"])
@login_required
@teacher_required
def teacher_approve_question(question_id):
    """Teacher can approve questions in their subject"""
    question = Question.query.get_or_404(question_id)
    
    # Check if question is in teacher's subject
    if question.subject != current_user.subject:
        abort(403)
    
    action = request.form.get('action')
    if action == 'approve':
        question.status = 'approved'
        question.approved_by = current_user.id
        flash("Question approved successfully", "success")
    elif action == 'reject':
        question.status = 'rejected'
        question.approved_by = current_user.id
        question.rejection_reason = request.form.get('rejection_reason')
        flash("Question rejected", "warning")
    
    db.session.commit()
    
    # Log activity
    log_activity(current_user, "moderate_question", f"{action}d question ID {question.id}")
    
    return redirect(url_for("teacher_question_bank"))


@app.route("/student/questions")
@login_required
@student_required
def student_questions():
    """Student can view and answer questions"""
    page = request.args.get('page', 1, type=int)
    per_page = 10
    
    # Get approved questions for student's subjects
    # For now, get questions from all subjects, but in production this should be filtered
    # based on student's enrolled subjects
    questions = Question.query.filter_by(status='approved').order_by(Question.created_at.desc()).paginate(page=page, per_page=per_page)
    
    # Get student's previous attempts
    attempts = {attempt.question_id: attempt for attempt in 
               QuestionAttempt.query.filter_by(student_id=current_user.id).all()}
    
    return render_template("student_questions.html", questions=questions, attempts=attempts)


@app.route("/student/questions/<int:question_id>/attempt", methods=["POST"])
@login_required
@student_required
def attempt_question(question_id):
    """Student submits answer to a question"""
    question = Question.query.get_or_404(question_id)
    
    if question.status != 'approved':
        abort(404)
    
    student_answer = request.form.get('answer')
    if not student_answer:
        flash("Please provide an answer", "danger")
        return redirect(url_for("student_questions"))
    
    # Check if correct
    is_correct = False
    if question.question_type == 'mcq':
        is_correct = student_answer.strip().upper() == question.correct_answer.strip().upper()
    elif question.question_type == 'true_false':
        is_correct = student_answer.lower() == question.correct_answer.lower()
    else:  # short_answer - for now, simple string match, but could be more sophisticated
        is_correct = student_answer.strip().lower() == question.correct_answer.strip().lower()
    
    # Record attempt
    attempt = QuestionAttempt(
        student_id=current_user.id,
        question_id=question_id,
        student_answer=student_answer,
        is_correct=is_correct
    )
    db.session.add(attempt)
    db.session.commit()
    
    # Log activity
    log_activity(current_user, "attempt_question", f"Answered question ID {question.id}, correct: {is_correct}")
    
    if is_correct:
        flash("Correct answer!", "success")
    else:
        flash(f"Incorrect. The correct answer is: {question.correct_answer}", "warning")
    
    return redirect(url_for("student_questions"))


@app.route("/teacher/quizzes")
@login_required
def teacher_quizzes():
    """Teacher can view and manage quizzes for their subject, Admin can view all"""
    if not (current_user.is_teacher() or current_user.is_admin()):
        abort(403)
    
    if current_user.is_admin():
        quizzes = Quiz.query.order_by(Quiz.created_at.desc()).all()
    else:
        # Teachers see quizzes for their subject
        quizzes = Quiz.query.filter_by(subject=current_user.subject).order_by(Quiz.created_at.desc()).all()
    
    return render_template("teacher_quizzes.html", quizzes=quizzes)


@app.route("/teacher/quizzes/new", methods=["GET", "POST"])
@login_required
def create_quiz():
    """Teacher can create new quizzes for their subject, Admin can create for any subject"""
    if not (current_user.is_teacher() or current_user.is_admin()):
        abort(403)
    
    form = QuizForm()
    
    # Set subject choices based on user role
    if current_user.is_admin():
        # Admin can choose any subject
        form.subject.choices = [(subject[0], subject[1]) for subject in app.config['LEARNING_AREAS']]
    else:
        # Teachers are limited to their subject
        form.subject.choices = [(current_user.subject, current_user.subject.replace('_', ' ').title())]
        form.subject.data = current_user.subject
    
    # Set questions choices based on subject
    subject = None
    if request.method == 'POST':
        subject = request.form.get('subject', current_user.subject if current_user.is_teacher() else None)
    else:
        subject = request.args.get('subject', current_user.subject if current_user.is_teacher() else None)
    
    if subject:
        questions = Question.query.filter_by(subject=subject, status='approved').all()
        form.questions.choices = [(str(q.id), f"{q.question_text[:50]}{'...' if len(q.question_text) > 50 else ''} ({q.difficulty.title()}, {q.question_type.upper()})") for q in questions]
    else:
        form.questions.choices = []
    
    if form.validate_on_submit():
        # Get approved questions for the selected subject
        questions = Question.query.filter_by(subject=form.subject.data, status='approved').all()
        selected_question_ids = [int(q) for q in form.questions.data if q.isdigit()]
        
        # Validate that selected questions exist and are approved
        valid_questions = [q for q in questions if q.id in selected_question_ids]
        
        quiz = Quiz(
            title=form.title.data,
            subject=form.subject.data,
            description=form.description.data,
            questions=[q.id for q in valid_questions],
            time_limit=form.time_limit.data,
            created_by=current_user.id
        )
        db.session.add(quiz)
        db.session.commit()
        
        # Log activity
        log_activity(current_user, "create_quiz", f"Created quiz '{quiz.title}' with {len(quiz.questions)} questions")
        
        flash("Quiz created successfully", "success")
        return redirect(url_for("teacher_quizzes"))
    
    # For GET request, populate questions based on subject
    if not subject:
        questions = []
    else:
        questions = Question.query.filter_by(subject=subject, status='approved').all()
    
    return render_template("quiz_form.html", form=form, available_questions=questions, quiz=None)


@app.route("/student/quizzes")
@login_required
@student_required
def student_quizzes():
    """Student can view available quizzes"""
    # Get student record
    student = Student.query.filter_by(student_number=current_user.username).first()
    if not student:
        flash("Student record not found", "danger")
        return redirect(url_for("student_dashboard"))
    
    # For now, show all active quizzes, but should filter by student's subjects
    quizzes = Quiz.query.filter_by(is_active=True).order_by(Quiz.created_at.desc()).all()
    
    # Get student's previous attempts
    attempts = {attempt.quiz_id: attempt for attempt in 
               QuizAttempt.query.filter_by(student_id=student.id).all()}
    
    return render_template("student_quizzes.html", quizzes=quizzes, attempts=attempts)


@app.route("/student/quizzes/<int:quiz_id>/take", methods=["GET", "POST"])
@login_required
@student_required
def take_quiz(quiz_id):
    """Student takes a quiz"""
    quiz = Quiz.query.get_or_404(quiz_id)
    
    if not quiz.is_active:
        abort(404)
    
    # Get student record
    student = Student.query.filter_by(student_number=current_user.username).first()
    if not student:
        flash("Student record not found", "danger")
        return redirect(url_for("student_dashboard"))
    
    # Check if student already completed this quiz
    completed_attempt = QuizAttempt.query.filter_by(student_id=student.id, quiz_id=quiz_id, status="completed").first()
    if completed_attempt:
        flash("You have already taken this quiz", "warning")
        return redirect(url_for("student_quizzes"))
    
    # Check for in-progress attempt
    attempt = QuizAttempt.query.filter_by(student_id=student.id, quiz_id=quiz_id, status="in_progress").first()
    if not attempt:
        # Create new attempt
        attempt = QuizAttempt(
            student_id=student.id,
            quiz_id=quiz_id,
            score=0.0,
            total_questions=len(quiz.questions),
            correct_answers=0,
            remaining_time=quiz.time_limit * 60 if quiz.time_limit else None
        )
        db.session.add(attempt)
        db.session.commit()
    
    questions = Question.query.filter(Question.id.in_(quiz.questions)).all()
    questions_dict = {q.id: q for q in questions}
    
    # Load saved answers if any
    saved_answers = {}
    if attempt.answers_json:
        import json
        saved_answers = json.loads(attempt.answers_json)
    
    if request.method == 'POST':
        # Process quiz submission
        answers = {}
        total_score = 0.0
        total_marks = 0.0
        question_results = {}
        
        for qid in quiz.questions:
            answer = request.form.get(f'answer_{qid}')
            if answer:
                question = questions_dict.get(int(qid))
                if question:
                    score = 0.0
                    is_correct = False
                    if question.question_type == 'mcq':
                        is_correct = answer.strip().upper() == question.correct_answer.strip().upper()
                        score = question.marks if is_correct else 0.0
                    elif question.question_type == 'true_false':
                        is_correct = answer.lower() == question.correct_answer.lower()
                        score = question.marks if is_correct else 0.0
                    elif question.question_type == 'short_answer':
                        # Flexible marking for short answer questions
                        score = calculate_short_answer_score(answer, question)
                        is_correct = score > 0  # Any score > 0 is considered attempted correctly
                    
                    total_score += score
                    total_marks += question.marks
                    
                    # Store question result for display
                    question_results[qid] = {
                        'student_answer': answer,
                        'score': score,
                        'max_marks': question.marks,
                        'correct_answer': question.correct_answer
                    }
                    
                    # Record individual question attempt
                    question_attempt = QuestionAttempt(
                        student_id=student.id,
                        question_id=qid,
                        student_answer=answer,
                        is_correct=is_correct,
                        score=score
                    )
                    db.session.add(question_attempt)
        
        # Update quiz attempt
        attempt.score = total_score
        attempt.correct_answers = sum(1 for result in question_results.values() if result['score'] > 0)
        attempt.completed_at = datetime.utcnow()
        attempt.time_taken = int((attempt.completed_at - attempt.started_at).total_seconds())
        attempt.status = "completed"
        attempt.answers_json = None  # Clear saved answers
        db.session.commit()
        
        # Log activity
        log_activity(current_user, "complete_quiz", f"Completed quiz '{quiz.title}' with score {total_score:.1f}/{total_marks:.1f}")
        
        # Store quiz results temporarily in session (expires in 2 hours)
        import time
        session['quiz_results'] = {
            'quiz_id': quiz_id,
            'quiz_title': quiz.title,
            'score': total_score,
            'total_marks': total_marks,
            'percentage': round((total_score / total_marks) * 100, 1) if total_marks > 0 else 0,
            'completed_at': datetime.utcnow().timestamp(),
            'question_results': question_results  # Store individual question results
        }
        session.modified = True
        
        return redirect(url_for("quiz_results"))
    
    return render_template("take_quiz.html", quiz=quiz, questions=questions_dict, attempt=attempt, saved_answers=saved_answers)


@app.route("/student/quizzes/<int:quiz_id>/save_progress", methods=["POST"])
@login_required
@student_required
def save_quiz_progress(quiz_id):
    """Auto-save quiz progress"""
    student = Student.query.filter_by(student_number=current_user.username).first()
    if not student:
        return jsonify({"success": False, "message": "Student not found"}), 400
    
    attempt = QuizAttempt.query.filter_by(student_id=student.id, quiz_id=quiz_id, status="in_progress").first()
    if not attempt:
        return jsonify({"success": False, "message": "No active attempt found"}), 400
    
    # Get answers from request
    answers = {}
    for key, value in request.form.items():
        if key.startswith('answer_'):
            qid = key.replace('answer_', '')
            answers[qid] = value
    
    # Save answers and remaining time
    import json
    attempt.answers_json = json.dumps(answers)
    attempt.remaining_time = int(request.form.get('remaining_time', 0))
    db.session.commit()
    
    return jsonify({"success": True})


@app.route("/quiz/results")
@login_required
@student_required
def quiz_results():
    """Display quiz results temporarily (for 2 hours)"""
    quiz_results = session.get('quiz_results')
    
    if not quiz_results:
        flash("No quiz results available", "warning")
        return redirect(url_for("student_quizzes"))
    
    # Check if results are still valid (within 2 hours)
    import time
    current_time = time.time()
    results_time = quiz_results.get('completed_at', 0)
    
    if current_time - results_time > 7200:  # 2 hours in seconds
        session.pop('quiz_results', None)
        flash("Quiz results have expired", "info")
        return redirect(url_for("student_quizzes"))
    
    # Get quiz and questions for detailed display
    quiz = Quiz.query.get_or_404(quiz_results['quiz_id'])
    questions = {}
    for q_id in quiz.questions:
        question = Question.query.get(q_id)
        if question:
            questions[q_id] = question
    
    # Format completion time for display
    import time
    completed_at_timestamp = quiz_results.get('completed_at', 0)
    completed_at_formatted = datetime.fromtimestamp(completed_at_timestamp).strftime('%Y-%m-%d %H:%M')
    
    return render_template("quiz_results.html", 
                         quiz_results=quiz_results, 
                         quiz=quiz, 
                         questions=questions,
                         completed_at_formatted=completed_at_formatted)


@app.route("/student/quiz-attempt/<int:attempt_id>/review")
@login_required
@student_required
def quiz_attempt_review(attempt_id):
    """Review a specific quiz attempt with detailed question breakdown"""
    # Get student record
    student = Student.query.filter_by(student_number=current_user.username).first()
    if not student:
        flash("Student record not found", "danger")
        return redirect(url_for("student_dashboard"))
    
    # Get the quiz attempt
    attempt = QuizAttempt.query.filter_by(id=attempt_id, student_id=student.id).first()
    if not attempt:
        flash("Quiz attempt not found", "danger")
        return redirect(url_for("student_dashboard"))
    
    # Get quiz details
    quiz = Quiz.query.get(attempt.quiz_id)
    if not quiz:
        flash("Quiz not found", "danger")
        return redirect(url_for("student_dashboard"))
    
    # Get questions
    questions = {}
    for q_id in quiz.questions:
        question = Question.query.get(q_id)
        if question:
            questions[q_id] = question
    
    # Get question attempts for this quiz attempt
    # Match attempts by quiz questions and time proximity to completion
    if attempt.completed_at:
        # Get attempts within 10 minutes of completion for the quiz questions
        from datetime import timedelta
        time_window_start = attempt.completed_at.replace(second=0, microsecond=0)  # Round down to minute
        time_window_end = time_window_start + timedelta(minutes=10)
        
        question_attempts = QuestionAttempt.query.filter(
            QuestionAttempt.student_id == student.id,
            QuestionAttempt.question_id.in_(quiz.questions),
            QuestionAttempt.attempted_at >= time_window_start,
            QuestionAttempt.attempted_at <= time_window_end
        ).order_by(QuestionAttempt.attempted_at.desc()).all()
        
        # Group by question_id, taking the most recent attempt for each question
        latest_attempts = {}
        for qa in question_attempts:
            if qa.question_id not in latest_attempts:
                latest_attempts[qa.question_id] = qa
    else:
        # Fallback: get the most recent attempts for these questions
        question_attempts = []
        for q_id in quiz.questions:
            latest_attempt = QuestionAttempt.query.filter_by(
                student_id=student.id,
                question_id=q_id
            ).order_by(QuestionAttempt.attempted_at.desc()).first()
            if latest_attempt:
                question_attempts.append(latest_attempt)
        
        latest_attempts = {qa.question_id: qa for qa in question_attempts}
    
    return render_template("quiz_attempt_review.html",
                         attempt=attempt,
                         quiz=quiz,
                         questions=questions,
                         question_attempts=latest_attempts)


@app.route("/teacher/quizzes/<int:quiz_id>")
@login_required
def quiz_detail(quiz_id):
    """View quiz details"""
    if not (current_user.is_teacher() or current_user.is_admin()):
        abort(403)
    
    quiz = Quiz.query.get_or_404(quiz_id)
    
    # Check permissions: admin can see all, teachers can see their subject
    if not current_user.is_admin() and quiz.subject != current_user.subject:
        abort(403)
    
    # Get questions for this quiz
    questions = {}
    for q_id in quiz.questions:
        question = Question.query.get(q_id)
        if question:
            questions[q_id] = question
    
    return render_template("quiz_detail.html", quiz=quiz, questions=questions)


@app.route("/teacher/quizzes/<int:quiz_id>/edit", methods=["GET", "POST"])
@login_required
def edit_quiz(quiz_id):
    """Edit existing quiz"""
    if not (current_user.is_teacher() or current_user.is_admin()):
        abort(403)
    
    quiz = Quiz.query.get_or_404(quiz_id)
    
    # Check permissions: admin can edit all, teachers can edit their subject quizzes
    if not current_user.is_admin() and quiz.subject != current_user.subject:
        abort(403)
    
    form = QuizForm()
    
    # Set subject choices based on user role
    if current_user.is_admin():
        # Admin can choose any subject
        form.subject.choices = [(subject[0], subject[1]) for subject in app.config['LEARNING_AREAS']]
    else:
        # Teachers are limited to their subject
        form.subject.choices = [(current_user.subject, current_user.subject.replace('_', ' ').title())]
    
    # Determine subject for questions
    subject = quiz.subject  # Default to current quiz subject
    if request.method == 'POST':
        subject = request.form.get('subject', quiz.subject)
    
    # Get available questions for the subject
    available_questions = Question.query.filter_by(
        subject=subject, 
        status='approved'
    ).all()
    
    # Set questions choices
    form.questions.choices = [(q.id, f"{q.question_text[:50]}...") for q in available_questions]
    
    if form.validate_on_submit():
        quiz.title = form.title.data
        quiz.description = form.description.data
        quiz.subject = form.subject.data
        quiz.time_limit = form.time_limit.data if form.time_limit.data else None
        quiz.is_active = form.is_active.data
        
        # Handle question selection
        selected_questions = request.form.getlist('questions')
        quiz.questions = [int(q) for q in selected_questions if q.isdigit()]
        
        db.session.commit()
        log_activity(current_user, "edit_quiz", f"Edited quiz '{quiz.title}'")
        flash("Quiz updated successfully", "success")
        return redirect(url_for("teacher_quizzes"))
    
    # Pre-populate form
    form.title.data = quiz.title
    form.description.data = quiz.description
    form.subject.data = quiz.subject
    form.time_limit.data = quiz.time_limit
    form.is_active.data = quiz.is_active
    form.questions.data = quiz.questions  # This is a list of question IDs
    
    return render_template("quiz_form.html", form=form, quiz=quiz, available_questions=available_questions)


@app.route("/teacher/quizzes/<int:quiz_id>/delete", methods=["POST"])
@login_required
def delete_quiz(quiz_id):
    """Delete a quiz"""
    if not (current_user.is_teacher() or current_user.is_admin()):
        abort(403)
    
    quiz = Quiz.query.get_or_404(quiz_id)
    
    # Check permissions: admin can delete all, teachers can delete their subject quizzes
    if not current_user.is_admin() and quiz.subject != current_user.subject:
        abort(403)
    
    quiz_title = quiz.title
    
    # Delete associated attempts
    QuizAttempt.query.filter_by(quiz_id=quiz_id).delete()
    
    # Delete the quiz
    db.session.delete(quiz)
    db.session.commit()
    
    log_activity(current_user, "delete_quiz", f"Deleted quiz '{quiz_title}'")
    flash(f"Quiz '{quiz_title}' deleted successfully", "success")
    return redirect(url_for("teacher_quizzes"))


@app.route("/teacher/quizzes/<int:quiz_id>/results")
@login_required
def quiz_results_view(quiz_id):
    """View results of a specific quiz"""
    if not (current_user.is_teacher() or current_user.is_admin()):
        abort(403)
    
    quiz = Quiz.query.get_or_404(quiz_id)
    
    # Check permissions: admin can see all, teachers can see their subject quizzes
    if not current_user.is_admin() and quiz.subject != current_user.subject:
        abort(403)
    
    # Get all attempts for this quiz
    attempts = QuizAttempt.query.filter_by(quiz_id=quiz_id).order_by(QuizAttempt.completed_at.desc()).all()
    
    # Calculate summary statistics
    summary_stats = {
        'total_attempts': len(attempts),
        'avg_score': 0.0,
        'highest_score': 0.0,
        'completed_count': 0
    }
    
    if attempts:
        percentages = []
        for attempt in attempts:
            percentage = attempt.get_percentage()
            percentages.append(percentage)
            summary_stats['highest_score'] = max(summary_stats['highest_score'], percentage)
            if attempt.completed_at:
                summary_stats['completed_count'] += 1
        
        summary_stats['avg_score'] = sum(percentages) / len(percentages)
    
    # Get student details
    student_ids = [attempt.student_id for attempt in attempts]
    students = {student.id: student for student in Student.query.filter(Student.id.in_(student_ids)).all()}
    
    # For any missing students (old data with User.id), try to find by student_number and map them
    missing_ids = [sid for sid in student_ids if sid not in students]
    if missing_ids:
        users = User.query.filter(User.id.in_(missing_ids)).all()
        user_dict = {user.id: user for user in users}
        for user_id in missing_ids:
            user = user_dict.get(user_id)
            if user:
                student = Student.query.filter_by(student_number=user.username).first()
                if student:
                    students[user_id] = student  # Map old User.id to Student object
    
    return render_template("quiz_results_view.html", quiz=quiz, attempts=attempts, students=students, summary_stats=summary_stats)


@app.route("/teacher/quiz-results")
@login_required
def teacher_quiz_results():
    """Teacher can view all quiz results for their subject, Admin can view all or filter by subject"""
    if not (current_user.is_teacher() or current_user.is_admin()):
        abort(403)
    
    # Get subject filter
    subject_filter = request.args.get("subject", "")
    
    # Get quizzes based on permissions and filter
    if current_user.is_admin():
        query = Quiz.query
        if subject_filter:
            query = query.filter_by(subject=subject_filter)
        quizzes = query.order_by(Quiz.created_at.desc()).all()
    else:
        quizzes = Quiz.query.filter_by(subject=current_user.subject).order_by(Quiz.created_at.desc()).all()
    
    # Get attempts for these quizzes
    quiz_ids = [quiz.id for quiz in quizzes]
    attempts = QuizAttempt.query.filter(QuizAttempt.quiz_id.in_(quiz_ids)).order_by(QuizAttempt.completed_at.desc()).all()
    
    # Group attempts by quiz and calculate summaries
    attempts_by_quiz = {}
    quiz_summaries = {}
    
    for attempt in attempts:
        quiz_id = attempt.quiz_id
        if quiz_id not in attempts_by_quiz:
            attempts_by_quiz[quiz_id] = []
            quiz_summaries[quiz_id] = {
                'total_attempts': 0,
                'avg_score': 0.0,
                'highest_score': 0.0,
                'completed_count': 0
            }
        attempts_by_quiz[quiz_id].append(attempt)
        
        # Update summary
        summary = quiz_summaries[quiz_id]
        summary['total_attempts'] += 1
        percentage = attempt.get_percentage()
        summary['avg_score'] += percentage
        summary['highest_score'] = max(summary['highest_score'], percentage)
        if attempt.completed_at:
            summary['completed_count'] += 1
    
    # Calculate final averages
    for quiz_id, summary in quiz_summaries.items():
        if summary['total_attempts'] > 0:
            summary['avg_score'] = summary['avg_score'] / summary['total_attempts']
    
    # Get student details
    student_ids = list(set(attempt.student_id for attempt in attempts))
    students = {student.id: student for student in Student.query.filter(Student.id.in_(student_ids)).all()}
    
    # For any missing students (old data with User.id), try to find by student_number and map them
    missing_ids = [sid for sid in student_ids if sid not in students]
    if missing_ids:
        users = User.query.filter(User.id.in_(missing_ids)).all()
        user_dict = {user.id: user for user in users}
        for user_id in missing_ids:
            user = user_dict.get(user_id)
            if user:
                student = Student.query.filter_by(student_number=user.username).first()
                if student:
                    students[user_id] = student  # Map old User.id to Student object
    
    # Get all subjects for filter dropdown
    all_subjects = app.config['LEARNING_AREAS']
    
    return render_template("teacher_quiz_results.html", 
                         quizzes=quizzes, 
                         attempts_by_quiz=attempts_by_quiz, 
                         students=students, 
                         quiz_summaries=quiz_summaries,
                         all_subjects=all_subjects,
                         subject_filter=subject_filter)


@app.route("/admin/archive-term", methods=["POST"])
@login_required
@admin_required
def archive_term():
    """Archive assessments for the previous term"""
    settings = Setting.query.first()
    if not settings:
        flash("No settings found", "danger")
        return redirect(url_for("admin_settings"))
    
    # Archive assessments not in current term
    assessments = Assessment.query.filter(
        (Assessment.term != settings.current_term) |
        (Assessment.academic_year != settings.current_academic_year)
    ).filter_by(archived=False).all()
    
    for assessment in assessments:
        assessment.archived = True
    
    db.session.commit()
    flash(f"Archived {len(assessments)} assessments from previous terms", "success")
    return redirect(url_for("admin_settings"))

# -------------------------
# API Endpoints
# -------------------------
@app.route("/api/student_search")
@login_required
def student_search():
    query = request.args.get("q", "").strip()
    
    if not query:
        return jsonify({"results": []})
    
    matches = Student.query.filter(
        (Student.student_number.ilike(f"%{query}%")) |
        (Student.first_name.ilike(f"%{query}%")) |
        (Student.last_name.ilike(f"%{query}%"))
    ).limit(10).all()
    
    results = [
        {
            "student_number": student.student_number,
            "name": student.full_name(),
            "reference_number": student.reference_number
        }
        for student in matches
    ]
    
    return jsonify({"results": results})

@app.route("/api/teacher/assessments")
@login_required
@teacher_required
def teacher_assessments_api():
    """Get assessments for teacher's subject - DIFFERENT NAME to avoid conflict"""
    if not current_user.subject:
        return jsonify({"assessments": []})
    
    assessments = Assessment.query.filter_by(
        subject=current_user.subject,
        teacher_id=current_user.id
    ).order_by(Assessment.date_recorded.desc()).limit(50).all()
    
    result = []
    for a in assessments:
        result.append({
            "student_name": a.student.full_name(),
            "student_number": a.student.student_number,
            "category": a.category,
            "score": a.score,
            "max_score": a.max_score,
            "percentage": a.get_percentage(),
            "class_name": a.class_name,
            "date": a.date_recorded.strftime("%Y-%m-%d")
        })
    
    return jsonify({"assessments": result})

# -------------------------
# Export Routes
# -------------------------
@app.route("/export/csv")
@login_required
def export_csv():
    assessments = Assessment.query.filter_by(archived=False)\
        .order_by(Assessment.date_recorded.desc()).all()
    
    # Create CSV in memory
    si = io.StringIO()
    writer = csv.writer(si)
    
    # Write header
    writer.writerow([
        "student_number",
        "name",
        "category",
        "subject",
        "score",
        "max_score",
        "percentage",
        "term",
        "academic_year",
        "session",
        "assessor",
        "teacher",
        "comments",
        "date_recorded"
    ])
    
    # Write data
    for assessment in assessments:
        teacher_name = assessment.assigned_teacher.username if assessment.assigned_teacher else "N/A"
        writer.writerow([
            assessment.student.student_number,
            assessment.student.full_name(),
            assessment.category,
            assessment.subject,
            assessment.score,
            assessment.max_score,
            f"{assessment.get_percentage():.2f}",
            assessment.term,
            assessment.academic_year,
            assessment.session,
            assessment.assessor,
            teacher_name,
            assessment.comments,
            assessment.date_recorded.strftime("%Y-%m-%d %H:%M:%S")
        ])
    
    # Convert to bytes
    mem = io.BytesIO()
    mem.write(si.getvalue().encode("utf-8"))
    mem.seek(0)
    
    return send_file(
        mem,
        as_attachment=True,
        download_name="assessments_export.csv",
        mimetype="text/csv"
    )
    
@app.route("/export/excel/assessment-template/<int:student_id>")
@login_required
def export_assessment_template(student_id):
    """Export student data to the assessment template Excel format"""
    student = Student.query.get_or_404(student_id)
    
    # Get all assessments for this student
    assessments = Assessment.query.filter_by(student_id=student.id, archived=False).all()
    
    # Create a template path
    template_path = os.path.join(app.config['TEMPLATE_FOLDER'], 'student_template.xlsx')
    
    # If template doesn't exist, create a default one
    if not os.path.exists(template_path):
        # You'll need to copy the actual template file here
        # For now, we'll create a placeholder
        flash("Template file not found. Please upload the template first.", "warning")
        return redirect(url_for('student_view', student_id=student_id))
    
    # Create output filename
    output_filename = f"{student.student_number}_{student.last_name}_assessment.xlsx"
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
    
    try:
        # Initialize template updater
        updater = AssessmentTemplateUpdater(template_path)
        updater.load_template()
        
        # Get student data in template format
        subject = None
        if current_user.is_teacher() and current_user.subject:
            subject = current_user.subject
        student_data = student.to_template_dict(subject)
        
        # Add student to template
        updater.add_student(10, student_data)
        
        # Save the updated workbook
        updater.save_workbook(output_path)
        
        # Send file to user
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except PermissionError:
        flash("The Excel template file is currently open in another program (like Excel). Please close it and try the export again.", "danger")
        return redirect(url_for('student_view', student_id=student_id))
    except Exception as e:
        app.logger.error(f"Error exporting assessment template: {str(e)}")
        flash(f"Error exporting assessment template: {str(e)}", "danger")
        return redirect(url_for('student_view', student_id=student_id))

# Add a route to upload template
@app.route("/upload/template", methods=["GET", "POST"])
@login_required
@admin_required
def upload_template():
    """Upload assessment template Excel file"""
    if request.method == 'POST':
        if 'template_file' not in request.files:
            flash('No file selected', 'danger')
            return redirect(request.url)
        
        file = request.files['template_file']
        if file.filename == '':
            flash('No file selected', 'danger')
            return redirect(request.url)
        
        if file and file.filename.endswith('.xlsx'):
            filename = secure_filename('student_template.xlsx')
            filepath = os.path.join(app.config['TEMPLATE_FOLDER'], filename)
            file.save(filepath)
            flash('Template uploaded successfully', 'success')
            return redirect(url_for('dashboard'))
    
    return render_template("upload_template.html")

@app.route("/export/student/<int:student_id>/csv")
@login_required
def export_student_csv(student_id):
    student = Student.query.get_or_404(student_id)
    
    subject = request.args.get('subject')
    
    # Filter assessments by subject if specified
    assessments = student.assessments
    if subject:
        assessments = [a for a in assessments if a.subject == subject]
    
    # Create CSV in memory
    si = io.StringIO()
    writer = csv.writer(si)
    
    # Write header
    writer.writerow([
        "category",
        "subject",
        "class",
        "score",
        "max_score",
        "percentage",
        "grade",
        "term",
        "academic_year",
        "session",
        "assessor",
        "teacher",
        "comments",
        "date_recorded"
    ])
    
    # Write data
    for assessment in assessments:
        teacher_name = assessment.assigned_teacher.username if assessment.assigned_teacher else "N/A"
        writer.writerow([
            assessment.category,
            assessment.subject,
            assessment.class_name,
            assessment.score,
            assessment.max_score,
            f"{assessment.get_percentage():.2f}",
            assessment.get_grade_letter(),
            assessment.term,
            assessment.academic_year,
            assessment.session,
            assessment.assessor,
            teacher_name,
            assessment.comments,
            assessment.date_recorded.strftime("%Y-%m-%d %H:%M:%S")
        ])
    
    # Convert to bytes
    mem = io.BytesIO()
    mem.write(si.getvalue().encode("utf-8"))
    mem.seek(0)
    
    subject_str = f"_{subject}" if subject else ""
    filename = f"{student.student_number}_{student.last_name}_assessments{subject_str}.csv"
    
    return send_file(
        mem,
        as_attachment=True,
        download_name=filename,
        mimetype="text/csv"
    )

# -------------------------
# Excel Export/Import Routes
# -------------------------
@app.route("/export/excel/student/<int:student_id>")
@login_required
def export_student_excel(student_id):
    """Export single student to Excel template"""
    student = Student.query.get_or_404(student_id)
    
    subject = request.args.get('subject')
    
    # Get or create template path
    template_path = os.path.join(app.config['TEMPLATE_FOLDER'], 'student_template.xlsx')
    
    # Create default template if it doesn't exist
    if not os.path.exists(template_path):
        create_default_template(template_path)
        flash("Default template created. You can customize it in templates_excel folder.", "info")
    
    # Create output file
    subject_str = f"_{subject}" if subject else ""
    output_filename = f"{student.student_number}_{student.last_name}_report{subject_str}.xlsx"
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
    
    try:
        # Get settings
        settings = Setting.query.first()
        
        # Initialize template updater
        updater = AssessmentTemplateUpdater(template_path)
        updater.load_template()
        
        # Update school info
        if settings:
            # Use teacher's subject if available, otherwise use requested subject or student study area
            export_subject = subject
            if hasattr(current_user, 'is_teacher') and current_user.is_teacher() and current_user.subject:
                export_subject = current_user.subject
            elif not export_subject:
                export_subject = student.study_area
            
            updater.update_school_info(
                subject=export_subject,
                term_year=f"{settings.current_term} {settings.current_academic_year}",
                form=student.class_name
            )
        
        # Get student data in template format
        student_data = student.to_template_dict(subject)
        
        # Add student to template
        updater.add_student(10, student_data)
        
        # Save the updated workbook
        updater.save_workbook(output_path)
        
        # Send file to user
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except PermissionError:
        flash("The Excel template file is currently open in another program (like Excel). Please close it and try the export again.", "danger")
        return redirect(url_for('student_view', student_id=student_id))
    except Exception as e:
        flash(f"Error exporting to Excel: {str(e)}", "danger")
        return redirect(url_for('student_view', student_id=student_id))

@app.route("/export/excel/all-students")
@login_required
def export_all_students_excel():
    """Export all students to Excel template"""
    subject = request.args.get('subject')
    class_name = request.args.get('class')
    
    # Filter students based on subject and class
    query = Student.query
    if subject:
        # Get students who have assessments in this subject
        subquery = db.session.query(Assessment.student_id).filter(Assessment.subject == subject).distinct()
        query = query.filter(Student.id.in_(subquery))
    if class_name:
        query = query.filter_by(class_name=class_name)
    
    students = query.order_by(Student.last_name, Student.first_name).all()
    
    # Get or create template path
    template_path = os.path.join(app.config['TEMPLATE_FOLDER'], 'student_template.xlsx')
    
    # Create default template if it doesn't exist
    if not os.path.exists(template_path):
        create_default_template(template_path)
    
    # Create output file
    subject_str = subject or "all_subjects"
    class_str = class_name or "all_classes"
    output_filename = f"students_{subject_str}_{class_str}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
    
    try:
        # Get settings
        settings = Setting.query.first()
        
        # Initialize template updater
        updater = AssessmentTemplateUpdater(template_path)
        updater.load_template()
        
        # Update school info
        form = class_name or "All Classes"
        subj = subject or "All Subjects"
        
        if settings:
            updater.update_school_info(
                subject=subj,
                term_year=f"{settings.current_term} {settings.current_academic_year}",
                form=form
            )
        
        # Get all students data in template format
        students_data = [student.to_template_dict() for student in students]
        
        # Add all students to template
        updater.add_students_batch(students_data)
        
        # Save the updated workbook
        updater.save_workbook(output_path)
        
        # Send file to user
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except PermissionError:
        flash("The Excel template file is currently open in another program (like Excel). Please close it and try the export again.", "danger")
        return redirect(url_for('students'))
    except Exception as e:
        flash(f"Error exporting to Excel: {str(e)}", "danger")
        return redirect(url_for('students'))

@app.route("/export/assessments/excel")
@login_required
def export_assessments_excel():
    """Export filtered assessments to Excel"""
    from openpyxl import Workbook
    
    subject = request.args.get('subject', '')
    class_name = request.args.get('class', '')
    category = request.args.get('category', '')
    
    # Build query based on filters
    query = Assessment.query.filter_by(archived=False)
    if subject:
        query = query.filter_by(subject=subject)
    if class_name:
        query = query.filter_by(class_name=class_name)
    if category:
        query = query.filter_by(category=category)
    
    assessments = query.order_by(Assessment.date_recorded.desc()).all()
    
    # Create output file
    filters = []
    if subject: filters.append(subject)
    if class_name: filters.append(class_name)
    if category: filters.append(category)
    filter_str = "_".join(filters) if filters else "all"
    output_filename = f"assessments_{filter_str}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
    
    try:
        # Create workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Assessments"
        
        # Headers
        headers = [
            "Student Number", "Student Name", "Subject", "Category", 
            "Score", "Max Score", "Percentage", "Grade", "Class", 
            "Term", "Academic Year", "Session", "Assessor", "Date Recorded"
        ]
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=header)
        
        # Data
        for row, assessment in enumerate(assessments, 2):
            teacher_name = assessment.assigned_teacher.username if assessment.assigned_teacher else "N/A"
            ws.cell(row=row, column=1, value=assessment.student.student_number)
            ws.cell(row=row, column=2, value=assessment.student.full_name())
            ws.cell(row=row, column=3, value=assessment.subject)
            ws.cell(row=row, column=4, value=assessment.category)
            ws.cell(row=row, column=5, value=assessment.score)
            ws.cell(row=row, column=6, value=assessment.max_score)
            ws.cell(row=row, column=7, value=round(assessment.get_percentage(), 2))
            ws.cell(row=row, column=8, value=assessment.get_grade_letter())
            ws.cell(row=row, column=9, value=assessment.class_name)
            ws.cell(row=row, column=10, value=assessment.term)
            ws.cell(row=row, column=11, value=assessment.academic_year)
            ws.cell(row=row, column=12, value=assessment.session)
            ws.cell(row=row, column=13, value=assessment.assessor)
            ws.cell(row=row, column=14, value=assessment.date_recorded.strftime("%Y-%m-%d %H:%M:%S"))
        
        # Save
        wb.save(output_path)
        
        # Send file
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype="application/vnd/openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        flash(f"Error exporting to Excel: {str(e)}", "danger")
        return redirect(url_for('assessments_list'))

@app.route("/import/excel", methods=["GET", "POST"])
@login_required
def import_excel():
    """Bulk import assessments from Excel file"""
    form = BulkImportForm()
    
    if form.validate_on_submit():
        file = form.excel_file.data
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        # Save uploaded file
        file.save(filepath)
        
        try:
            # Import assessments
            importer = ExcelBulkImporter(filepath)
            assessments_data = importer.import_assessments()
            
            # Process and save assessments
            success_count = 0
            error_count = 0
            errors = []
            
            for data in assessments_data:
                try:
                    # Find student
                    student = Student.query.filter_by(
                        student_number=data['student_number']
                    ).first()
                    
                    if not student:
                        errors.append(f"Student {data['student_number']} not found")
                        error_count += 1
                        continue
                    
                    # Check if assessment already exists for this student, category, subject, term, academic_year, session
                    existing_assessment = Assessment.query.filter_by(
                        student_id=student.id,
                        category=data['category'],
                        subject=data['subject'],
                        term=data['term'],
                        academic_year=data.get('academic_year'),
                        session=data['session']
                    ).first()
                    
                    if existing_assessment:
                        errors.append(f"Assessment for {data['category']} in {data['subject']} already exists for student {data['student_number']} in the same term, academic year, and session")
                        error_count += 1
                        continue
                    
                    # Create assessment
                    assessment = Assessment(
                        student=student,
                        category=data['category'],
                        subject=data['subject'],
                        score=float(data['score']),
                        max_score=float(data['max_score']),
                        term=data['term'],
                        academic_year=data.get('academic_year'),
                        session=data['session'],
                        assessor=data['assessor'],
                        teacher_id=current_user.id if hasattr(current_user, 'is_teacher') and current_user.is_teacher() else None,
                        comments=data['comments']
                    )
                    db.session.add(assessment)
                    success_count += 1
                    
                except Exception as e:
                    errors.append(f"Row error: {str(e)}")
                    error_count += 1
            
            # Commit all changes
            db.session.commit()
            
            # Clean up uploaded file
            os.remove(filepath)
            
            # Show results
            flash(f"Successfully imported {success_count} assessments", "success")
            if error_count > 0:
                flash(f"{error_count} errors occurred: {'; '.join(errors[:5])}", "warning")
            
            return redirect(url_for('assessments_list'))
            
        except Exception as e:
            db.session.rollback()
            if os.path.exists(filepath):
                os.remove(filepath)
            flash(f"Error importing file: {str(e)}", "danger")
    
    return render_template("import_excel.html", form=form)

@app.route("/download/template/<template_type>")
@login_required
def download_template(template_type):
    """Download Excel template"""
    if template_type == "student":
        template_path = os.path.join(app.config['TEMPLATE_FOLDER'], 'student_template.xlsx')
        filename = "student_assessment_template.xlsx"
    elif template_type == "import":
        template_path = os.path.join(app.config['TEMPLATE_FOLDER'], 'import_template.xlsx')
        filename = "bulk_import_template.xlsx"
    elif template_type == "student_import":
        template_path = os.path.join(app.config['TEMPLATE_FOLDER'], 'student_import_template.xlsx')
        filename = "student_bulk_import_template.xlsx"
    else:
        abort(404)
    
    # For import template, use the existing one, do not create default
    if template_type == "import" and not os.path.exists(template_path):
        flash("Import template not found. Please contact administrator.", "danger")
        return redirect(url_for('import_excel'))
    
    # Create template if it doesn't exist (for student template)
    if template_type == "student" and not os.path.exists(template_path):
        create_default_template(template_path)
    elif template_type == "student_import" and not os.path.exists(template_path):
        create_student_import_template(template_path)
    
    return send_file(
        template_path,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/download/question_template")
@login_required
@teacher_required
def download_question_template():
    """Download question import template"""
    template_path = os.path.join(app.config['TEMPLATE_FOLDER'], 'question_import_template.xlsx')
    
    # Create template if it doesn't exist
    if not os.path.exists(template_path):
        create_question_import_template(template_path)
    
    return send_file(
        template_path,
        as_attachment=True,
        download_name="question_bulk_import_template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -------------------------
# Error Handlers
# -------------------------
@app.errorhandler(403)
def forbidden(e):
    return render_template("403.html"), 403

@app.errorhandler(404)
def not_found(e):
    return render_template("404.html"), 404

@app.errorhandler(500)
def internal_error(e):
    db.session.rollback()
    return render_template("500.html"), 500

# -------------------------
# Context Processors
# -------------------------
@app.context_processor
def inject_config():
    """Make config values available in templates"""
    return {
        'CATEGORY_LABELS': app.config['CATEGORY_LABELS'],
        'ASSESSMENT_WEIGHTS': app.config['ASSESSMENT_WEIGHTS'],
        'LEARNING_AREAS': app.config['LEARNING_AREAS'],
        'CLASS_LEVELS': app.config['CLASS_LEVELS']
    }

# -------------------------
# Run Application
# -------------------------
if __name__ == "__main__":
    print("\n" + "="*60)
    print("EduAssess Module")
    print("="*60)
    print(f"Environment: {env}")
    print(f"Database: {app.config['SQLALCHEMY_DATABASE_URI']}")
    print(f"Access at: http://127.0.0.1:5000")
    print("="*60 + "\n")
    
    app.run(
        debug=app.config.get('DEBUG', True), 
        host='127.0.0.1', 
        port=5000,
        use_reloader=True
    )