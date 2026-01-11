from datetime import datetime
from db import db
from flask_login import UserMixin
import json

class SubjectArea:
    """Helper class to categorize subjects"""
    CORE_SUBJECTS = ['Mathematics', 'English Language', 'General Science', 'Social Studies']
    SCIENCES = ['Biology', 'Chemistry', 'Physics', 'Additional Mathematics']
    ARTS_HUMANITIES = ['History', 'Geography', 'Economics', 'Government', 'Lit in English']
    BUSINESS = ['Business Management', 'Accounting', 'Computing in Business']
    TECHNICAL_VOCATIONAL = [
        'ICT', 'Design and Communication Technology', 
        'Food and Nutrition', 'Clothing and Textile', 'Management in Living'
    ]
    CREATIVE_ARTS = [
        'Arts and Design Foundation', 'Arts and Design Studio', 
        'Music', 'Arts and Design'
    ]
    LANGUAGES = ['French', 'Ghanaian Language']
    RELIGIOUS_STUDIES = ['Christian Religious Studies', 'Religious and Moral Education']
    PHYSICAL_EDUCATION = ['Physical Education and Health']


class User(UserMixin, db.Model):
    __tablename__ = "users"
    
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(20), default="teacher")
    subject = db.Column(db.String(100), nullable=True)
    class_name = db.Column(db.String(50), nullable=True)
    classes = db.Column(db.Text, nullable=True)  # JSON string of multiple classes for teachers
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    # Relationships
    assessments = db.relationship(
        "Assessment",
        backref="assigned_teacher",
        foreign_keys="Assessment.teacher_id",
        lazy=True
    )
    
    def get_classes_list(self):
        """Get list of classes assigned to this teacher"""
        if self.classes:
            try:
                return json.loads(self.classes)
            except json.JSONDecodeError:
                return []
        # Fallback to old single class_name field for backward compatibility
        elif self.class_name:
            return [self.class_name]
        return []
    
    def set_classes_list(self, classes_list):
        """Set list of classes for this teacher"""
        if classes_list and isinstance(classes_list, list):
            self.classes = json.dumps(classes_list)
        else:
            self.classes = None
    
    def check_password(self, password, bcrypt):
        return bcrypt.check_password_hash(self.password_hash, password)
    
    def is_admin(self):
        return self.role == "admin"
    
    def is_teacher(self):
        return self.role == "teacher"
    
    def is_student(self):
        return self.role == "student"
    
    def get_subject_display(self):
        """Return formatted subject name"""
        if not self.subject:
            return None
        return self.subject.replace('_', ' ').title()
    
    def __repr__(self):
        return f"<User id={self.id}>"


class Student(UserMixin, db.Model):
    __tablename__ = "students"
    
    id = db.Column(db.Integer, primary_key=True)
    student_number = db.Column(db.String(50), unique=True, nullable=False, index=True)
    first_name = db.Column(db.String(120), nullable=False)
    last_name = db.Column(db.String(120), nullable=False)
    middle_name = db.Column(db.String(120), nullable=True)
    class_name = db.Column(db.String(50), nullable=True)
    reference_number = db.Column(db.String(50), unique=True, nullable=True, index=True)
    date_of_birth = db.Column(db.Date, nullable=True)
    study_area = db.Column(db.String(50), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    def get_assessments_for_template(self, subject=None):
        """Get assessments formatted for template export"""
        query = Assessment.query.filter_by(student_id=self.id)
        
        if subject:
            query = query.filter_by(subject=subject)
        
        assessments = query.all()
        
        # Initialize template data with zeros
        template_data = {
            'ica1': 0,
            'ica2': 0,
            'icp1': 0,
            'icp2': 0,
            'gp1': 0,
            'gp2': 0,
            'practical': 0,
            'mid_term': 0,
            'end_term': 0
        }
        
        # Map assessments to template columns
        for assessment in assessments:
            category = assessment.category
            if category in template_data:
                # For template, we take the latest assessment score for each category
                # In a real implementation, you might want to aggregate or select specific ones
                template_data[category] = assessment.score
        
        return template_data
    
    def to_template_dict(self, subject=None):
        """Convert student data to template dictionary format"""
        # Format name as Surname Firstname Othername
        name_parts = [self.last_name, self.first_name]
        if self.middle_name:
            name_parts.append(self.middle_name)
        formatted_name = ' '.join(name_parts)
        
        return {
            'name': formatted_name,
            'student_number': self.student_number,
            'ref_id': self.reference_number or self.student_number,
            'study_area': (subject or self.study_area or "Not Specified").upper(),
            **self.get_assessments_for_template(subject)
        }
    
    # Relationships
    assessments = db.relationship(
        "Assessment",
        backref="student",
        cascade="all, delete-orphan",
        lazy=True,
        order_by="Assessment.date_recorded.desc()"
    )
    
    def full_name(self):
        """Return full name of student"""
        if self.middle_name:
            return f"{self.first_name} {self.middle_name} {self.last_name}"
        return f"{self.first_name} {self.last_name}"
    
    def get_assessment_summary(self, subject=None, teacher_id=None):
        assessments = self.assessments
        if subject:
            assessments = [a for a in assessments if a.subject == subject]
        if teacher_id:
            assessments = [a for a in assessments if a.teacher_id == teacher_id]
        
        summary = {}
        for assessment in assessments:
            cat = assessment.category
            if cat not in summary:
                summary[cat] = {
                    "count": 0,
                    "total_score": 0.0,
                    "total_max": 0.0,
                    "avg_percent": 0.0
                }
            summary[cat]["count"] += 1
            summary[cat]["total_score"] += assessment.score
            summary[cat]["total_max"] += assessment.max_score
        
        for cat, data in summary.items():
            if data["count"] > 0:
                data["avg_percent"] = data["total_score"] / data["count"]  # Average raw score
        
        return summary
    
    def get_subject_summary(self, teacher_id=None):
        """Get summary by subject"""
        assessments = self.assessments
        if teacher_id:
            assessments = [a for a in assessments if a.teacher_id == teacher_id]
            
        summary = {}
        for assessment in assessments:
            subject = assessment.subject
            if subject not in summary:
                summary[subject] = {
                    "count": 0,
                    "total_score": 0.0,
                    "total_max": 0.0,
                    "avg_percent": 0.0,
                    "assessments": []
                }
            summary[subject]["count"] += 1
            summary[subject]["total_score"] += assessment.score
            summary[subject]["total_max"] += assessment.max_score
            summary[subject]["assessments"].append(assessment)
        
        for subject, data in summary.items():
            if data["count"] > 0:
                data["avg_percent"] = data["total_score"] / data["count"]  # Average raw score
        
        return summary
    
    def calculate_final_grade(self, weights=None, subject=None, teacher_id=None):
        """Calculate final grade using raw scores to match Excel template calculation"""
        summary = self.get_assessment_summary(subject, teacher_id)
        
        # Template calculation logic - use raw scores directly
        # Class assessments: ica1, ica2, icp1, icp2, gp1, gp2, practical, mid_term
        class_categories = ['ica1', 'ica2', 'icp1', 'icp2', 'gp1', 'gp2', 'practical', 'mid_term']
        class_raw_total = 0.0
        
        for cat in class_categories:
            if cat in summary:
                class_raw_total += summary[cat]["total_score"]
        
        # Class total points (P in template) = min(500, sum of raw scores)
        class_total_points = min(500.0, class_raw_total)
        
        # Class percentage (Q in template) = class_total_points / 500 * 100
        class_percent = (class_total_points / 500.0) * 100
        
        # Class score contribution (R in template) = min(50, roundup(class_percent / 2, 0))
        class_score = min(50.0, round(class_percent / 2))
        
        # Exam assessment: end_term
        exam_raw_score = 0.0
        if 'end_term' in summary:
            exam_raw_score = summary['end_term']["total_score"]
        
        # Exam score contribution (T in template) = min(50, roundup(exam_raw_score / 2, 0))
        exam_score = min(50.0, round(exam_raw_score / 2))
        
        # Final grade (U in template) = min(100, class_score + exam_score)
        final_grade = min(100.0, class_score + exam_score)
        
        return round(final_grade, 2)  # Round to 2 decimal places
    
    def get_gpa_and_grade(self):
        """Calculate GPA and Grade Letter to match Excel template"""
        final_percent = self.calculate_final_grade()
        
        if final_percent is None:
            return {"gpa": "N/A", "grade": "N/A"}
        
        # GPA calculation matching Excel sheet grading scale
        if final_percent >= 80:
            gpa = "4.0"
            grade = "A1"
        elif final_percent >= 70:
            gpa = "3.5"
            grade = "B2"
        elif final_percent >= 65:
            gpa = "3.0"
            grade = "B3"
        elif final_percent >= 60:
            gpa = "2.5"
            grade = "C4"
        elif final_percent >= 55:
            gpa = "2.0"
            grade = "C5"
        elif final_percent >= 50:
            gpa = "1.5"
            grade = "C6"
        elif final_percent >= 45:
            gpa = "1.0"
            grade = "D7"
        elif final_percent >= 40:
            gpa = "0.5"
            grade = "E8"
        else:
            gpa = "0.0"
            grade = "F9"
        
        return {"gpa": gpa, "grade": grade}
    
    def __repr__(self):
        return f"<Student {self.student_number}: {self.full_name()}>"


class Assessment(db.Model):
    __tablename__ = "assessments"
    
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey("students.id"), nullable=False, index=True)
    category = db.Column(db.String(20), nullable=False, index=True)
    subject = db.Column(db.String(120), nullable=False, index=True)
    class_name = db.Column(db.String(50), nullable=True, index=True)
    score = db.Column(db.Float, nullable=False)
    max_score = db.Column(db.Float, nullable=False, default=100.0)
    term = db.Column(db.String(32), nullable=True)
    academic_year = db.Column(db.String(32), nullable=True)
    session = db.Column(db.String(32), nullable=True)
    assessor = db.Column(db.String(120), nullable=True)
    teacher_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    comments = db.Column(db.Text, nullable=True)
    date_recorded = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    archived = db.Column(db.Boolean, default=False, index=True)
    
    def get_percentage(self):
        # Return raw score as percentage to match user expectation
        return self.score
    
    def get_grade_letter(self):
        percentage = self.get_percentage()
        if percentage >= 90:
            return "A+"
        elif percentage >= 80:
            return "A"
        elif percentage >= 75:
            return "B+"
        elif percentage >= 70:
            return "B"
        elif percentage >= 65:
            return "C+"
        elif percentage >= 60:
            return "C"
        elif percentage >= 55:
            return "D+"
        elif percentage >= 50:
            return "D"
        else:
            return "F"
    
    def get_grade_point(self):
        percentage = self.get_percentage()
        if percentage >= 80:
            return 4.0
        elif percentage >= 75:
            return 3.5
        elif percentage >= 70:
            return 3.0
        elif percentage >= 65:
            return 2.5
        elif percentage >= 60:
            return 2.0
        elif percentage >= 55:
            return 1.5
        elif percentage >= 50:
            return 1.0
        else:
            return 0.0
    
    def get_subject_display(self):
        """Return formatted subject name"""
        return self.subject.replace('_', ' ').title()
    
    def __repr__(self):
        return f"<Assessment {self.category} - {self.subject}: {self.score}/{self.max_score}>"


class Setting(db.Model):
    __tablename__ = "settings"
    
    id = db.Column(db.Integer, primary_key=True)
    current_term = db.Column(db.String(32), nullable=False, default='term1')
    current_academic_year = db.Column(db.String(32), nullable=False, default='2024-2025')
    current_session = db.Column(db.String(32), nullable=False, default='First Term')
    assessment_active = db.Column(db.Boolean, default=True)
    
    def __repr__(self):
        return f"<Setting term={self.current_term}, year={self.current_academic_year}>"


class ActivityLog(db.Model):
    __tablename__ = "activity_logs"
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    action = db.Column(db.String(100), nullable=False, index=True)
    details = db.Column(db.Text, nullable=True)
    ip_address = db.Column(db.String(45), nullable=True)  # IPv6 addresses can be up to 45 chars
    timestamp = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    
    # Relationship
    user = db.relationship("User", backref="activity_logs", lazy=True)
    
    def __repr__(self):
        return f"<ActivityLog {self.user.username} - {self.action} at {self.timestamp}>"


class Question(db.Model):
    __tablename__ = "questions"
    
    id = db.Column(db.Integer, primary_key=True)
    subject = db.Column(db.String(120), nullable=False, index=True)
    question_text = db.Column(db.Text, nullable=False)
    question_type = db.Column(db.String(20), nullable=False, default="mcq")  # mcq, true_false, short_answer
    options = db.Column(db.JSON, nullable=True)  # For MCQ: ["A", "B", "C", "D"]
    correct_answer = db.Column(db.String(500), nullable=False)  # For MCQ: "A", for true_false: "True"/"False", for short_answer: the answer
    marks = db.Column(db.Float, nullable=False, default=1.0)  # Marks for the question
    keywords = db.Column(db.JSON, nullable=True)  # For short_answer: list of keywords for flexible marking
    difficulty = db.Column(db.String(20), nullable=False, default="medium")  # easy, medium, hard
    explanation = db.Column(db.Text, nullable=True)  # Optional explanation for the answer
    created_by = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False)
    approved_by = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    status = db.Column(db.String(20), nullable=False, default="pending")  # pending, approved, rejected
    rejection_reason = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    # Relationships
    creator = db.relationship("User", foreign_keys=[created_by], backref="created_questions")
    approver = db.relationship("User", foreign_keys=[approved_by], backref="approved_questions")
    
    def is_approved(self):
        return self.status == "approved"
    
    def can_edit(self, user):
        """Check if user can edit this question"""
        if user.is_admin():
            return True
        if user.is_teacher() and user.id == self.created_by and self.status == "pending":
            return True
        return False
    
    def can_approve(self, user):
        """Check if user can approve/reject this question"""
        return user.is_admin() or (user.is_teacher() and user.subject == self.subject)
    
    def get_subject_display(self):
        """Return formatted subject name"""
        return self.subject.replace('_', ' ').title()
    
    @property
    def normalized_options(self):
        """Return options as a list, handling legacy string format"""
        if not self.options:
            return []
        if isinstance(self.options, list):
            return self.options
        if isinstance(self.options, str):
            # Split by newlines and filter empty lines
            return [line.strip() for line in self.options.split('\n') if line.strip()]
        return []
    
    def __repr__(self):
        return f"<Question {self.id} - {self.subject}: {self.question_text[:50]}...>"


class QuestionAttempt(db.Model):
    __tablename__ = "question_attempts"
    
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey("students.id"), nullable=False)
    question_id = db.Column(db.Integer, db.ForeignKey("questions.id"), nullable=False)
    student_answer = db.Column(db.String(500), nullable=False)
    is_correct = db.Column(db.Boolean, nullable=False)
    score = db.Column(db.Float, nullable=False, default=0.0)
    attempted_at = db.Column(db.DateTime, default=datetime.utcnow)
    time_taken = db.Column(db.Integer, nullable=True)  # Time in seconds
    
    # Relationships
    student = db.relationship("Student", backref="question_attempts")
    question = db.relationship("Question", backref="attempts")
    
    def __repr__(self):
        return f"<QuestionAttempt student={self.student_id} question={self.question_id} correct={self.is_correct}>"


class Quiz(db.Model):
    __tablename__ = "quizzes"
    
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    subject = db.Column(db.String(120), nullable=False, index=True)
    description = db.Column(db.Text, nullable=True)
    questions = db.Column(db.JSON, nullable=False)  # List of question IDs
    time_limit = db.Column(db.Integer, nullable=True)  # Time limit in minutes
    created_by = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False)
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    # Relationships
    creator = db.relationship("User", foreign_keys=[created_by], backref="created_quizzes")
    
    def get_subject_display(self):
        """Return formatted subject name"""
        return self.subject.replace('_', ' ').title()
    
    def __repr__(self):
        return f"<Quiz {self.title} - {self.subject}>"


class QuizAttempt(db.Model):
    __tablename__ = "quiz_attempts"
    
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey("students.id"), nullable=False)
    quiz_id = db.Column(db.Integer, db.ForeignKey("quizzes.id"), nullable=False)
    score = db.Column(db.Float, nullable=False)
    total_questions = db.Column(db.Integer, nullable=False)
    correct_answers = db.Column(db.Integer, nullable=False)
    started_at = db.Column(db.DateTime, default=datetime.utcnow)
    completed_at = db.Column(db.DateTime, nullable=True)
    time_taken = db.Column(db.Integer, nullable=True)  # Time in seconds
    status = db.Column(db.String(20), default="in_progress")  # in_progress, completed
    answers_json = db.Column(db.Text, nullable=True)  # Store partial answers as JSON
    remaining_time = db.Column(db.Integer, nullable=True)  # Remaining time in seconds
    
    # Relationships
    student = db.relationship("Student", backref="quiz_attempts")
    quiz = db.relationship("Quiz", backref="attempts")
    
    def get_percentage(self):
        if self.total_questions > 0:
            return (self.correct_answers / self.total_questions) * 100
        return 0
    
    def __repr__(self):
        return f"<QuizAttempt student={self.student_id} quiz={self.quiz_id} score={self.score}>"


def init_db(app, bcrypt):
    db.init_app(app)
    
    with app.app_context():
        db.create_all()
        
        # Create default settings if not exist
        if not Setting.query.first():
            default_settings = Setting(
                current_term='term1',
                current_academic_year='2024-2025',
                current_session='First Term'
            )
            db.session.add(default_settings)
            db.session.commit()
        
        if User.query.count() == 0:
            default_username = app.config.get("DEFAULT_ADMIN_USERNAME", "admin")
            default_password = app.config.get("DEFAULT_ADMIN_PASSWORD", "Admin@123")
            
            hashed = bcrypt.generate_password_hash(default_password).decode("utf-8")
            admin = User(
                username=default_username,
                password_hash=hashed,
                role="admin"
            )
            db.session.add(admin)
            db.session.commit()
            
            print(f"\n{'='*60}")
            print(f"Created default admin account:")
            print(f"  Username: {default_username}")
            print(f"  Password: {default_password}")
            print(f"  ** CHANGE THIS PASSWORD IMMEDIATELY **")
            print(f"{'='*60}\n")