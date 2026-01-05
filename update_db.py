import sys
import os
sys.path.append(os.path.dirname(__file__))

from app import app, db
from models import User, Student, Assessment

def update_database():
    """Update database schema without losing data"""
    with app.app_context():
        print("Checking database schema...")
        
        # Check if the database exists
        try:
            # Try to query existing tables
            existing_users = User.query.first()
            print(f"Found existing database with {User.query.count()} users")
            print(f"Found {Student.query.count()} students")
            print(f"Found {Assessment.query.count()} assessments")
            
            # Check if new columns exist
            from sqlalchemy import inspect
            inspector = inspect(db.engine)
            
            # Check columns in users table
            user_columns = [col['name'] for col in inspector.get_columns('users')]
            print("Current user columns:", user_columns)
            
            # Add missing columns if needed
            if 'subject' not in user_columns:
                print("Adding 'subject' column to users table...")
                db.engine.execute('ALTER TABLE users ADD COLUMN subject VARCHAR(100)')
            
            if 'class_name' not in user_columns:
                print("Adding 'class_name' column to users table...")
                db.engine.execute('ALTER TABLE users ADD COLUMN class_name VARCHAR(50)')
            
            # Check columns in students table
            student_columns = [col['name'] for col in inspector.get_columns('students')]
            print("Current student columns:", student_columns)
            
            if 'middle_name' not in student_columns:
                print("Adding 'middle_name' column to students table...")
                db.engine.execute('ALTER TABLE students ADD COLUMN middle_name VARCHAR(120)')
            
            if 'class_name' not in student_columns:
                print("Adding 'class_name' column to students table...")
                db.engine.execute('ALTER TABLE students ADD COLUMN class_name VARCHAR(50)')
            
            if 'reference_number' not in student_columns:
                print("Adding 'reference_number' column to students table...")
                db.engine.execute('ALTER TABLE students ADD COLUMN reference_number VARCHAR(50)')
            
            # Check columns in questions table
            question_columns = [col['name'] for col in inspector.get_columns('questions')]
            print("Current question columns:", question_columns)
            
            if 'marks' not in question_columns:
                print("Adding 'marks' column to questions table...")
                db.engine.execute('ALTER TABLE questions ADD COLUMN marks FLOAT DEFAULT 1.0')
            
            if 'keywords' not in question_columns:
                print("Adding 'keywords' column to questions table...")
                db.engine.execute('ALTER TABLE questions ADD COLUMN keywords TEXT')
            
            # Check columns in question_attempts table
            attempt_columns = [col['name'] for col in inspector.get_columns('question_attempts')]
            print("Current question_attempt columns:", attempt_columns)
            
            if 'score' not in attempt_columns:
                print("Adding 'score' column to question_attempts table...")
                db.engine.execute('ALTER TABLE question_attempts ADD COLUMN score FLOAT DEFAULT 0.0')
                # Create index
                db.engine.execute('CREATE INDEX IF NOT EXISTS ix_students_reference_number ON students (reference_number)')
            
            if 'date_of_birth' not in student_columns:
                print("Adding 'date_of_birth' column to students table...")
                db.engine.execute('ALTER TABLE students ADD COLUMN date_of_birth DATE')
            
            if 'gender' not in student_columns:
                print("Adding 'gender' column to students table...")
                db.engine.execute('ALTER TABLE students ADD COLUMN gender VARCHAR(10)')
            
            # Check columns in assessments table
            assessment_columns = [col['name'] for col in inspector.get_columns('assessments')]
            print("Current assessment columns:", assessment_columns)
            
            if 'class_name' not in assessment_columns:
                print("Adding 'class_name' column to assessments table...")
                db.engine.execute('ALTER TABLE assessments ADD COLUMN class_name VARCHAR(50)')
                # Create index
                db.engine.execute('CREATE INDEX IF NOT EXISTS ix_assessments_class_name ON assessments (class_name)')
            
            if 'teacher_id' not in assessment_columns:
                print("Adding 'teacher_id' column to assessments table...")
                db.engine.execute('ALTER TABLE assessments ADD COLUMN teacher_id INTEGER')
                # Add foreign key constraint
                db.engine.execute('ALTER TABLE assessments ADD CONSTRAINT fk_assessments_teacher_id FOREIGN KEY (teacher_id) REFERENCES users (id)')
            
            if 'academic_year' not in assessment_columns:
                print("Adding 'academic_year' column to assessments table...")
                db.engine.execute('ALTER TABLE assessments ADD COLUMN academic_year VARCHAR(32)')
            
            print("\nDatabase update completed successfully!")
            print("="*60)
            
        except Exception as e:
            print(f"Error: {e}")
            print("Creating new database...")
            
            # Create all tables from scratch
            db.drop_all()
            db.create_all()
            
            # Create default admin
            default_username = app.config.get("DEFAULT_ADMIN_USERNAME", "admin")
            default_password = app.config.get("DEFAULT_ADMIN_PASSWORD", "Admin@123")
            
            from flask_bcrypt import Bcrypt
            bcrypt = Bcrypt()
            hashed = bcrypt.generate_password_hash(default_password).decode("utf-8")
            admin = User(
                username=default_username,
                password_hash=hashed,
                role="admin"
            )
            db.session.add(admin)
            db.session.commit()
            
            print(f"\nCreated new database with default admin account:")
            print(f"  Username: {default_username}")
            print(f"  Password: {default_password}")
            print("="*60)

if __name__ == "__main__":
    update_database()