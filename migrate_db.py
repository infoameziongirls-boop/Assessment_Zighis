import sys
import os
sys.path.append(os.path.dirname(__file__))

from app import app, db, bcrypt
from models import User, Student, Assessment, Setting, ActivityLog, Question, QuestionAttempt, Quiz, QuizAttempt

def migrate_database():
    with app.app_context():
        # Add new columns if they don't exist
        # This is a simple migration - in production use Alembic
        
        # For SQLite, we need to recreate tables
        # This will drop existing data, so backup first!
        print("WARNING: This will recreate all tables and lose existing data!")
        response = input("Continue? (y/n): ")
        
        if response.lower() != 'y':
            print("Migration cancelled.")
            return
        
        # Drop all tables
        db.drop_all()
        
        # Create all tables with new schema
        db.create_all()
        
        # Create default admin
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
        
        print("Database migrated successfully!")
        print(f"Admin account: {default_username} / {default_password}")

if __name__ == "__main__":
    migrate_database()