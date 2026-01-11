"""
Migration script to add multiple classes support for teachers.
This script adds a 'classes' column to the users table and migrates
single class_name values to the new JSON format if needed.

Run this after updating the code but before using the application:
    python migrations_teacher_classes.py
"""

from app import app, db
from models import User
import json
import sys

def migrate_to_multiple_classes():
    """Migrate teacher class assignments from single class_name to multiple classes"""
    
    with app.app_context():
        try:
            # Check if classes column exists
            inspector = db.inspect(db.engine)
            columns = [col['name'] for col in inspector.get_columns('users')]
            
            if 'classes' not in columns:
                print("Adding 'classes' column to users table...")
                with db.engine.connect() as conn:
                    # Add the classes column
                    conn.execute(db.text('ALTER TABLE users ADD COLUMN classes TEXT'))
                    conn.commit()
                print("✓ 'classes' column added successfully")
            else:
                print("✓ 'classes' column already exists")
            
            # Migrate existing data from class_name to classes
            teachers = User.query.filter_by(role='teacher').all()
            migrated_count = 0
            
            for teacher in teachers:
                if teacher.class_name and not teacher.classes:
                    # Convert single class_name to classes JSON array
                    teacher.set_classes_list([teacher.class_name])
                    migrated_count += 1
            
            if migrated_count > 0:
                db.session.commit()
                print(f"✓ Migrated {migrated_count} teacher(s) to multiple classes format")
            else:
                print("✓ No migration needed - all teachers already using multiple classes format")
            
            print("\nMigration completed successfully!")
            return True
            
        except Exception as e:
            print(f"✗ Migration failed: {str(e)}", file=sys.stderr)
            db.session.rollback()
            return False

if __name__ == '__main__':
    print("Starting migration to support multiple teacher classes...\n")
    success = migrate_to_multiple_classes()
    sys.exit(0 if success else 1)
