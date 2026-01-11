# Quick Setup Guide - EDUASSESS MODULE PRO

## Step 1: Run Migration
```bash
# Activate virtual environment first
cd c:\Users\HP\Documents\school_assess_app_EXPERIMENTAL_ver_1
python migrations_teacher_classes.py
```

Expected output:
```
Starting migration to support multiple teacher classes...

Adding 'classes' column to users table...
✓ 'classes' column added successfully
✓ Migrated X teacher(s) to multiple classes format

Migration completed successfully!
```

## Step 2: Restart Application
After migration, restart your Flask application:
```bash
# Ctrl+C to stop current server
python app.py
```

## Step 3: Test the Feature

### Test 1: Create New Teacher with Multiple Classes
1. Login as admin
2. Go to **Settings → User Management**
3. Click **Add New User**
4. Fill in:
   - Username: `testteacher`
   - Password: `Test@123`
   - Role: `teacher`
   - Subject: Select any subject
   - Classes: Hold Ctrl and select 2-3 classes
5. Click **Create User**
6. Verify in user list that the teacher shows multiple classes

### Test 2: Edit Existing Teacher's Classes
1. Go to **User Management**
2. Find a teacher user
3. Click **Edit** button
4. In Classes field, select multiple classes
5. Click **Update User**
6. Verify the changes appear in the user list

### Test 3: Assign Classes via Subject Assignment
1. Go to **User Management**
2. Find a teacher
3. Click the **Book Icon** (📖)
4. Select subject and multiple classes
5. Click **Save Assignment**
6. Verify in user list

## Multi-Select Instructions for Users

### On Windows/Linux:
- Hold **Ctrl** and click to select multiple options
- Click once to select
- Click again (while holding Ctrl) to deselect

### On Mac:
- Hold **Cmd** (⌘) and click to select multiple options
- Click once to select
- Click again (while holding Cmd) to deselect

## Files Modified

1. `models.py` - Added `classes` column and helper methods
2. `app.py` - Updated forms and route handlers
3. `templates/user_form.html` - Multiple class selection
4. `templates/edit_user.html` - Multiple class selection
5. `templates/teacher_subject.html` - Multiple class selection
6. `templates/users.html` - Display multiple classes
7. `migrations_teacher_classes.py` - New migration script

## Troubleshooting

### Issue: "classes" column already exists error
- This is fine, the migration handles this gracefully

### Issue: Classes not showing after edit
- Verify the migration completed successfully
- Check that the form is being filled correctly: `form.classes.data = list_of_classes`

### Issue: Old single class still shows
- The system automatically checks both old and new fields
- After first edit with new form, the new field will be used

### To Reset (Optional):
If you need to test the migration again:
1. Delete the migration script execution
2. Drop the `classes` column: `ALTER TABLE users DROP COLUMN classes;`
3. Run migration again

## API/Code Usage

### Get teacher's classes (in templates):
```django
{% set classes = teacher.get_classes_list() %}
{% if classes %}
    Classes: {{ classes|join(', ') }}
{% endif %}
```

### Get teacher's classes (in Python):
```python
teacher = User.query.get(teacher_id)
classes = teacher.get_classes_list()  # Returns list
```

### Set teacher's classes (in Python):
```python
teacher = User.query.get(teacher_id)
teacher.set_classes_list(['Form 1', 'Form 2'])  # Sets JSON
db.session.commit()
```

## Support

For issues or questions, check:
1. Migration log output
2. Database connection
3. Form validation errors in browser console
4. Application logs
