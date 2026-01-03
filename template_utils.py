# template_utils.py
import os
from template_updater import AssessmentTemplateUpdater
from flask import current_app

def export_students_to_template(students_data, subject=None, class_name=None):
    """
    Export multiple students to assessment template
    
    Args:
        students_data: List of student dictionaries or Student objects
        subject: Subject to filter assessments
        class_name: Class to filter
        
    Returns:
        Path to the generated Excel file
    """
    # Get template path
    template_path = os.path.join(
        current_app.config['TEMPLATE_FOLDER'],
        current_app.config['ASSESSMENT_TEMPLATE_FILE']
    )
    
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template file not found: {template_path}")
    
    # Create output file
    from datetime import datetime
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_filename = f"assessment_export_{timestamp}.xlsx"
    output_path = os.path.join(current_app.config['UPLOAD_FOLDER'], output_filename)
    
    # Initialize template updater
    updater = AssessmentTemplateUpdater(template_path)
    updater.load_template()
    
    # Update school info if needed
    # updater.update_school_info(school_name="School Name", subject=subject)
    
    # Add students in batch
    template_students_data = []
    for i, student in enumerate(students_data):
        if hasattr(student, 'to_template_dict'):
            student_data = student.to_template_dict(subject)
        else:
            student_data = student
        
        template_students_data.append(student_data)
    
    updater.add_students_batch(template_students_data)
    
    # Save the workbook
    updater.save_workbook(output_path)
    
    return output_path, output_filename