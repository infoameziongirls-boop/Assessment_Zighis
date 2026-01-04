import os

class Config:
    """Base configuration"""
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'dev-secret-key-change-in-production'
    SQLALCHEMY_DATABASE_URI = os.environ.get('DATABASE_URL') or 'sqlite:///assessment.db'
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    
    # User roles
    USER_ROLES = [
        ('admin', 'Administrator'),
        ('teacher', 'Teacher')
    ]
    
    # Template configuration
    TEMPLATE_FOLDER = 'templates_excel'
    ASSESSMENT_TEMPLATE_FILE = 'student_template.xlsx'
    
    # Template mapping - map your assessment categories to template columns
    TEMPLATE_CATEGORY_MAPPING = {
        'individual_assignment': 'ica',
        'quiz': 'ica',
        'homework': 'ica',
        'project': 'icp',
        'research_paper': 'icp',
        'group_assignment': 'gp',
        'presentation': 'gp',
        'lab_report': 'practical',
        'midterm_exam': 'mid_term',
        'final_exam': 'end_term'
    }
    
    # Assessment categories - updated to match Excel template columns
    ASSESSMENT_CATEGORIES = [
        ('ica1', 'Individual Assessment 1'),
        ('ica2', 'Individual Assessment 2'),
        ('icp1', 'Individual Class Project 1'),
        ('icp2', 'Individual Class Project 2'),
        ('gp1', 'Group Project/Research 1'),
        ('gp2', 'Group Project/Research 2'),
        ('practical', 'Practical Portfolio'),
        ('mid_term', 'Mid-Semester Exam'),
        ('end_term', 'End of Term Exam')
    ]
    
    # Category max scores
    CATEGORY_MAX_SCORES = {
        'ica1': 50.0,
        'ica2': 50.0,
        'icp1': 50.0,
        'icp2': 50.0,
        'gp1': 50.0,
        'gp2': 50.0,
        'practical': 100.0,
        'mid_term': 100.0,
        'end_term': 100.0
    }
    
    # Terms
    TERMS = [
        ('term1', 'Term 1'),
        ('term2', 'Term 2'),
        ('term3', 'Term 3')
    ]
    
    # Study/Learning Areas
    STUDY_AREAS = [
        ('science_a', 'SCIENCE A'),
        ('science_b', 'SCIENCE B'),
        ('visual_performing_arts', 'VISUAL AND PERFORMING ARTS'),
        ('home_economics_a', 'HOME ECONOMICS A'),
        ('home_economics_b', 'HOME ECONOMICS B'),
        ('home_economics_c', 'HOME ECONOMICS C'),
        ('home_economics_e', 'HOME ECONOMICS E'),
        ('home_economics_f', 'HOME ECONOMICS F'),
        ('general_arts_1a', 'GENERAL ARTS 1_A'),
        ('general_arts_2b', 'GENERAL ARTS 2_B'),
        ('general_arts_3a_c', 'GENERAL ARTS 3a_C'),
        ('general_arts_3b_d', 'GENERAL ARTS 3b_D'),
        ('general_arts_4a_e', 'GENERAL ARTS 4a_E'),
        ('general_arts_4b_f', 'GENERAL ARTS 4b_F'),
        ('general_arts_5a_g', 'GENERAL ARTS 5a_G'),
        ('general_arts_5b_h', 'GENERAL ARTS 5b_H'),
        ('general_arts_6a_i', 'GENERAL ARTS 6a_I'),
        ('general_arts_6b_j', 'GENERAL ARTS 6b_J'),
        ('business_a', 'BUSINESS A'),
        ('business_b', 'BUSINESS B'),
        ('business_c', 'BUSINESS C'),
        ('business_d', 'BUSINESS D')
    ]
    
    # Complete Learning Areas/Subjects
    LEARNING_AREAS = [
        ('mathematics', 'Mathematics'),
        ('social_studies', 'Social Studies'),
        ('general_science', 'General Science'),
        ('english_language', 'English Language'),
        ('ict', 'ICT'),
        ('physical_education_health', 'Physical Education and Health'),
        ('music', 'Music'),
        ('additional_mathematics', 'Additional Mathematics'),
        ('biology', 'Biology'),
        ('chemistry', 'Chemistry'),
        ('physics', 'Physics'),
        ('geography', 'Geography'),
        ('economics', 'Economics'),
        ('french', 'French'),
        ('arts_design_studio', 'Arts and Design Studio'),
        ('arts_design_foundation', 'Arts and Design Foundation'),
        ('design_communication_technology', 'Design and Communication Technology'),
        ('clothing_textile', 'Clothing and Textile'),
        ('management_in_living', 'Management in Living'),
        ('food_nutrition', 'Food and Nutrition'),
        ('lit_in_english', 'Lit in English'),
        ('christian_religious_studies', 'Christian Religious Studies'),
        ('history', 'History'),
        ('ghanaian_language', 'Ghanaian Language'),
        ('religious_moral_education', 'Religious and Moral Education'),
        ('government', 'Government'),
        ('business_management', 'Business Management'),
        ('accounting', 'Accounting'),
        ('computing_in_business', 'Computing in Business')
    ]
    
    # Class levels
    CLASS_LEVELS = [
        ('form1', 'Form 1'),
        ('form2', 'Form 2'),
        ('form3', 'Form 3')
    ]
    
    # Assessment weights for final grade calculation - updated for new categories
    ASSESSMENT_WEIGHTS = {
        'ica1': 0.1,
        'ica2': 0.1,
        'icp1': 0.1,
        'icp2': 0.1,
        'gp1': 0.1,
        'gp2': 0.1,
        'practical': 0.15,
        'mid_term': 0.15,
        'end_term': 0.15
    }
    
    # Category display labels - updated to match new categories
    CATEGORY_LABELS = {
        'ica1': 'Individual Assessment 1',
        'ica2': 'Individual Assessment 2',
        'icp1': 'Individual Class Project 1',
        'icp2': 'Individual Class Project 2',
        'gp1': 'Group Project/Research 1',
        'gp2': 'Group Project/Research 2',
        'practical': 'Practical Portfolio',
        'mid_term': 'Mid-Semester Exam',
        'end_term': 'End of Term Exam'
    }
    
    # Pagination
    ASSESSMENTS_PER_PAGE = 25
    STUDENTS_PER_PAGE = 30
    
    # Default admin credentials
    DEFAULT_ADMIN_USERNAME = 'admin'
    DEFAULT_ADMIN_PASSWORD = 'Admin@123'
    
    # Student login credentials
    DEFAULT_STUDENT_PASSWORD = 'Student@123'


class DevelopmentConfig(Config):
    """Development configuration"""
    DEBUG = True
    SQLALCHEMY_DATABASE_URI = 'sqlite:///assessment_dev.db'


class TestingConfig(Config):
    """Testing configuration"""
    TESTING = True
    SQLALCHEMY_DATABASE_URI = 'sqlite:///assessment_test.db'
    WTF_CSRF_ENABLED = False


class ProductionConfig(Config):
    """Production configuration"""
    DEBUG = False
    SQLALCHEMY_DATABASE_URI = os.environ.get('DATABASE_URL')


config = {
    'development': DevelopmentConfig,
    'testing': TestingConfig,
    'production': ProductionConfig,
    'default': DevelopmentConfig
}