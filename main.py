# main_enhanced_gui.py - Student Tracker with Modern GUI
# Programmed by: Younes Bennacer
# Enhanced Professional Edition with Android Permissions & Native File Picker

import os
import sqlite3
import pandas as pd
from datetime import datetime, timedelta
from collections import defaultdict
import logging
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.spinner import Spinner
from kivy.uix.progressbar import ProgressBar
from kivy.core.window import Window
from kivy.metrics import dp, sp
from kivy.graphics import Color, Rectangle, Line, RoundedRectangle, Ellipse
from kivy.clock import Clock
from kivy.animation import Animation
from kivy.properties import StringProperty, NumericProperty, ListProperty
import threading

# ============================================
# ANDROID PERMISSIONS & FILE PICKER HANDLER
# ============================================
from kivy.utils import platform

if platform == 'android':
    from android.permissions import request_permissions, Permission, check_permission
    from android.storage import primary_external_storage_path
    from android import activity
    from jnius import autoclass, cast
    
    # Android classes for native file picker
    PythonActivity = autoclass('org.kivy.android.PythonActivity')
    Intent = autoclass('android.content.Intent')
    Uri = autoclass('android.net.Uri')
    ContentResolver = autoclass('android.content.ContentResolver')
    Environment = autoclass('android.os.Environment')
    VERSION = autoclass('android.os.Build$VERSION')
    Settings = autoclass('android.provider.Settings')
    
    # Define all required permissions
    REQUIRED_PERMISSIONS = [
        Permission.READ_EXTERNAL_STORAGE,
        Permission.WRITE_EXTERNAL_STORAGE,
    ]
    
    # Global variable to store file picker callback
    _file_picker_callback = None
    
    def on_activity_result(request_code, result_code, intent):
        """
        Handle result from Android file picker.
        This is called when user selects a file.
        """
        global _file_picker_callback
        
        if request_code == 42:  # File picker request code
            if result_code == -1 and intent is not None:  # RESULT_OK = -1
                try:
                    uri = intent.getData()
                    if uri:
                        # Get actual file path from URI
                        file_path = get_path_from_uri(uri)
                        
                        if file_path and _file_picker_callback:
                            _file_picker_callback(file_path)
                        else:
                            logger.error("Could not get file path from URI")
                            if _file_picker_callback:
                                _file_picker_callback(None)
                except Exception as e:
                    logger.error(f"Error processing file picker result: {str(e)}")
                    if _file_picker_callback:
                        _file_picker_callback(None)
            else:
                logger.info("File picker canceled or failed")
                if _file_picker_callback:
                    _file_picker_callback(None)
    
    def get_path_from_uri(uri):
        """
        Convert Android content URI to actual file path.
        Copies file from content:// to app cache if necessary.
        """
        try:
            # Get content resolver
            context = PythonActivity.mActivity
            content_resolver = context.getContentResolver()
            
            # Get file name from URI
            cursor = None
            try:
                # Try to get display name
                OpenableColumns = autoclass('android.provider.OpenableColumns')
                cursor = content_resolver.query(uri, None, None, None, None)
                
                if cursor and cursor.moveToFirst():
                    name_index = cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME)
                    if name_index >= 0:
                        file_name = cursor.getString(name_index)
                    else:
                        file_name = "imported_file.xlsx"
                else:
                    file_name = "imported_file.xlsx"
            finally:
                if cursor:
                    cursor.close()
            
            # Create cache file path
            cache_dir = context.getCacheDir().getAbsolutePath()
            cache_file_path = os.path.join(cache_dir, file_name)
            
            # Copy content to cache file
            input_stream = content_resolver.openInputStream(uri)
            
            # Write to cache file using Python
            with open(cache_file_path, 'wb') as f:
                File = autoclass('java.io.File')
                FileOutputStream = autoclass('java.io.FileOutputStream')
                
                # Read from input stream
                buffer_size = 8192
                buffer = bytearray(buffer_size)
                
                while True:
                    bytes_read = input_stream.read(buffer, 0, buffer_size)
                    if bytes_read <= 0:
                        break
                    f.write(buffer[:bytes_read])
            
            input_stream.close()
            
            logger.info(f"File copied to cache: {cache_file_path}")
            return cache_file_path
            
        except Exception as e:
            logger.error(f"Error converting URI to path: {str(e)}")
            return None
    
    def open_android_file_picker(callback):
        """
        Open Android's native file picker using SAF (Storage Access Framework).
        This is the PROPER way to access files on modern Android.
        """
        global _file_picker_callback
        _file_picker_callback = callback
        
        try:
            # Register activity result handler
            activity.bind(on_activity_result=on_activity_result)
            
            # Create intent for file picking
            intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
            intent.addCategory(Intent.CATEGORY_OPENABLE)
            
            # Set MIME type for Excel files
            intent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            # Also allow older .xls format
            mime_types = [
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/vnd.ms-excel"
            ]
            intent.putExtra(Intent.EXTRA_MIME_TYPES, mime_types)
            
            # Start activity for result
            PythonActivity.mActivity.startActivityForResult(intent, 42)
            
        except Exception as e:
            logger.error(f"Error opening file picker: {str(e)}")
            if callback:
                callback(None)
    
    def request_android_permissions():
        """
        Request all necessary Android permissions at app startup.
        This is the CATALYST that enables permissions in Settings.
        """
        try:
            # Check if we already have permissions
            has_read = check_permission(Permission.READ_EXTERNAL_STORAGE)
            has_write = check_permission(Permission.WRITE_EXTERNAL_STORAGE)
            
            if not has_read or not has_write:
                logger.info("Requesting storage permissions...")
                request_permissions(REQUIRED_PERMISSIONS, permission_callback)
            else:
                logger.info("Storage permissions already granted")
                
            # For Android 11+, check if we need MANAGE_EXTERNAL_STORAGE
            if VERSION.SDK_INT >= 30:  # Android 11+
                if not Environment.isExternalStorageManager():
                    logger.warning("MANAGE_EXTERNAL_STORAGE not granted - showing dialog")
                    # Schedule dialog to avoid blocking startup
                    Clock.schedule_once(lambda dt: show_manage_storage_dialog(), 2.0)
                else:
                    logger.info("MANAGE_EXTERNAL_STORAGE already granted")
                    
        except Exception as e:
            logger.error(f"Error requesting permissions: {str(e)}")
    
    def permission_callback(permissions, grant_results):
        """
        Callback when user responds to permission request.
        """
        for permission, granted in zip(permissions, grant_results):
            if granted:
                logger.info(f"Permission granted: {permission}")
            else:
                logger.warning(f"Permission denied: {permission}")
                Clock.schedule_once(lambda dt: show_permission_denied_dialog(permission), 0.5)
    
    def show_manage_storage_dialog():
        """
        Show dialog informing user about MANAGE_EXTERNAL_STORAGE permission.
        For Android 11+, this permission must be granted manually in Settings.
        """
        try:
            def open_settings(instance):
                # Create intent to open app settings
                intent = Intent(Settings.ACTION_MANAGE_APP_ALL_FILES_ACCESS_PERMISSION)
                uri = Uri.parse(f"package:{PythonActivity.mActivity.getPackageName()}")
                intent.setData(uri)
                PythonActivity.mActivity.startActivity(intent)
                popup.dismiss()
            
            content = BoxLayout(orientation='vertical', spacing=dp(15), padding=dp(20))
            
            message = Label(
                text=(
                    "[b]Additional Permission for Android 11+[/b]\n\n"
                    "To import/export Excel files on Android 11+, "
                    "please enable 'All Files Access' permission.\n\n"
                    "Note: You can also use the built-in file picker "
                    "which works without this permission."
                ),
                markup=True,
                halign='center',
                valign='middle',
                color=(0.098, 0.110, 0.129, 1)
            )
            message.bind(size=message.setter('text_size'))
            content.add_widget(message)
            
            btn_layout = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(10))
            
            open_btn = Button(
                text='Open Settings',
                background_color=(0.204, 0.459, 0.765, 1),
                color=(1, 1, 1, 1)
            )
            open_btn.bind(on_press=open_settings)
            btn_layout.add_widget(open_btn)
            
            later_btn = Button(
                text='Use File Picker Instead',
                background_color=(0.435, 0.451, 0.478, 1),
                color=(1, 1, 1, 1)
            )
            later_btn.bind(on_press=lambda x: popup.dismiss())
            btn_layout.add_widget(later_btn)
            
            content.add_widget(btn_layout)
            
            popup = Popup(
                title='File Access Permission',
                content=content,
                size_hint=(0.85, 0.5),
                auto_dismiss=False
            )
            popup.open()
            
        except Exception as e:
            logger.error(f"Error showing manage storage dialog: {str(e)}")
    
    def show_permission_denied_dialog(permission):
        """
        Show dialog when a permission is denied.
        """
        content = BoxLayout(orientation='vertical', spacing=dp(15), padding=dp(20))
        
        message = Label(
            text=(
                f"[b]Permission Denied[/b]\n\n"
                f"Storage permission is needed for import/export features.\n\n"
                f"You can still use the app's native file picker, "
                f"or enable permissions in:\n"
                f"Settings > Apps > Student Tracker Pro > Permissions"
            ),
            markup=True,
            halign='center',
            valign='middle',
            color=(0.098, 0.110, 0.129, 1)
        )
        message.bind(size=message.setter('text_size'))
        content.add_widget(message)
        
        ok_btn = Button(
            text='OK',
            size_hint=(None, None),
            size=(dp(120), dp(45)),
            pos_hint={'center_x': 0.5},
            background_color=(0.204, 0.459, 0.765, 1),
            color=(1, 1, 1, 1)
        )
        
        popup = Popup(
            title='Information',
            content=content,
            size_hint=(0.8, 0.4)
        )
        
        ok_btn.bind(on_press=popup.dismiss)
        content.add_widget(ok_btn)
        popup.open()
    
    def get_external_storage_path():
        """
        Get the external storage path for file operations.
        """
        try:
            return primary_external_storage_path()
        except:
            return '/storage/emulated/0'
else:
    # Non-Android platforms
    def request_android_permissions():
        pass
    
    def get_external_storage_path():
        return os.path.expanduser('~')
    
    def open_android_file_picker(callback):
        # On non-Android, callback with None
        if callback:
            callback(None)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('student_tracker.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ============================================
# MODERN COLOR SCHEME - Professional & Elegant
# ============================================
# Primary Colors - Modern Blue Gradient
PRIMARY_DARK = (0.098, 0.278, 0.478, 1)      # Deep blue #19478a
PRIMARY_COLOR = (0.204, 0.459, 0.765, 1)      # Royal blue #3475c3
PRIMARY_LIGHT = (0.282, 0.565, 0.847, 1)      # Sky blue #4890d8

# Accent Colors
ACCENT_COLOR = (0.157, 0.706, 0.647, 1)       # Teal #28b4a5
ACCENT_LIGHT = (0.204, 0.804, 0.741, 1)       # Light teal #34cdbd

# Background Colors
BACKGROUND_DARK = (0.945, 0.953, 0.961, 1)    # Very light gray #f1f3f5
BACKGROUND_COLOR = (0.976, 0.980, 0.984, 1)   # Almost white #f9fafb
CARD_COLOR = (1, 1, 1, 1)                      # Pure white

# Text Colors
TEXT_PRIMARY = (0.098, 0.110, 0.129, 1)       # Dark gray #191c21
TEXT_SECONDARY = (0.435, 0.451, 0.478, 1)     # Medium gray #6f737a
TEXT_LIGHT = (0.647, 0.667, 0.690, 1)         # Light gray #a5aab0

# Semantic Colors
SUCCESS_COLOR = (0.125, 0.749, 0.408, 1)      # Green #20bf68
WARNING_COLOR = (1, 0.643, 0.020, 1)          # Orange #ffa405
ERROR_COLOR = (0.910, 0.176, 0.227, 1)        # Red #e82d3a
INFO_COLOR = (0.204, 0.459, 0.765, 1)         # Blue #3475c3

# Special Colors
HEADER_GRADIENT_START = (0.141, 0.376, 0.647, 1)  # Dark blue
HEADER_GRADIENT_END = (0.204, 0.459, 0.765, 1)    # Royal blue

# UI Element Colors
BUTTON_PRIMARY = PRIMARY_COLOR
BUTTON_SECONDARY = ACCENT_COLOR
BUTTON_HOVER = PRIMARY_LIGHT
SHADOW_COLOR = (0, 0, 0, 0.1)

# ============================================
# CONFIGURATION
# ============================================
class Config:
    """Application configuration"""
    # App Info
    APP_NAME = "Student Tracker Pro"
    APP_VERSION = "2.1"
    DEVELOPER = "Younes Bennacer"
    YEAR = "2024"
    
    # Excel import settings
    POSSIBLE_SHEET_NAMES = ['note', 'noteDataTable1', 'Sheet1', 'Feuil1', 'notes']
    REQUIRED_COLUMNS = ['Matricule', 'Nom', 'Prénom']
    
    # Pagination
    STUDENTS_PER_PAGE = 50
    
    # Backup
    AUTO_BACKUP_INTERVAL = 3600
    BACKUP_FOLDER = 'backups'
    
    # Validation
    MATRICULE_LENGTH = 12
    MIN_SCORE = 0
    MAX_SCORE = 20
    
    # Database
    DB_NAME = 'student_tracker.db'

# ============================================
# UTILITY FUNCTIONS
# ============================================
def validate_matricule(matricule):
    """Validate student matricule format"""
    if not matricule:
        return False, "Matricule cannot be empty"
    
    matricule = str(matricule).strip()
    
    if len(matricule) != Config.MATRICULE_LENGTH:
        return False, f"Matricule must be {Config.MATRICULE_LENGTH} characters"
    
    if not matricule.isdigit():
        return False, "Matricule must contain only numbers"
    
    return True, ""

def validate_score(score):
    """Validate score/grade"""
    if score is None or score == '':
        return True, ""
    
    try:
        score_float = float(score)
        if score_float < Config.MIN_SCORE or score_float > Config.MAX_SCORE:
            return False, f"Score must be between {Config.MIN_SCORE} and {Config.MAX_SCORE}"
        return True, ""
    except ValueError:
        return False, "Score must be a number"

# ============================================
# ENHANCED DATABASE HANDLER
# ============================================
class StudentTrackerDB:
    """Enhanced database handler"""
    
    def __init__(self, db_name=Config.DB_NAME):
        # On Android, store database in external storage
        if platform == 'android':
            storage_path = get_external_storage_path()
            app_folder = os.path.join(storage_path, 'StudentTrackerPro')
            os.makedirs(app_folder, exist_ok=True)
            self.db_name = os.path.join(app_folder, db_name)
        else:
            self.db_name = db_name
            
        self.init_database()
        logger.info(f"Database initialized: {self.db_name}")
    
    def init_database(self):
        """Initialize database with required tables and indexes"""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        
        try:
            # Students table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS students (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    matricule TEXT UNIQUE NOT NULL,
                    nom TEXT NOT NULL,
                    prenom TEXT NOT NULL,
                    section TEXT,
                    groupe TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # Classes table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS classes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    course_name TEXT NOT NULL,
                    subject_name TEXT,
                    class_date DATE NOT NULL,
                    groupe TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # Attendance table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS attendance (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    student_id INTEGER NOT NULL,
                    class_id INTEGER NOT NULL,
                    status TEXT CHECK(status IN ('Present', 'Absent', 'Absent Justifié')) DEFAULT 'Present',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
                    FOREIGN KEY(class_id) REFERENCES classes(id) ON DELETE CASCADE,
                    UNIQUE(student_id, class_id)
                )
            ''')
            
            # Marks table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS marks (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    student_id INTEGER NOT NULL,
                    class_id INTEGER NOT NULL,
                    score REAL CHECK(score >= 0 AND score <= 20),
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
                    FOREIGN KEY(class_id) REFERENCES classes(id) ON DELETE CASCADE,
                    UNIQUE(student_id, class_id)
                )
            ''')
            
            # Comments table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS comments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    student_id INTEGER NOT NULL,
                    class_id INTEGER NOT NULL,
                    comment TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
                    FOREIGN KEY(class_id) REFERENCES classes(id) ON DELETE CASCADE
                )
            ''')
            
            # Create indexes for performance
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_student_matricule ON students(matricule)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_student_groupe ON students(groupe)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_attendance_student ON attendance(student_id)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_attendance_class ON attendance(class_id)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_marks_student ON marks(student_id)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_marks_class ON marks(class_id)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_comments_student ON comments(student_id)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_comments_class ON comments(class_id)')
            
            conn.commit()
            logger.info("Database tables and indexes created successfully")
            
        except sqlite3.Error as e:
            logger.error(f"Database initialization error: {str(e)}")
            conn.rollback()
        finally:
            conn.close()
    
    def add_student(self, matricule, nom, prenom, section=None, groupe=None):
        """Add a new student"""
        # Validate matricule
        is_valid, error_msg = validate_matricule(matricule)
        if not is_valid:
            return False, error_msg
        
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
                INSERT INTO students (matricule, nom, prenom, section, groupe)
                VALUES (?, ?, ?, ?, ?)
            ''', (matricule.strip(), nom.strip(), prenom.strip(), section, groupe))
            
            conn.commit()
            logger.info(f"Student added: {matricule} - {nom} {prenom}")
            return True, "Student added successfully"
            
        except sqlite3.IntegrityError:
            return False, f"Student with matricule {matricule} already exists"
        except sqlite3.Error as e:
            logger.error(f"Error adding student: {str(e)}")
            return False, f"Database error: {str(e)}"
        finally:
            conn.close()
    
    def get_all_students(self, groupe=None, search_term=None, offset=0, limit=50):
        """Get all students with optional filtering and pagination"""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        
        try:
            query = "SELECT * FROM students WHERE 1=1"
            params = []
            
            if groupe:
                query += " AND groupe = ?"
                params.append(groupe)
            
            if search_term:
                query += " AND (matricule LIKE ? OR nom LIKE ? OR prenom LIKE ?)"
                search_pattern = f"%{search_term}%"
                params.extend([search_pattern, search_pattern, search_pattern])
            
            query += " ORDER BY nom, prenom LIMIT ? OFFSET ?"
            params.extend([limit, offset])
            
            cursor.execute(query, params)
            students = cursor.fetchall()
            
            # Get total count
            count_query = "SELECT COUNT(*) FROM students WHERE 1=1"
            count_params = []
            
            if groupe:
                count_query += " AND groupe = ?"
                count_params.append(groupe)
            
            if search_term:
                count_query += " AND (matricule LIKE ? OR nom LIKE ? OR prenom LIKE ?)"
                count_params.extend([search_pattern, search_pattern, search_pattern])
            
            cursor.execute(count_query, count_params)
            total_count = cursor.fetchone()[0]
            
            return students, total_count
            
        except sqlite3.Error as e:
            logger.error(f"Error fetching students: {str(e)}")
            return [], 0
        finally:
            conn.close()
    
    def get_student_by_id(self, student_id):
        """Get student by ID"""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        
        try:
            cursor.execute("SELECT * FROM students WHERE id = ?", (student_id,))
            return cursor.fetchone()
        except sqlite3.Error as e:
            logger.error(f"Error fetching student: {str(e)}")
            return None
        finally:
            conn.close()
    
    def update_student(self, student_id, matricule, nom, prenom, section=None, groupe=None):
        """Update student information"""
        # Validate matricule
        is_valid, error_msg = validate_matricule(matricule)
        if not is_valid:
            return False, error_msg
        
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
                UPDATE students
                SET matricule = ?, nom = ?, prenom = ?, section = ?, groupe = ?, 
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            ''', (matricule.strip(), nom.strip(), prenom.strip(), section, groupe, student_id))
            
            conn.commit()
            
            if cursor.rowcount > 0:
                logger.info(f"Student updated: {student_id}")
                return True, "Student updated successfully"
            else:
                return False, "Student not found"
                
        except sqlite3.IntegrityError:
            return False, f"Student with matricule {matricule} already exists"
        except sqlite3.Error as e:
            logger.error(f"Error updating student: {str(e)}")
            return False, f"Database error: {str(e)}"
        finally:
            conn.close()
    
    def delete_student(self, student_id):
        """Delete a student and all related records"""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        
        try:
            cursor.execute("DELETE FROM students WHERE id = ?", (student_id,))
            conn.commit()
            
            if cursor.rowcount > 0:
                logger.info(f"Student deleted: {student_id}")
                return True, "Student and all related records deleted successfully"
            else:
                return False, "Student not found"
                
        except sqlite3.Error as e:
            logger.error(f"Error deleting student: {str(e)}")
            return False, f"Database error: {str(e)}"
        finally:
            conn.close()
    
    def get_all_groupes(self):
        """Get all unique group names"""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        
        try:
            cursor.execute("SELECT DISTINCT groupe FROM students WHERE groupe IS NOT NULL ORDER BY groupe")
            groupes = [row[0] for row in cursor.fetchall()]
            return groupes
        except sqlite3.Error as e:
            logger.error(f"Error fetching groups: {str(e)}")
            return []
        finally:
            conn.close()
    
    def import_from_excel(self, file_path, groupe_name=None, progress_callback=None):
        """Import students from Excel file"""
        try:
            logger.info(f"Attempting to import from: {file_path}")
            
            # Check if file exists
            if not os.path.exists(file_path):
                return False, f"File not found: {file_path}", 0
            
            # Read Excel file
            excel_file = pd.ExcelFile(file_path)
            sheet_name = None
            
            # Find the correct sheet
            for possible_name in Config.POSSIBLE_SHEET_NAMES:
                if possible_name in excel_file.sheet_names:
                    sheet_name = possible_name
                    break
            
            if sheet_name is None:
                sheet_name = excel_file.sheet_names[0]
            
            logger.info(f"Reading sheet: {sheet_name}")
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Validate required columns
            missing_columns = [col for col in Config.REQUIRED_COLUMNS if col not in df.columns]
            if missing_columns:
                return False, f"Missing required columns: {', '.join(missing_columns)}", 0
            
            # Import students
            success_count = 0
            error_count = 0
            total = len(df)
            
            for idx, row in df.iterrows():
                try:
                    matricule = str(row['Matricule']).strip()
                    nom = str(row['Nom']).strip()
                    prenom = str(row['Prénom']).strip()
                    section = str(row.get('Section', '')).strip() if 'Section' in row else None
                    
                    # Use provided group name or from Excel
                    if groupe_name:
                        groupe = groupe_name
                    else:
                        groupe = str(row.get('Groupe', '')).strip() if 'Groupe' in row else None
                    
                    success, _ = self.add_student(matricule, nom, prenom, section, groupe)
                    
                    if success:
                        success_count += 1
                    else:
                        error_count += 1
                    
                    if progress_callback:
                        progress = (idx + 1) / total
                        progress_callback(progress)
                        
                except Exception as e:
                    logger.error(f"Error importing row {idx}: {str(e)}")
                    error_count += 1
            
            message = f"Import complete: {success_count} students added"
            if error_count > 0:
                message += f", {error_count} errors"
            
            logger.info(message)
            return True, message, success_count
            
        except Exception as e:
            error_msg = f"Error reading Excel file: {str(e)}"
            logger.error(error_msg)
            return False, error_msg, 0
    
    def export_to_excel(self, output_path, groupe=None):
        """Export students to Excel file"""
        try:
            conn = sqlite3.connect(self.db_name)
            
            if groupe:
                query = "SELECT * FROM students WHERE groupe = ? ORDER BY nom, prenom"
                df = pd.read_sql_query(query, conn, params=(groupe,))
            else:
                query = "SELECT * FROM students ORDER BY nom, prenom"
                df = pd.read_sql_query(query, conn)
            
            conn.close()
            
            # Save to Excel
            df.to_excel(output_path, index=False, sheet_name='Students')
            
            return True, f"Data exported successfully to {output_path}"
            
        except Exception as e:
            logger.error(f"Excel export error: {str(e)}")
            return False, f"Error exporting to Excel: {str(e)}"
    
    def backup_database(self):
        """Create a backup of the database"""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_filename = f"backup_{timestamp}.db"
            
            # Determine backup path
            if platform == 'android':
                storage_path = get_external_storage_path()
                backup_folder = os.path.join(storage_path, 'StudentTrackerPro', Config.BACKUP_FOLDER)
            else:
                backup_folder = Config.BACKUP_FOLDER
            
            os.makedirs(backup_folder, exist_ok=True)
            backup_path = os.path.join(backup_folder, backup_filename)
            
            # Copy database file
            import shutil
            shutil.copy2(self.db_name, backup_path)
            
            logger.info(f"Database backed up to: {backup_path}")
            return True, backup_path
            
        except Exception as e:
            logger.error(f"Backup error: {str(e)}")
            return False, f"Backup failed: {str(e)}"
    
    def get_student_statistics(self, student_id):
        """Get comprehensive statistics for a student"""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        
        try:
            stats = {
                'total_classes': 0,
                'present_count': 0,
                'absent_count': 0,
                'justified_count': 0,
                'attendance_rate': 0,
                'total_marks': 0,
                'average_score': 0,
                'highest_score': 0,
                'lowest_score': 0
            }
            
            # Attendance statistics
            cursor.execute('''
                SELECT status, COUNT(*) 
                FROM attendance 
                WHERE student_id = ? 
                GROUP BY status
            ''', (student_id,))
            
            for status, count in cursor.fetchall():
                if status == 'Present':
                    stats['present_count'] = count
                elif status == 'Absent':
                    stats['absent_count'] = count
                elif status == 'Absent Justifié':
                    stats['justified_count'] = count
            
            stats['total_classes'] = stats['present_count'] + stats['absent_count'] + stats['justified_count']
            
            if stats['total_classes'] > 0:
                stats['attendance_rate'] = (stats['present_count'] / stats['total_classes']) * 100
            
            # Marks statistics
            cursor.execute('''
                SELECT COUNT(*), AVG(score), MAX(score), MIN(score)
                FROM marks
                WHERE student_id = ?
            ''', (student_id,))
            
            result = cursor.fetchone()
            if result and result[0] > 0:
                stats['total_marks'] = result[0]
                stats['average_score'] = round(result[1], 2) if result[1] else 0
                stats['highest_score'] = result[2] if result[2] else 0
                stats['lowest_score'] = result[3] if result[3] else 0
            
            return stats
            
        except sqlite3.Error as e:
            logger.error(f"Error getting student statistics: {str(e)}")
            return None
        finally:
            conn.close()

# ============================================
# CUSTOM UI COMPONENTS
# ============================================
class ModernButton(Button):
    """Modern styled button with hover effect"""
    
    def __init__(self, button_color=BUTTON_PRIMARY, **kwargs):
        super().__init__(**kwargs)
        self.button_color = button_color
        self.background_normal = ''
        self.background_color = button_color
        self.color = (1, 1, 1, 1)
        self.font_size = sp(14)
        self.bold = True
        self.size_hint_y = None
        self.height = dp(45)
        
        # Add rounded corners
        with self.canvas.before:
            Color(*button_color)
            self.rect = RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(8)])
        
        self.bind(pos=self._update_rect, size=self._update_rect)
    
    def _update_rect(self, instance, value):
        self.rect.pos = self.pos
        self.rect.size = self.size
    
    def on_press(self):
        # Darken on press
        darker_color = tuple(c * 0.8 for c in self.button_color[:3]) + (1,)
        self.background_color = darker_color
    
    def on_release(self):
        # Restore original color
        self.background_color = self.button_color

class ModernCard(BoxLayout):
    """Modern card container with shadow"""
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.padding = dp(15)
        self.spacing = dp(10)
        
        with self.canvas.before:
            # Shadow
            Color(*SHADOW_COLOR)
            self.shadow = RoundedRectangle(
                pos=(self.x + dp(2), self.y - dp(2)),
                size=self.size,
                radius=[dp(10)]
            )
            
            # Card background
            Color(*CARD_COLOR)
            self.rect = RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(10)])
        
        self.bind(pos=self._update_rect, size=self._update_rect)
    
    def _update_rect(self, instance, value):
        self.rect.pos = self.pos
        self.rect.size = self.size
        self.shadow.pos = (self.x + dp(2), self.y - dp(2))
        self.shadow.size = self.size

class ModernLabel(Label):
    """Modern styled label"""
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.color = TEXT_PRIMARY
        self.font_size = sp(14)

class HeaderLabel(Label):
    """Header label with gradient background"""
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.color = (1, 1, 1, 1)
        self.font_size = sp(16)
        self.bold = True
        self.size_hint_y = None
        self.height = dp(50)
        
        with self.canvas.before:
            Color(*HEADER_GRADIENT_START)
            self.rect = Rectangle(pos=self.pos, size=self.size)
        
        self.bind(pos=self._update_rect, size=self._update_rect)
    
    def _update_rect(self, instance, value):
        self.rect.pos = self.pos
        self.rect.size = self.size

class LoadingPopup(Popup):
    """Loading popup with progress bar"""
    
    def __init__(self, **kwargs):
        content = BoxLayout(orientation='vertical', spacing=dp(15), padding=dp(20))
        
        self.message_label = Label(
            text='Loading...',
            color=TEXT_PRIMARY,
            font_size=sp(14)
        )
        content.add_widget(self.message_label)
        
        self.progress_bar = ProgressBar(max=1, value=0, size_hint_y=None, height=dp(30))
        content.add_widget(self.progress_bar)
        
        super().__init__(
            content=content,
            size_hint=(0.7, 0.3),
            auto_dismiss=False,
            **kwargs
        )
    
    def update_progress(self, value, message=""):
        self.progress_bar.value = value
        if message:
            self.message_label.text = message

class ConfirmationDialog(Popup):
    """Confirmation dialog"""
    
    def __init__(self, message, on_yes, on_no=None, **kwargs):
        content = BoxLayout(orientation='vertical', spacing=dp(15), padding=dp(20))
        
        # Message
        msg_label = Label(
            text=message,
            color=TEXT_PRIMARY,
            font_size=sp(14),
            halign='center',
            valign='middle'
        )
        msg_label.bind(size=msg_label.setter('text_size'))
        content.add_widget(msg_label)
        
        # Buttons
        btn_layout = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(10))
        
        yes_btn = ModernButton(
            text='Yes',
            button_color=SUCCESS_COLOR
        )
        yes_btn.bind(on_press=lambda x: self._on_yes(on_yes))
        btn_layout.add_widget(yes_btn)
        
        no_btn = ModernButton(
            text='No',
            button_color=ERROR_COLOR
        )
        no_btn.bind(on_press=lambda x: self._on_no(on_no))
        btn_layout.add_widget(no_btn)
        
        content.add_widget(btn_layout)
        
        super().__init__(
            title='Confirmation',
            content=content,
            size_hint=(0.7, 0.4),
            **kwargs
        )
    
    def _on_yes(self, callback):
        if callback:
            callback()
        self.dismiss()
    
    def _on_no(self, callback):
        if callback:
            callback()
        self.dismiss()

def show_success(message, title='Success'):
    """Show success popup"""
    content = BoxLayout(orientation='vertical', spacing=dp(15), padding=dp(20))
    
    msg_label = Label(
        text=message,
        color=TEXT_PRIMARY,
        font_size=sp(14),
        halign='center',
        valign='middle'
    )
    msg_label.bind(size=msg_label.setter('text_size'))
    content.add_widget(msg_label)
    
    ok_btn = ModernButton(
        text='OK',
        size_hint=(None, None),
        size=(dp(120), dp(45)),
        pos_hint={'center_x': 0.5},
        button_color=SUCCESS_COLOR
    )
    
    popup = Popup(
        title=title,
        content=content,
        size_hint=(0.7, 0.3)
    )
    
    ok_btn.bind(on_press=popup.dismiss)
    content.add_widget(ok_btn)
    popup.open()

def show_error(message, title='Error'):
    """Show error popup"""
    content = BoxLayout(orientation='vertical', spacing=dp(15), padding=dp(20))
    
    msg_label = Label(
        text=message,
        color=ERROR_COLOR,
        font_size=sp(14),
        halign='center',
        valign='middle'
    )
    msg_label.bind(size=msg_label.setter('text_size'))
    content.add_widget(msg_label)
    
    ok_btn = ModernButton(
        text='OK',
        size_hint=(None, None),
        size=(dp(120), dp(45)),
        pos_hint={'center_x': 0.5},
        button_color=ERROR_COLOR
    )
    
    popup = Popup(
        title=title,
        content=content,
        size_hint=(0.7, 0.3)
    )
    
    ok_btn.bind(on_press=popup.dismiss)
    content.add_widget(ok_btn)
    popup.open()

def show_info(message, title='Information'):
    """Show info popup"""
    content = BoxLayout(orientation='vertical', spacing=dp(15), padding=dp(20))
    
    msg_label = Label(
        text=message,
        color=TEXT_PRIMARY,
        font_size=sp(14),
        halign='center',
        valign='middle'
    )
    msg_label.bind(size=msg_label.setter('text_size'))
    content.add_widget(msg_label)
    
    ok_btn = ModernButton(
        text='OK',
        size_hint=(None, None),
        size=(dp(120), dp(45)),
        pos_hint={'center_x': 0.5},
        button_color=INFO_COLOR
    )
    
    popup = Popup(
        title=title,
        content=content,
        size_hint=(0.7, 0.3)
    )
    
    ok_btn.bind(on_press=popup.dismiss)
    content.add_widget(ok_btn)
    popup.open()

# ============================================
# MAIN SCREEN
# ============================================
class MainScreen(Screen):
    """Main application screen"""
    
    def __init__(self, db, **kwargs):
        super().__init__(**kwargs)
        self.db = db
        self.current_page = 0
        self.students_per_page = Config.STUDENTS_PER_PAGE
        self.total_students = 0
        self.selected_groupe = None
        self.search_mode = False
        self.search_term = ""
        self.pending_import_groupe = None  # Store group name for import
        
        self.build_ui()
        self.refresh_groups()
    
    def build_ui(self):
        """Build the main UI"""
        main_layout = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(15))
        main_layout.canvas.before.clear()
        
        with main_layout.canvas.before:
            Color(*BACKGROUND_COLOR)
            Rectangle(pos=main_layout.pos, size=main_layout.size)
        
        # Header
        header = self.create_header()
        main_layout.add_widget(header)
        
        # Controls section
        controls = self.create_controls()
        main_layout.add_widget(controls)
        
        # Students list
        self.students_container = BoxLayout(orientation='vertical', size_hint_y=0.8)
        main_layout.add_widget(self.students_container)
        
        # Pagination
        self.pagination_layout = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(10))
        main_layout.add_widget(self.pagination_layout)
        
        self.add_widget(main_layout)
    
    def create_header(self):
        """Create header with app title"""
        header_card = ModernCard(size_hint_y=None, height=dp(80))
        
        title_layout = BoxLayout(orientation='vertical', spacing=dp(5))
        
        title = Label(
            text=f'[b]{Config.APP_NAME}[/b]',
            markup=True,
            color=PRIMARY_COLOR,
            font_size=sp(24),
            size_hint_y=None,
            height=dp(35)
        )
        
        subtitle = Label(
            text=f'Professional Student Management System - v{Config.APP_VERSION}',
            color=TEXT_SECONDARY,
            font_size=sp(12),
            size_hint_y=None,
            height=dp(25)
        )
        
        title_layout.add_widget(title)
        title_layout.add_widget(subtitle)
        
        header_card.add_widget(title_layout)
        
        return header_card
    
    def create_controls(self):
        """Create control buttons"""
        controls_card = ModernCard(size_hint_y=None, height=dp(150))
        controls_layout = BoxLayout(orientation='vertical', spacing=dp(10))
        
        # Group selector and search
        top_row = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(10))
        
        top_row.add_widget(ModernLabel(text='Select Group:', size_hint_x=0.2))
        
        self.groupe_spinner = Spinner(
            text='All Groups',
            values=['All Groups'],
            size_hint_x=0.3,
            background_color=CARD_COLOR,
            color=TEXT_PRIMARY,
            font_size=sp(14)
        )
        self.groupe_spinner.bind(text=self.on_groupe_selected)
        top_row.add_widget(self.groupe_spinner)
        
        self.search_input = TextInput(
            hint_text='Search (Matricule, Name...)',
            size_hint_x=0.4,
            multiline=False,
            background_color=(0.95, 0.95, 0.96, 1),
            foreground_color=TEXT_PRIMARY,
            font_size=sp(14),
            padding=[dp(10), dp(12)]
        )
        self.search_input.bind(on_text_validate=self.search_students)
        top_row.add_widget(self.search_input)
        
        search_btn = ModernButton(
            text='Search',
            size_hint_x=0.15,
            button_color=INFO_COLOR
        )
        search_btn.bind(on_press=self.search_students)
        top_row.add_widget(search_btn)
        
        controls_layout.add_widget(top_row)
        
        # Action buttons
        action_row = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(10))
        
        add_btn = ModernButton(
            text='➕ Add Student',
            button_color=SUCCESS_COLOR
        )
        add_btn.bind(on_press=self.show_add_student_dialog)
        action_row.add_widget(add_btn)
        
        import_btn = ModernButton(
            text='📥 Import Excel',
            button_color=PRIMARY_COLOR
        )
        import_btn.bind(on_press=self.show_import_dialog)
        action_row.add_widget(import_btn)
        
        export_btn = ModernButton(
            text='📤 Export Excel',
            button_color=ACCENT_COLOR
        )
        export_btn.bind(on_press=self.export_data)
        action_row.add_widget(export_btn)
        
        backup_btn = ModernButton(
            text='💾 Backup',
            button_color=WARNING_COLOR
        )
        backup_btn.bind(on_press=self.backup_database)
        action_row.add_widget(backup_btn)
        
        refresh_btn = ModernButton(
            text='🔄 Refresh',
            button_color=INFO_COLOR
        )
        refresh_btn.bind(on_press=lambda x: self.refresh_data())
        action_row.add_widget(refresh_btn)
        
        controls_layout.add_widget(action_row)
        
        controls_card.add_widget(controls_layout)
        
        return controls_card
    
    def refresh_groups(self):
        """Refresh available groups"""
        groupes = self.db.get_all_groupes()
        self.groupe_spinner.values = ['All Groups'] + groupes
    
    def on_groupe_selected(self, spinner, text):
        """Handle group selection"""
        if text == 'All Groups':
            self.selected_groupe = None
        else:
            self.selected_groupe = text
        
        self.current_page = 0
        self.load_students()
    
    def load_students(self):
        """Load and display students"""
        offset = self.current_page * self.students_per_page
        
        if self.search_mode:
            students, total = self.db.get_all_students(
                groupe=self.selected_groupe,
                search_term=self.search_term,
                offset=offset,
                limit=self.students_per_page
            )
        else:
            students, total = self.db.get_all_students(
                groupe=self.selected_groupe,
                offset=offset,
                limit=self.students_per_page
            )
        
        self.total_students = total
        self.display_students(students)
        self.update_pagination()
    
    def display_students(self, students):
        """Display students in a scrollable list"""
        self.students_container.clear_widgets()
        
        if not students:
            no_data = Label(
                text='No students found',
                color=TEXT_SECONDARY,
                font_size=sp(16)
            )
            self.students_container.add_widget(no_data)
            return
        
        # Create scrollable view
        scroll = ScrollView(size_hint=(1, 1))
        students_grid = GridLayout(cols=1, spacing=dp(5), size_hint_y=None, padding=dp(5))
        students_grid.bind(minimum_height=students_grid.setter('height'))
        
        # Header
        header = BoxLayout(size_hint_y=None, height=dp(40), spacing=dp(5))
        header_fields = ['ID', 'Matricule', 'Last Name', 'First Name', 'Section', 'Group', 'Actions']
        
        for field in header_fields:
            lbl = HeaderLabel(text=f'[b]{field}[/b]', markup=True)
            header.add_widget(lbl)
        
        students_grid.add_widget(header)
        
        # Student rows
        for idx, student in enumerate(students):
            row = self.create_student_row(student, idx)
            students_grid.add_widget(row)
        
        scroll.add_widget(students_grid)
        self.students_container.add_widget(scroll)
    
    def create_student_row(self, student, index):
        """Create a single student row"""
        row = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(5))
        
        # Alternate row colors
        bg_color = CARD_COLOR if index % 2 == 0 else BACKGROUND_DARK
        
        with row.canvas.before:
            Color(*bg_color)
            row.rect = RoundedRectangle(pos=row.pos, size=row.size, radius=[dp(5)])
        
        row.bind(pos=lambda instance, value: setattr(row.rect, 'pos', instance.pos))
        row.bind(size=lambda instance, value: setattr(row.rect, 'size', instance.size))
        
        # Student data
        fields = [
            str(student[0]),  # ID
            str(student[1]),  # Matricule
            str(student[2]),  # Nom
            str(student[3]),  # Prenom
            str(student[4] or ''),  # Section
            str(student[5] or '')  # Groupe
        ]
        
        for field in fields:
            lbl = Label(
                text=field,
                color=TEXT_PRIMARY,
                font_size=sp(13)
            )
            row.add_widget(lbl)
        
        # Action buttons
        actions = BoxLayout(spacing=dp(5))
        
        view_btn = Button(
            text='👁',
            size_hint_x=0.33,
            background_color=INFO_COLOR,
            color=(1, 1, 1, 1),
            font_size=sp(16)
        )
        view_btn.bind(on_press=lambda x: self.view_student_details(student))
        actions.add_widget(view_btn)
        
        edit_btn = Button(
            text='✏',
            size_hint_x=0.33,
            background_color=WARNING_COLOR,
            color=(1, 1, 1, 1),
            font_size=sp(16)
        )
        edit_btn.bind(on_press=lambda x: self.show_edit_student_dialog(student))
        actions.add_widget(edit_btn)
        
        delete_btn = Button(
            text='🗑',
            size_hint_x=0.33,
            background_color=ERROR_COLOR,
            color=(1, 1, 1, 1),
            font_size=sp(16)
        )
        delete_btn.bind(on_press=lambda x: self.confirm_delete_student(student))
        actions.add_widget(delete_btn)
        
        row.add_widget(actions)
        
        return row
    
    def update_pagination(self):
        """Update pagination controls"""
        self.pagination_layout.clear_widgets()
        
        total_pages = (self.total_students + self.students_per_page - 1) // self.students_per_page
        
        if total_pages <= 1:
            return
        
        # Previous button
        prev_btn = ModernButton(
            text='← Previous',
            size_hint_x=0.2,
            button_color=PRIMARY_COLOR if self.current_page > 0 else TEXT_LIGHT
        )
        if self.current_page > 0:
            prev_btn.bind(on_press=lambda x: self.change_page(-1))
        self.pagination_layout.add_widget(prev_btn)
        
        # Page info
        page_info = Label(
            text=f'Page {self.current_page + 1} of {total_pages} ({self.total_students} students)',
            color=TEXT_PRIMARY,
            font_size=sp(14)
        )
        self.pagination_layout.add_widget(page_info)
        
        # Next button
        next_btn = ModernButton(
            text='Next →',
            size_hint_x=0.2,
            button_color=PRIMARY_COLOR if self.current_page < total_pages - 1 else TEXT_LIGHT
        )
        if self.current_page < total_pages - 1:
            next_btn.bind(on_press=lambda x: self.change_page(1))
        self.pagination_layout.add_widget(next_btn)
    
    def change_page(self, direction):
        """Change current page"""
        self.current_page += direction
        self.load_students()
    
    def search_students(self, instance):
        """Search for students"""
        search_text = self.search_input.text.strip()
        
        if search_text:
            self.search_mode = True
            self.search_term = search_text
            self.current_page = 0
            self.load_students()
        else:
            self.clear_search()
    
    def clear_search(self):
        """Clear search and reload all students"""
        self.search_mode = False
        self.search_term = ""
        self.search_input.text = ""
        self.current_page = 0
        self.load_students()
    
    def show_add_student_dialog(self, instance):
        """Show dialog to add new student"""
        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(15))
        
        # Input fields
        fields = {}
        field_names = [
            ('Matricule', True),
            ('Last Name', True),
            ('First Name', True),
            ('Section', False),
            ('Group', False)
        ]
        
        for field_name, required in field_names:
            field_card = ModernCard(size_hint_y=None, height=dp(60))
            field_layout = BoxLayout(spacing=dp(10))
            
            label_text = f'{field_name}:' + (' *' if required else '')
            field_layout.add_widget(ModernLabel(
                text=label_text,
                size_hint_x=0.3
            ))
            
            text_input = TextInput(
                size_hint_x=0.7,
                multiline=False,
                background_color=(0.95, 0.95, 0.96, 1),
                foreground_color=TEXT_PRIMARY,
                font_size=sp(14),
                padding=[dp(10), dp(12)]
            )
            fields[field_name] = text_input
            field_layout.add_widget(text_input)
            
            field_card.add_widget(field_layout)
            content.add_widget(field_card)
        
        # Buttons
        btn_layout = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        
        def do_add(instance):
            matricule = fields['Matricule'].text.strip()
            nom = fields['Last Name'].text.strip()
            prenom = fields['First Name'].text.strip()
            section = fields['Section'].text.strip() or None
            groupe = fields['Group'].text.strip() or None
            
            if not matricule or not nom or not prenom:
                show_error("Please fill in all required fields")
                return
            
            success, message = self.db.add_student(matricule, nom, prenom, section, groupe)
            
            if success:
                popup.dismiss()
                show_success(message)
                self.refresh_groups()
                self.refresh_data()
            else:
                show_error(message)
        
        add_btn = ModernButton(
            text='Add',
            button_color=SUCCESS_COLOR
        )
        add_btn.bind(on_press=do_add)
        btn_layout.add_widget(add_btn)
        
        cancel_btn = ModernButton(
            text='Cancel',
            button_color=ERROR_COLOR
        )
        cancel_btn.bind(on_press=lambda x: popup.dismiss())
        btn_layout.add_widget(cancel_btn)
        
        content.add_widget(btn_layout)
        
        popup = Popup(
            title='Add New Student',
            content=content,
            size_hint=(0.8, 0.7)
        )
        popup.open()
    
    def show_edit_student_dialog(self, student):
        """Show dialog to edit student"""
        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(15))
        
        # Input fields
        fields = {}
        field_names = [
            ('Matricule', True, student[1]),
            ('Last Name', True, student[2]),
            ('First Name', True, student[3]),
            ('Section', False, student[4] or ''),
            ('Group', False, student[5] or '')
        ]
        
        for field_name, required, default_value in field_names:
            field_card = ModernCard(size_hint_y=None, height=dp(60))
            field_layout = BoxLayout(spacing=dp(10))
            
            label_text = f'{field_name}:' + (' *' if required else '')
            field_layout.add_widget(ModernLabel(
                text=label_text,
                size_hint_x=0.3
            ))
            
            text_input = TextInput(
                text=str(default_value),
                size_hint_x=0.7,
                multiline=False,
                background_color=(0.95, 0.95, 0.96, 1),
                foreground_color=TEXT_PRIMARY,
                font_size=sp(14),
                padding=[dp(10), dp(12)]
            )
            fields[field_name] = text_input
            field_layout.add_widget(text_input)
            
            field_card.add_widget(field_layout)
            content.add_widget(field_card)
        
        # Buttons
        btn_layout = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        
        def do_update(instance):
            matricule = fields['Matricule'].text.strip()
            nom = fields['Last Name'].text.strip()
            prenom = fields['First Name'].text.strip()
            section = fields['Section'].text.strip() or None
            groupe = fields['Group'].text.strip() or None
            
            if not matricule or not nom or not prenom:
                show_error("Please fill in all required fields")
                return
            
            success, message = self.db.update_student(
                student[0], matricule, nom, prenom, section, groupe
            )
            
            if success:
                popup.dismiss()
                show_success(message)
                self.refresh_groups()
                self.refresh_data()
            else:
                show_error(message)
        
        update_btn = ModernButton(
            text='Update',
            button_color=SUCCESS_COLOR
        )
        update_btn.bind(on_press=do_update)
        btn_layout.add_widget(update_btn)
        
        cancel_btn = ModernButton(
            text='Cancel',
            button_color=ERROR_COLOR
        )
        cancel_btn.bind(on_press=lambda x: popup.dismiss())
        btn_layout.add_widget(cancel_btn)
        
        content.add_widget(btn_layout)
        
        popup = Popup(
            title='Edit Student',
            content=content,
            size_hint=(0.8, 0.7)
        )
        popup.open()
    
    def view_student_details(self, student):
        """View detailed student information"""
        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(15))
        
        # Student info card
        info_card = ModernCard()
        info_layout = GridLayout(cols=2, spacing=dp(10))
        
        fields = [
            ('ID', student[0]),
            ('Matricule', student[1]),
            ('Last Name', student[2]),
            ('First Name', student[3]),
            ('Section', student[4] or 'N/A'),
            ('Group', student[5] or 'N/A'),
            ('Created', student[6]),
            ('Last Updated', student[7])
        ]
        
        for label, value in fields:
            info_layout.add_widget(ModernLabel(text=f'[b]{label}:[/b]', markup=True))
            info_layout.add_widget(ModernLabel(text=str(value)))
        
        info_card.add_widget(info_layout)
        content.add_widget(info_card)
        
        # Statistics card
        stats = self.db.get_student_statistics(student[0])
        
        if stats:
            stats_card = ModernCard()
            
            details_text = f"""[b]Statistics:[/b]

[b]Attendance:[/b]
• Total Classes: {stats['total_classes']}
• Present: {stats['present_count']}
• Absent: {stats['absent_count']}
• Justified Absences: {stats['justified_count']}
• Attendance Rate: {stats['attendance_rate']:.1f}%

[b]Academic Performance:[/b]
• Total Marks: {stats['total_marks']}
• Average Score: {stats['average_score']:.2f}/20
• Highest Score: {stats['highest_score']:.2f}/20
• Lowest Score: {stats['lowest_score']:.2f}/20
"""
            
            details_label = Label(
                text=details_text,
                markup=True,
                color=TEXT_PRIMARY,
                font_size=sp(13),
                halign='left',
                valign='top'
            )
            details_label.bind(size=details_label.setter('text_size'))
            stats_card.add_widget(details_label)
            
            content.add_widget(stats_card)
        
        # Close button
        close_btn = ModernButton(
            text='Close',
            size_hint=(None, None),
            size=(dp(150), dp(45)),
            pos_hint={'center_x': 0.5},
            button_color=PRIMARY_COLOR
        )
        
        popup = Popup(
            title='Student Details',
            content=content,
            size_hint=(0.8, 0.8)
        )
        
        close_btn.bind(on_press=popup.dismiss)
        content.add_widget(close_btn)
        
        popup.open()
    
    def confirm_delete_student(self, student):
        """Show confirmation before deleting student"""
        message = f"Are you sure you want to delete:\n\n{student[3]} {student[2]}\nMatricule: {student[1]}\n\nThis will delete all related records!"
        
        ConfirmationDialog(
            message=message,
            on_yes=lambda: self.delete_student(student[0])
        ).open()
    
    def delete_student(self, student_id):
        """Delete a student"""
        success, message = self.db.delete_student(student_id)
        
        if success:
            show_success(message)
            self.refresh_data()
        else:
            show_error(message)
    
    def show_import_dialog(self, instance):
        """Show dialog for Excel import with group name input"""
        content = BoxLayout(orientation='vertical', spacing=dp(15), padding=dp(20))
        
        # Instructions
        instructions = Label(
            text=(
                "[b]Import Students from Excel[/b]\n\n"
                "You can optionally specify a group name that will be "
                "assigned to all imported students."
            ),
            markup=True,
            halign='center',
            valign='middle',
            color=TEXT_PRIMARY,
            size_hint_y=None,
            height=dp(80)
        )
        instructions.bind(size=instructions.setter('text_size'))
        content.add_widget(instructions)
        
        # Group name input
        group_card = ModernCard(size_hint_y=None, height=dp(70))
        group_layout = BoxLayout(spacing=dp(10))
        group_layout.add_widget(ModernLabel(
            text='Group Name (optional):',
            size_hint_x=0.4
        ))
        group_input = TextInput(
            hint_text='e.g., Groupe10',
            size_hint_x=0.6,
            multiline=False,
            background_color=(0.95, 0.95, 0.96, 1),
            foreground_color=TEXT_PRIMARY,
            font_size=sp(14),
            padding=[dp(10), dp(12)]
        )
        group_layout.add_widget(group_input)
        group_card.add_widget(group_layout)
        content.add_widget(group_card)
        
        # Buttons
        btn_layout = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        
        def do_browse(instance):
            # Store group name
            self.pending_import_groupe = group_input.text.strip() or None
            popup.dismiss()
            
            # Open file picker
            if platform == 'android':
                # Use native Android file picker
                open_android_file_picker(self.handle_file_selection)
            else:
                # Use Kivy file chooser for desktop
                self.open_desktop_file_chooser()
        
        browse_btn = ModernButton(
            text='Browse Files',
            button_color=PRIMARY_COLOR
        )
        browse_btn.bind(on_press=do_browse)
        btn_layout.add_widget(browse_btn)
        
        cancel_btn = ModernButton(
            text='Cancel',
            button_color=ERROR_COLOR
        )
        cancel_btn.bind(on_press=lambda x: popup.dismiss())
        btn_layout.add_widget(cancel_btn)
        
        content.add_widget(btn_layout)
        
        popup = Popup(
            title='Import Excel File',
            content=content,
            size_hint=(0.8, 0.5)
        )
        popup.open()
    
    def handle_file_selection(self, file_path):
        """Handle file selected from Android file picker"""
        if file_path:
            logger.info(f"File selected: {file_path}")
            self.import_excel(file_path, self.pending_import_groupe)
        else:
            show_error("No file selected or file access failed")
    
    def open_desktop_file_chooser(self):
        """Open desktop file chooser (for non-Android platforms)"""
        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        
        # File chooser
        file_chooser = FileChooserListView(
            path=os.path.expanduser('~'),
            filters=['*.xlsx', '*.xls'],
            size_hint_y=0.85
        )
        content.add_widget(file_chooser)
        
        # Buttons
        btn_layout = BoxLayout(size_hint_y=None, height=dp(55), spacing=dp(10))
        
        def do_import(instance):
            if file_chooser.selection:
                file_path = file_chooser.selection[0]
                popup.dismiss()
                self.import_excel(file_path, self.pending_import_groupe)
            else:
                show_error("Please select a file")
        
        import_btn = ModernButton(
            text='Import',
            button_color=SUCCESS_COLOR
        )
        import_btn.bind(on_press=do_import)
        btn_layout.add_widget(import_btn)
        
        cancel_btn = ModernButton(
            text='Cancel',
            button_color=ERROR_COLOR
        )
        cancel_btn.bind(on_press=lambda x: popup.dismiss())
        btn_layout.add_widget(cancel_btn)
        
        content.add_widget(btn_layout)
        
        popup = Popup(
            title='Select Excel File',
            content=content,
            size_hint=(0.9, 0.9)
        )
        popup.open()
    
    def import_excel(self, file_path, groupe_name):
        """Import Excel file with progress indicator"""
        loading = LoadingPopup(title='Importing Students...')
        loading.open()
        
        def update_progress(value):
            loading.update_progress(value, f'Importing... {int(value * 100)}%')
        
        def do_import():
            success, message, count = self.db.import_from_excel(
                file_path,
                groupe_name,
                progress_callback=update_progress
            )
            
            Clock.schedule_once(lambda dt: self._import_complete(loading, success, message), 0)
        
        thread = threading.Thread(target=do_import)
        thread.start()
    
    def _import_complete(self, loading_popup, success, message):
        """Handle import completion"""
        loading_popup.dismiss()
        
        if success:
            show_success(message, 'Import Successful')
            self.refresh_groups()
            self.refresh_data()
        else:
            show_error(message, 'Import Failed')
    
    def export_data(self, instance):
        """Export current data to Excel"""
        if not self.selected_groupe:
            show_error("Please select a group first")
            return
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"export_{self.selected_groupe}_{timestamp}.xlsx"
        
        # Determine export path
        if platform == 'android':
            storage_path = get_external_storage_path()
            export_folder = os.path.join(storage_path, 'StudentTrackerPro', 'exports')
        else:
            export_folder = 'exports'
        
        os.makedirs(export_folder, exist_ok=True)
        output_path = os.path.join(export_folder, filename)
        
        success, message = self.db.export_to_excel(output_path, self.selected_groupe)
        
        if success:
            show_success(message, 'Export Successful')
        else:
            show_error(message, 'Export Failed')
    
    def backup_database(self, instance):
        """Create database backup"""
        success, result = self.db.backup_database()
        
        if success:
            show_success(f"Database backed up successfully!\n\n{result}", 'Backup Complete')
        else:
            show_error(result, 'Backup Failed')
    
    def refresh_data(self):
        """Refresh current view"""
        if self.search_mode:
            self.clear_search()
        elif self.selected_groupe or True:  # Always load
            self.load_students()

# ============================================
# APPLICATION
# ============================================
class StudentTrackerApp(App):
    """Enhanced Kivy application with modern GUI"""
    
    def build(self):
        self.title = f'{Config.APP_NAME} - by {Config.DEVELOPER}'
        
        # Request Android permissions FIRST - this is the CATALYST!
        if platform == 'android':
            Clock.schedule_once(lambda dt: request_android_permissions(), 0.5)
        
        self.db = StudentTrackerDB()
        
        # Set window properties
        Window.size = (1280, 820)
        Window.minimum_width = 1000
        Window.minimum_height = 700
        Window.clearcolor = BACKGROUND_COLOR
        
        # Create screen manager
        sm = ScreenManager()
        sm.add_widget(MainScreen(name='main', db=self.db))
        
        # Schedule auto-backup
        Clock.schedule_interval(
            lambda dt: self.auto_backup(),
            Config.AUTO_BACKUP_INTERVAL
        )
        
        logger.info(f"Application started successfully - {Config.APP_NAME} by {Config.DEVELOPER}")
        return sm
    
    def auto_backup(self):
        """Perform automatic database backup"""
        success, result = self.db.backup_database()
        if success:
            logger.info(f"Auto-backup completed: {result}")
        else:
            logger.error(f"Auto-backup failed: {result}")
    
    def on_stop(self):
        """Cleanup when app closes"""
        logger.info("Application closing")
        self.db.backup_database()

if __name__ == '__main__':
    StudentTrackerApp().run()
