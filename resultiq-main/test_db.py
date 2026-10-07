from app import create_app
from models import db, User, UploadedFile, Report, ActivityLog

app = create_app('testing')

with app.app_context():
    print("1. Creating tables (if they don't exist)...")
    db.create_all()
    
    try:
        print("2. Checking for existing test user...")
        test_user = User.query.filter_by(email="faculty@resultiq.com").first()
        
        if test_user:
            print(f"[INFO] User already exists (ID: {test_user.id}). Reusing.")
        else:
            print("[INFO] User does not exist. Creating new test user...")
            test_user = User(
                name="Test Faculty",
                email="faculty@resultiq.com",
                username="faculty",
                department="Computer Science",
                designation="Professor"
            )
            test_user.set_password("securepassword")
            db.session.add(test_user)
            db.session.flush() # flush to get the user ID for related records
            
            print(f"[INFO] Created User: {test_user.id}")
            
            print("Creating test file...")
            test_file = UploadedFile(
                user_id=test_user.id,
                original_filename="grades.xlsx",
                # Generate unique stored filename to avoid clashes
                stored_filename=f"grades_{test_user.id}.xlsx", 
                file_size=1024,
                file_path=f"/uploads/grades_{test_user.id}.xlsx",
                is_valid=True,
                worksheets={"Sheet1": ["ID", "Name", "Score"]}
            )
            db.session.add(test_file)
            db.session.flush()
            print(f"[INFO] Created UploadedFile: {test_file.id}")
            
            print("Creating test report...")
            test_report = Report(
                user_id=test_user.id,
                uploaded_file_id=test_file.id,
                original_filename="grades.xlsx",
                report_name="Final Exam Report",
                sheet_name="Sheet1",
                requested_columns=["ID", "Score"],
                report_path=f"/reports/report_{test_file.id}.xlsx",
                status="completed"
            )
            db.session.add(test_report)
            
            print("Creating activity log...")
            log = ActivityLog(
                user_id=test_user.id,
                activity_type="REPORT_GENERATED",
                description="Generated Final Exam Report"
            )
            db.session.add(log)
            
            # Commit the transaction safely
            db.session.commit()
            print("[INFO] Database transaction committed successfully.")
            
        print("\n--- Verifying Relationships ---")
        user = User.query.filter_by(email="faculty@resultiq.com").first()
        print(f"User retrieved: {user.name}")
        print(f"User's files count: {len(user.uploaded_files)}")
        print(f"User's reports count: {len(user.reports)}")
        print(f"User's activity logs count: {len(user.activity_logs)}")
        
        if user.reports:
            report = user.reports[0]
            print(f"Report's owner: {report.user.name}")
            print(f"Report's source file: {report.uploaded_file.original_filename}")
        
        print("\n[SUCCESS] All validations passed.")
        
    except Exception as e:
        db.session.rollback()
        print(f"\n[FAILURE] An error occurred: {e}")
        print("[INFO] Transaction rolled back.")
