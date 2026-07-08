from datetime import datetime, timezone
from enum import unique
from flask import current_app

from sqlalchemy.engine import default

from . import db, login_manager
from flask_login import UserMixin
from itsdangerous import URLSafeTimedSerializer


class User(UserMixin, db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255))
    email = db.Column(db.String(255), unique=True, nullable=False)
    phone = db.Column(db.String(20))
    profile_pic = db.Column(db.String(512))
    password = db.Column(db.Text, nullable=True)
    
    is_admin = db.Column(db.Boolean, default=False)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)

    # Foreign key to Department
    department_id = db.Column(db.Integer, db.ForeignKey('departments.id'))

    tasks = db.relationship('Task', back_populates='owner')
    reports = db.relationship('Report', back_populates='owner')
    department = db.relationship('Department', back_populates='users')

    def __repr__(self):
        return f'<User {self.email}>'

    def get_reset_token(self):
        serializer = URLSafeTimedSerializer(current_app.config['SECRET_KEY'])
        return serializer.dumps(self.email, salt='password-reset-salt')

    @staticmethod
    def verify_reset_token(token, expiration=1800):  # 1800 seconds = 30 minutes
        serializer = URLSafeTimedSerializer(current_app.config['SECRET_KEY'])
        try:
            email = serializer.loads(
                token,
                salt='password-reset-salt',
                max_age=expiration
            )
        except:
            return None

        return User.query.filter_by(email=email).first()


class PasswordResetCode(db.Model):
    """Stores a single active reset code per user email."""
    __tablename__ = "password_reset_codes"
 
    id         = db.Column(db.Integer, primary_key=True)
    email      = db.Column(db.String(255), nullable=False, index=True)
    code       = db.Column(db.String(6), nullable=False)
    expires_at = db.Column(db.DateTime(timezone=True), nullable=False)
    used       = db.Column(db.Boolean, default=False, nullable=False)
 
    def is_valid(self) -> bool:
        return (
            not self.used
            and datetime.now(timezone.utc) < self.expires_at
        )


class Department(db.Model):
    __tablename__ = 'departments'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255), nullable=False)
    code = db.Column(db.String(20), nullable=False)
    description = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)

    # Relationship to Users
    users = db.relationship('User', back_populates='department')
    def __repr__(self):
        return f'<Department {self.name}>'


class Category(db.Model):
    __tablename__ = 'categories'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False)

    tasks = db.relationship('Task', back_populates='category')

    def __repr__(self):
        return f'<Category {self.name}>'


class Task(db.Model):
    __tablename__ = 'tasks'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    date = db.Column(db.Date, nullable=False)
    description = db.Column(db.Text)
    remarks = db.Column(db.Text, nullable=True)
    is_complete = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)

    category_id = db.Column(db.Integer, db.ForeignKey('categories.id'), nullable=False)
    category = db.relationship('Category', back_populates='tasks')

    owner = db.relationship('User', back_populates='tasks')

    @property
    def effective_category_name(self):
        return self.category.name if self.category else "Others"

    def __repr__(self):
        return f'<Task {self.id} {self.effective_category_name} {self.date}>'



class ReportType(db.Model):
    __tablename__ = 'reporttype'
    id = db.Column(db.Integer, primary_key=True)
    type = db.Column(db.String(15))
    
    reports = db.relationship('Report', back_populates='report_type')



class Report(db.Model):
    __tablename__ = 'reports'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)

    report_type_id = db.Column(db.Integer, db.ForeignKey('reporttype.id'), nullable=False)
    report_type = db.relationship('ReportType', back_populates='reports')

    # Filename of the saved report file
    report_file = db.Column(db.String(255), unique=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)

    owner = db.relationship('User', back_populates='reports')

    @property
    def title(self):
        type_label = self.report_type.type.title() if self.report_type else 'Report'
        return f"{type_label} Report: {self.start_date.strftime('%b %d, %Y')} – {self.end_date.strftime('%b %d, %Y')}"

    def __repr__(self):
        report_type_label = self.report_type.type if self.report_type else 'unknown'
        return f'<Report {self.id} {report_type_label} {self.start_date} {self.end_date}>'



@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))




