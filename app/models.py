from datetime import datetime
from enum import unique

from sqlalchemy.engine import default

from . import db, login_manager
from flask_login import UserMixin



class User(UserMixin, db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255))
    email = db.Column(db.String(255), unique=True, nullable=False)
    phone = db.Column(db.String(20))
    profile_pic = db.Column(db.String(512))
    
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


class Task(db.Model):
    __tablename__ = 'tasks'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    date = db.Column(db.Date, nullable=False)
    task_type = db.Column(db.String(255), nullable=False)
    description = db.Column(db.Text)
    remarks = db.Column(db.Text, nullable=True)
    is_complete = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)

    owner = db.relationship('User', back_populates='tasks')

    def __repr__(self):
        return f'<Task {self.id} {self.task_type} {self.date}>'



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

    def __repr__(self):
        report_type_label = self.report_type.type if self.report_type else 'unknown'
        return f'<Report {self.id} {report_type_label} {self.start_date} {self.end_date}>'





"""
class WeeklyReport(db.Model):
    __tablename__ = 'weeklyreports'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)
    report_type = db.Column(db.String(255))
    report_content = db.Column(db.Text)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)

    owner = db.relationship('User', back_populates='reports')

    def __repr__(self):
        return f'<Report {self.id} {self.report_type} {self.start_date} {self.end_date}>'


class QuarterlyReport(db.Model):
    __tablename__ = 'quarterlyreports'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)
    report_type = db.Column(db.String(255))
    report_content = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)

    report_file = db.Column(db.String(255))

    owner = db.relationship('User', back_populates='reports')


    def __repr__(self):
        return f'<Report {self.id} {self.report_type} {self.start_date} {self.end_date}>'

"""



@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))




