from flask_wtf import FlaskForm
from wtforms import EmailField, StringField, TextAreaField, BooleanField, DateField, SelectField, SubmitField, FileField
from wtforms.fields import PasswordField
from wtforms.validators import DataRequired, Email, EqualTo, Optional, Length
from flask_ckeditor import CKEditorField

from app.models import Department


class TaskForm(FlaskForm):
    date = DateField('Date', validators=[DataRequired()])
    category_id = SelectField('Category', coerce=int, validators=[DataRequired()])
    description = TextAreaField('Description', validators=[DataRequired()])
    remarks = TextAreaField('Remarks', validators=[Optional()])
    is_complete = BooleanField('Completed')
    submit = SubmitField('Save')

    def __init__(self, *args, **kwargs):
        super(TaskForm, self).__init__(*args, **kwargs)
        from .models import Category
        self.category_id.choices = [
            (c.id, c.name) for c in Category.query.order_by(Category.name).all()
        ]


class EmailForm(FlaskForm):
    email = EmailField('Email', validators=[DataRequired()])
    submit = SubmitField('Send Email')


class CodeForm(FlaskForm):
    code = StringField('Code', validators=[DataRequired(), Length(min=6, max=6)])
    submit = SubmitField('Submit Code')


class LoginForm(FlaskForm):
    email = EmailField("Email", validators=[DataRequired()])
    password = PasswordField("Password", validators=[DataRequired()])
    submit = SubmitField('Login')


class RegisterForm(FlaskForm):
    name = StringField('Full name', validators=[DataRequired()])
    email = EmailField("Email", validators=[DataRequired(), Email()])
    phone_no = StringField('Phone Number', validators=[DataRequired()])
    department_id = SelectField('Department', coerce=int, validators=[DataRequired()])
    password = PasswordField("Password", validators=[DataRequired()])
    password_2 = PasswordField(
        "Confirm Password",
        validators=[DataRequired(), EqualTo('password')]
    )
    submit = SubmitField('Register')

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        from .models import Department
        self.department_id.choices = [
            (d.id, d.name) for d in Department.query.order_by(Department.name).all()
        ]


class PasswordForm(FlaskForm):
    password = PasswordField("Password", validators=[DataRequired()])
    password_2 = PasswordField(
        "Confirm Password",
        validators=[DataRequired(), EqualTo('password')]
    )
    submit = SubmitField('Submit')


class SetPasswordForm(FlaskForm):
    email = SelectField(
        'Email',
        choices=[],
        coerce=int,
        validators=[DataRequired()]
    )
    password = PasswordField("Password", validators=[DataRequired()])
    password_2 = PasswordField(
        "Confirm Password",
        validators=[DataRequired(), EqualTo('password')]
    )
    submit = SubmitField('Set Password')

    def __init__(self, *args, **kwargs):
        super(SetPasswordForm, self).__init__(*args, **kwargs)
        from .models import User
        self.email.choices = [(usr.id, usr.email) for usr in User.query.all()]


class ResetPasswordForm(FlaskForm):
    password = PasswordField("Password", validators=[DataRequired()])
    password_2 = PasswordField(
        "Confirm Password",
        validators=[DataRequired(), EqualTo('password')]
    )
    submit = SubmitField('Submit')




class WeeklyReportForm(FlaskForm):
    start_date = DateField('Start date', validators=[DataRequired()])
    end_date = DateField('End date', validators=[DataRequired()])
    report_type = SelectField('Report type', choices=[], validators=[DataRequired()])
    submit = SubmitField('Generate Report')

    def __init__(self, *args, **kwargs):
        super(WeeklyReportForm, self).__init__(*args, **kwargs)
        from .models import ReportType
        # Populate report type choices when form is instantiated (within app context)
        self.report_type.choices = [(rt.id, rt.type.title()) for rt in ReportType.query.all()]


class QuarterlyReportForm(FlaskForm):
    start_date = DateField('Start date', validators=[DataRequired()])
    end_date = DateField('End date', validators=[DataRequired()])
    department = SelectField('Department', choices=[], coerce=int, validators=[DataRequired()])
    submit = SubmitField('Generate Report')

    def __init__(self, *args, **kwargs):
        super(QuarterlyReportForm, self).__init__(*args, **kwargs)
        from .models import Department
        # Populate department choices when form is instantiated (within app context)
        self.department.choices = [
            (d.id, d.name) for d in Department.query.order_by(Department.name).all()
        ]

class UserForm(FlaskForm):
    phone = StringField('Phone', validators=[DataRequired()])
    department = SelectField('Department', choices=[], validators=[DataRequired()])
    #profile_pic = FileField('Profile picture (optional)', validators=[FileAllowed(['jpg', 'png', 'jpeg'], 'Images only!')])
    submit = SubmitField('Create Profile')

    def __init__(self, *args, **kwargs):
        super(UserForm, self).__init__(*args, **kwargs)
        from .models import Department
        # Populate department choices when form is instantiated (within app context)
        self.department.choices = [(d.id, d.name) for d in Department.query.all()]


class DepartmentForm(FlaskForm):
    name = StringField('Name', validators=[DataRequired()], render_kw={"placeholder": "Enter College/Department/Unit name"})
    code = StringField('Department Code', render_kw={"placeholder": "e.g. CoHSS"})
    description = TextAreaField('Description', validators=[Optional()])
    submit = SubmitField('Create Department')


class CategoryForm(FlaskForm):
    name = StringField('Category Name', validators=[DataRequired(), Length(max=100)], render_kw={"placeholder": "e.g. Printers and Peripherals"})
    submit = SubmitField('Create Category')



class EditReportForm(FlaskForm):
    report = CKEditorField('Report Content')
    submit = SubmitField('Submit')


class StatsForm(FlaskForm):
    start_date = DateField('Start date', validators=[DataRequired()])
    end_date = DateField('End date', validators=[DataRequired()])
    department = SelectField('Department', choices=[], coerce=int, validators=[DataRequired()])
    submit = SubmitField('Generate Stats')

    def __init__(self, *args, **kwargs):
        super(StatsForm, self).__init__(*args, **kwargs)
        from .models import Department
        # Populate department choices when form is instantiated (within app context)
        self.department.choices = [(d.id, d.name) for d in Department.query.order_by(Department.name).all()]


