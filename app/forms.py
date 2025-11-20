from flask_wtf import FlaskForm
from wtforms import StringField, TextAreaField, BooleanField, DateField, SelectField, SubmitField, FileField
from wtforms.validators import DataRequired, Email, Optional
from flask_ckeditor import CKEditorField



class TaskForm(FlaskForm):
    date = DateField('Date', validators=[DataRequired()])
    task_type = StringField('Task type', validators=[DataRequired()])
    description = TextAreaField('Description', validators=[DataRequired()])
    remarks = TextAreaField('Remarks', validators=[Optional()])
    is_complete = BooleanField('Completed')
    submit = SubmitField('Save')


class ReportForm(FlaskForm):
    start_date = DateField('Start date', validators=[DataRequired()])
    end_date = DateField('End date', validators=[DataRequired()])
    report_type = SelectField('Report type', choices=[], validators=[DataRequired()])
    submit = SubmitField('Generate Report')

    def __init__(self, *args, **kwargs):
        super(ReportForm, self).__init__(*args, **kwargs)
        from .models import ReportType
        # Populate report type choices when form is instantiated (within app context)
        self.report_type.choices = [(rt.id, rt.type.title()) for rt in ReportType.query.all()]


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



class EditReportForm(FlaskForm):
    report = CKEditorField('Report Content')
    submit = SubmitField('Submit')


