import os
import re
import html
from collections import Counter
from types import SimpleNamespace
from datetime import datetime, timedelta
from flask import Blueprint, render_template, redirect, url_for, flash, request, current_app, send_file, session, jsonify
from flask_login import login_user, logout_user, login_required, current_user
from io import BytesIO
from authlib.common.security import generate_token
from . import db, oauth
from .models import *
from .forms import *
from .ai_integration import generate_weekly_report_docs, generate_quarterly_report_docs
import secrets
from flask_wtf.csrf import CSRF
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from bs4 import BeautifulSoup
import os
import re
from werkzeug.security import generate_password_hash, check_password_hash
from itsdangerous import URLSafeTimedSerializer
#from flask_mail import Message
from app import mail

from flask_mail import Message
import string
import random
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import base64
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from xhtml2pdf import pisa



main_bp = Blueprint('main', __name__)
auth_bp = Blueprint('auth', __name__)


csrf = CSRF()


def _is_privileged_admin(user):
    """Single source of truth for admin-level access.

    Mirrors the bootstrap-admin fallback already relied on by base.html's
    nav (owner email / user id 1) so server-side checks can't diverge from
    what the UI implies is admin-only.
    """
    return bool(
        user.is_authenticated
        and (user.is_admin or user.email == "niiakoadjei@gmail.com" or user.id == 1)
    )


def admin_required(view_func):
    from functools import wraps

    @wraps(view_func)
    def wrapped(*args, **kwargs):
        if not _is_privileged_admin(current_user):
            flash('Permission Denied', 'danger')
            return redirect(url_for('main.dashboard'))
        return view_func(*args, **kwargs)
    return wrapped


def _iter_document_blocks(document):
    body = document.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document)


def _format_run_html(run):
    text = run.text
    if not text:
        return ""

    chunk = html.escape(text)

    if run.bold:
        chunk = f"<strong>{chunk}</strong>"
    if run.italic:
        chunk = f"<em>{chunk}</em>"
    if run.underline:
        chunk = f"<u>{chunk}</u>"

    style_bits = []
    if run.font.color and run.font.color.rgb:
        style_bits.append(f"color: #{run.font.color.rgb};")
    if run.font.size:
        try:
            size_pt = run.font.size.pt
            if size_pt:
                style_bits.append(f"font-size: {size_pt:.1f}pt;")
        except AttributeError:
            pass

    if style_bits:
        chunk = f"<span style=\"{''.join(style_bits)}\">{chunk}</span>"

    return chunk


def _paragraph_to_html(paragraph):
    runs_html = [_format_run_html(run) for run in paragraph.runs]
    runs_html = [chunk for chunk in runs_html if chunk]
    if not runs_html:
        return ""

    style_name = (paragraph.style.name or "").lower()
    heading_match = re.match(r'heading\s+([1-6])', style_name)
    tag = f"h{heading_match.group(1)}" if heading_match else "p"

    align_map = {
        WD_ALIGN_PARAGRAPH.LEFT: "left",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.CENTER: "center",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    }
    styles = []
    if paragraph.alignment in align_map:
        styles.append(f"text-align: {align_map[paragraph.alignment]};")

    style_attr = f" style=\"{''.join(styles)}\"" if styles else ""
    return f"<{tag}{style_attr}>{''.join(runs_html)}</{tag}>"


def _table_to_html(table):
    row_chunks = []
    for row in table.rows:
        cell_chunks = []
        for cell in row.cells:
            cell_content = []
            for paragraph in cell.paragraphs:
                para_html = _paragraph_to_html(paragraph)
                if para_html:
                    cell_content.append(para_html)
            cell_chunks.append(f"<td>{''.join(cell_content) or '&nbsp;'}</td>")
        row_chunks.append(f"<tr>{''.join(cell_chunks)}</tr>")
    if not row_chunks:
        return ""

    table_html = "".join(row_chunks)
    return f"<table class=\"table table-bordered table-sm\"><tbody>{table_html}</tbody></table>"


def _document_to_html(document):
    html_parts = []
    for block in _iter_document_blocks(document):
        if isinstance(block, Paragraph):
            para_html = _paragraph_to_html(block)
            if para_html:
                html_parts.append(para_html)
        elif isinstance(block, Table):
            table_html = _table_to_html(block)
            if table_html:
                html_parts.append(table_html)
    return "\n".join(html_parts)


def _html_to_paragraphs(html_content):
    if not html_content:
        return []

    working = re.sub(r'</(p|div|h[1-6])\s*>', '\n\n', html_content, flags=re.IGNORECASE)
    working = re.sub(r'<br\s*/?>', '\n', working, flags=re.IGNORECASE)
    working = re.sub(r'</tr\s*>', '\n', working, flags=re.IGNORECASE)
    working = re.sub(r'</td\s*>', ' ', working, flags=re.IGNORECASE)
    working = re.sub(r'<[^>]+>', '', working)
    working = html.unescape(working)

    lines = [line.strip() for line in working.splitlines()]
    paragraphs = []
    buffer = []

    for line in lines:
        if line:
            buffer.append(line)
        elif buffer:
            paragraphs.append(' '.join(buffer))
            buffer = []

    if buffer:
        paragraphs.append(' '.join(buffer))

    return paragraphs


# Session timeout route...
SESSION_TIMEOUT_MINUTES = 15

@main_bp.before_app_request
def enforce_session_timeout():
    """Automatically log users out after 15 minutes of inactivity."""
    if not current_user.is_authenticated:
        session.pop('last_activity', None)
        return

    now = datetime.utcnow()
    last_activity = session.get('last_activity')

    if last_activity is not None:
        last_activity_dt = datetime.fromtimestamp(last_activity)
        if now - last_activity_dt > timedelta(minutes=SESSION_TIMEOUT_MINUTES):
            session.pop('last_activity', None)
            logout_user()
            flash('Session timed out after 15 minutes of inactivity.', 'warning')
            return redirect(url_for('auth.login'))

    session['last_activity'] = now.timestamp()


@main_bp.context_processor
def inject_current_year():
    return dict(current_year=datetime.now().year)

@auth_bp.context_processor
def inject_current_year():
    from datetime import datetime
    return dict(current_year=datetime.now().year)


@main_bp.context_processor
def inject_admin_flag():
    return dict(is_privileged_admin=_is_privileged_admin(current_user))

@auth_bp.context_processor
def inject_admin_flag():
    return dict(is_privileged_admin=_is_privileged_admin(current_user))


@main_bp.route('/')
def index():
    return render_template('index.html')


@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    form = LoginForm()
    if form.validate_on_submit():
        # Find user by email
        user = User.query.filter_by(email=form.email.data.strip()).first()
        # Check if user exists and password matches
        if user and check_password_hash(user.password, form.password.data):
            login_user(user)  # Log the user in
            flash('Login successful!', 'success')
            return redirect(url_for('main.dashboard'))  # Change to your dashboard
        else:
            flash('Invalid email or password.', 'danger')
            return redirect(url_for('main.index'))
    return render_template('signin.html', form=form, title='Welcome Back')



def check_phone_number(phone_number):
    # Must not be empty
    if not phone_number:
        return False
    # Allow leading +
    if phone_number[0] == '+':
        phone_number = phone_number[1:]
    # After removing +, all characters must be digits
    return phone_number.isdigit()


@auth_bp.route('/register', methods=['GET', 'POST'])
def register():
    form = RegisterForm()
    # Populate department choices correctly
    form.department_id.choices = [
        (dept.id, dept.name) for dept in Department.query.order_by(Department.name).all()
    ]
    if form.validate_on_submit():
        phone_num = form.phone_no.data
        if check_phone_number(phone_num) == False:
            flash('Invalid phone number. Enter a valid number', 'danger')
            return redirect(request.referrer)
        # Check if email already exists
        existing_user = User.query.filter_by(email=form.email.data).first()
        if existing_user:
            flash('Email already registered.', 'danger')
            return render_template('register.html', form=form, title='Create Account')
        # Hash password using bcrypt-compatible method
        hashed_password = generate_password_hash(form.password.data, method='scrypt', salt_length=8)
        new_user = User(
            name=form.name.data,
            email=form.email.data,
            phone=phone_num,
            department_id=form.department_id.data,
            password=hashed_password
        )
        db.session.add(new_user)
        db.session.commit()
        flash('Account created successfully!', 'success')
        return redirect(url_for('auth.login'))

    return render_template('signup.html', form=form, title='Create Account')


#############################################
############# ── Helpers ────################

def generate_code():
    code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
    #print(code)
    return code
 
 
def send_reset_email(recipient: str, code: str) -> None:
    server = smtplib.SMTP(os.getenv('SMTP_SERVER'), os.getenv('SMTP_PORT'))
    sender_email = os.getenv('MAIL_USERNAME')
    sender_password = os.getenv('MAIL_PASSWORD')
    # Create the email
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = recipient
    message["Subject"] = "Test Email from Python"

    html = f"""
    <html>
    <body>
        <p>Hello,</p>
        <p>Your password reset code is:</p>
        <h2>{code}</h2>
        <br>
        <p>If you did not request this, please ignore this email.</p>
    </body>
    </html>
    """

    message.attach(MIMEText(html, "html"))

    try:
        server.starttls()  # Secure the connection
        server.login(sender_email, sender_password)

        # Send email
        server.send_message(message)
        print("Email sent successfully!")

    except Exception as e:
        print(f"Error: {e}")

    finally:
        server.quit()

 
"""
@main_bp.route("/password-reset/request", methods=["GET","POST"])
def request_password_reset():
    form = EmailForm()

    # data  = request.get_json(silent=True) or {}
    # email = (data.get("email") or "").strip().lower()

    if form.validate_on_submit():
        email = form.email.data.strip()
 
        if not email:
            flash('Email does not exist', 'danger')
            return redirect(request.referrer)
    
        # Invalidate any previous unused code for this email
        PasswordResetCode.query.filter_by(email=email, used=False).update({"used": True})
        db.session.commit()
    
        # Create new code
        code = generate_code()
        expires_at = datetime.now(timezone.utc) + timedelta(minutes=15)
    
        reset_entry = PasswordResetCode(email=email, code=code, expires_at=expires_at)
        db.session.add(reset_entry)
        db.session.commit()
    
        # Send email (fire-and-forget; wrap in try/except so DB record isn't lost)
        try:
            send_reset_email(email, code)
            return redirect(url_for('main.verify_reset_code', email=email))
        except Exception as exc:
            flash(f'An error occurred: {exc}')
            app.logger.error("Failed to send reset email to %s: %s", email, exc)
            # Still return a generic success so attackers can't detect failures
            # Log/alert internally instead
            return redirect(request.referrer)
    
    # Generic response — never confirm whether the email exists
    return render_template('reset_link.html', title='Reset link')
"""

@auth_bp.route('/forgot-password', methods=['GET', 'POST'])
def forgot_password():
    form = EmailForm()

    if form.validate_on_submit():
        email_input = (form.email.data or "").strip().lower()

        user = User.query.filter_by(email=email_input).first()

        # always behave the same (security best practice)
        if not user:
            flash('If the email exists, a reset code has been sent.', 'info')
            return redirect(url_for('auth.forgot_password'))

        email = user.email

        # invalidate old codes
        PasswordResetCode.query.filter_by(email=email, used=False).update({"used": True})
        db.session.commit()

        # create new code
        code = generate_code()
        #expires_at = datetime.now() + timedelta(minutes=15)
        # Timezone-aware datetime
        expires_at = datetime.now(timezone.utc) + timedelta(minutes=15)

        reset_entry = PasswordResetCode(email=email, code=code, expires_at=expires_at)
        db.session.add(reset_entry)
        db.session.commit()

        try:
            send_reset_email(email, code)

            # STORE EMAIL IN SESSION HERE
            session["reset_email"] = email
            return redirect(url_for('auth.verify_reset_code'))

        except Exception as exc:
            current_app.logger.exception("Failed to send reset email for %s", email)
            flash("Something went wrong. Try again.", "danger")
            return redirect(url_for('auth.forgot_password'))

    return render_template('forgot_password.html', form=form)


 
@auth_bp.route("/reset-password/verify", methods=["GET", "POST"])
def verify_reset_code():

    form = CodeForm()
    email = session.get("reset_email")

    if not email:
        flash("Session expired. Restart password reset.", "danger")
        return redirect(url_for("auth.forgot_password"))

    if form.validate_on_submit():
        code = form.code.data

        if not code:
            flash("Code missing", "danger")
            return redirect(url_for("auth.verify_reset_code"))

        entry = (
            PasswordResetCode.query
            .filter_by(email=email, code=code, used=False)
            .order_by(PasswordResetCode.expires_at.desc())
            .first()
        )

        if not entry:
            flash("Invalid code!", "danger")
            return redirect(url_for("auth.verify_reset_code"))

        if not entry.is_valid():
            flash("Code expired. Request a new one.", "danger")
            return redirect(url_for("auth.forgot_password"))

        entry.used = True
        db.session.commit()

        # mark verified
        session["password_reset_verified"] = True

        return redirect(url_for("auth.set_new_password"))

    return render_template("verify_code.html", form=form)


@auth_bp.route("/password/set-password", methods=["GET", "POST"])
def set_new_password():

    # 🔐 Guard: must come from verified step
    if not session.get("password_reset_verified"):
        flash("Unauthorized access. Please verify your code first.", "danger")
        return redirect(url_for("auth.forgot_password"))

    email = session.get("reset_email")

    if not email:
        flash("Session expired. Please restart password reset.", "danger")
        return redirect(url_for("auth.forgot_password"))

    form = PasswordForm()

    if form.validate_on_submit():
        user = User.query.filter_by(email=email).first()

        if not user:
            flash("User does not exist", "danger")
            return redirect(url_for("main.request_password_reset"))

        password = form.password.data.strip()
        user.password = generate_password_hash(password, method="scrypt", salt_length=8)

        db.session.commit()

        # Clear session after success (VERY IMPORTANT)
        session.pop("reset_email", None)
        session.pop("password_reset_verified", None)

        flash("Password reset successful", "success")
        return redirect(url_for("auth.login"))

    return render_template("set_password.html", form=form)


@auth_bp.route('/set_password', methods=['GET', 'POST'])
@login_required
@admin_required
def set_password():
    form = SetPasswordForm()

    if form.validate_on_submit():
        usr_id = form.email.data
        user = User.query.get_or_404(usr_id)
        password = form.password.data
        password_2 = form.password_2.data
        if password != password_2:
            flash('Passwords do not match', 'danger')
            return render_template('set_password.html', form=form, title='Create Password')
        pwd_hash = generate_password_hash(
            password,
            method='scrypt',
            salt_length=8
        )
        user.password = pwd_hash
        db.session.commit()

        flash('Password successfully created', 'success')
        return redirect(url_for('auth.login'))  # recommended
    return render_template('set_password.html', form=form, title='Create Password')


@auth_bp.route('/reset-password/<token>', methods=['GET', 'POST'])
def reset_password(token):

    user = User.verify_reset_token(token)

    if not user:
        flash('The reset link is invalid or has expired', 'danger')
        return redirect(url_for('auth.forgot_password'))

    form = ResetPasswordForm()

    if form.validate_on_submit():
        user.set_password(form.password.data)
        db.session.commit()

        flash('Your password has been updated', 'success')
        return redirect(url_for('main.index'))

    return render_template(
        'reset_password.html',
        form=form,
        title='Reset Password'
    )


@auth_bp.route('/login/ngrok')
def ngrok_login():
    # Generate a unique nonce for CSRF protection during the OAuth flow
    session['nonce'] = generate_token()
    redirect_uri = "https://madyson-predorsal-georgeann.ngrok-free.dev/auth/callback"
    #redirect_uri = os.getenv('REDIRECT_URI', _external=True)
    #redirect_uri = url_for('auth.auth_callback', _external=True)
    return oauth.google.authorize_redirect(redirect_uri, nonce=session['nonce'])


@auth_bp.route('/auth/callback')
def auth_callback():
    # Check if this is an OAuth callback with the required parameters
    if 'code' not in request.args and 'error' not in request.args:
        # This might be a validation request or direct access - return 200 OK
        return 'OK', 200
    
    # Handle OAuth errors
    if 'error' in request.args:
        error = request.args.get('error')
        flash(f'Google OAuth error: {error}', 'danger')
        return redirect(url_for('main.index'))
    
    try:
        token = oauth.google.authorize_access_token()
        if not token:
            flash('Failed to obtain access token from Google.', 'danger')
            return redirect(url_for('main.index'))
        
        # Get the nonce from session and remove it (one-time use)
        nonce = session.pop('nonce', None)
        if not nonce:
            flash('Session expired. Please try logging in again.', 'danger')
            return redirect(url_for('main.index'))
        
        userinfo = oauth.google.parse_id_token(token, nonce=nonce)
        if not userinfo:
            flash('Failed to fetch user info from Google.', 'danger')
            return redirect(url_for('main.index'))

        email = userinfo.get('email')
        name = userinfo.get('name')
        picture = userinfo.get('picture')

        if not email:
            flash('Email not provided by Google.', 'danger')
            return redirect(url_for('main.index'))

        user = User.query.filter_by(email=email).first()
        if not user:
            # Store user info in session for profile creation
            session['pending_user'] = {
                'email': email,
                'name': name,
                'profile_pic': picture
            }
            flash('Please complete your profile to continue.', 'info')
            return redirect(url_for('main.create_profile'))

        login_user(user)
        flash('Logged in successfully.', 'success')
        return redirect(url_for('main.dashboard'))
    except Exception as e:
        flash(f'Authentication error: {str(e)}', 'danger')
        return redirect(url_for('main.index'))



@main_bp.route('/create-profile', methods=['GET', 'POST'])
def create_profile():
    # Check if there's pending user info in session
    pending_user = session.get('pending_user')
    if not pending_user:
        flash('No pending registration found. Please log in again.', 'warning')
        return redirect(url_for('main.index'))
    
    form = UserForm()
    
    if form.validate_on_submit():
        # Check if user already exists (race condition check)
        existing_user = User.query.filter_by(email=pending_user['email']).first()
        if existing_user:
            session.pop('pending_user', None)
            flash('User already exists. Please log in.', 'info')
            return redirect(url_for('main.index'))
        
        # Create new user
        new_user = User(
            name=pending_user['name'],
            email=pending_user['email'],
            profile_pic=pending_user.get('profile_pic'),
            phone=form.phone.data,
            department_id=form.department.data,
        )
        try:
            db.session.add(new_user)
            db.session.commit()
            session.pop('pending_user', None)  # Clear session
            flash('User profile created successfully', 'success')
            login_user(new_user)
            return redirect(url_for('main.dashboard'))
        except Exception as e:
            db.session.rollback()
            flash(f'An error occurred: {e}', 'danger')
            return redirect(url_for('main.index'))
    
    # Pre-fill name and email in the form if available
    return render_template('userform.html', form=form, name=pending_user.get('name'), email=pending_user.get('email'))



@auth_bp.route('/logout')
@login_required
def logout():
    logout_user()
    flash('You have been logged out.', 'info')
    return redirect(url_for('main.index'))


@main_bp.route('/create_department', methods=['GET', 'POST'])
@login_required
def create_department():
    if not _is_privileged_admin(current_user):
        if request.method == 'POST' and request.is_json:
            return jsonify({
                'success': False,
                'message': 'Only admins can create departments'
            }), 403
        flash('Permission Denied: only admins can create departments.', 'danger')
        return redirect(url_for('main.dashboard'))

    # Handle AJAX POST request (from JavaScript)
    if request.method == 'POST' and request.is_json:
        try:
            data = request.get_json()
            name = data.get('name')
            code = data.get('code')
            description = data.get('description', '')

            # Validate required fields
            if not name or not code:
                return jsonify({
                    'success': False,
                    'message': 'Department name and code are required'
                }), 400

            # Check if department already exists
            existing_dept = Department.query.filter_by(name=name).first()
            if existing_dept:
                return jsonify({
                    'success': False,
                    'message': 'Department already exists'
                }), 400

            # Check if code already exists
            existing_code = Department.query.filter_by(code=code).first()
            if existing_code:
                return jsonify({
                    'success': False,
                    'message': 'Department code already exists'
                }), 400

            # Create new department
            new_department = Department(
                name=name,
                code=code,
                description=description
            )
            
            db.session.add(new_department)
            db.session.commit()

            return jsonify({
                'success': True,
                'department_id': new_department.id,
                'message': 'Department created successfully'
            }), 201

        except Exception as e:
            db.session.rollback()
            return jsonify({
                'success': False,
                'message': f'An error occurred: {str(e)}'
            }), 500

    # Handle regular GET request (if someone navigates to the URL directly)
    form = DepartmentForm()
    
    # Handle regular form POST request
    if form.validate_on_submit():
        name = form.name.data
        code = form.code.data
        description = form.description.data
        
        # Check if department already exists
        if Department.query.filter_by(name=name).first():
            flash('Department already exists.', 'info')
            return redirect(url_for('main.create_department'))
        
        # Check if code already exists
        if Department.query.filter_by(code=code).first():
            flash('Department code already exists.', 'info')
            return redirect(url_for('main.create_department'))
        
        try:
            # Fixed: Use Department model, not User
            new_department = Department(
                name=name,
                code=code,
                description=description
            )
            db.session.add(new_department)
            db.session.commit()
            flash('Department successfully created.', 'success')
            return redirect(url_for('main.index'))
        except Exception as e:
            db.session.rollback()
            flash(f"An error occurred: {e}", 'danger')
            return redirect(url_for('main.create_department'))
    
    return render_template('department_form.html', form=form)



@main_bp.route('/department/create-new')
@login_required
@admin_required
def create_new_department():
    form = DepartmentForm()
    if form.validate_on_submit():
        try:
            db.session.add(
                Department(
                    name=form.name.data,
                    code=form.code.data,
                    description=form.description.data
                )
            )
            db.session.commit()
            flash("Department created successfully.", 'success')
        except Exception as e:
            flash(f'An error occurred: {e}', 'danger')
        return redirect(url_for('main.home'))
    return render_template('department_form.html', form=form)


@main_bp.route('/department/edit/<int:dept_id>')
@login_required
def edit_department(dept_id):
    if not _is_privileged_admin(current_user):
        flash("Permission Denied", 'danger')
        return redirect(request.referrer)
    
    department = Department.query.get_or_404(dept_id)
    if not department:
        flash('Department does not exist', 'warning')
        return redirect(request.referrer)

    return redirect(request.referrer)


@main_bp.route('/department/delete/<int:dept_id>')
@login_required
def delete_department(dept_id):
    if not _is_privileged_admin(current_user):
        flash("Permission Denied", 'danger')
        return redirect(request.referrer)
    
    department = Department.query.get_or_404(dept_id)
    if not department:
        flash('Department does not exist', 'warning')
        return redirect(request.referrer)
    
    # Delete department object
    db.session.delete(department)
    db.session.commit()

    return redirect(request.referrer)
    
    

@main_bp.route('/category/create', methods=['GET', 'POST'])
@login_required
def create_category():
    form = CategoryForm()
    if form.validate_on_submit():
        name = form.name.data.strip()

        if Category.query.filter_by(name=name).first():
            flash('Category already exists.', 'info')
            return redirect(url_for('main.create_category'))

        try:
            db.session.add(Category(name=name))
            db.session.commit()
            flash('Category created successfully.', 'success')
        except Exception as e:
            db.session.rollback()
            flash(f'An error occurred: {e}', 'danger')
        return redirect(url_for('main.create_category'))

    return render_template(
        'report_form.html', form=form,
        heading='New Category',
        subtitle='Add a new task category',
    )



@main_bp.route('/dashboard')
@login_required
def dashboard():
    tasks = Task.query.filter_by(user_id=current_user.id).order_by(Task.is_complete.asc(), Task.date.desc()).limit(10).all()
    return render_template('dashboard.html', tasks=tasks)



@main_bp.route('/admin-dashboard')
@login_required
@admin_required
def admin_dashboard():
    # Get statistics
    total_users = User.query.count()
    total_departments = Department.query.count()
    total_tasks = Task.query.count()
    completed_tasks = Task.query.filter_by(is_complete=True).count()
    total_reports = Report.query.count()
    
    # Get recent users
    recent_users = User.query.order_by(User.created_at.desc()).limit(5).all()
    
    # Get all departments with user count
    departments = db.session.query(
        Department,
        db.func.count(User.id).label('user_count')
    ).outerjoin(User).group_by(Department.id).all()
    
    # Get task statistics by category. Categorization isn't a DB column for
    # legacy tasks (it falls back to keyword classification), so this can't
    # be a SQL-level GROUP BY — group in Python via effective_category_name
    # instead, to avoid the same fragmentation the reports/stats pages had.
    category_counts = Counter(t.effective_category_name for t in Task.query.all())
    task_stats = sorted(
        (SimpleNamespace(category=name, count=count) for name, count in category_counts.items()),
        key=lambda s: s.count,
        reverse=True,
    )
    
    # Get recent reports
    recent_reports = Report.query.order_by(Report.created_at.desc()).limit(10).all()
    
    # Get all users for management
    all_users = User.query.order_by(User.name).all()
    
    # Calculate completion rate
    completion_rate = (completed_tasks / total_tasks * 100) if total_tasks > 0 else 0
    
    return render_template(
        'admin.html',
        total_users=total_users,
        total_departments=total_departments,
        total_tasks=total_tasks,
        completed_tasks=completed_tasks,
        total_reports=total_reports,
        recent_users=recent_users,
        departments=departments,
        task_stats=task_stats,
        recent_reports=recent_reports,
        all_users=all_users,
        completion_rate=completion_rate
    )


@main_bp.route('/toggle-admin/<int:user_id>', methods=['POST'])
#@csrf.exempt
@login_required
def toggle_admin(user_id):
    """Toggle admin status for a user"""
    
    # Check if current user has permission to modify admin status
    if not _is_privileged_admin(current_user):
        return jsonify({
            'success': False,
            'error': 'Unauthorized: Only admins can modify user roles'
        }), 403
    
    # Prevent users from removing their own admin status
    if current_user.id == user_id:
        return jsonify({
            'success': False,
            'error': 'You cannot modify your own admin status'
        }), 400
    
    try:
        # Get the user from database
        user = User.query.get(user_id)
        
        if not user:
            return jsonify({
                'success': False,
                'error': 'User not found'
            }), 404
        
        # Get the new admin status from request
        data = request.get_json()
        is_admin = data.get('is_admin', False)
        
        # Update user admin status
        user.is_admin = is_admin
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': f'User admin status updated to {is_admin}',
            'user_id': user_id,
            'is_admin': is_admin
        }), 200
        
    except Exception as e:
        db.session.rollback()
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500



@main_bp.route('/task/new', methods=['GET', 'POST'])
@login_required
def new_task():
    form = TaskForm()
    if form.validate_on_submit():
        task = Task(
            user_id=current_user.id,
            date=form.date.data,
            category_id=form.category_id.data,
            description=form.description.data,
            remarks=form.remarks.data,
            is_complete=form.is_complete.data,
        )
        try:
            db.session.add(task)
            db.session.commit()
            flash('Task created successfully!', 'success')
            return redirect(url_for('main.dashboard'))
        except Exception as e:
            flash(f"An error occurred: {e}", 'danger')
            return redirect(url_for('main.dashboard'))
    return render_template('new_task.html', form=form, is_edit=False)



@main_bp.route('/task/edit/<int:id>', methods=['GET', 'POST'])
@login_required
def edit_task(id):
    task = Task.query.get_or_404(id)
    
    # Ensure the task belongs to the current user
    if task.user_id != current_user.id:
        flash('You do not have permission to edit this task.', 'danger')
        return redirect(url_for('main.dashboard'))
    
    form = TaskForm(obj=task)
    
    if form.validate_on_submit():
        task.date = form.date.data
        task.category_id = form.category_id.data
        task.description = form.description.data
        task.remarks = form.remarks.data
        task.is_complete = form.is_complete.data
        task.updated_at = datetime.utcnow()
        
        try:
            db.session.commit()
            flash('Task updated successfully!', 'success')
            return redirect(url_for('main.dashboard'))
        except Exception as e:
            db.session.rollback()
            flash(f"An error occurred: {e}", 'danger')
            return redirect(url_for('main.dashboard'))
    
    return render_template('task_form.html', form=form, task=task, is_edit=True)



def _save_and_send_report(docx_bytes, filename, user_id, start, end, report_type_slug, error_redirect):
    report_type = ReportType.query.filter_by(type=report_type_slug).first()

    random_suffix = secrets.token_hex(8)
    name, ext = os.path.splitext(filename)
    final_filename = f"{name}_{random_suffix}{ext}"
    downloads_dir = os.path.join(current_app.root_path, 'downloads')
    os.makedirs(downloads_dir, exist_ok=True)
    file_path = os.path.join(downloads_dir, final_filename)

    try:
        with open(file_path, 'wb') as f:
            f.write(docx_bytes)

        new_report = Report(
            user_id=user_id,
            start_date=start,
            end_date=end,
            report_type_id=report_type.id,
            report_file=final_filename,
        )
        db.session.add(new_report)
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        flash(f"An error occurred while saving the report: {e}", 'danger')
        return redirect(error_redirect)

    return send_file(
        file_path,
        as_attachment=True,
        download_name=final_filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )


@main_bp.route('/report/weekly', methods=['GET', 'POST'])
@login_required
def report_weekly():
    form = WeeklyReportForm()
    if form.validate_on_submit():
        start = form.start_date.data
        end = form.end_date.data
        tasks = Task.query.filter(
            Task.user_id == current_user.id,
            Task.date >= start,
            Task.date <= end,
        ).order_by(Task.date).all()

        docx_bytes, filename = generate_weekly_report_docs(tasks, current_user, start, end)

        return _save_and_send_report(
            docx_bytes, filename, current_user.id, start, end, 'weekly',
            url_for('main.report_weekly'),
        )

    return render_template(
        'report_form.html', form=form,
        heading='Generate Weekly Report',
        subtitle='Select a date range to generate your personal weekly report',
    )


@main_bp.route('/report/quarterly', methods=['GET', 'POST'])
@login_required
def report_quarterly():
    form = QuarterlyReportForm()
    if form.validate_on_submit():
        start = form.start_date.data
        end = form.end_date.data
        department = Department.query.get_or_404(form.department.data)

        user_ids = [u.id for u in User.query.filter_by(department_id=department.id).all()]
        tasks = []
        if user_ids:
            tasks = Task.query.filter(
                Task.user_id.in_(user_ids),
                Task.date >= start,
                Task.date <= end,
            ).order_by(Task.date).all()

        docx_bytes, filename = generate_quarterly_report_docs(tasks, department, start, end)

        return _save_and_send_report(
            docx_bytes, filename, current_user.id, start, end, 'quarterly',
            url_for('main.report_quarterly'),
        )

    return render_template(
        'report_form.html', form=form,
        heading='Generate Quarterly Report',
        subtitle='Select a department and date range to generate a holistic quarterly report',
    )


@main_bp.route('/view-report/<int:id>', methods=['GET', 'POST'])
@login_required
def view_report(id):
    report = Report.query.get_or_404(id)
    if not report:
        flash("Report not found", 'warning')
        return redirect(request.referrer)
    if not _is_privileged_admin(current_user) and report.user_id != current_user.id:
        flash("unauthorized. You can only view your reports", 'danger')
        return redirect(request.referrer)

    report_filename = report.report_file

    # Search downloads_dir for file
    downloads_dir = os.path.join(current_app.root_path, 'downloads')
    file_path = os.path.join(downloads_dir, report_filename)

    if not os.path.exists(file_path):
        flash("Report file not found on server.", 'warning')
        return redirect(request.referrer or url_for('main.dashboard'))

    try:
        document = Document(file_path)
    except Exception as e:
        current_app.logger.exception("Failed to open report document: %s", e)
        flash("Unable to open report file.", 'danger')
        return redirect(request.referrer or url_for('main.dashboard'))

    doc_paragraphs = [para.text.strip() for para in document.paragraphs if para.text.strip()]
    doc_tables = []

    for table in document.tables:
        parsed_table = []
        for row in table.rows:
            parsed_table.append([cell.text.strip() for cell in row.cells])
        if parsed_table:
            doc_tables.append(parsed_table)

    return render_template(
        'view_report.html',
        title='View Report',
        report_filename=report_filename,
        report_id=id,
        doc_paragraphs=doc_paragraphs,
        doc_tables=doc_tables
    )


def _document_to_html(document):
    """Convert Word document to HTML for Summernote editor"""
    html_parts = []
    
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            html_parts.append('<p><br></p>')
            continue
        
        # Get paragraph style
        style = 'normal'
        if paragraph.style.name.startswith('Heading'):
            level = paragraph.style.name.replace('Heading ', '')
            if level.isdigit():
                style = f'h{level}'
        
        # Check for bold/italic
        runs = paragraph.runs
        if runs:
            is_bold = any(run.bold for run in runs)
            is_italic = any(run.italic for run in runs)
            
            formatted_text = text
            if is_bold:
                formatted_text = f'<strong>{formatted_text}</strong>'
            if is_italic:
                formatted_text = f'<em>{formatted_text}</em>'
            
            # Check alignment
            alignment = ''
            if paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                alignment = ' style="text-align: center;"'
            elif paragraph.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                alignment = ' style="text-align: right;"'
            elif paragraph.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                alignment = ' style="text-align: justify;"'
            
            if style.startswith('h'):
                html_parts.append(f'<{style}{alignment}>{formatted_text}</{style}>')
            else:
                html_parts.append(f'<p{alignment}>{formatted_text}</p>')
        else:
            html_parts.append(f'<p>{text}</p>')
    
    # Handle tables
    for table in document.tables:
        html_parts.append('<table class="table table-bordered">')
        for row in table.rows:
            html_parts.append('<tr>')
            for cell in row.cells:
                cell_text = cell.text.strip()
                html_parts.append(f'<td>{cell_text}</td>')
            html_parts.append('</tr>')
        html_parts.append('</table>')
    
    return '\n'.join(html_parts)


def _html_to_paragraphs(html_content):
    """Convert HTML content from Summernote back to plain text paragraphs"""
    soup = BeautifulSoup(html_content, 'html.parser')
    paragraphs = []
    
    # Process all text-containing elements
    for element in soup.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'div']):
        text = element.get_text(strip=True)
        if text:
            paragraphs.append(text)
    
    # Handle tables
    for table in soup.find_all('table'):
        for row in table.find_all('tr'):
            cells = row.find_all(['td', 'th'])
            if cells:
                row_text = ' | '.join(cell.get_text(strip=True) for cell in cells)
                if row_text:
                    paragraphs.append(row_text)
    
    return paragraphs


def _html_to_document(html_content):
    """Convert HTML content to Word document with formatting preserved"""
    soup = BeautifulSoup(html_content, 'html.parser')
    doc = Document()
    
    for element in soup.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li']):
        text = element.get_text(strip=True)
        if not text:
            doc.add_paragraph()  # Empty paragraph for spacing
            continue
        
        # Determine style based on tag
        if element.name.startswith('h'):
            level = element.name[1]
            paragraph = doc.add_heading(text, level=int(level))
        elif element.name == 'li':
            paragraph = doc.add_paragraph(text, style='List Bullet')
        else:
            paragraph = doc.add_paragraph()
            
            # Process inline formatting
            for content in element.contents:
                if hasattr(content, 'name'):
                    run = paragraph.add_run(content.get_text())
                    
                    # Apply formatting
                    if content.name in ['strong', 'b']:
                        run.bold = True
                    if content.name in ['em', 'i']:
                        run.italic = True
                    if content.name == 'u':
                        run.underline = True
                else:
                    paragraph.add_run(str(content))
           
            # Apply alignment
            style = element.get('style', '')
            if 'text-align: center' in style:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif 'text-align: right' in style:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif 'text-align: justify' in style:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    # Handle tables
    for table_element in soup.find_all('table'):
        rows = table_element.find_all('tr')
        if rows:
            num_cols = max(len(row.find_all(['td', 'th'])) for row in rows)
            table = doc.add_table(rows=len(rows), cols=num_cols)
            table.style = 'Table Grid'
            
            for i, row in enumerate(rows):
                cells = row.find_all(['td', 'th'])
                for j, cell in enumerate(cells):
                    table.rows[i].cells[j].text = cell.get_text(strip=True)
    
    return doc


@main_bp.route('/edit-report/<int:id>', methods=['GET', 'POST'])
@login_required
def edit_report(id):
    form = EditReportForm()

    report = Report.query.get_or_404(id)
    
    # Authorization check
    if not _is_privileged_admin(current_user) and report.user_id != current_user.id:
        flash("Unauthorized. You can only edit your own reports", 'danger')
        return redirect(request.referrer or url_for('main.dashboard'))

    report_filename = report.report_file
    downloads_dir = os.path.join(current_app.root_path, 'downloads')
    file_path = os.path.join(downloads_dir, report_filename)

    if not os.path.exists(file_path):
        flash("Report file not found on server.", 'warning')
        return redirect(request.referrer or url_for('main.dashboard'))

    if request.method == 'POST':
        edited_content = request.form.get('edited_content', '').strip()

        if not edited_content or edited_content == '<p><br></p>':
            flash("Edited content cannot be empty.", 'warning')
            return redirect(request.url)

        try:
            # Convert HTML to Word document with formatting
            updated_document = _html_to_document(edited_content)
            
            # Save the document
            updated_document.save(file_path)
            
            flash("Report updated successfully.", 'success')
            return redirect(url_for('main.view_report', id=id))
            
        except Exception as e:
            current_app.logger.exception("Failed to save edited report: %s", e)
            flash(f"Unable to save edited report: {str(e)}", 'danger')
            return redirect(request.url)

    # GET request - load document
    try:
        document = Document(file_path)
        editable_content = _document_to_html(document)
        form.report.data = editable_content
    except Exception as e:
        current_app.logger.exception("Failed to open report document: %s", e)
        flash("Unable to open report file.", 'danger')
        return redirect(request.referrer or url_for('main.dashboard'))

    return render_template('edit_report.html', title='Edit Report', report_filename=report_filename, 
                           report_id=report.id, editable_content=editable_content, form=form)



@main_bp.route('/report/<int:id>/download', methods=['GET'])
@login_required
def download_report(id):
    report = Report.query.get_or_404(id)
    if not report:
        flash("Report not found", 'warning')
        return redirect(request.referrer)

    if not _is_privileged_admin(current_user) and report.user_id != current_user.id:
        flash("unauthorized. You can only view your reports", 'danger')
        return redirect(request.referrer)

    report_filename = report.report_file

    # Search downloads_dir for file
    downloads_dir = os.path.join(current_app.root_path, 'downloads')
    file_path = os.path.join(downloads_dir, report_filename)

    if not os.path.exists(file_path):
        flash("Report file not found on server.", 'warning')
        return redirect(request.referrer or url_for('main.dashboard'))

    return send_file(file_path, as_attachment=True, download_name=report_filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


@main_bp.route('/reports/mine')
@login_required
def my_reports():
    reports = (
        Report.query
        .filter_by(user_id=current_user.id)
        .order_by(Report.created_at.desc())
        .all()
    )
    return render_template('my_reports.html', reports=reports)


@main_bp.route('/report/<int:id>/delete', methods=['POST'])
@login_required
def delete_report(id):
    report = Report.query.get_or_404(id)

    if not _is_privileged_admin(current_user) and report.user_id != current_user.id:
        flash("Unauthorized. You can only delete your own reports.", 'danger')
        return redirect(url_for('main.my_reports'))

    downloads_dir = os.path.join(current_app.root_path, 'downloads')
    file_path = os.path.join(downloads_dir, report.report_file)

    try:
        db.session.delete(report)
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        flash(f"An error occurred while deleting the report: {e}", 'danger')
        return redirect(url_for('main.my_reports'))

    if os.path.exists(file_path):
        try:
            os.remove(file_path)
        except OSError:
            current_app.logger.exception("Failed to remove report file: %s", file_path)

    flash("Report deleted successfully.", 'success')
    return redirect(url_for('main.my_reports'))


# =====================================================
# STATISTICS / ANALYTICS
# =====================================================

def _parse_date(value):
    if not value:
        return None
    try:
        return datetime.strptime(value, '%Y-%m-%d').date()
    except ValueError:
        return None


def _fig_to_base64(fig):
    buf = BytesIO()
    fig.savefig(buf, format='png', bbox_inches='tight', dpi=110)
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode('utf-8')


def _style_x_labels(ax):
    plt.setp(ax.get_xticklabels(), rotation=30, ha='right')


def _make_bar_chart(labels, values, title, xlabel, ylabel):
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.bar(labels, values, color='#4C72B0')
    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    _style_x_labels(ax)
    fig.tight_layout()
    return _fig_to_base64(fig)


def _make_pie_chart(values, labels, title):
    fig, ax = plt.subplots(figsize=(5, 4))
    colors = ['#55A868', '#C44E52', '#4C72B0', '#8172B2', '#CCB974']
    ax.pie(values, labels=labels, autopct='%1.1f%%', colors=colors[:len(values)])
    ax.set_title(title)
    fig.tight_layout()
    return _fig_to_base64(fig)


def _make_line_chart(labels, values, title, xlabel, ylabel):
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.plot(labels, values, marker='o', color='#4C72B0')
    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    _style_x_labels(ax)
    fig.tight_layout()
    return _fig_to_base64(fig)


def _build_stats_context(department_id, start, end):
    department = Department.query.get_or_404(department_id)

    users = User.query.filter_by(department_id=department_id).all()
    user_ids = [u.id for u in users]
    user_names = {u.id: u.name for u in users}

    tasks = []
    if user_ids:
        tasks = Task.query.filter(
            Task.user_id.in_(user_ids),
            Task.date >= start,
            Task.date <= end,
        ).order_by(Task.date).all()

    context = {
        'department': department,
        'start_date': start,
        'end_date': end,
        'has_data': bool(tasks),
        'total_tasks': len(tasks),
    }

    if not tasks:
        return context

    df = pd.DataFrame([{
        'date': t.date,
        'category': t.effective_category_name,
        'is_complete': bool(t.is_complete),
        'user': user_names.get(t.user_id, 'Unknown'),
    } for t in tasks])

    total_tasks = len(df)
    completed = int(df['is_complete'].sum())
    pending = total_tasks - completed
    completion_rate = round((completed / total_tasks) * 100, 1) if total_tasks else 0

    by_category = (
        df.groupby('category')
        .agg(total=('category', 'count'), completed=('is_complete', 'sum'))
        .reset_index()
        .sort_values('total', ascending=False)
    )
    by_category['completion_rate'] = (by_category['completed'] / by_category['total'] * 100).round(1)

    by_user = (
        df.groupby('user')
        .agg(total=('user', 'count'), completed=('is_complete', 'sum'))
        .reset_index()
        .sort_values('total', ascending=False)
    )
    by_user['completion_rate'] = (by_user['completed'] / by_user['total'] * 100).round(1)

    weekly = df.copy()
    weekly['week'] = pd.to_datetime(weekly['date']).dt.to_period('W').apply(lambda p: p.start_time.date())
    by_week = weekly.groupby('week').size().reset_index(name='count').sort_values('week')

    charts = {
        'category_chart': _make_bar_chart(
            by_category['category'].tolist(), by_category['total'].tolist(),
            'Tasks by Category', 'Category', 'Tasks',
        ),
        'completion_chart': _make_pie_chart(
            [completed, pending], ['Completed', 'Pending'], 'Completion Status',
        ),
        'user_chart': _make_bar_chart(
            by_user['user'].tolist(), by_user['total'].tolist(),
            'Tasks by User', 'User', 'Tasks',
        ),
        'timeline_chart': _make_line_chart(
            [d.isoformat() for d in by_week['week']], by_week['count'].tolist(),
            'Tasks Over Time (Weekly)', 'Week Starting', 'Tasks',
        ),
    }

    context.update({
        'completed': completed,
        'pending': pending,
        'completion_rate': completion_rate,
        'by_category': by_category.to_dict('records'),
        'by_user': by_user.to_dict('records'),
        'charts': charts,
    })

    return context


@main_bp.route('/stats', methods=['GET', 'POST'])
@login_required
def generate_stats():
    form = StatsForm()
    if form.validate_on_submit():
        return redirect(url_for(
            'main.view_stats',
            department_id=form.department.data,
            start=form.start_date.data.isoformat(),
            end=form.end_date.data.isoformat(),
        ))
    return render_template('stats_form.html', form=form)


@main_bp.route('/stats/view')
@login_required
def view_stats():
    department_id = request.args.get('department_id', type=int)
    start = _parse_date(request.args.get('start'))
    end = _parse_date(request.args.get('end'))

    if not department_id or not start or not end:
        flash('Missing or invalid statistics parameters.', 'danger')
        return redirect(url_for('main.generate_stats'))

    context = _build_stats_context(department_id, start, end)
    return render_template('stats.html', **context)


@main_bp.route('/stats/download')
@login_required
def download_stats():
    department_id = request.args.get('department_id', type=int)
    start = _parse_date(request.args.get('start'))
    end = _parse_date(request.args.get('end'))

    if not department_id or not start or not end:
        flash('Missing or invalid statistics parameters.', 'danger')
        return redirect(url_for('main.generate_stats'))

    context = _build_stats_context(department_id, start, end)
    html_content = render_template('stats_pdf.html', **context)

    pdf_buffer = BytesIO()
    pisa_status = pisa.CreatePDF(html_content, dest=pdf_buffer)

    if pisa_status.err:
        flash('Failed to generate PDF report.', 'danger')
        return redirect(url_for(
            'main.view_stats',
            department_id=department_id,
            start=start.isoformat(),
            end=end.isoformat(),
        ))

    pdf_buffer.seek(0)
    safe_name = (context['department'].code or context['department'].name).replace(' ', '_')
    filename = f"stats_{safe_name}_{start}_{end}.pdf"

    return send_file(
        pdf_buffer,
        as_attachment=True,
        download_name=filename,
        mimetype='application/pdf',
    )



