import os
import re
import html
from datetime import datetime, timedelta
from flask import Blueprint, render_template, redirect, url_for, flash, request, current_app, send_file, session, jsonify
from flask_login import login_user, logout_user, login_required, current_user
from io import BytesIO
from authlib.common.security import generate_token
from . import db, oauth
from .models import *
from .forms import *
from .ai_integration import generate_report_docs
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


main_bp = Blueprint('main', __name__)
auth_bp = Blueprint('auth', __name__)


csrf = CSRF()


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



@main_bp.route('/')
def index():
    return render_template('index.html')


@auth_bp.route('/login')
def login():
    # Generate a unique nonce for CSRF protection during the OAuth flow
    session['nonce'] = generate_token()
    redirect_uri = url_for('auth.auth_callback', _external=True)
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
def create_department():
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



@main_bp.route('/dashboard')
@login_required
def dashboard():
    tasks = Task.query.filter_by(user_id=current_user.id).order_by(Task.is_complete.asc(), Task.date.desc()).limit(10).all()
    return render_template('dashboard.html', tasks=tasks)



@main_bp.route('/admin-dashboard')
@login_required
def admin_dashboard():
    # Check if user is admin (you may want to add an is_admin field to User model)
    # For now, we'll allow all logged-in users, but you should add proper authorization
    
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
    
    # Get task statistics by type
    task_stats = db.session.query(
        Task.task_type,
        db.func.count(Task.id).label('count')
    ).group_by(Task.task_type).all()
    
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
    if not current_user.is_admin:
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
            task_type=form.task_type.data,
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
    return render_template('task_form.html', form=form, is_edit=False)



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
        task.task_type = form.task_type.data
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



@main_bp.route('/report', methods=['GET', 'POST'])
@login_required
def report():
    form = ReportForm()
    if form.validate_on_submit():
        start = form.start_date.data
        end = form.end_date.data
        rtype = int(form.report_type.data)
        tasks = Task.query.filter(Task.user_id == current_user.id, Task.date >= start, Task.date <= end).order_by(Task.date).all()

        # Send to AI agent aggregator and return a docx in memory
        docx_bytes, filename = generate_report_docs(tasks, current_user, start, end, rtype)

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
                user_id=current_user.id,
                start_date=start,
                end_date=end,
                report_type_id=rtype,
                report_file=final_filename,
            )
            db.session.add(new_report)
            db.session.commit()
        except Exception as e:
            db.session.rollback()
            flash(f"An error occurred while saving the report: {e}", 'danger')
            return redirect(url_for('main.report'))

        return send_file(
            file_path,
            as_attachment=True,
            download_name=final_filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    return render_template('report.html', form=form)


@main_bp.route('/view-report/<int:id>', methods=['GET', 'POST'])
@login_required
def view_report(id):
    report = Report.query.get_or_404(id)
    if not report:
        flash("Report not found", 'warning')
        return redirect(request.referrer)
    if not current_user.is_admin and report.user_id != current_user.id:
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
    if not current_user.is_admin and report.user_id != current_user.id:
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

    return render_template(
        'edit_rep.html',
        title='Edit Report',
        report_filename=report_filename,
        report_id=report.id,
        editable_content=editable_content,
        form=form
    )



"""
@main_bp.route('/edit-report/<int:id>', methods=['GET', 'POST'])
@login_required
def edit_report(id):
    report = Report.query.get_or_404(id)
    if not report:
        flash("Report not found", 'warning')
        return redirect(request.referrer)
    if not current_user.is_admin and report.user_id != current_user.id:
        flash("unauthorized. You can only view your reports", 'danger')
        return redirect(request.referrer)

    report_filename = report.report_file

    # Search downloads_dir for file
    downloads_dir = os.path.join(current_app.root_path, 'downloads')
    file_path = os.path.join(downloads_dir, report_filename)

    if not os.path.exists(file_path):
        flash("Report file not found on server.", 'warning')
        return redirect(request.referrer or url_for('main.dashboard'))

    if request.method == 'POST':
        edited_content = request.form.get('edited_content', '').strip()

        if not edited_content:
            flash("Edited content cannot be empty.", 'warning')
            return redirect(request.url)

        paragraph_texts = _html_to_paragraphs(edited_content)

        if not paragraph_texts:
            flash("No valid text found to save.", 'warning')
            return redirect(request.url)

        updated_document = Document()
        for paragraph in paragraph_texts:
            updated_document.add_paragraph(paragraph)

        try:
            updated_document.save(file_path)
        except Exception as e:
            current_app.logger.exception("Failed to save edited report document: %s", e)
            flash("Unable to save edited report.", 'danger')
            return redirect(request.url)

        flash("Report updated successfully.", 'success')
        return redirect(url_for('main.view_report', id=id))

    try:
        document = Document(file_path)
    except Exception as e:
        current_app.logger.exception("Failed to open report document: %s", e)
        flash("Unable to open report file.", 'danger')
        return redirect(request.referrer or url_for('main.dashboard'))

    editable_content = _document_to_html(document)

    return render_template(
        'edit_rep.html',
        title='Edit Report',
        report_filename=report_filename,
        report_id=id,
        editable_content=editable_content
    )

"""

@main_bp.route('/report/<int:id>/download', methods=['GET'])
@login_required
def download_report(id):
    report = Report.query.get_or_404(id)
    if not report:
        flash("Report not found", 'warning')
        return redirect(request.referrer)

    if not current_user.is_admin and report.user_id != current_user.id:
        flash("unauthorized. You can only view your reports", 'danger')
        return redirect(request.referrer)

    report_filename = report.report_file

    # Search downloads_dir for file
    downloads_dir = os.path.join(current_app.root_path, 'downloads')
    file_path = os.path.join(downloads_dir, report_filename)

    if not os.path.exists(file_path):
        flash("Report file not found on server.", 'warning')
        return redirect(request.referrer or url_for('main.dashboard'))

    return send_file(
        file_path,
        as_attachment=True,
        download_name=report_filename,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )



"""
Old routes
@main_bp.route('/report', methods=['GET', 'POST'])
@login_required
def report():
    form = ReportForm()
    if form.validate_on_submit():
        start = form.start_date.data
        end = form.end_date.data
        rtype = form.report_type.data
        tasks = Task.query.filter(Task.user_id == current_user.id, Task.date >= start, Task.date <= end).order_by(Task.date).all()

        # Send to AI agent aggregator and return a docx in memory
        docx_bytes, filename = generate_report_docs(tasks, current_user, start, end, rtype)

        return send_file(
            BytesIO(docx_bytes),
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    return render_template('report.html', form=form)


@main_bp.route('/report', methods=['GET', 'POST'])
@login_required
def report():
    form = ReportForm()
    if form.validate_on_submit():
        start = form.start_date.data
        end = form.end_date.data
        rtype = form.report_type.data
        tasks = Task.query.filter(Task.user_id == current_user.id, Task.date >= start, Task.date <= end).order_by(Task.date).all()

        # Send to AI agent aggregator and return a docx in memory
        docx_bytes, filename = generate_report_docs(tasks, current_user, start, end, rtype)

        return send_file(
            BytesIO(docx_bytes),
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    return render_template('report.html', form=form)




"""

