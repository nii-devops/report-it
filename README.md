# Report-IT

A Flask-based reporting tool for IT teams. Log daily tasks as you complete them, then generate polished, AI-written weekly or quarterly reports as downloadable Word documents — no more manually writing up what you did every week.

## Features

- **User accounts** — email/password registration and login, plus Google OAuth sign-in
- **Password recovery** — email-based reset codes with expiry
- **Task logging** — record daily IT tasks with type, description, status, and remarks
- **AI-generated reports** — turn a date range of tasks into a structured `.docx` report:
  - **Weekly reports**: grouped, human-readable summary of the week's work
  - **Quarterly reports**: holistic report grouped by category/theme, with a table of contents, key metrics, trends, challenges, and recommendations
  - Uses Anthropic as the primary provider with automatic fallback to OpenAI if it's unavailable
- **Report management** — view, edit (rich text), re-download, or regenerate past reports
- **Departments** — organize users by department/unit
- **Admin dashboard** — org-wide stats, user/department overview, task breakdown, and admin role management
- **Session timeout** — automatic logout after 15 minutes of inactivity

## Tech Stack

- **Backend**: Flask, Flask-SQLAlchemy, Flask-Login, Flask-WTF, Flask-Migrate
- **Auth**: Authlib (Google OAuth), Werkzeug password hashing
- **Database**: SQLite (default/local) or PostgreSQL (via `psycopg`/`psycopg2`)
- **AI providers**: Anthropic (Claude) and OpenAI
- **Documents**: `python-docx` for report generation/editing
- **Email**: Flask-Mail / smtplib for password reset codes
- **Frontend**: Jinja2 templates, Bootstrap-Flask, CKEditor

## Prerequisites

- Python 3.11+
- A SQLite file (default) or a reachable PostgreSQL database
- API keys for Anthropic and/or OpenAI (at least one is required for report generation)
- (Optional) Google OAuth client credentials, for "Sign in with Google"
- (Optional) SMTP credentials, for password reset emails

## Setup

1. **Clone the repository**

   ```bash
   git clone <repository-url>
   cd REPORT_IT
   ```

2. **Create and activate a virtual environment**

   ```bash
   python -m venv virt
   # Windows
   virt\Scripts\activate
   # macOS/Linux
   source virt/bin/activate
   ```

3. **Install dependencies**

   ```bash
   pip install -r requirements.txt
   ```

4. **Configure environment variables**

   Create a `.env` file in the project root:

   ```env
   SECRET_KEY=your-secret-key

   # Database (defaults to local sqlite:///reportit.db if unset)
   DATABASE_URL=postgresql://user:password@host:5432/dbname

   # AI providers (at least one required)
   ANTHROPIC_API_KEY=your-anthropic-key
   OPENAI_API_KEY=your-openai-key

   # Google OAuth (optional)
   GOOGLE_CLIENT_ID=your-google-client-id
   GOOGLE_CLIENT_SECRET=your-google-client-secret

   # Mail / password reset (optional)
   MAIL_SERVER=smtp.gmail.com
   MAIL_PORT=587
   MAIL_USERNAME=your-email@example.com
   MAIL_PASSWORD=your-email-app-password
   SMTP_SERVER=smtp.gmail.com
   SMTP_PORT=587
   ```

5. **Run the app**

   ```bash
   python run.py
   ```

   Tables and the default report types (`weekly`, `quarterly`) are created automatically on first startup. The app is served at `http://127.0.0.1:5000` by default.

## Project Structure

```text
REPORT_IT/
├── app/
│   ├── __init__.py          # App factory, extension init, table/seed setup
│   ├── ai_integration.py    # AI prompt building + report generation (docx)
│   ├── forms.py             # WTForms definitions
│   ├── models.py            # SQLAlchemy models (User, Task, Report, Department, ...)
│   ├── routes.py            # Route handlers (auth + main blueprints)
│   ├── static/               # CSS, JS, images, uploads
│   ├── templates/            # Jinja2 templates
│   └── downloads/            # Generated report files (gitignored)
├── config.py                 # App configuration (Config/DevelopmentConfig/ProductionConfig)
├── run.py                    # App entry point
├── requirements.txt
└── instance/                 # Local SQLite DB (gitignored)
```

## Usage

1. Register an account (or sign in with Google) and select your department.
2. Log tasks as you complete them from the dashboard.
3. Go to **Generate Report**, pick a date range and report type (weekly/quarterly), and submit.
4. The app compiles your tasks, sends them to the AI provider, and returns a formatted `.docx` report — saved to your report history and downloadable at any time.
5. Admins can access **Admin Dashboard** for organization-wide statistics and user/department management.
