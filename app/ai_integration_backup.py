"""
Handles sending Task data to multiple AI providers (ChatGPT/OpenAI, Anthropic/Claude, Grok) and building a combined report.
This file provides a generic implementation and placeholders for each provider. Add provider-specific SDK calls and
API keys via environment variables. Keep in mind each provider has its own rate limits and request format.
"""
import os
import json
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import requests
from io import BytesIO


OPENAI_API_KEY = os.environ.get('OPENAI_API_KEY')
ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY')
GROK_API_KEY = os.environ.get('GROK_API_KEY')
AI_TIMEOUT = int(os.environ.get('AI_AGENT_TIMEOUT_SECONDS', 30))


def _tasks_to_text(tasks):
    """Convert tasks to structured text with clear separators""""""
Handles sending Task data to multiple AI providers (ChatGPT/OpenAI, Anthropic/Claude, Grok) and building a combined report.
This file provides a generic implementation and placeholders for each provider. Add provider-specific SDK calls and
API keys via environment variables. Keep in mind each provider has its own rate limits and request format.
"""
import os
import re
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import requests
from io import BytesIO


OPENAI_API_KEY = os.environ.get('OPENAI_API_KEY')
ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY')
GROK_API_KEY = os.environ.get('GROK_API_KEY')
AI_TIMEOUT = int(os.environ.get('AI_AGENT_TIMEOUT_SECONDS', 30))


def _tasks_to_text(tasks):
    """Convert tasks to structured text with clear separators"""
    lines = []
    lines.append(f"TOTAL TASKS IN PERIOD: {len(tasks)}")
    lines.append("=" * 80)
    
    for idx, t in enumerate(tasks, 1):
        lines.append(f"\nTASK #{idx}")
        lines.append(f"Date: {t.date.isoformat()}")
        lines.append(f"Type: {t.task_type}")
        lines.append(f"Status: {'COMPLETED' if t.is_complete else 'PENDING'}")
        lines.append(f"Description: {t.description}")
        if t.remarks:
            lines.append(f"Remarks: {t.remarks}")
        lines.append("-" * 80)
    
    return '\n'.join(lines)


def _build_prompt(tasks_text, report_type, user_name, start_date, end_date):
    """Build a comprehensive prompt that ensures all tasks are addressed"""
    
    if report_type == 'quarterly':
        prompt = f"""You are an IT professional writing a comprehensive quarterly report for {user_name}.

REPORTING PERIOD: {start_date.isoformat()} to {end_date.isoformat()}

CRITICAL INSTRUCTIONS:
1. You MUST address EVERY SINGLE TASK listed below - do not skip or summarize any task
2. Group tasks by type/category for better organization
3. For each task, mention: date, what was done, current status, and any remarks
4. Create a professional quarterly report structure with:
   - Executive Summary (high-level overview of ALL activities)
   - Introduction
   - Detailed Findings (organized by task type, covering ALL tasks)
   - Key Metrics (completion rates, task distribution)
   - Challenges and Solutions (from remarks)
   - Recommendations for next quarter
   - Conclusion

TASKS DATA:
{tasks_text}

Remember: Every task listed above must appear in your report. Do not omit any task, even if it seems minor.
Format your response in markdown for proper formatting."""

    else:  # weekly report
        prompt = f"""You are an IT professional writing a detailed weekly report for {user_name}.

REPORTING PERIOD: {start_date.isoformat()} to {end_date.isoformat()}

CRITICAL INSTRUCTIONS:
1. You MUST mention EVERY SINGLE TASK listed below by date and description
2. Create a day-by-day breakdown if helpful
3. For each task, include: date, task type, description, status, and any remarks
4. Structure the report with:
   - Week Overview (summary of ALL activities)
   - Daily/Task Breakdown (detailed list of ALL tasks)
   - Completed vs Pending Tasks
   - Issues Encountered (from remarks)
   - Action Items for Next Week
   - Metrics Summary

TASKS DATA:
{tasks_text}

Remember: List every single task above in your report. Account for all {tasks_text.split('TOTAL TASKS IN PERIOD:')[1].split()[0]} tasks.
Format your response in markdown for proper formatting."""

    return prompt


def _call_openai(prompt):
    """Call OpenAI API with better parameters for comprehensive reports"""
    headers = {'Authorization': f'Bearer {OPENAI_API_KEY}', 'Content-Type': 'application/json'}
    body = {
        'model': 'gpt-4o-mini',
        'messages': [
            {
                'role': 'system', 
                'content': 'You are a meticulous IT report writer who ensures every task is documented. Never skip or omit tasks from your reports.'
            },
            {
                'role': 'user', 
                'content': prompt
            }
        ],
        'max_tokens': 3000,
        'temperature': 0.7,
    }
    resp = requests.post('https://api.openai.com/v1/chat/completions', headers=headers, json=body, timeout=AI_TIMEOUT)
    resp.raise_for_status()
    data = resp.json()
    return data['choices'][0]['message']['content']


def _call_anthropic(prompt):
    """Call Anthropic API - Updated to use Messages API"""
    headers = {
        'x-api-key': ANTHROPIC_API_KEY, 
        'anthropic-version': '2023-06-01',
        'Content-Type': 'application/json'
    }
    body = {
        'model': 'claude-3-5-sonnet-20241022',
        'max_tokens': 3000,
        'messages': [
            {
                'role': 'user',
                'content': prompt
            }
        ],
        'system': 'You are a thorough IT report writer. You must include every single task in your reports without exception.'
    }
    resp = requests.post('https://api.anthropic.com/v1/messages', headers=headers, json=body, timeout=AI_TIMEOUT)
    resp.raise_for_status()
    data = resp.json()
    return data['content'][0]['text']


def _call_grok(prompt):
    """Call Grok API - placeholder for Grok-style API"""
    headers = {'Authorization': f'Bearer {GROK_API_KEY}', 'Content-Type': 'application/json'}
    body = {
        'messages': [
            {
                'role': 'system',
                'content': 'You are a comprehensive IT report writer who documents every task without omission.'
            },
            {
                'role': 'user',
                'content': prompt
            }
        ],
        'model': 'grok-beta',
        'max_tokens': 3000
    }
    resp = requests.post('https://api.x.ai/v1/chat/completions', headers=headers, json=body, timeout=AI_TIMEOUT)
    resp.raise_for_status()
    data = resp.json()
    return data['choices'][0]['message']['content']


def _select_best_response(responses, tasks_count):
    """Select the response that appears most complete"""
    valid_responses = [r for r in responses if not r.startswith('OpenAI failed') and not r.startswith('Anthropic failed') and not r.startswith('Grok failed')]
    
    if not valid_responses:
        return "Error: All AI providers failed to generate a report."
    
    # Simple heuristic: prefer longer responses as they likely cover more tasks
    best = max(valid_responses, key=lambda x: (str(tasks_count) in x, len(x)))
    return best


def _parse_markdown_to_docx(doc, markdown_text):
    """Parse markdown text and add properly formatted content to Word document"""
    lines = markdown_text.split('\n')
    i = 0
    
    while i < len(lines):
        line = lines[i].strip()
        
        if not line:
            i += 1
            continue
        
        # Handle headings
        if line.startswith('# '):
            heading_text = line.replace('# ', '').strip()
            doc.add_heading(heading_text, level=1)
        elif line.startswith('## '):
            heading_text = line.replace('## ', '').strip()
            doc.add_heading(heading_text, level=2)
        elif line.startswith('### '):
            heading_text = line.replace('### ', '').strip()
            doc.add_heading(heading_text, level=3)
        
        # Handle bullet points
        elif line.startswith('- ') or line.startswith('* '):
            bullet_text = line[2:].strip()
            p = doc.add_paragraph(style='List Bullet')
            _add_formatted_text(p, bullet_text)
        
        # Handle numbered lists
        elif re.match(r'^\d+\.', line):
            list_text = re.sub(r'^\d+\.\s*', '', line).strip()
            p = doc.add_paragraph(style='List Number')
            _add_formatted_text(p, list_text)
        
        # Handle regular paragraphs
        else:
            p = doc.add_paragraph()
            _add_formatted_text(p, line)
            
            # Set font size for regular paragraphs
            for run in p.runs:
                run.font.size = Pt(11)
        
        i += 1


def _add_formatted_text(paragraph, text):
    """Add text to paragraph with bold and italic formatting"""
    # Pattern to match **bold** and *italic*
    parts = re.split(r'(\*\*.*?\*\*|\*.*?\*)', text)
    
    for part in parts:
        if not part:
            continue
            
        if part.startswith('**') and part.endswith('**'):
            # Bold text
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        elif part.startswith('*') and part.endswith('*'):
            # Italic text
            run = paragraph.add_run(part[1:-1])
            run.italic = True
        else:
            # Normal text
            paragraph.add_run(part)


def generate_report_docs(tasks, user, start_date, end_date, report_type):
    """
    Given a list of Task objects, send them to multiple AI providers and produce a .docx bytes object.
    Returns: (bytes, filename)
    """
    if not tasks:
        # Build an empty report doc
        doc = Document()
        doc.add_heading(f'{report_type.title()} Report - {user.name}', level=1)
        doc.add_paragraph('No tasks found in the selected period.')
        bio = BytesIO()
        doc.save(bio)
        return bio.getvalue(), f'report_{report_type}_{start_date}_{end_date}.docx'

    # Convert tasks to structured text
    tasks_text = _tasks_to_text(tasks)
    
    # Build comprehensive prompt
    prompt = _build_prompt(tasks_text, report_type, user.name, start_date, end_date)

    responses = []
    # Call each provider with error handling
    if OPENAI_API_KEY:
        try:
            responses.append(_call_openai(prompt))
        except Exception as e:
            responses.append(f'OpenAI failed: {e}')
    
    if ANTHROPIC_API_KEY:
        try:
            responses.append(_call_anthropic(prompt))
        except Exception as e:
            responses.append(f'Anthropic failed: {e}')
    
    if GROK_API_KEY:
        try:
            responses.append(_call_grok(prompt))
        except Exception as e:
            responses.append(f'Grok failed: {e}')

    # Select best response
    final_text = _select_best_response(responses, len(tasks))

    # Create Word document with proper formatting
    doc = Document()
    
    # Add document header with metadata
    title = doc.add_heading(f'{report_type.title()} Report', level=1)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # Metadata section
    metadata_para = doc.add_paragraph()
    metadata_para.add_run(f'Author: ').bold = True
    metadata_para.add_run(f'{user.name} ({user.email})\n')
    metadata_para.add_run(f'Period: ').bold = True
    metadata_para.add_run(f'{start_date.isoformat()} to {end_date.isoformat()}\n')
    metadata_para.add_run(f'Total Tasks: ').bold = True
    metadata_para.add_run(f'{len(tasks)}\n')
    
    completed = sum(1 for t in tasks if t.is_complete)
    metadata_para.add_run(f'Completed: ').bold = True
    metadata_para.add_run(f'{completed} | ')
    metadata_para.add_run(f'Pending: ').bold = True
    metadata_para.add_run(f'{len(tasks) - completed}')
    
    doc.add_paragraph()  # Empty line
    
    # Parse and add AI-generated content with proper formatting
    _parse_markdown_to_docx(doc, final_text)

    # Save to bytes
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    filename = f'report_{report_type}_{start_date.isoformat()}_{end_date.isoformat()}.docx'
    return bio.getvalue(), filename
    lines = []
    lines.append(f"TOTAL TASKS IN PERIOD: {len(tasks)}")
    lines.append("=" * 80)
    
    for idx, t in enumerate(tasks, 1):
        lines.append(f"\nTASK #{idx}")
        lines.append(f"Date: {t.date.isoformat()}")
        lines.append(f"Type: {t.task_type}")
        lines.append(f"Status: {'COMPLETED' if t.is_complete else 'PENDING'}")
        lines.append(f"Description: {t.description}")
        if t.remarks:
            lines.append(f"Remarks: {t.remarks}")
        lines.append("-" * 80)
    
    return '\n'.join(lines)


def _build_prompt(tasks_text, report_type, user_name, start_date, end_date):
    """Build a comprehensive prompt that ensures all tasks are addressed"""
    
    if report_type == 'quarterly':
        prompt = f"""You are an IT professional writing a comprehensive quarterly report for {user_name}.

REPORTING PERIOD: {start_date.isoformat()} to {end_date.isoformat()}

CRITICAL INSTRUCTIONS:
1. You MUST address EVERY SINGLE TASK listed below - do not skip or summarize any task
2. Group tasks by type/category for better organization
3. For each task, mention: date, what was done, current status, and any remarks
4. Create a professional quarterly report structure with:
   - Executive Summary (high-level overview of ALL activities)
   - Table of Contents
   - Introduction
   - Detailed Findings (organized by task type, covering ALL tasks)
   - Key Metrics (completion rates, task distribution)
   - Challenges and Solutions (from remarks)
   - Recommendations for next quarter. Reserch and add industry best practices
   - Conclusion

TASKS DATA:
{tasks_text}

Remember: Every task listed above must appear in your report. Do not omit any task, even if it seems minor."""

    else:  # weekly report
        prompt = f"""You are an IT professional writing a detailed weekly report for {user_name}.

REPORTING PERIOD: {start_date.isoformat()} to {end_date.isoformat()}

CRITICAL INSTRUCTIONS:
1. You MUST mention EVERY SINGLE TASK listed below by date and description
2. Group similar tasks together and add counts where necessary, but DO NOT create a day-by-day breakdown
3. For each task or group, include: task type, merge description and remarks and generate humanized paragraph(s), and status
4. Structure the report with:
   - Week Overview (summary of ALL activities)
   - Task/Task group Breakdown (detailed list of ALL tasks)
   - Completed vs Pending Tasks
   - Issues Encountered (where applicable)
   - Action Items for Later Week(s)
   - Metrics Summary

TASKS DATA:
{tasks_text}

Remember: List every single task above in your report. Account for all {tasks_text.split('TOTAL TASKS IN PERIOD:')[1].split()[0]} tasks."""

    return prompt


def _call_openai(prompt):
    """Call OpenAI API with better parameters for comprehensive reports"""
    headers = {'Authorization': f'Bearer {OPENAI_API_KEY}', 'Content-Type': 'application/json'}
    body = {
        'model': 'gpt-4o-mini',
        'messages': [
            {
                'role': 'system', 
                'content': 'You are a meticulous IT report writer who ensures every task is documented. Never skip or omit tasks from your reports.'
            },
            {
                'role': 'user', 
                'content': prompt
            }
        ],
        'max_tokens': 3000,
        'temperature': 0.7,
    }
    resp = requests.post('https://api.openai.com/v1/chat/completions', headers=headers, json=body, timeout=AI_TIMEOUT)
    resp.raise_for_status()
    data = resp.json()
    return data['choices'][0]['message']['content']


def _call_anthropic(prompt):
    """Call Anthropic API - Updated to use Messages API"""
    headers = {
        'x-api-key': ANTHROPIC_API_KEY, 
        'anthropic-version': '2023-06-01',
        'Content-Type': 'application/json'
    }
    body = {
        'model': 'claude-3-5-sonnet-20241022',
        'max_tokens': 3000,
        'messages': [
            {
                'role': 'user',
                'content': prompt
            }
        ],
        'system': 'You are a thorough IT report writer. You must include every single task in your reports without exception.'
    }
    resp = requests.post('https://api.anthropic.com/v1/messages', headers=headers, json=body, timeout=AI_TIMEOUT)
    resp.raise_for_status()
    data = resp.json()
    return data['content'][0]['text']


def _call_grok(prompt):
    """Call Grok API - placeholder for Grok-style API"""
    headers = {'Authorization': f'Bearer {GROK_API_KEY}', 'Content-Type': 'application/json'}
    body = {
        'messages': [
            {
                'role': 'system',
                'content': 'You are a comprehensive IT report writer who documents every task without omission.'
            },
            {
                'role': 'user',
                'content': prompt
            }
        ],
        'model': 'grok-beta',
        'max_tokens': 3000
    }
    resp = requests.post('https://api.x.ai/v1/chat/completions', headers=headers, json=body, timeout=AI_TIMEOUT)
    resp.raise_for_status()
    data = resp.json()
    return data['choices'][0]['message']['content']


def _select_best_response(responses, tasks_count):
    """Select the response that appears most complete"""
    valid_responses = [r for r in responses if not r.startswith('OpenAI failed') and not r.startswith('Anthropic failed') and not r.startswith('Grok failed')]
    
    if not valid_responses:
        return "Error: All AI providers failed to generate a report."
    
    # Simple heuristic: prefer longer responses as they likely cover more tasks
    # Also check if response mentions task count
    best = max(valid_responses, key=lambda x: (str(tasks_count) in x, len(x)))
    return best


def generate_report_docs(tasks, user, start_date, end_date, report_type):
    """
    Given a list of Task objects, send them to multiple AI providers and produce a .docx bytes object.
    Returns: (bytes, filename)
    """
    if not tasks:
        # Build an empty report doc
        doc = Document()
        doc.add_heading(f'{report_type.title()} Report - {user.name}', level=1)
        doc.add_paragraph('No tasks found in the selected period.')
        bio = BytesIO()
        doc.save(bio)
        return bio.getvalue(), f'report_{report_type}_{start_date}_{end_date}.docx'

    # Convert tasks to structured text
    tasks_text = _tasks_to_text(tasks)
    
    # Build comprehensive prompt
    prompt = _build_prompt(tasks_text, report_type, user.name, start_date, end_date)

    responses = []
    # Call each provider with error handling
    if OPENAI_API_KEY:
        try:
            responses.append(_call_openai(prompt))
        except Exception as e:
            responses.append(f'OpenAI failed: {e}')
    
    if ANTHROPIC_API_KEY:
        try:
            responses.append(_call_anthropic(prompt))
        except Exception as e:
            responses.append(f'Anthropic failed: {e}')
    
    if GROK_API_KEY:
        try:
            responses.append(_call_grok(prompt))
        except Exception as e:
            responses.append(f'Grok failed: {e}')

    # Select best response
    final_text = _select_best_response(responses, len(tasks))

    # Convert final_text into a nicely formatted docx
    doc = Document()
    
    # Title
    title = doc.add_heading(f'{report_type.title()} Report', level=1)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # Metadata
    doc.add_paragraph(f'Author: {user.name} ({user.email})')
    doc.add_paragraph(f'Period: {start_date.isoformat()} to {end_date.isoformat()}')
    doc.add_paragraph(f'Total Tasks: {len(tasks)}')
    completed = sum(1 for t in tasks if t.is_complete)
    doc.add_paragraph(f'Completed: {completed} | Pending: {len(tasks) - completed}')
    doc.add_paragraph('')

    # Add AI-generated content
    for line in final_text.split('\n'):
        if line.strip():
            # Check if it's a heading (contains #, all caps, or ends with :)
            if line.strip().startswith('#'):
                doc.add_heading(line.replace('#', '').strip(), level=2)
            elif line.strip().isupper() and len(line.strip()) < 60:
                doc.add_heading(line.strip(), level=2)
            else:
                p = doc.add_paragraph(line.strip())
                for run in p.runs:
                    run.font.size = Pt(11)

    # Save to bytes
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    filename = f'report_{report_type}_{start_date.isoformat()}_{end_date.isoformat()}.docx'
    return bio.getvalue(), filename


