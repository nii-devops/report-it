"""
Handles sending Task data to multiple AI providers (ChatGPT/OpenAI, Anthropic/Claude, Grok) and building a combined report.
This file provides a generic implementation and placeholders for each provider. Add provider-specific SDK calls and
API keys via environment variables. Keep in mind each provider has its own rate limits and request format.
"""
import os
import json
from datetime import datetime
from docx import Document
from docx.shared import Pt
import requests
from io import BytesIO



OPENAI_API_KEY = os.environ.get('OPENAI_API_KEY')
ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY')
GROK_API_KEY = os.environ.get('GROK_API_KEY')
AI_TIMEOUT = int(os.environ.get('AI_AGENT_TIMEOUT_SECONDS', 30))


def _tasks_to_text(tasks):
    lines = []
    for t in tasks:
        lines.append(f"Date: {t.date.isoformat()} | Type: {t.task_type} | Complete: {t.is_complete}\nDescription: {t.description}\nRemarks: {t.remarks}\n---")
        return '\n'.join(lines)


def _call_openai(prompt):
    # Example using OpenAI Chat completions (replace with official client if installed)
    headers = {'Authorization': f'Bearer {OPENAI_API_KEY}', 'Content-Type': 'application/json'}
    body = {
    'model': 'gpt-4o-mini',
    'messages': [{'role': 'user', 'content': prompt}],
    'max_tokens': 1500,
    }
    resp = requests.post('https://api.openai.com/v1/chat/completions', headers=headers, json=body, timeout=AI_TIMEOUT)
    resp.raise_for_status()
    data = resp.json()
    # This assumes response structure; adapt as provider docs indicate
    return data['choices'][0]['message']['content']


def _call_anthropic(prompt):
    # Example placeholder for Claude via Anthropic API
    headers = {'x-api-key': ANTHROPIC_API_KEY, 'Content-Type': 'application/json'}
    body = { 'prompt': prompt, 'max_tokens_to_sample': 2000 }
    resp = requests.post('https://api.anthropic.com/v1/complete', headers=headers, json=body, timeout=AI_TIMEOUT)
    resp.raise_for_status()
    return resp.json().get('completion')


def _call_grok(prompt):
    # Placeholder for Grok-style API (adjust URL & auth)
    headers = {'Authorization': f'Bearer {GROK_API_KEY}', 'Content-Type': 'application/json'}
    body = {'input': prompt}
    resp = requests.post('https://api.grok.ai/v1/generate', headers=headers, json=body, timeout=AI_TIMEOUT)
    resp.raise_for_status()
    return resp.json().get('output')


def _merge_responses(responses, report_type):
    # Simple heuristic: prefer the most 'complete' response (longest) and combine notes
    combined = '\n\n'.join(responses)
    # For quarterly report we request a full template conversion
    if report_type == 'quarterly':
        # Ask providers to restructure into executive summary, toc, etc.
        combined = 'Please convert the following into a professional quarterly report with: Executive Summary, Table of Contents, Introduction, Findings, Recommendations, Conclusion.\n\n' + combined
    else:
        combined = 'Please convert the following into a concise weekly IT report. Include highlights, issues, and action items.\n\n' + combined
    # Use OpenAI one last time to tidy and format
    final = _call_openai(combined)
    return final



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

    raw = _tasks_to_text(tasks)

    responses = []
    # Call each provider (safe guards - wrap in try/except so one provider failing doesn't kill the flow)
    try:
        responses.append(_call_openai(raw))
    except Exception as e:
        responses.append(f'OpenAI failed: {e}')

    try:
        responses.append(_call_anthropic(raw))
    except Exception as e:
        responses.append(f'Anthropic failed: {e}')

    try:
        responses.append(_call_grok(raw))
    except Exception as e:
        responses.append(f'Grok failed: {e}')

    # Merge / synthesize
    final_text = _merge_responses(responses, report_type)

    # Convert final_text into a nicely formatted docx
    doc = Document()
    doc.add_heading(f'{report_type.title()} Report', level=1)
    doc.add_paragraph(f'Author: {user.name} ({user.email})')
    doc.add_paragraph(f'Period: {start_date.isoformat()} to {end_date.isoformat()}')
    doc.add_paragraph('')

    # Split into paragraphs reasonably
    for para in final_text.split('\n\n'):
        p = doc.add_paragraph(para.strip())
    for run in p.runs:
        font = run.font
        font.size = Pt(11)

    # Save to bytes
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    filename = f'report_{report_type}_{start_date.isoformat()}_{end_date.isoformat()}.docx'
    return bio.getvalue(), filename


