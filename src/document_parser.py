import os
import re
from typing import Dict, List

from docx import Document


def parse_txt(path: str) -> Dict:
    with open(path, 'r', encoding='utf-8') as f:
        text = f.read()
    return parse_text_content(text)


def parse_md(path: str) -> Dict:
    with open(path, 'r', encoding='utf-8') as f:
        text = f.read()
    return parse_text_content(text)


def parse_docx(path: str) -> Dict:
    doc = Document(path)
    title = None
    sections: List[Dict] = []
    current = None
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        style = para.style.name.lower()
        # naive heading detection
        if 'title' in style or 'heading' in style:
            level = 1 if 'heading 1' in style or 'title' in style else 2
            if current:
                sections.append(current)
            current = {'level': level, 'heading': text, 'content': ''}
            if title is None:
                title = text
        else:
            if current is None:
                current = {'level': 0, 'heading': '', 'content': text}
            else:
                if current['content']:
                    current['content'] += '\n' + text
                else:
                    current['content'] = text
    if current:
        sections.append(current)
    return {'title': title or os.path.basename(path), 'sections': sections}


def parse_text_content(text: str) -> Dict:
    lines = text.splitlines()
    title = None
    sections: List[Dict] = []
    current = None
    for line in lines:
        line = line.rstrip()
        if not line:
            # blank line => new paragraph
            continue
        # heading detection for markdown
        m = re.match(r'^(#{1,6})\s*(.*)', line)
        if m:
            level = len(m.group(1))
            heading = m.group(2).strip()
            if current:
                sections.append(current)
            current = {'level': level, 'heading': heading, 'content': ''}
            if title is None:
                title = heading
            continue
        # bullet list
        m2 = re.match(r'^[-\*\u2022]\s+(.*)', line)
        if m2:
            bullet = m2.group(1).strip()
            if current is None:
                current = {'level': 0, 'heading': '', 'content': bullet}
            else:
                if current['content']:
                    current['content'] += '\n- ' + bullet
                else:
                    current['content'] = '- ' + bullet
            continue
        # regular paragraph
        if current is None:
            current = {'level': 0, 'heading': '', 'content': line}
        else:
            if current['content']:
                current['content'] += '\n' + line
            else:
                current['content'] = line
    if current:
        sections.append(current)
    return {'title': title or 'Document', 'sections': sections}


def parse_document(path: str) -> Dict:
    ext = os.path.splitext(path)[1].lower()
    if ext == '.txt':
        return parse_txt(path)
    elif ext == '.md':
        return parse_md(path)
    elif ext == '.docx':
        return parse_docx(path)
    else:
        raise ValueError(f'Unsupported file type: {ext}')
