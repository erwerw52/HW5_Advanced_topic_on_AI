import os
import json
import logging
from typing import Dict, List

logger = logging.getLogger(__name__)
GEMINI_API_KEY = os.getenv('GEMINI_API_KEY')
GEMINI_AVAILABLE = bool(GEMINI_API_KEY)

if not GEMINI_AVAILABLE:
    logger.warning('GEMINI_API_KEY not set; using heuristic summarizer fallback')

# Constants for auto-pagination
MAX_BULLETS_PER_SLIDE = 5
MAX_CHARS_PER_BULLET = 80


def split_into_bullets(text: str) -> List[str]:
    """Split text into bullet points."""
    import re
    # Split by newlines, bullet markers, or sentences
    lines = re.split(r'\n+|(?:^|\n)\s*[-•*]\s*', text)
    bullets = []
    for line in lines:
        line = line.strip()
        if not line:
            continue
        # If line is too long, split by sentences
        if len(line) > MAX_CHARS_PER_BULLET:
            sentences = re.split(r'(?<=[.!?。！？])\s*', line)
            for s in sentences:
                s = s.strip()
                if s:
                    bullets.append(s[:MAX_CHARS_PER_BULLET])
        else:
            bullets.append(line)
    return bullets


def paginate_bullets(heading: str, bullets: List[str]) -> List[Dict]:
    """Split bullets into multiple slides if exceeding MAX_BULLETS_PER_SLIDE."""
    slides = []
    total = len(bullets)
    if total == 0:
        return [{'type': 'content', 'heading': heading, 'bullets': []}]
    
    page_count = (total + MAX_BULLETS_PER_SLIDE - 1) // MAX_BULLETS_PER_SLIDE
    for i in range(page_count):
        start = i * MAX_BULLETS_PER_SLIDE
        end = start + MAX_BULLETS_PER_SLIDE
        page_bullets = bullets[start:end]
        page_heading = heading if page_count == 1 else f'{heading} ({i+1}/{page_count})'
        slides.append({'type': 'content', 'heading': page_heading, 'bullets': page_bullets})
    return slides


def heuristic_analyze(structured: Dict) -> Dict:
    """Fallback heuristic analysis when Gemini is not available."""
    title = structured.get('title', 'Document')
    sections = structured.get('sections', [])

    # Naive theme detection based on keywords
    text_all = ' '.join([s.get('heading', '') + ' ' + s.get('content', '') for s in sections])
    theme = '簡約'
    if any(k in text_all.lower() for k in ['research', '論文', '實驗', '研究', 'experiment', 'methodology']):
        theme = '學術研究'
    elif any(k in text_all.lower() for k in ['api', 'implementation', 'architecture', 'design', '系統', '技術']):
        theme = '專業技術'

    slides = []
    # Title slide
    slides.append({'type': 'title', 'title': title, 'subtitle': ''})

    # Content slides with auto-pagination
    for sec in sections:
        heading = sec.get('heading', '')
        content = sec.get('content', '')
        if not heading and not content:
            continue
        bullets = split_into_bullets(content)
        paginated = paginate_bullets(heading, bullets)
        slides.extend(paginated)

    # Ending slide
    slides.append({'type': 'ending', 'title': '感謝聆聽，敬請指教'})

    return {'theme': theme, 'slides': slides, 'gemini_used': False}


def gemini_analyze(structured: Dict) -> Dict:
    """Use Gemini API for intelligent content analysis and summarization."""
    import google.generativeai as genai
    
    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel('gemini-2.5-flash')
    
    title = structured.get('title', 'Document')
    sections = structured.get('sections', [])
    
    # Prepare content for Gemini
    content_text = f"標題: {title}\n\n"
    for sec in sections:
        heading = sec.get('heading', '')
        content = sec.get('content', '')
        if heading:
            content_text += f"## {heading}\n"
        if content:
            content_text += f"{content}\n\n"
    
    prompt = f"""請分析以下文檔內容，並生成適合製作 PPT 簡報的結構化內容。

文檔內容：
{content_text}

請以 JSON 格式回傳，格式如下：
{{{{
    "theme": "專業技術" 或 "學術研究" 或 "簡約",
    "slides": [
        {{{{"type": "title", "title": "簡報標題", "subtitle": "副標題（可選）"}}}},
        {{{{"type": "content", "heading": "章節標題", "bullets": ["重點1", "重點2", "重點3"]}}}},
        ...
        {{{{"type": "ending", "title": "結尾語"}}}}
    ]
}}}}

要求：
1. 每張內容頁最多 10 個重點
2. 每個重點不超過 30 個字
3. 將長內容自動分成多張投影片
4. 根據內容特性選擇適合的主題風格
5. 確保結構清晰、重點明確

只回傳 JSON，不要有其他文字。"""

    try:
        response = model.generate_content(prompt)
        response_text = response.text.strip()
        
        # Extract JSON from response (handle markdown code blocks)
        if '```json' in response_text:
            response_text = response_text.split('```json')[1].split('```')[0]
        elif '```' in response_text:
            response_text = response_text.split('```')[1].split('```')[0]
        
        result = json.loads(response_text)
        result['gemini_used'] = True
        return result
        
    except Exception as e:
        logger.error(f'Gemini API error: {e}, falling back to heuristic')
        return heuristic_analyze(structured)


def analyze_document(structured: Dict) -> Dict:
    """Analyze a parsed document and return theme, recommended template, and slides structure."""
    if GEMINI_AVAILABLE:
        return gemini_analyze(structured)
    else:
        return heuristic_analyze(structured)
