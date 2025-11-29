from typing import Dict


TEMPLATES = {
    '專業技術': {
        'primary_color': '#0B63B8',
        'secondary_color': '#1F3A93',
        'font': 'Calibri',
        'title_size': 44,
        'subtitle_size': 24,
        'content_size': 20
    },
    '學術研究': {
        'primary_color': '#34568B',
        'secondary_color': '#6B7B8C',
        'font': 'Times New Roman',
        'title_size': 40,
        'subtitle_size': 20,
        'content_size': 18
    },
    '簡約': {
        'primary_color': '#000000',
        'secondary_color': '#666666',
        'font': 'Arial',
        'title_size': 42,
        'subtitle_size': 20,
        'content_size': 18
    }
}


def get_template(theme: str) -> Dict:
    return TEMPLATES.get(theme, TEMPLATES['簡約'])
