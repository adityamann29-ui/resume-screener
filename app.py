# app.py - main Streamlit application
# run with: streamlit run app.py

import time
import re
from io import BytesIO

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

import pandas as pd
import plotly.express as px
import streamlit as st

from helper import (
    parse_file, light_clean, clean_for_scoring, load_skills,
    load_bias_terms, load_cv_courses, get_all_details, find_skills,
    apply_bias_mask, detect_bias_terms, compute_match_indices, get_skill_gap,
    predict_category, MATCH_INDEX_CONFIG, extract_email, extract_phone,
)

st.set_page_config(page_title='Resume Scanner', layout='wide')

# The dark theme is configured in .streamlit/config.toml as well as here,
# so Streamlit's canvas-based tables use the same readable palette.
st.markdown('''
<style>
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600&family=Inter:wght@400;500;600;700&display=swap');
:root {
    --bg: #06110C; --charcoal: #171717; --cream: #EDE6D6;
    --latte: #B7A89A; --crimson: #A31C1B; --crimson-text: #D6545A;
    --glass-fill: rgba(255,255,255,0.07); --glass-border: rgba(255,255,255,0.14);
}
html, body, [class*="stApp"], p, button, button *, label, input, textarea,
[data-testid="stMarkdownContainer"], [data-testid="stCaptionContainer"] {font-family: 'Inter', sans-serif;}
.stApp {
    background-color: var(--bg);
    background-image: radial-gradient(ellipse 70% 45% at 100% 0%, rgba(163,28,27,0.30), transparent 68%),
                      radial-gradient(ellipse 65% 50% at 0% 100%, rgba(40,90,60,0.35), transparent 70%);
    background-attachment: fixed;
    color: var(--cream);
}
[data-testid="stAppViewContainer"], [data-testid="stHeader"] {background: transparent;}
[data-testid="stDecoration"] {display: none;}
.block-container {max-width: 1240px; padding-top: 2.7rem; padding-bottom: 4rem;}
h1, h2, h3, h4 {font-family: 'Fraunces', Georgia, serif !important; font-weight: 400 !important;
                color: var(--cream) !important; letter-spacing: -0.025em; line-height: 1.12;}
p, label, [data-testid="stMarkdownContainer"], [data-testid="stWidgetLabel"] {color: var(--cream);}
[data-testid="stCaptionContainer"] p, small {color: var(--latte) !important; opacity: 1 !important;}
.hero {text-align: left; padding: 3rem 0 2.4rem;}
.hero-eyebrow, .section-label {color: var(--crimson-text); background: var(--bg);
    display: inline-block; font-family: 'Inter', sans-serif; font-size: 0.76rem;
    font-weight: 600; letter-spacing: 0.10em; text-transform: uppercase;}
.hero-title {font-size: clamp(2.8rem, 5vw, 4.5rem) !important; max-width: 860px;
             margin: 1rem 0 0.8rem !important; line-height: 1.05 !important;}
.hero-copy {color: var(--latte); font-size: 1.05rem; line-height: 1.65; max-width: 640px;}
.tile-number {color: var(--crimson-text); font: 600 20px/1 'Inter', sans-serif;
              letter-spacing: 0.04em; margin-bottom: 0.45rem;}
.tile-copy {color: var(--latte); line-height: 1.55; min-height: 3.2rem;}
/* Only the card's own bordered wrapper gets the glass treatment. */
[data-testid="stVerticalBlockBorderWrapper"]:has(> div > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .tile-number) {
    background: var(--glass-fill); border: 1px solid var(--glass-border); border-radius: 14px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.18); -webkit-backdrop-filter: blur(12px);
    backdrop-filter: blur(12px); transition: transform 150ms ease, border-color 150ms ease;
}
[data-testid="stVerticalBlockBorderWrapper"]:has(> div > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .tile-number):hover {
    border-color: rgba(255,255,255,0.28); transform: translateY(-2px);
}
[data-testid="stVerticalBlockBorderWrapper"]:has(> div > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .featured-marker) {
    background: rgba(163,28,27,0.22); border-color: rgba(214,84,90,0.45);
}
@supports not ((backdrop-filter: blur(12px)) or (-webkit-backdrop-filter: blur(12px))) {
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .tile-number) {background: rgba(255,255,255,0.09);}
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .featured-marker) {background: rgba(163,28,27,0.22);}
}
[data-testid="stVerticalBlockBorderWrapper"]:has(> div > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .tile-number) button {
    background: transparent !important; border: 0 !important; border-bottom: 1px solid var(--crimson-text) !important;
    border-radius: 0; color: var(--cream) !important; padding: 0 0 3px !important;
    width: auto !important; min-height: 1.8rem; box-shadow: none !important;
}
div[data-testid="stMetric"] {background: transparent; padding: 16px 0 8px;
    border: 0; border-top: 1px solid rgba(255,255,255,0.18); border-radius: 0;}
div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] * {
    font-family: 'Fraunces', Georgia, serif !important; color: var(--cream);
}
div[data-testid="stMetricLabel"] p {color: var(--latte); opacity: 1;}
.stButton > button, .stDownloadButton > button {border-radius: 8px; font-weight: 500; min-height: 2.75rem;
                    transition: border-color 150ms ease, transform 150ms ease;}
.stButton > button:hover, .stDownloadButton > button:hover {border-color: var(--crimson-text); transform: translateY(-2px);}
[data-testid="baseButton-primary"] {background: var(--crimson); color: var(--cream); border-color: var(--crimson);}
[data-testid="baseButton-secondary"] {background: var(--glass-fill); color: var(--cream); border-color: var(--glass-border);}
button:focus-visible, a:focus-visible, input:focus-visible, textarea:focus-visible {
    outline: 2px solid var(--crimson-text) !important; outline-offset: 3px;}
[data-testid="stFileUploaderDropzone"] {background: rgba(255,255,255,0.07); border: 1px solid var(--glass-border);}
[data-testid="stFileUploaderDropzone"] svg {color: var(--crimson-text) !important;}
[data-baseweb="input"], [data-baseweb="textarea"] {background: #17221B !important; border-color: rgba(255,255,255,0.25) !important;}
[data-baseweb="input"] > div, [data-baseweb="textarea"] > div {background: #17221B !important;}
input, textarea {color: var(--cream) !important;}
input::placeholder, textarea::placeholder {color: var(--latte) !important; opacity: 1 !important;}
[data-testid="stExpander"] {border-color: var(--glass-border); background: var(--glass-fill);}
section[data-testid="stSidebar"] {background: var(--charcoal); border-right: 1px solid rgba(255,255,255,0.08);}
@media (min-width: 768px) {
    section[data-testid="stSidebar"] {min-width: 220px !important; max-width: 220px !important;}
}
section[data-testid="stSidebar"] [data-testid="stSidebarContent"] {padding-top: 1.2rem;}
section[data-testid="stSidebar"] .stButton > button {
    font-size: 0.875rem; justify-content: flex-start; text-align: left; background: transparent !important;
    border: 1px solid transparent !important; border-left: 2px solid transparent !important;
    color: var(--latte) !important; padding-left: 0.85rem; box-shadow: none !important;
}
section[data-testid="stSidebar"] .stButton > button p {font-size: 0.875rem;}
section[data-testid="stSidebar"] .stButton > button[data-testid="baseButton-primary"] {
    background: rgba(255,255,255,0.07) !important; border-left-color: var(--crimson-text) !important;
    color: var(--cream) !important;
}
section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {color: var(--latte) !important;}
.sidebar-brand {font: 400 1.45rem/1.15 'Fraunces', Georgia, serif; color: var(--cream);}
.sidebar-subtitle {color: var(--latte); font: 400 0.78rem/1.4 'Inter', sans-serif; margin-top: 0.3rem;}
#MainMenu {visibility: hidden;} footer {visibility: hidden;}
[data-testid="stHorizontalBlock"]:has(.dashboard-marker) {flex-wrap: wrap;}
[data-testid="stHorizontalBlock"]:has(.dashboard-marker) > [data-testid="column"] {
    flex: 0 0 calc((100% - 4rem) / 3) !important; width: calc((100% - 4rem) / 3) !important;
}
@media (min-width: 640px) and (max-width: 1100px) {
    [data-testid="stHorizontalBlock"]:has(.dashboard-marker) > [data-testid="column"] {
        flex: 0 0 calc((100% - 2rem) / 2) !important; width: calc((100% - 2rem) / 2) !important;
    }
}
@media (max-width: 639px) {
    [data-testid="stHorizontalBlock"]:has(.dashboard-marker) > [data-testid="column"] {
        flex: 0 0 100% !important; width: 100% !important;
    }
}
@media (prefers-reduced-motion: reduce) { * {transition: none !important;} }
</style>
''', unsafe_allow_html=True)

for key, value in {
    'page': 'home', 'recruiter_result': None, 'personal_result': None,
    'rec_jd': '', 'my_jd': '',
}.items():
    if key not in st.session_state:
        st.session_state[key] = value

# Keep compatible recruiter results and migrate a previous personal match, if any.
existing = st.session_state.recruiter_result
if existing and existing.get('scoring_version') != 'coverage-v1':
    st.session_state.recruiter_result = None
existing = st.session_state.personal_result
if existing and existing.get('review_version') != 'review-v6':
    comparison = existing if existing.get('scoring_version') == 'coverage-v1' else existing.get('comparison')
    st.session_state.personal_result = {
        'review_version': 'review-v6', 'review': None, 'comparison': comparison,
        'time': existing.get('time', ''), 'comparison_error': None,
    }


def goto(page):
    st.session_state.page = page
    st.rerun()


def sidebar(workspace):
    pages = {
        'recruiter': {
            'overview': 'Dashboard', 'screening': 'Resume screening',
            'rankings': 'Candidate rankings', 'skills': 'Skills and analysis',
            'bias': 'Bias detection', 'reports': 'Reports',
        },
        'personal': {
            'dashboard': 'Dashboard', 'analyze': 'Analyze my resume',
            'insights': 'Resume insights', 'improve': 'Skill improvement',
            'learning': 'Learning hub',
        }
    }
    subtitle = 'Recruiter workspace' if workspace == 'recruiter' else 'My resume workspace'
    with st.sidebar:
        st.markdown('<div class="sidebar-brand">Resume Scanner</div>'
                    f'<div class="sidebar-subtitle">{subtitle}</div>', unsafe_allow_html=True)
        st.write('')
        for page, label in pages[workspace].items():
            if st.button(label, key='nav_' + page, use_container_width=True,
                         type='primary' if st.session_state.page == page else 'secondary'):
                goto(page)
        st.markdown('---')
        st.caption('Match indices support review, not automated decisions.' if workspace == 'recruiter'
                   else 'A checklist is a starting point, not a hiring prediction.')
        if st.button('← Home', key='home_btn', use_container_width=True):
            goto('home')


def landing():
    st.markdown('''<div class="hero">
        <div class="hero-eyebrow">RESUME SCANNER  ·  STUDENT PROJECT</div>
        <h1 class="hero-title">Thoughtful screening.<br>Clearer next steps.</h1>
        <div class="hero-copy">A simple space to compare skills, explore gaps and make more informed decisions.</div>
    </div>''', unsafe_allow_html=True)
    st.markdown('<p class="section-label">Choose your workspace</p>', unsafe_allow_html=True)
    left, right = st.columns(2, gap='large')
    with left:
        with st.container(border=True):
            st.markdown('<div class="tile-number">01</div>', unsafe_allow_html=True)
            st.subheader('Recruiter Workspace')
            st.markdown('<div class="tile-copy">Screen multiple resumes, compare match details and review skills.</div>', unsafe_allow_html=True)
            if st.button('Open Recruiter Workspace →', key='open_recruiter', type='primary', use_container_width=True):
                goto('overview')
    with right:
        with st.container(border=True):
            st.markdown('<div class="tile-number">02</div>', unsafe_allow_html=True)
            st.subheader('My Resume Workspace')
            st.markdown('<div class="tile-copy">Review your resume for clarity, useful sections and next steps.</div>', unsafe_allow_html=True)
            if st.button('Open My Resume Workspace →', key='open_personal', use_container_width=True):
                goto('dashboard')
    st.write('')
    st.caption('Made for thoughtful review. Match indices are not hiring decisions or probabilities.')
    with st.expander('About this project'):
        st.write('A student project using document parsing, TF-IDF, MiniLM embeddings and weighted keyword coverage.')
        st.caption('Text-based PDF and DOCX files only. Scanned PDFs and images are not supported.')


def recruiter_nav(page):
    # Recruiter pages use clear buttons rather than a second sidebar menu.
    left, middle, _ = st.columns([1.5, 1.1, 6])
    if left.button('← Dashboard', key='rec_dashboard_' + page, use_container_width=True):
        goto('overview')
    if middle.button('Home', key='rec_home_' + page, use_container_width=True):
        goto('home')


def education_text(items):
    # The keyword extractor can find the same degree more than once.
    # Show each spelling only once without deleting the extracted records.
    shown = []
    seen = set()
    for item in items or []:
        degree = str(item.get('degree', '')).strip()
        key = ''.join(char for char in degree.casefold() if char.isalnum())
        if key and key not in seen:
            seen.add(key)
            shown.append(degree)
    return '; '.join(shown) if shown else 'Not detected'


def mask_keep_lines(text, bias_terms):
    # The coverage rule needs line breaks, so mask each line separately.
    # This can be slower, and masking is still approximate, not anonymity.
    return '\n'.join(apply_bias_mask(line, bias_terms) for line in text.split('\n'))


def process(files, jd, mask):
    # Keep the details and scores together so sorting cannot mix resumes up.
    skills, skills_df = load_skills()
    bias_terms = load_bias_terms() if mask else set()
    parsed = []
    for f in files:
        try:
            raw = parse_file(f.name, f.getvalue())
            if not raw.strip():
                st.warning('Could not extract text from ' + f.name + '. Scanned PDFs need OCR.')
                continue
            if len(raw) > 200000:
                st.warning(f'{f.name} is too long to process (limit: 200,000 characters).')
                continue
            parsed.append({'filename': f.name, 'raw': raw})
        except Exception as e:
            st.warning(f'Could not read {f.name}: {e}')
    if not parsed:
        raise ValueError('No readable resumes were uploaded.')

    score_jd = mask_keep_lines(jd, bias_terms) if mask else jd
    jd_skills = find_skills(light_clean(score_jd), skills)
    if not jd_skills:
        raise ValueError('No skills found in the job description. Add recognizable skills from skills.csv.')

    items, scoring_texts = [], []
    for i, p in enumerate(parsed):
        text = p['raw']
        scoring_text = mask_keep_lines(text, bias_terms) if mask else text
        details = get_all_details(text, skills)
        if mask:
            details['skills'] = find_skills(scoring_text, skills)
        category, confidence = predict_category(scoring_text)
        details['category'] = category
        details['category_conf'] = confidence
        bias_found = detect_bias_terms(text, bias_terms) if mask else []
        label = f'Candidate {i + 1}' if mask else (details['name'] or p['filename'])
        document = f'Document {i + 1}' if mask else p['filename']
        if mask:
            details['name'] = label
            details['email'] = None
            details['phone'] = None
        items.append({'candidate': label, 'document': document, 'details': details,
                      'bias_terms': bias_found})
        scoring_texts.append(scoring_text)

    # This is the new coverage / semantic / TF-IDF scoring path.
    scores = compute_match_indices(scoring_texts, score_jd)
    for item, score in zip(items, scores):
        coverage = score['coverage_details']
        missing = coverage['missing']
        # Group labels can also use the existing YouTube search fallback.
        links = get_skill_gap([], missing, skills_df)['youtube_links']
        item['gap'] = {
            'matched': coverage['matched'],
            'partial': coverage['partial'],
            'missing': missing,
            'extra': sorted(set(coverage['resume_skills']) - set(jd_skills)),
            'total_required': len(coverage['requirements']),
            'total_matched': len(coverage['matched']) + len(coverage['partial']),
            'youtube_links': links,
        }
        item['scores'] = score
    items.sort(key=lambda item: item['scores']['match_index'], reverse=True)
    return {'items': items, 'jd_skills': jd_skills, 'masked': mask,
            'scoring_version': 'coverage-v1',
            'time': time.strftime('%d %b %Y, %H:%M')}


def experience_display(details):
    # Prefer the duration exactly as stated, including months, ranges and +.
    stated = details.get('experience_text')
    if stated:
        return stated
    # Earlier session results might only contain a numeric value.
    years = details.get('experience')
    return f'{years:g} years' if years is not None else 'Not detected'


def results_dataframe(result):
    rows = []
    for i, item in enumerate(result['items'], 1):
        d, score, gap = item['details'], item['scores'], item['gap']
        rows.append({
            'Rank': i, 'Document': item['document'], 'Candidate': item['candidate'],
            'Email': d['email'] or 'Not shown', 'Phone': d['phone'] or 'Not shown',
            'Experience stated': experience_display(d),
            'Experience status': d.get('experience_status') or 'Not detected',
            'Education': education_text(d['education']),
            'Match Index': score['match_index'],
            'Keyword Coverage': score['coverage_score'],
            'Semantic': score['semantic_score'], 'TF-IDF': score['tfidf_score'],
            'Semantic Raw': score['semantic_raw'], 'TF-IDF Raw': score['tfidf_raw'],
            'Scaling Used': score['scaling_used'],
            'Before Experience': score['combined_before_experience'],
            'JD Required Years': score['required_years'] if score['required_years'] is not None else 'Not detected',
            'Total Experience Detected': score['total_experience_detected'] or 'Not detected',
            'Experience Fit': score['experience_fit'] if score['experience_fit'] is not None else 'Not calculated',
            'Experience Factor': round(score['experience_factor'], 4),
            'Experience Note': score['experience_note'],
            'Full Credit': ', '.join(gap['matched']),
            'Half Credit': ', '.join(gap['partial']),
            'Missing Requirements': ', '.join(gap['missing']),
            'Learning Links': '; '.join(gap['youtube_links'].get(skill, '') for skill in gap['missing']),
        })
    return pd.DataFrame(rows)


def safe_csv(df):
    # Spreadsheets may execute text beginning with =, +, -, or @ as formulas.
    clean = df.copy()
    for column in clean.columns:
        clean[column] = clean[column].map(
            lambda x: "'" + x if isinstance(x, str) and x.lstrip().startswith(('=', '+', '-', '@')) else x
        )
    return clean.to_csv(index=False).encode('utf-8-sig')


def jd_input(key):
    # A separate widget key lets us fill the text area from an uploaded JD.
    upload = st.file_uploader('Or upload a job description', type=['txt', 'pdf', 'docx'], key=key + '_file')
    if upload and st.button('Use uploaded job description', key=key + '_use'):
        try:
            if upload.name.lower().endswith('.txt'):
                text = upload.getvalue().decode('utf-8-sig')
            else:
                text = parse_file(upload.name, upload.getvalue())
            if not text.strip():
                st.warning('The file did not contain readable text.')
            else:
                st.session_state[key] = text
        except Exception as e:
            st.error(f'Could not read the job description: {e}')
    return st.text_area('Job description', height=220, key=key,
                        placeholder='Paste the job description here...')


# Section labels are intentionally limited: unknown layouts may need manual review.
PERSONAL_HEADINGS = {
    'contact': ('contact', 'contact information', 'contact details'),
    'education': ('education', 'academic background', 'academic qualifications', 'education and qualifications', 'qualifications', 'academic history'),
    'skills': ('skills', 'technical skills', 'key skills', 'core skills', 'skills and tools', 'technical proficiencies', 'core competencies', 'technologies', 'tools and technologies'),
    'experience': ('work experience', 'experience', 'professional experience', 'teaching experience', 'additional experience', 'relevant experience', 'employment', 'employment history', 'work history', 'internships', 'internship experience', 'professional background'),
    'projects': ('projects', 'project experience', 'academic projects', 'personal projects', 'selected projects', 'key projects', 'project work'),
    'summary': ('summary', 'professional summary', 'profile', 'professional profile', 'objective', 'career objective', 'about me'),
    'other': ('certifications', 'certificates', 'certifications and training', 'certification and training', 'training and certifications', 'training', 'professional development', 'personal details', 'personal information', 'declaration', 'achievements', 'awards', 'publications', 'volunteering', 'leadership', 'activities', 'interests', 'references', 'relevant coursework'),
}


def personal_heading_group(label, lookup):
    name = label.strip().casefold()
    if name in lookup:
        return lookup[name]
    # Short heading-like phrases can include other kinds of experience.
    # Avoid numeric durations and obvious sentences, but layouts can be ambiguous.
    words = name.split()
    if (len(name) <= 55 and 1 <= len(words) <= 5
            and re.search(r'\bexperience\b', name)
            and not re.search(r'[0-9.!?;()]', name)
            and not re.match(r'^(?:i|my|we|our)\b', name)
            and not re.search(r'\b(?:one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|several|many)\s+(?:years?|months?)\b', name)):
        return 'experience'
    return None


def split_personal_sections(text):
    # Also recognize "Heading: text" and simple DOCX table rows ("Heading | text").
    lookup = {name: group for group, names in PERSONAL_HEADINGS.items() for name in names}
    sections = {group: [] for group in PERSONAL_HEADINGS}
    sections['contact'] = []  # The text before the first heading is the contact/header area.
    headings = []
    blocks = []
    current = 'contact'
    current_block = None
    for raw_line in text.splitlines():
        line = re.sub(r'\s+', ' ', raw_line).strip()
        if not line:
            continue
        line = re.sub(r'^[•●▪◆\-*]+\s*', '', line).strip()
        if not line:
            continue
        label, content = line, ''
        inline = re.match(r'^([^:|]{1,55})\s*[:|]\s*(.*)$', line)
        if inline and personal_heading_group(inline.group(1), lookup):
            label, content = inline.group(1).strip(), inline.group(2).strip()
        else:
            label = line.rstrip(':–—- ').strip()
        group = personal_heading_group(label, lookup)
        if group:
            current = group
            headings.append(group)
            current_block = {'label': label.casefold(), 'group': group, 'lines': []}
            blocks.append(current_block)
            if content:
                sections[current].append(content)
                current_block['lines'].append(content)
        else:
            sections[current].append(line)
            if current_block is not None:
                current_block['lines'].append(line)
    return sections, headings, blocks


# Personal reviews read DOCX paragraphs and tables in the order they appear.
# The recruiter still uses the shared parse_file function unchanged.
def personal_text(filename, file_bytes):
    if not filename.lower().endswith('.docx'):
        return parse_file(filename, file_bytes)
    doc = Document(BytesIO(file_bytes))
    parts = []
    for element in doc.element.body.iterchildren():
        if isinstance(element, CT_P):
            text = Paragraph(element, doc).text
            if text.strip():
                parts.append(text)
        elif isinstance(element, CT_Tbl):
            table = Table(element, doc)
            for row in table.rows:
                # A merged cell can appear more than once in python-docx rows.
                cells, seen = [], set()
                for cell in row.cells:
                    if cell._tc in seen:
                        continue
                    seen.add(cell._tc)
                    if cell.text.strip():
                        cells.append(cell.text.strip())
                if cells:
                    parts.append(' | '.join(cells))
    return '\n'.join(parts)


# Explicit forms avoid using a broad suffix rule that can match unrelated words.
PERSONAL_ACTION_FORMS = {
    'build builds building built', 'create creates creating created',
    'develop develops developing developed', 'design designs designing designed',
    'implement implements implementing implemented', 'improve improves improving improved',
    'lead leads leading led', 'manage manages managing managed',
    'organize organizes organizing organized', 'organise organises organising organised',
    'analyze analyzes analyzing analyzed', 'analyse analyses analysing analysed',
    'deliver delivers delivering delivered', 'test tests testing tested',
    'automate automates automating automated', 'collaborate collaborates collaborating collaborated',
    'launch launches launching launched', 'reduce reduces reducing reduced',
    'increase increases increasing increased', 'support supports supporting supported',
    'research researches researching researched', 'contribute contributes contributing contributed',
    'deploy deploys deploying deployed', 'write writes writing wrote written',
    'coordinate coordinates coordinating coordinated', 'establish establishes establishing established',
    'produce produces producing produced', 'publish publishes publishing published',
    'maintain maintains maintaining maintained', 'optimize optimizes optimizing optimized',
    'optimise optimises optimising optimised', 'make makes making made',
    'prepare prepares preparing prepared', 'present presents presenting presented',
    'clean cleans cleaning cleaned', 'fix fixes fixing fixed', 'use uses using used',
    'teach teaches teaching taught', 'explain explains explaining explained',
    'help helps helping helped', 'assist assists assisting assisted',
    'conduct conducts conducting conducted', 'supervise supervises supervising supervised',
    'handle handles handling handled', 'work works working worked',
    'speak speaks speaking spoke spoken',
}
PERSONAL_ACTION_WORDS = {word for forms in PERSONAL_ACTION_FORMS for word in forms.split()}
PERSONAL_SKILL_SHARE_MIN = 0.30  # Preliminary heuristic, not an HR standard.
PERSONAL_SKILL_CHECK_MIN = 6  # Skip the proportion when too few listed skills are recognized.
DETAILS_ON_REQUEST = re.compile(r'\b(?:details|information)\s+(?:are\s+|is\s+)?(?:available\s+)?(?:on|upon)\s+request\b', re.I)


def review_resume(upload):
    """Experimental, English-oriented text checks, not an HR readiness prediction."""
    raw = personal_text(upload.name, upload.getvalue())
    if not raw.strip():
        return {'filename': upload.name, 'details': {'name': None, 'email': None, 'phone': None,
                'skills': [], 'education': [], 'experience': None, 'experience_text': None,
                'experience_status': 'Not detected'}, 'word_count': 0,
                'band': 'Could not check reliably', 'areas': [], 'adjustments': [],
                'reliability': 'No text could be extracted. Scanned or image-only PDFs are not supported.'}
    if len(raw) > 200000:
        raise ValueError('The resume is too long to process (limit: 200,000 characters).')
    text = light_clean(raw)
    words = re.findall(r"[^\W\d_]+(?:[’'-][^\W\d_]+)*", text, re.UNICODE)
    sections, headings, blocks = split_personal_sections(text)
    # Fallback experience headings may organize text, but are not strong enough
    # evidence to make an uncertain extraction reliable.
    known_labels = {name for names in PERSONAL_HEADINGS.values() for name in names}
    known_heading_groups = {block['group'] for block in blocks if block['label'] in known_labels}
    # Contact text can be either before the first heading or under a Contact heading.
    section_text = {key: '\n'.join(lines) for key, lines in sections.items()}
    nonspace = [c for c in text if not c.isspace()]
    odd_chars = sum(c == '\ufffd' or (ord(c) < 32 and not c.isspace()) for c in nonspace)
    letter_words = [w.casefold() for w in words if len(w) >= 3]
    vowel_words = sum(bool(re.search(r'[aeiou]', w)) for w in letter_words)
    reliability = None
    if len(words) < 50:
        reliability = 'Fewer than 50 words were extracted. Check that the file is complete and contains selectable text.'
    elif nonspace and odd_chars / len(nonspace) > 0.03:
        reliability = 'The extracted text contains many damaged characters.'
    elif len(letter_words) < 15 or vowel_words / max(len(letter_words), 1) < 0.35:
        reliability = 'The extracted text may be garbled or use a language this English-oriented checker cannot assess.'
    elif len(known_heading_groups) < 2:
        reliability = 'The app could not identify enough section boundaries reliably in this layout.'

    skills, _ = load_skills()
    details = get_all_details(raw, skills)
    # Do not turn an unreliable extraction into a confident score or band.
    if reliability:
        return {'filename': upload.name, 'details': details, 'word_count': len(words),
                'band': 'Could not check reliably', 'areas': [], 'adjustments': [],
                'reliability': reliability}

    def content_words(value):
        return re.findall(r"[^\W\d_]+(?:[’'-][^\W\d_]+)*", value, re.UNICODE)

    def meaningful(value):
        # Placeholder-only content is not treated as a filled section.
        cleaned = re.sub(r'\b(?:lorem ipsum|placeholder|insert here|your name|your text|n/?a|tbd)\b', '', value, flags=re.I)
        cleaned = DETAILS_ON_REQUEST.sub('', cleaned)
        if DETAILS_ON_REQUEST.search(value) and len(content_words(cleaned)) < 7:
            return False
        return len(content_words(cleaned)) >= 3

    contact_text = section_text['contact']
    education = section_text['education']
    skill_text = section_text['skills']
    work = '\n'.join((section_text['experience'], section_text['projects'])).strip()
    contact_found = bool(extract_email(contact_text) or extract_phone(contact_text))
    edu_words = content_words(education)
    education_has_body = meaningful(education)
    edu_institution = bool(re.search(r'\b(university|college|institute|school|academy|iit|nit|iiit)\b', education, re.I))
    skills_has_body = meaningful(skill_text) or bool(find_skills(skill_text, skills))
    skill_hits = find_skills(skill_text, skills)

    # Look for skill use only in substantive project/work descriptions, not a
    # bare skills list or a line that merely offers details on request.
    work_lines = sections['experience'] + sections['projects']
    thin_lines = []
    usable_lines = []
    for line in work_lines:
        cleaned = DETAILS_ON_REQUEST.sub('', line).strip(' ,;:.|-')
        if DETAILS_ON_REQUEST.search(line) and len(content_words(cleaned)) < 7:
            thin_lines.append(line)
        elif meaningful(cleaned):
            usable_lines.append(cleaned)
    work_words = content_words('\n'.join(usable_lines))
    # Identical descriptions count once, regardless of repetition.
    distinct_work = list(dict.fromkeys(re.sub(r'\s+', ' ', x).casefold() for x in usable_lines))
    action_lines = [line for line in distinct_work if len(content_words(line)) >= 7
                    and any(word.casefold() in PERSONAL_ACTION_WORDS for word in content_words(line))]
    context_lines = [line for line in action_lines if len(content_words(line)) >= 9 and re.search(r'\b(using|with|for|to|through|by|across|from|based on)\b', line)]
    output_lines = [line for line in action_lines if len(content_words(line)) >= 9 and re.search(r'\b(application|app|website|dashboard|report|tool|system|service|model|analysis|dataset|data|research|design|process|pipeline|product|platform|feature|workflow|documentation|campaign|prototype|solution|program|software|automation|testing)\b', line)]
    description_lines = [line for line in distinct_work if len(content_words(line)) >= 7
                         and (any(word.casefold() in PERSONAL_ACTION_WORDS for word in content_words(line))
                              or re.search(r'\b(responsible for|using|with|for|through|by|across|based on)\b', line))]
    # A line listing many skills is not reliable evidence that each was used.
    # Count distinct recognized skills per line, not repeated mentions.
    skill_in_descriptions = set()
    dense_skill_lines = []
    enough_listed_skills = len(skill_hits) >= PERSONAL_SKILL_CHECK_MIN
    if enough_listed_skills:
        for line in description_lines:
            line_skills = find_skills(line, skills)
            if len(line_skills) >= 6:
                dense_skill_lines.append(line)
            else:
                skill_in_descriptions.update(line_skills)
    applied_skills = sorted(set(skill_hits) & skill_in_descriptions)
    skill_share = len(applied_skills) / len(skill_hits) if enough_listed_skills else None
    low_skill_share = skill_share is not None and skill_share < PERSONAL_SKILL_SHARE_MIN

    all_content = [line for group in sections.values() for line in group if meaningful(line)]
    normalized = [re.sub(r'\s+', ' ', line).strip().casefold() for line in all_content]
    unique_content = list(dict.fromkeys(normalized))
    duplicate_lines = len(normalized) - len(unique_content)
    duplicate_ratio = duplicate_lines / len(normalized) if normalized else 0
    unique_words = sum(len(content_words(line)) for line in unique_content)
    long_lines = sum(len(content_words(line)) > 80 for line in all_content)
    placeholders = bool(re.search(r'\b(lorem ipsum|placeholder|insert here|your name|your text|tbd|can\s+be\s+added|to\s+be\s+added)\b', text, re.I))
    content_groups = sum((skills_has_body if group == 'skills' else meaningful(section_text[group]))
                         for group in ('education', 'skills', 'experience', 'projects'))
    filled_heading_groups = sum(meaningful(section_text[group]) or (group == 'skills' and skills_has_body)
                                for group in set(headings))
    empty_groups = []
    # These optional sections can legitimately contain a short list or one item.
    short_optional = {'personal details', 'personal information',
                      'declaration', 'certifications and training', 'certification and training',
                      'training and certifications', 'training', 'professional development'}
    for block in blocks:
        body = '\n'.join(block['lines'])
        block_has_skills = block['group'] == 'skills' and bool(find_skills(body, skills))
        cleaned_optional = re.sub(r'\b(?:lorem ipsum|placeholder|insert here|your name|your text|n/?a|tbd)\b', '', body, flags=re.I)
        cleaned_optional = DETAILS_ON_REQUEST.sub('', cleaned_optional)
        optional_has_content = block['label'] in short_optional and bool(content_words(cleaned_optional))
        if not meaningful(body) and not block_has_skills and not optional_has_content:
            empty_groups.append(block['label'])

    # Each check describes the observable signal, not a claim about quality or truth.
    def check(title, points, maximum, reason, advice):
        return {'title': title, 'points': points, 'maximum': maximum,
                'reason': reason, 'advice': advice}

    section_checks = [
        check('Contact content', 6 if contact_found else 0, 6,
              'An email or phone was recognized in the contact/header area.' if contact_found else 'No email or phone was recognized in the contact/header area.',
              'Check that your preferred contact method is clear and correct.'),
        check('Education content', 8 if education_has_body and edu_institution else (4 if education_has_body else 0), 8,
              'Education text and an institution indicator were recognized.' if education_has_body and edu_institution else ('Education text was recognized, but an institution indicator was not.' if education_has_body else 'No substantive education text was recognized under an Education heading.'),
              'Include relevant qualifications and institutions; verify the parser did not miss them.'),
        check('Skills content', (min(7, 3 + len(skill_hits)) if skill_hits else (3 if skills_has_body else 0)), 7,
              f'{len(skill_hits)} listed skills were recognized in the Skills section.' if skills_has_body else 'No substantive Skills section content was recognized.',
              'List relevant skills you can explain; this app does not recognize every skill.'),
        check('Project or experience content', 9 if len(work_words) >= 20 else (4 if len(work_words) >= 7 else 0), 9,
              f'{len(work_words)} words were recognized in Projects and Experience sections.',
              'Describe relevant projects, internships or work, including what you did.'),
    ]
    evidence_checks = [
        check('Described actions', 15 if len(action_lines) >= 2 else (8 if action_lines else 0), 15,
              f'{len(action_lines)} distinct work or project descriptions contain a recognized action verb and at least seven words.',
              'Explain what you did in distinct project or work descriptions; the verb list is limited.'),
        check('Context in descriptions', 10 if len(context_lines) >= 2 else (5 if context_lines else 0), 10,
              f'{len(context_lines)} descriptions contain a recognized context signal; this is only a rough text check.',
              'Add relevant context, such as the purpose, approach or tools used.'),
        check('Output described', 5 if output_lines else 0, 5,
              'A description contains a recognized output or work-product term.' if output_lines else 'No recognized work-product term was found in a descriptive work or project line.',
              'Where relevant, describe what the work produced. Numbers are not required.'),
    ]
    clarity_checks = [
        check('Content length', 6 if 150 <= unique_words <= 1200 else (3 if 75 <= unique_words < 150 else 0), 6,
              f'{unique_words} words remain after exact repeated content lines are counted once.',
              'Use enough relevant detail without unnecessary repetition; length alone does not determine quality.'),
        check('Repeated content', 6 if duplicate_lines == 0 else (3 if duplicate_ratio < 0.2 else 0), 6,
              f'{duplicate_lines} repeated content line(s) detected.',
              'Remove repeated descriptions and tailor each entry to the work actually done.'),
        check('Dense extracted lines', 4 if not long_lines else 0, 4,
              f'{long_lines} extracted content line(s) exceed 80 words.',
              'Review long blocks for readability; PDF extraction can merge lines.'),
        check('Template placeholders', 0 if placeholders else 4, 4,
              'Possible placeholder text was detected.' if placeholders else 'No common placeholder text was detected.',
              'Replace template placeholders with truthful information.'),
    ]
    structure_checks = [
        check('Substantive core sections', 8 if content_groups >= 3 else (4 if content_groups >= 2 else 0), 8,
              f'{content_groups} of Education, Skills, Experience and Projects have recognizable content.',
              'Organize relevant content under clear headings; Experience and Projects are alternatives.'),
        check('Section separation', 6 if filled_heading_groups >= 3 and content_groups >= 2 else (3 if content_groups >= 1 else 0), 6,
              f'{filled_heading_groups} recognizable section types have content, with {content_groups} substantive core section(s).',
              'Use clear headings and check that content is extracted in the intended section.'),
        check('Filled section bodies', 6 if not empty_groups and content_groups >= 2 else (3 if content_groups >= 1 and len(empty_groups) < 2 else 0), 6,
              'No empty recognized sections were found.' if not empty_groups else 'Empty or very thin recognized sections: ' + ', '.join(empty_groups) + '.',
              'Fill or remove empty sections; heading recognition can miss unusual formats.'),
    ]
    areas = [
        {'title': 'Essential information and sections', 'max': 30, 'checks': section_checks},
        {'title': 'Project or experience descriptions', 'max': 30, 'checks': evidence_checks},
        {'title': 'Clarity and conciseness', 'max': 20, 'checks': clarity_checks},
        {'title': 'Structure and extractability', 'max': 20, 'checks': structure_checks},
    ]
    for area in areas:
        area['points'] = sum(item['points'] for item in area['checks'])

    adjustments = []
    if placeholders:
        adjustments.append({'points': 0, 'reason': 'Possible template text was detected. Review it; ordinary descriptions of future work can also match.'})
    if unique_words < 150:
        adjustments.append({'points': -8, 'reason': 'Fewer than 150 distinct-content words: the text may be too thin for a full review.'})
    if duplicate_lines >= 2:
        adjustments.append({'points': -min(15, 2 * duplicate_lines), 'reason': f'{duplicate_lines} repeated content lines may inflate the resume.'})
    if empty_groups:
        adjustments.append({'points': -min(12, 4 * len(empty_groups)), 'reason': 'Recognized empty or very thin section(s): ' + ', '.join(empty_groups) + '.'})
    if thin_lines:
        adjustments.append({'points': 0, 'reason': f'{len(thin_lines)} project or experience line(s) offer details on request but contain too little other detail to count as descriptions.'})
    if dense_skill_lines:
        adjustments.append({'points': 0, 'reason': f'{len(dense_skill_lines)} description line(s) with six or more recognized skills were excluded from the skill-use check. This is a rough filter; document extraction can combine valid lines.'})
    if low_skill_share:
        adjustments.append({'points': -5, 'reason': f'Only {len(applied_skills)} of {len(skill_hits)} recognized listed skills appear in recognized substantive project or experience descriptions; this is below the provisional 30% check. The app can miss synonyms and skills not in its list.'})
    total = max(0, sum(a['points'] for a in areas) + sum(a['points'] for a in adjustments))
    cap_reasons = []
    if len(work_words) < 7:
        total = min(total, 59)
        cap_reasons.append('No substantive project or experience content was recognized.')
    if len(empty_groups) >= 2:
        total = min(total, 59)
        cap_reasons.append('Two or more recognized sections appear empty or very thin.')
    elif empty_groups:
        total = min(total, 79)
        cap_reasons.append('A recognized section appears empty or very thin.')
    if duplicate_lines >= 2 and duplicate_ratio >= 0.2:
        total = min(total, 59)
        cap_reasons.append('Repeated content accounts for a substantial share of extracted lines.')
    if unique_words < 150:
        total = min(total, 69)
        cap_reasons.append('The distinct extracted content is under 150 words.')
    if low_skill_share:
        total = min(total, 79)
        cap_reasons.append('Less than 30% of recognized listed skills were found in recognized substantive descriptions; wording or extraction may have been missed.')
    every_area_passes = all(area['points'] / area['max'] >= 0.6 for area in areas)
    if total >= 80 and every_area_passes:
        band = 'Strong'
    elif total >= 60:
        band = 'Good'
    else:
        band = 'Needs work'
    if total >= 80 and not every_area_passes:
        cap_reasons.append('Strong requires at least 60% in each of the four areas.')
    if enough_listed_skills:
        observation = (f'{len(applied_skills)} of {len(skill_hits)} recognized listed skills appear in recognized substantive '
                       'project or experience descriptions. The provisional check is 30%; other wording may be missed.')
    else:
        observation = (f'The skill-use check was skipped: {len(skill_hits)} distinct listed skill(s) were recognized '
                       f'in the Skills section; at least {PERSONAL_SKILL_CHECK_MIN} are needed. '
                       'The app may not recognize relevant skills or different wording.')
    return {'filename': upload.name, 'details': details, 'word_count': len(words),
            'band': band, 'areas': areas, 'adjustments': adjustments, 'observations': [observation],
            'cap_reasons': cap_reasons, 'reliability': None}


def personal_analyze_page():
    st.header('Analyze My Resume')
    st.caption('Upload your resume for a text-based review. No job description is required.')
    upload = st.file_uploader('Your resume', type=['pdf', 'docx'], key='personal_files')
    st.caption('Text-based PDF or DOCX only; scanned or image-only files cannot be read.')
    with st.expander('Optional: compare with a job description'):
        compare = st.checkbox('Compare this resume with a job description', key='personal_compare')
        jd = jd_input('my_jd') if compare else ''
        st.caption('This optional comparison is separate from the personal resume review.')
    if st.button('Analyze My Resume', type='primary', use_container_width=True):
        if not upload:
            st.error('Please upload your resume.')
        elif compare and not jd.strip():
            st.error('Add a job description or uncheck the optional comparison.')
        elif compare and len(jd) > 50000:
            st.error('Please use a job description under 50,000 characters.')
        else:
            try:
                with st.spinner('Reading your resume...'):
                    review = review_resume(upload)
            except Exception as e:
                st.error(f'Analysis could not finish: {e}')
                st.caption('Previous results, if any, are still available.')
            else:
                comparison, comparison_error = None, None
                if compare and review['reliability']:
                    comparison_error = 'Comparison was skipped because the resume text could not be checked reliably.'
                elif compare:
                    try:
                        with st.spinner('Comparing against the job description...'):
                            comparison = process([upload], jd, False)
                    except Exception as e:
                        comparison_error = str(e)
                st.session_state.personal_result = {
                    'review_version': 'review-v6', 'review': review, 'comparison': comparison,
                    'comparison_error': comparison_error, 'time': time.strftime('%d %b %Y, %H:%M'),
                }
                goto('insights')


def analyze_page(personal=False):
    if personal:
        personal_analyze_page()
        return
    st.header('Analyze My Resume' if personal else 'Resume Screening')
    st.caption('Upload PDF or DOCX resumes and provide a job description.')
    left, right = st.columns(2)
    with left:
        jd = jd_input('my_jd' if personal else 'rec_jd')
    with right:
        upload = st.file_uploader('Resume files', type=['pdf', 'docx'],
                                  accept_multiple_files=not personal,
                                  key='personal_files' if personal else 'recruiter_files')
        files = ([upload] if upload else []) if personal else upload
        if files:
            st.write(f'{len(files)} file(s) selected')
            for f in files:
                st.caption(f.name)
        mask = False
        if not personal:
            mask = st.checkbox('Mask identity hints before scoring',
                               help='Attempts to remove names, gender words and contact information. It may also remove skill terms. This does not guarantee fairness or anonymity.')
    if st.button('Analyze My Resume' if personal else 'Analyze Resumes', type='primary', use_container_width=True):
        if not files or not jd.strip():
            st.error('Please upload a resume and provide a job description.')
        elif len(files) > 20 or len(jd) > 50000:
            st.error('Please use at most 20 resumes and a job description under 50,000 characters.')
        else:
            try:
                with st.spinner('Reading files and loading models if needed. The first run may take longer...'):
                    result = process(files, jd, mask)
            except Exception as e:
                st.error(f'Analysis could not finish: {e}')
                st.caption('Check the app logs for technical details. Previous results, if any, are still available.')
            else:
                # st.rerun() uses an internal exception; don't catch it as an error.
                st.session_state['personal_result' if personal else 'recruiter_result'] = result
                goto('insights' if personal else 'rankings')


def need_result(key):
    result = st.session_state[key]
    if not result:
        st.info('No analysis yet. Upload your resume to get started.' if key == 'personal_result'
                else 'No analysis yet. Upload resumes and a job description to get started.')
    return result


def score_metrics(item):
    score = item['scores']
    c1, c2, c3, c4 = st.columns(4)
    c1.metric('Match index', f"{score['match_index']:.1f}")
    c2.metric('Keyword coverage', f"{score['coverage_score']:.1f}%")
    semantic_label = 'Semantic (scaled)' if score['scaling_used'] else 'Semantic (raw)'
    tfidf_label = 'TF-IDF (scaled)' if score['scaling_used'] else 'TF-IDF (raw)'
    c3.metric(semantic_label, f"{score['semantic_score']:.1f}%")
    c4.metric(tfidf_label, f"{score['tfidf_score']:.1f}%")
    weights = MATCH_INDEX_CONFIG
    st.caption(f"Index before experience = {weights['coverage_weight'] * 100:g}% coverage + "
               f"{weights['semantic_weight'] * 100:g}% semantic + {weights['tfidf_weight'] * 100:g}% TF-IDF. "
               'Final index = that result × experience factor. It is not a probability.')
    st.write(f"**Before experience:** {score['combined_before_experience']:.1f}  |  "
             f"**Experience factor:** {score['experience_factor']:.3f}")
    stated = score['total_experience_detected'] or 'Not detected'
    required = score['required_years']
    requirement = f'{required:g} years' if required is not None else 'Not detected / not applied'
    st.caption(f"JD minimum: {requirement}. Total experience detected: {stated}. "
               f"{score['experience_note']}. A detected duration is not verified total or relevant experience.")


def show_gap(gap):
    a, b, c, d = st.columns(4)
    a.write('**Full credit**')
    a.write(', '.join(gap['matched']) or 'None')
    b.write('**Half credit**')
    b.write(', '.join(gap['partial']) or 'None')
    c.write('**Not detected**')
    c.write(', '.join(gap['missing']) or 'None')
    d.write('**Additional skills**')
    d.write(', '.join(gap['extra']) or 'None detected')
    st.caption('Credit is based on a simple line-format rule, not verified proficiency. '
               'A missing skill may still be present in a different wording.')


def overview():
    if st.button('← Home', key='rec_overview_home'):
        goto('home')
    st.markdown('<p class="section-label">Recruiter workspace</p>', unsafe_allow_html=True)
    st.title('Your screening dashboard')
    st.write('Choose a task to get started. Results stay in this session.')
    result = st.session_state.recruiter_result
    if result:
        df = results_dataframe(result)
        resume_count = len(df)
        top_index = f"{df['Match Index'].max():.1f}"
        gap_count = sum(len(item['gap']['missing']) for item in result['items'])
    else:
        resume_count, top_index, gap_count = 0, '—', '—'
    a, b, c = st.columns(3)
    a.metric('Resumes screened', resume_count)
    b.metric('Top match index', top_index)
    c.metric('Skill gaps (mentions)', gap_count)

    cards = [
        ('01', 'Resume screening', 'Upload resumes and compare them with a job description.', 'screening', 'Start screening'),
        ('02', 'Candidate rankings', 'Review match indices, document details and full rankings.', 'rankings', 'View rankings'),
        ('03', 'Skills and analysis', 'Explore detected requirements and candidate skill gaps.', 'skills', 'Explore skills'),
        ('04', 'Bias detection', 'Review the experimental text-based masking and term checks.', 'bias', 'View bias checks'),
        ('05', 'Reports and analytics', 'See score charts and download the screening report.', 'reports', 'View reports'),
    ]
    columns = st.columns(len(cards), gap='medium')
    for column, card in zip(columns, cards):
        number, title, description, page, button = card
        with column:
            with st.container(border=True):
                marker = 'tile-number dashboard-marker' + (' featured-marker' if page == 'screening' else '')
                st.markdown(f'<div class="{marker}">{number}</div>', unsafe_allow_html=True)
                st.subheader(title)
                st.markdown(f'<div class="tile-copy">{description}</div>', unsafe_allow_html=True)
                if st.button(button + ' →', key='rec_tile_' + page, use_container_width=True):
                    goto(page)

    if result:
        with st.expander('Latest screening summary'):
            st.caption('Last screening in this session: ' + result['time'])
            left, right = st.columns(2)
            with left:
                st.subheader('Top match scores')
                st.dataframe(df[['Rank', 'Candidate', 'Match Index']].head(5),
                             hide_index=True, use_container_width=True)
            with right:
                st.subheader('Match index distribution')
                fig = px.histogram(df, x='Match Index', nbins=8, range_x=[0, 100],
                                   color_discrete_sequence=['#D6545A'])
                st.plotly_chart(style_chart(fig), use_container_width=True)
    st.caption('Match indices are uncalibrated aids for human review, not hiring decisions.')


def rankings():
    st.header('Candidate Rankings')
    result = need_result('recruiter_result')
    if not result:
        return
    if result['masked']:
        st.info('Identity hints were masked for scoring. Masking may also remove skill terms and change scores; it is not a fairness or anonymity guarantee.')
    df = results_dataframe(result)
    cols = ['Rank', 'Document', 'Candidate', 'Email', 'Phone', 'Experience stated', 'Experience status',
            'Education', 'Match Index']
    st.dataframe(df[cols], hide_index=True, use_container_width=True)
    st.caption('Experience status is based on one detected duration, not verified total experience. '
               'Fresher means less than 12 months stated. Review ambiguous or missing information.')
    st.download_button('Download full rankings CSV', safe_csv(df), 'candidate_rankings.csv', 'text/csv')
    st.subheader('Candidate details')
    for rank, item in enumerate(result['items'], 1):
        with st.expander(f"#{rank}  {item['candidate']} — {item['scores']['match_index']:.1f}"):
            score_metrics(item)
            show_gap(item['gap'])


def skill_analysis():
    st.header('Skills & Analysis')
    result = need_result('recruiter_result')
    if not result:
        return
    requirements = result['items'][0]['scores']['coverage_details']['requirements']
    labels = [r['label'] + (' (preferred)' if r['preferred'] else '') for r in requirements]
    st.write('**Detected JD requirements:** ' + ', '.join(labels))
    for rank, item in enumerate(result['items'], 1):
        gap = item['gap']
        with st.expander(f"#{rank} {item['candidate']} — {gap['total_matched']}/{gap['total_required']} requirements detected"):
            show_gap(gap)
            if gap['missing']:
                st.write('**Learning resources**')
                for skill in gap['missing']:
                    st.link_button(f'Search YouTube: {skill}', gap['youtube_links'][skill])


def bias_page():
    st.header('Bias Detection')
    st.caption('This is a simple text-based check, not a fairness audit. A detected term is not necessarily discriminatory.')
    result = need_result('recruiter_result')
    if not result:
        return
    if not result['masked']:
        st.warning('The last analysis did not use masking. Re-run screening with the checkbox selected.')
        return
    for item in result['items']:
        with st.expander(item['candidate']):
            st.write('Detected terms: ' + (', '.join(item['bias_terms']) or 'None detected'))
    st.caption('Names and terms may be missed; other identity clues may remain in the text.')


def style_chart(fig):
    # Keep Plotly's chart surfaces and labels consistent with the app theme.
    fig.update_layout(paper_bgcolor='#06110C', plot_bgcolor='#06110C',
                      font=dict(family='Inter, sans-serif', color='#EDE6D6'),
                      legend=dict(font=dict(color='#EDE6D6')))
    fig.update_xaxes(gridcolor='rgba(237,230,214,0.15)', zerolinecolor='rgba(237,230,214,0.2)')
    fig.update_yaxes(gridcolor='rgba(237,230,214,0.15)', zerolinecolor='rgba(237,230,214,0.2)')
    return fig


def reports():
    st.header('Reports & Analytics')
    result = need_result('recruiter_result')
    if not result:
        return
    df = results_dataframe(result)
    a, b, c = st.columns(3)
    a.metric('Candidates analyzed', len(df))
    b.metric('Average match index', f"{df['Match Index'].mean():.1f}")
    c.metric('Highest match index', f"{df['Match Index'].max():.1f}")
    # Rank makes labels unique even if two candidates have the same name.
    df['Chart label'] = df.apply(lambda row: f"#{row['Rank']} {row['Candidate']}", axis=1)
    fig = px.bar(df.sort_values('Match Index'), x='Match Index', y='Chart label',
                 orientation='h', range_x=[0, 100], color_discrete_sequence=['#D6545A'])
    st.plotly_chart(style_chart(fig), use_container_width=True)
    # These bands are only for viewing the distribution, not decision thresholds.
    bands = pd.cut(df['Match Index'], bins=[-0.01, 40, 70, 100],
                   labels=['0–40', '>40–70', '>70–100'], include_lowest=True)
    counts = bands.value_counts().sort_index().reset_index()
    counts.columns = ['Score band', 'Candidates']
    left, right = st.columns(2)
    with left:
        st.subheader('Score bands')
        st.plotly_chart(style_chart(px.bar(counts, x='Score band', y='Candidates',
                        color_discrete_sequence=['#D6545A'])), use_container_width=True)
    with right:
        st.subheader('Most detected candidate skills')
        all_skills = [skill for item in result['items'] for skill in item['details']['skills']]
        if all_skills:
            skill_counts = pd.Series(all_skills).value_counts().head(10).reset_index()
            skill_counts.columns = ['Skill', 'Candidates']
            st.plotly_chart(style_chart(px.bar(skill_counts, x='Candidates', y='Skill', orientation='h',
                            color_discrete_sequence=['#B7A89A'])), use_container_width=True)
        else:
            st.caption('No skills detected.')
    st.caption('Bands are illustrative. Scores are not validated hiring recommendations.')
    st.download_button('Download report CSV', safe_csv(df.drop(columns=['Chart label'])), 'report.csv', 'text/csv')


def personal_dashboard():
    st.header('My Dashboard')
    result = st.session_state.personal_result
    review = result.get('review') if result else None
    if not review:
        st.write('Upload your resume for a text-based review and practical suggestions. No job description is needed.')
    else:
        st.metric('Personal resume review', review['band'])
        st.caption('Latest review in this session: ' + result['time'])
        st.caption('Checks basic structure and evidence in extracted text; bands do not rank overall resume quality or predict hiring.')
    a, b = st.columns(2)
    if a.button('Analyze My Resume', type='primary', use_container_width=True):
        goto('analyze')
    if b.button('Browse Learning Hub', use_container_width=True):
        goto('learning')


def insights():
    st.header('Resume Insights')
    result = need_result('personal_result')
    if not result:
        return
    review = result.get('review')
    if review:
        st.metric('Personal resume review', review['band'])
        st.caption('These experimental, English-oriented checks cover basic structure and evidence only; the bands do not rank overall resume quality and are not an HR or ATS rating. '
                   'The app cannot verify claims, judge visual layout, assess overall writing quality, or detect AI authorship.')
        st.write(f"**File:** {review['filename']}  |  **Extracted words:** {review['word_count']}")
        if review['reliability']:
            st.warning('Could not check reliably: ' + review['reliability'])
            st.caption('Try a selectable-text PDF or DOCX with clearly identified sections, or ask a person to review the original.')
        else:
            st.subheader('Area breakdown and checklist')
            st.caption('Strong requires at least 60% in every area and an overall result of at least 80 after adjustments and caps. '
                       'Good starts at 60; otherwise the band is Needs work. These cutoffs are provisional, not HR standards.')
            for area in review['areas']:
                with st.expander(f"{area['title']} — {area['points']}/{area['max']}", expanded=True):
                    for item in area['checks']:
                        st.write(f"**{item['title']} ({item['points']}/{item['maximum']})** — {item['reason']}")
                        if item['points'] < item['maximum']:
                            st.caption(item['advice'])
            st.subheader('Adjustments and limits')
            for observation in review.get('observations', []):
                st.write(observation)
            if review['adjustments']:
                for item in review['adjustments']:
                    st.warning(item['reason'])
            else:
                st.write('No additional adjustments were triggered.')
            for reason in review['cap_reasons']:
                st.write('Band limit: ' + reason)
        d = review['details']
        st.subheader('Extracted information')
        a, b, c = st.columns(3)
        a.write('**Name:** ' + (d['name'] or 'Not detected'))
        b.write('**Email:** ' + (d['email'] or 'Not detected'))
        c.write('**Phone:** ' + (d['phone'] or 'Not detected'))
        a.write('**Education:** ' + education_text(d['education']))
        b.write('**Experience mentioned:** ' + experience_display(d))
        b.write('**Experience status:** ' + (d.get('experience_status') or 'Not detected'))
        st.write('**Recognized skills:** ' + (', '.join(d['skills']) or 'None detected'))
        st.caption('Extraction may miss details or put them in the wrong section. Verify the original file before sharing.')
    else:
        st.info('Upload your resume again to use the updated personal review. Any previous job-description comparison is retained below.')
    if result.get('comparison_error'):
        st.warning('The optional job-description comparison could not finish: ' + result['comparison_error'])
    comparison = result.get('comparison')
    if comparison:
        st.subheader('Optional job-description comparison')
        st.caption('Separate from the personal review; a match index is not a hiring prediction.')
        score_metrics(comparison['items'][0])
        show_gap(comparison['items'][0]['gap'])


def improvement():
    st.header('Skill Improvement')
    result = need_result('personal_result')
    if not result:
        return
    review = result.get('review')
    if review:
        st.subheader('Resume improvements to consider')
        if review['reliability']:
            st.warning('Could not check reliably: ' + review['reliability'])
        else:
            suggestions = [item for area in review['areas'] for item in area['checks'] if item['points'] < item['maximum']]
            for item in suggestions:
                st.write(f"**{item['title']}:** {item['advice']}")
            for item in review['adjustments']:
                st.write(item['reason'])
            if not suggestions and not review['adjustments']:
                st.write('The checks did not flag an issue. This is not a guarantee of resume quality or hiring success.')
        st.caption('Suggestions are not universal requirements. Add only accurate information relevant to your goals.')
        st.subheader('Before sharing')
        st.write('- Open the original file and check the layout, alignment, spacing and readability.\n'
                 '- Proofread spelling and grammar, and verify dates, contact details and claims.\n'
                 '- Ask a trusted person for a second review; include only information you can explain.')
    comparison = result.get('comparison')
    if comparison:
        gap = comparison['items'][0]['gap']
        st.subheader('Optional job-description skill gaps')
        if not gap['missing']:
            st.write('No missing skills were detected among the recognized job-description requirements.')
        else:
            for skill in gap['missing']:
                st.write('**' + skill + '**')
                st.link_button('Search YouTube for ' + skill, gap['youtube_links'][skill])
            st.caption('Search links are not individually curated video recommendations.')
    elif review:
        st.caption('A skill-gap comparison is available only if you choose to provide a job description.')


def learning():
    st.header('Learning Hub')
    st.caption('CV and resume-writing resources. Check each resource before relying on it.')
    courses = load_cv_courses()
    if courses.empty:
        st.info('No course data found in data/cv_course.csv.')
        return
    for _, row in courses.iterrows():
        title = str(row.get('title', 'Untitled'))
        topic = str(row.get('topic', 'General'))
        url = str(row.get('youtube_url', '')).strip()
        with st.expander(title + ' — ' + topic):
            if url.startswith(('https://www.youtube.com/', 'https://youtu.be/')):
                st.video(url)
                st.link_button('Open on YouTube', url)
            else:
                st.caption('Video link unavailable.')


page = st.session_state.page
if page == 'home':
    # No empty sidebar on the landing page.
    st.markdown('<style>[data-testid="stSidebar"], [data-testid="collapsedControl"] {display:none;}</style>', unsafe_allow_html=True)
    landing()
elif page in ['overview', 'screening', 'rankings', 'skills', 'bias', 'reports']:
    if page == 'overview':
        # The dashboard uses its cards; inner pages retain their existing sidebar.
        st.markdown('<style>[data-testid="stSidebar"], [data-testid="collapsedControl"] {display:none;}</style>', unsafe_allow_html=True)
    else:
        sidebar('recruiter')
    {'overview': overview, 'screening': analyze_page, 'rankings': rankings,
     'skills': skill_analysis, 'bias': bias_page, 'reports': reports}[page]()
else:
    sidebar('personal')
    if page == 'analyze':
        analyze_page(personal=True)
    else:
        {'dashboard': personal_dashboard, 'insights': insights,
         'improve': improvement, 'learning': learning}.get(page, personal_dashboard)()
