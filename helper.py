# helper.py - backend for the resume scanner
# parsing, extraction, scoring and skill gap logic are kept together here

import io
import os
import re
from urllib.parse import quote_plus

import joblib
import numpy as np
import pandas as pd
import pdfplumber
import spacy
import streamlit as st
import torch
from docx import Document
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from sentence_transformers import SentenceTransformer


# These paths are relative to this file, not the user's current folder.
BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, 'data')


@st.cache_resource
def load_bert():
    return SentenceTransformer('sentence-transformers/all-MiniLM-L6-v2', device='cpu')


@st.cache_resource
def load_spacy():
    # The model is installed through requirements.txt.
    return spacy.load('en_core_web_sm')


@st.cache_resource
def load_classifier():
    path = os.path.join(BASE, 'models', 'classifier.pkl')
    if not os.path.isfile(path):
        return None
    try:
        # Only load a model file created by you or another trusted source.
        return joblib.load(path)
    except Exception as e:
        print('Could not load classifier:', e)
        return None


def parse_file(filename, file_bytes):
    """Read text from a PDF or DOCX. Image-only PDFs need OCR."""
    name = filename.lower()
    if name.endswith('.pdf'):
        parts = []
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    parts.append(text)
        return '\n'.join(parts)
    if name.endswith('.docx'):
        doc = Document(io.BytesIO(file_bytes))
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                parts.append(' | '.join(cell.text for cell in row.cells if cell.text.strip()))
        return '\n'.join(parts)
    raise ValueError('Only PDF and DOCX files are supported.')


def light_clean(text):
    # Keep newlines and Unicode for name and education extraction.
    text = (text or '').replace('\x00', ' ')
    return re.sub(r'[^\S\n]+', ' ', text).strip()


def clean_for_scoring(text):
    if not text or not text.strip():
        return ''
    text = re.sub(r'https?://\S+|www\.\S+', ' ', text, flags=re.I)
    text = re.sub(r'[\w.+-]+@[\w.-]+\.[a-zA-Z]{2,}', ' ', text)
    # Avoid treating date ranges as phone numbers.
    text = re.sub(r'(?<!\w)(?:\+\d{1,3}[ .-]?)?(?:\d[ .()-]?){10,13}(?!\w)', ' ', text)
    text = re.sub(r'[^\w\s.+#/-]', ' ', text.lower())
    doc = load_spacy()(text)
    return ' '.join(t.lemma_.lower() for t in doc
                    if not t.is_space and not t.is_punct and not t.is_stop)


@st.cache_data
def load_skills(path=None):
    path = path or os.path.join(DATA, 'skills.csv')
    df = pd.read_csv(path, keep_default_na=False)
    required = {'skill', 'category', 'aliases', 'weight', 'match_mode', 'youtube_url'}
    if not required.issubset(df.columns):
        raise ValueError('skills.csv needs these columns: ' + ', '.join(sorted(required)))
    # Keep the original spelling for case-sensitive skill rows.
    df['_skill_exact'] = df['skill'].astype(str).str.strip()
    df['skill'] = df['_skill_exact'].str.lower()
    if (df['skill'] == '').any() or df['skill'].duplicated().any():
        raise ValueError('skills.csv has blank or duplicate skills.')
    df['weight'] = pd.to_numeric(df['weight'], errors='coerce')
    if df['weight'].isna().any() or (df['weight'] <= 0).any():
        raise ValueError('Every skill needs a positive numeric weight.')
    df['match_mode'] = df['match_mode'].astype(str).str.strip().str.lower()
    if not df['match_mode'].isin(['standard', 'case_sensitive']).all():
        raise ValueError('Unknown match_mode in skills.csv.')
    return set(df['skill']), df


@st.cache_data
def load_bias_terms(path=None):
    path = path or os.path.join(DATA, 'bias_terms.txt')
    with open(path, encoding='utf-8') as f:
        return {line.strip().lower() for line in f if line.strip() and not line.startswith('#')}


@st.cache_data
def load_cv_courses(path=None):
    path = path or os.path.join(DATA, 'cv_course.csv')
    if not os.path.isfile(path):
        return pd.DataFrame(columns=['title', 'youtube_url', 'topic'])
    return pd.read_csv(path, keep_default_na=False)


def _term_in_text(text, term, case_sensitive=False):
    # Lookarounds allow skills with punctuation such as C++, C# and Node.js.
    term = term.strip()
    if not term:
        return False
    pattern = r'(?<![\w+#])' + re.escape(term) + r'(?![\w+#])'
    return bool(re.search(pattern, text, flags=0 if case_sensitive else re.IGNORECASE))


def _skill_terms(row):
    """Use original case for case-sensitive names without breaking old CSVs."""
    aliases = [a.strip() for a in str(row['aliases']).split('|') if a.strip()]
    sensitive = row['match_mode'] == 'case_sensitive'
    if not sensitive:
        return [row['skill']] + aliases, False
    original = str(row.get('_skill_exact', row['skill'])).strip()
    # Older CSV rows use a lowercase ID and put the intended spelling in aliases.
    # Preserve that convention, while supporting a capitalized original name.
    terms = list(aliases)
    if original != original.lower() or not terms:
        terms.insert(0, original)
    return list(dict.fromkeys(terms)), True


def find_skills(text, skills_set):
    # Read aliases and case rules from the same file as the weights.
    _, df = load_skills()
    found = []
    for _, row in df.iterrows():
        skill = row['skill']
        if skill not in skills_set:
            continue
        terms, case_sensitive = _skill_terms(row)
        # The short JavaScript alias "js" should not match a framework's
        # .js suffix or the "js" in the separate alias "node js".
        # An independently written "JS" or "JavaScript" still counts.
        def term_matches(term):
            if skill == 'javascript' and term.lower() == 'js':
                candidate_text = re.sub(
                    r'(?<![\w+#])node\s+js(?![\w+#])',
                    ' ', text or '', flags=re.IGNORECASE
                )
                flags = 0 if case_sensitive else re.IGNORECASE
                pattern = r'(?<![\w.+#])' + re.escape(term) + r'(?![\w+#])'
                return bool(re.search(pattern, candidate_text, flags=flags))
            return _term_in_text(text or '', term, case_sensitive)

        if any(term_matches(term) for term in terms):
            found.append(skill)
    return sorted(found)


def extract_email(text):
    m = re.search(r'\b[\w.+-]+@[\w.-]+\.[a-zA-Z]{2,}\b', text)
    return m.group(0) if m else None


def extract_phone(text):
    # Simple heuristic, so users should verify the result.
    pattern = r'(?<!\w)(?:\+\d{1,3}[ .-]?)?(?:\d[ .()-]?){10,13}(?!\w)'
    for match in re.finditer(pattern, text):
        value = match.group(0).strip()
        if 10 <= len(re.sub(r'\D', '', value)) <= 13:
            return value
    return None


def extract_name(text):
    # Resume titles and skills are sometimes incorrectly tagged as PERSON.
    # Only consider short name-like lines near the start, before sections.
    if not text.strip():
        return None

    _, df = load_skills()
    skill_terms = set(df['skill'].str.lower())
    for aliases in df['aliases']:
        skill_terms.update(a.strip().lower() for a in str(aliases).split('|') if a.strip())

    section_titles = {
        'summary', 'professional summary', 'profile', 'objective',
        'technical skills', 'skills', 'education', 'experience',
        'work experience', 'projects', 'certifications', 'contact',
    }
    not_name_words = {
        'resume', 'curriculum', 'vitae', 'summary', 'profile', 'objective',
        'skills', 'experience', 'education', 'projects', 'certifications',
        'engineer', 'developer', 'analyst', 'scientist', 'manager',
        'consultant', 'designer', 'architect', 'specialist', 'administrator',
        'intern', 'internship', 'student', 'fresher', 'professional',
        'software', 'data', 'senior', 'junior', 'lead', 'candidate',
    }
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for line in lines[:6]:
        normal = re.sub(r'\s+', ' ', line).strip()
        lower = normal.lower().rstrip(':').strip()
        # A standalone contact heading can precede the person's name.
        if lower in {'contact', 'contact information', 'name'}:
            continue
        if lower in section_titles:
            break
        # Phone/email/address lines are contact details, not name candidates.
        if re.match(r'^(?:phone|email|address)(?:\s|:|$)', normal, flags=re.I):
            continue
        # Remove a leading name/contact label and common honorifics only.
        normal = re.sub(
            r'^(?:name|contact(?:\s+information)?)(?=\s|:|\||-)'
            r'\s*[:|\-]?\s*', '', normal, flags=re.I
        )
        normal = re.sub(
            r'^(?:mr|mrs|ms|miss|dr)\.?(?=\s|$)\s*',
            '', normal, flags=re.I
        ).strip()
        lower = normal.lower().rstrip(':')
        if lower in skill_terms or any(c in normal for c in '@|,/\\'):
            continue
        words = normal.split()
        if not 2 <= len(words) <= 4:
            continue
        if any(word.lower().strip('.:-') in not_name_words for word in words):
            continue
        # Allow common initials, hyphens and apostrophes, but not digits.
        if all(re.fullmatch(r"[^\W\d_]+(?:[.'’\-][^\W\d_]+)*\.?", word, re.UNICODE)
               for word in words):
            return normal
    return None


@st.cache_data
def load_education_keywords(path=None):
    path = path or os.path.join(DATA, 'education_keywords.txt')
    with open(path, encoding='utf-8') as f:
        return [line.strip() for line in f if line.strip() and not line.startswith('#')]


def extract_education(text):
    # This is a keyword check, not a full qualification parser.
    result = []
    terms = sorted(load_education_keywords(), key=len, reverse=True)
    occupied = []
    for term in terms:
        pattern = r'(?<!\w)' + re.escape(term) + r'(?!\w)'
        for match in re.finditer(pattern, text, flags=re.I):
            if not any(match.start() < end and match.end() > start for start, end in occupied):
                occupied.append((match.start(), match.end()))
                result.append({'degree': match.group(0), 'field': 'Not specified'})
    return result


# Words accepted for explicitly stated durations.
_DURATION_ONES = {
    'zero': 0, 'one': 1, 'two': 2, 'three': 3, 'four': 4,
    'five': 5, 'six': 6, 'seven': 7, 'eight': 8, 'nine': 9,
    'ten': 10, 'eleven': 11, 'twelve': 12, 'thirteen': 13,
    'fourteen': 14, 'fifteen': 15, 'sixteen': 16,
    'seventeen': 17, 'eighteen': 18, 'nineteen': 19,
}
_DURATION_TENS = {
    'twenty': 20, 'thirty': 30, 'forty': 40, 'fifty': 50,
    'sixty': 60, 'seventy': 70, 'eighty': 80, 'ninety': 90,
}


def _duration_number(value):
    """Convert a number or simple number words into a number."""
    value = value.lower()
    try:
        return float(value)
    except ValueError:
        parts = re.split(r'[ -]+', value)
        if len(parts) == 1:
            return float(_DURATION_ONES.get(parts[0], _DURATION_TENS.get(parts[0], 0)))
        return float(_DURATION_TENS[parts[0]] + _DURATION_ONES[parts[1]])


def _duration_pattern(unit):
    one_to_nine = '|'.join(list(_DURATION_ONES)[1:10])
    words = (r'(?:' + '|'.join(_DURATION_TENS) + r')'
             r'(?:[ -](?:' + one_to_nine + r'))?'
             r'|(?:' + '|'.join(_DURATION_ONES) + r')')
    number = r'(?:\d+(?:\.\d+)?|' + words + r')'
    return (r'(?<![\w.+–—-])(?P<first>' + number + r')'
            r'(?:(?P<plus>\s*\+)|\s*(?:-|–|—|\bto\b)\s*'
            r'(?P<second>' + number + r'))?'
            r'\s*(?P<unit>' + unit + r')\b')


def _experience_info(text):
    """Return (exact years or None, stated duration or None, status).

    This is a simple extraction of an explicitly stated duration, not a
    calculation of total work experience from employment dates.
    """
    text = text or ''
    year_pattern = (_duration_pattern(r'years?|yrs?') + r'\s*'
                    r'(?:of\s+)?(?:work\s+|professional\s+)?'
                    r'(?:experience|exp)\b')
    match = re.search(year_pattern, text, flags=re.I)
    is_months = False

    if not match:
        # Months must be close to an experience/internship term, rather than
        # treating every course or project duration as employment experience.
        month_pattern = _duration_pattern(r'months?|mos?')
        for candidate in re.finditer(month_pattern, text, flags=re.I):
            line_start = text.rfind('\n', 0, candidate.start()) + 1
            previous_start = text.rfind('\n', 0, max(0, line_start - 1)) + 1
            line_end = text.find('\n', candidate.end())
            if line_end == -1:
                line_end = len(text)
            context = text[previous_start:line_end]
            if re.search(r'\b(?:intern(?:ship)?|experience|employment|work|worked|job)\b',
                         context, flags=re.I):
                match = candidate
                is_months = True
                break

    if not match:
        return None, None, 'Not detected'

    display = text[match.start('first'):match.end('unit')].strip()
    first = _duration_number(match.group('first'))
    second = _duration_number(match.group('second')) if match.group('second') else None
    plus = bool(match.group('plus'))
    threshold = 12 if is_months else 1

    # Do not force an ambiguous range or lower-bound claim into one category.
    if second is not None and second < first:
        status = 'Needs review'
    elif first >= threshold:
        status = 'Experienced'
    elif plus or (second is not None and second >= threshold):
        status = 'Needs review'
    else:
        status = 'Fresher'

    # Preserve the old numeric interface for exact years only. The app will
    # later use experience_text for months, ranges and '+' without conversion.
    exact_years = None if is_months or plus or second is not None else first
    return exact_years, display, status


def extract_experience(text):
    # Backward-compatible numeric value; ranges/+ are not exact values.
    return _experience_info(text)[0]


def get_all_details(raw_text, skills_set):
    text = light_clean(raw_text)
    experience, experience_text, experience_status = _experience_info(text)
    return {
        'name': extract_name(text),
        'email': extract_email(text),
        'phone': extract_phone(text),
        'skills': find_skills(text, skills_set),
        'education': extract_education(text),
        'experience': experience,
        'experience_text': experience_text,
        'experience_status': experience_status,
    }


def predict_category(raw_text):
    model = load_classifier()
    if model is None:
        return None, None
    try:
        pred = model.predict([raw_text])[0]
        probability = model.predict_proba([raw_text])[0]
        return str(pred), round(float(max(probability)) * 100, 1)
    except Exception as e:
        print('Classifier prediction failed:', e)
        return None, None


def detect_bias_terms(text, bias_terms=None):
    if bias_terms is None:
        bias_terms = load_bias_terms()
    return sorted(t for t in bias_terms if _term_in_text(text or '', t))


def apply_bias_mask(text, bias_terms=None):
    if not text:
        return ''
    if bias_terms is None:
        bias_terms = load_bias_terms()
    doc = load_spacy()(text)
    # Process in reverse so entity offsets stay correct.
    for ent in reversed(doc.ents):
        if ent.label_ == 'PERSON':
            text = text[:ent.start_char] + ' ' + text[ent.end_char:]
    text = re.sub(r'[\w.+-]+@[\w.-]+\.[a-zA-Z]{2,}', ' ', text)
    text = re.sub(r'https?://\S+|www\.\S+', ' ', text, flags=re.I)
    text = re.sub(r'(?<!\w)(?:\+\d{1,3}[ .-]?)?(?:\d[ .()-]?){10,13}(?!\w)', ' ', text)
    text = re.sub(r'\b\d{1,2}\s*years?\s*old\b', ' ', text, flags=re.I)
    text = re.sub(r'\b(?:date of birth|dob)\s*[:\-]?\s*[\d/.-]+', ' ', text, flags=re.I)
    for term in sorted(bias_terms, key=len, reverse=True):
        text = re.sub(r'(?<!\w)' + re.escape(term) + r'(?!\w)', ' ', text, flags=re.I)
    return re.sub(r'\s+', ' ', text).strip()


def calc_tfidf(resume_texts, jd_text):
    if not jd_text.strip() or not all(t.strip() for t in resume_texts):
        raise ValueError('A resume or job description is empty after cleaning.')
    vectorizer = TfidfVectorizer(max_features=5000, ngram_range=(1, 2),
                                 max_df=1.0, sublinear_tf=True)
    matrix = vectorizer.fit_transform([jd_text] + resume_texts)
    return np.clip(cosine_similarity(matrix[1:], matrix[:1]).flatten(), 0, 1)


def _embed_long_texts(texts, model):
    """Mean-pool token chunks so MiniLM can use more than its usual 256 tokens."""
    tokenizer = model.tokenizer
    seq_limit = model.max_seq_length
    chunk_size = seq_limit - tokenizer.num_special_tokens_to_add(pair=False)
    max_chunks = 64
    if chunk_size <= 0:
        raise ValueError('Invalid MiniLM sequence length.')

    # Keep token IDs instead of decoding chunks back to text. That way chunk
    # boundaries cannot accidentally change the number of tokens.
    all_chunks = []
    owners = []
    lengths = []
    for doc_index, text in enumerate(texts):
        # overflow returns every chunk instead of silently dropping later tokens
        encoded = tokenizer(text, add_special_tokens=False, truncation=True,
                            max_length=chunk_size, stride=0,
                            return_overflowing_tokens=True,
                            return_attention_mask=False, return_token_type_ids=False)
        chunks = encoded['input_ids']
        if not chunks or not chunks[0]:
            raise ValueError('A resume or job description has no usable tokens for semantic scoring.')
        if len(chunks) > max_chunks:
            label = 'Job description' if doc_index == 0 else f'Resume {doc_index}'
            raise ValueError(f'{label} is too long for semantic processing ({len(chunks)} chunks; limit {max_chunks}).')
        for chunk in chunks:
            all_chunks.append(tokenizer.prepare_for_model(chunk, add_special_tokens=True,
                                                           truncation=False, padding=False))
            owners.append(doc_index)
            lengths.append(len(chunk))

    # Weighted mean of normalized chunk embeddings, then normalize per document.
    sums = np.zeros((len(texts), model.get_sentence_embedding_dimension()), dtype=np.float64)
    batch_size = 8
    for start in range(0, len(all_chunks), batch_size):
        end = min(start + batch_size, len(all_chunks))
        batch = tokenizer.pad(all_chunks[start:end], padding=True, return_tensors='pt')
        batch = {key: value.to(model.device) for key, value in batch.items()}
        with torch.inference_mode():
            vectors = model(batch)['sentence_embedding']
            vectors = torch.nn.functional.normalize(vectors, p=2, dim=1).cpu().numpy()
        for offset, vector in enumerate(vectors):
            index = start + offset
            sums[owners[index]] += vector * lengths[index]
    norms = np.linalg.norm(sums, axis=1, keepdims=True)
    if np.any(norms == 0):
        raise ValueError('Could not create a semantic embedding for one of the documents.')
    return sums / norms


def calc_semantic(resume_texts, jd_text):
    if not resume_texts:
        return np.array([], dtype=float)
    if not jd_text.strip() or not all(t.strip() for t in resume_texts):
        raise ValueError('A resume or job description is empty after cleaning.')
    model = load_bert()
    embeddings = _embed_long_texts([jd_text] + resume_texts, model)
    similarity = cosine_similarity(embeddings[1:], embeddings[:1]).flatten()
    # Negative cosine is treated as zero. No +1 / 2 remapping.
    return np.clip(similarity, 0.0, 1.0)


def calc_weighted_skill_overlap(resume_skills, jd_skills, skills_df):
    if not jd_skills:
        raise ValueError('No skills detected in the job description. Add recognizable skills and try again.')
    weights = dict(zip(skills_df['skill'], skills_df['weight']))
    required = set(jd_skills)
    found = set(resume_skills)
    total = sum(weights.get(skill, 1.0) for skill in required)
    matched = sum(weights.get(skill, 1.0) for skill in required if skill in found)
    return matched / total


def compute_scores(resume_clean_texts, jd_clean_text, resume_skills_list, jd_skills, skills_df):
    if len(resume_clean_texts) != len(resume_skills_list):
        raise ValueError('Each resume must have one skills list.')
    if not resume_clean_texts:
        return []
    if not jd_skills:
        raise ValueError('No skills detected in the job description. Add recognizable skills and try again.')
    tfidf = calc_tfidf(resume_clean_texts, jd_clean_text)
    semantic = calc_semantic(resume_clean_texts, jd_clean_text)
    results = []
    for i in range(len(resume_clean_texts)):
        skill = calc_weighted_skill_overlap(resume_skills_list[i], jd_skills, skills_df)
        final = 0.40 * tfidf[i] + 0.40 * semantic[i] + 0.20 * skill
        results.append({
            'tfidf_score': round(float(tfidf[i]) * 100, 1),
            'semantic_score': round(float(semantic[i]) * 100, 1),
            'skill_score': round(float(skill) * 100, 1),
            'combined_score': round(float(final) * 100, 1),
        })
    return results


def get_skill_gap(resume_skills, jd_skills, skills_df=None):
    resume, required = set(resume_skills), set(jd_skills)
    matched = sorted(resume & required)
    missing = sorted(required - resume)
    extra = sorted(resume - required)
    links = {}
    if skills_df is not None and not skills_df.empty:
        for _, row in skills_df.iterrows():
            url = str(row.get('youtube_url', '')).strip()
            if row['skill'] in missing and url.startswith(('https://www.youtube.com/', 'https://youtu.be/')):
                links[row['skill']] = url
    # No curated links in the supplied CSV yet; link to search results instead.
    for skill in missing:
        if skill not in links:
            links[skill] = 'https://www.youtube.com/results?search_query=' + quote_plus(skill + ' tutorial')
    return {
        'matched': matched, 'missing': missing, 'extra': extra,
        'match_pct': round(len(matched) / len(required) * 100, 1) if required else 0,
        'total_required': len(required), 'total_matched': len(matched),
        'total_missing': len(missing), 'youtube_links': links,
    }


# ─── EXPERIMENTAL MATCH INDEX ───────────────────────────────────
# Scaling is deliberately off until benchmark ranges have been supplied.
# These ranges scale scores; they are not probability calibration.
MATCH_INDEX_CONFIG = {
    'use_scaling': False,
    'semantic_low': None,
    'semantic_high': None,
    'tfidf_low': None,
    'tfidf_high': None,
    'coverage_weight': 0.40,
    'semantic_weight': 0.40,
    'tfidf_weight': 0.20,
    'preferred_weight_factor': 0.5,
}


def _skill_occurrences(text, skills_set, skills_df):
    """Find skill positions for simple, explicit 'skill or skill' groups."""
    occurrences = []
    for _, row in skills_df.iterrows():
        skill = row['skill']
        if skill not in skills_set:
            continue
        terms, sensitive = _skill_terms(row)
        for term in terms:
            if skill == 'javascript' and term.lower() == 'js':
                pattern = r'(?<![\w.+#])' + re.escape(term) + r'(?![\w+#])'
            else:
                pattern = r'(?<![\w+#])' + re.escape(term) + r'(?![\w+#])'
            for match in re.finditer(pattern, text, flags=0 if sensitive else re.I):
                # Keep the same Node.js / node js protection as find_skills.
                if (skill == 'javascript' and term.lower() == 'js'
                        and re.search(r'(?<![\w+#])node\s+$',
                                      text[:match.start()], flags=re.I)):
                    continue
                occurrences.append((match.start(), match.end(), skill))
    return sorted(set(occurrences), key=lambda x: (x[0], -(x[1] - x[0]), x[2]))


def _preferred_context(text, start, end):
    """Approximate optional requirements from their sentence or bullet."""
    boundaries = list(re.finditer(r'[\n;]|(?<=[.!?])\s+', text))
    left, right = 0, len(text)
    for boundary in boundaries:
        if boundary.end() <= start:
            left = boundary.end()
        elif boundary.start() >= end:
            right = boundary.start()
            break
    part = text[left:right]
    return bool(re.search(
        r'\b(?:preferred|preferable|desirable|optional|advantage|advantageous|'
        r'bonus|nice[ -]to[ -]have|good[ -]to[ -]have|a plus|an asset|'
        r'highly\s+valued|helpful)\b', part, flags=re.I
    ))


def jd_skill_requirements(jd_text, skills_set, skills_df, preferred_factor=0.5):
    """Build skill requirements, grouping directly written 'X or Y'.

    For alternatives with different CSV weights, the group uses the highest
    alternative's weight. Only explicit adjacent 'or' is recognized; this is
    not a general language parser. A separately repeated skill remains separate.
    """
    jd_text = jd_text or ''
    if not 0 <= preferred_factor <= 1:
        raise ValueError('Preferred-skill factor must be between 0 and 1.')
    occurrences = _skill_occurrences(jd_text, skills_set, skills_df)
    # Use the longest occurrence at an overlapping position for grouping.
    selected = []
    for occurrence in occurrences:
        if not selected or occurrence[0] >= selected[-1][1]:
            selected.append(occurrence)

    groups = []
    i = 0
    while i < len(selected) - 1:
        chain = [selected[i]]
        j = i
        while (j + 1 < len(selected)
               and re.fullmatch(r'[ \t]+or[ \t]+',
                                jd_text[selected[j][1]:selected[j + 1][0]],
                                flags=re.I)):
            chain.append(selected[j + 1])
            j += 1
        if len(chain) > 1 and len({item[2] for item in chain}) > 1:
            groups.append(chain)
            i = j + 1
        else:
            i += 1

    weights = dict(zip(skills_df['skill'], skills_df['weight']))
    requirements = []
    grouped_spans = []
    for group in groups:
        alternatives = list(dict.fromkeys(item[2] for item in group))
        grouped_spans.extend((start, end) for start, end, _ in group)
        preferred = all(_preferred_context(jd_text, start, end)
                        for start, end, _ in group)
        base_weight = max(float(weights[skill]) for skill in alternatives)
        requirements.append({
            'skills': alternatives,
            'label': ' or '.join(alternatives),
            'base_weight': base_weight,
            'preferred': preferred,
            'weight': base_weight * (preferred_factor if preferred else 1.0),
        })

    # Keep separate mentions outside a group as separate requirements.
    outside = {}
    for start, end, skill in occurrences:
        if not any(start < group_end and end > group_start
                   for group_start, group_end in grouped_spans):
            outside.setdefault(skill, []).append(_preferred_context(jd_text, start, end))
    for skill in sorted(outside):
        # A separately stated non-preferred mention takes priority.
        preferred = all(outside[skill])
        base_weight = float(weights[skill])
        requirements.append({'skills': [skill], 'label': skill,
                             'base_weight': base_weight, 'preferred': preferred,
                             'weight': base_weight * (preferred_factor if preferred else 1.0)})
    return requirements


def keyword_coverage(resume_text, requirements, skills_set):
    """Score detected JD requirements using full/half-credit line rules."""
    text = light_clean(resume_text)
    resume_skills = set(find_skills(text, skills_set))
    descriptive_lines = [
        line for line in text.splitlines()
        if len(line.split()) >= 8 and line.count(',') < 4 and '|' not in line
    ]
    body_skills = set(find_skills('\n'.join(descriptive_lines), skills_set))
    total = sum(requirement['weight'] for requirement in requirements)
    if total <= 0:
        raise ValueError('No skills detected in the job description.')

    earned = 0.0
    matched, partial, missing = [], [], []
    for requirement in requirements:
        alternatives = set(requirement['skills'])
        if alternatives & body_skills:
            credit = 1.0
            matched.append(requirement['label'])
        elif alternatives & resume_skills:
            credit = 0.5
            partial.append(requirement['label'])
        else:
            credit = 0.0
            missing.append(requirement['label'])
        earned += requirement['weight'] * credit
    return {
        'coverage': earned / total,
        'matched': matched,
        'partial': partial,
        'missing': missing,
        'resume_skills': sorted(resume_skills),
        'body_skills': sorted(body_skills),
        'requirements': requirements,
    }


def _scale_match_component(value, low, high):
    if low is None or high is None or high <= low:
        raise ValueError('Set valid low and high scaling ranges before enabling match-index scaling.')
    return float(np.clip((value - low) / (high - low), 0.0, 1.0))


def _jd_required_years(jd_text):
    """Detect a simple stated year minimum; ambiguous wording is not inferred."""
    text = jd_text or ''
    found = []
    for match in re.finditer(_duration_pattern(r'years?|yrs?'), text, flags=re.I):
        line_start = text.rfind('\n', 0, match.start()) + 1
        line_end = text.find('\n', match.end())
        if line_end == -1:
            line_end = len(text)
        context = text[line_start:line_end]
        # Optional, negated or maximum durations should not become minimums.
        if re.search(r'\b(?:preferred|preferable|desirable|optional|not required|not mandatory|'
                     r'no minimum|up to|at most|maximum|less than)\b', context, re.I):
            continue
        # A plus sign can signal a minimum; otherwise require an experience cue.
        if not match.group('plus') and not re.search(r'\b(?:experience|exp)\b', context, re.I):
            continue
        first = _duration_number(match.group('first'))
        second = _duration_number(match.group('second')) if match.group('second') else None
        if second is not None and second < first:
            return None
        found.append(first)
    # Repeated mentions of the same minimum are not conflicting requirements.
    # Different minimums, or fresher mixed with one, still need review.
    distinct = set(found)
    if len(distinct) != 1 or re.search(r'\bfreshers?\b', text, re.I):
        return None
    return next(iter(distinct))


def _experience_factor(jd_text, resume_text):
    """Return a cautious multiplier based on stated total duration, if usable."""
    required = _jd_required_years(jd_text)
    _, stated, _ = _experience_info(light_clean(resume_text))
    result = {
        'required_years': required,
        'total_experience_detected': stated,
        'experience_fit': None,
        'experience_factor': 1.0,
        'experience_note': 'Experience not detected' if not stated else 'No usable JD experience minimum detected',
    }
    if required is None or required <= 0 or not stated:
        return result
    match = re.fullmatch(_duration_pattern(r'years?|yrs?|months?|mos?'), stated, flags=re.I)
    if not match:
        result['experience_note'] = 'Experience needs review; no penalty applied'
        return result
    first = _duration_number(match.group('first'))
    second = _duration_number(match.group('second')) if match.group('second') else None
    if match.group('unit').lower().startswith(('month', 'mo')):
        first /= 12
        if second is not None:
            second /= 12
    # Uncertain ranges/plus claims are not treated as exact experience.
    if match.group('plus') or second is not None:
        result['experience_note'] = 'Experience range or minimum needs review; no penalty applied'
        return result
    fit = min(1.0, first / required)
    result['experience_fit'] = fit
    result['experience_factor'] = 0.7 + 0.3 * fit
    result['experience_note'] = 'Based on a stated duration; relevance and total not verified'
    return result


def compute_match_indices(resume_texts, jd_text, config=None):
    """Experimental scoring for raw resumes and JD; does not change old scoring.

    TF-IDF uses the existing cleaned text and batch fitting. Semantic embeddings
    use lightly cleaned text. The caller is responsible for applying any mask
    consistently before supplying text to this function.
    """
    config = MATCH_INDEX_CONFIG if config is None else config
    if not resume_texts:
        return []
    weights = [config['coverage_weight'], config['semantic_weight'], config['tfidf_weight']]
    if any(weight < 0 for weight in weights) or not np.isclose(sum(weights), 1.0):
        raise ValueError('Match-index weights must be non-negative and total 1.0.')
    skills, skills_df = load_skills()
    requirements = jd_skill_requirements(light_clean(jd_text), skills, skills_df,
                                         config.get('preferred_weight_factor', 0.5))
    if not requirements:
        raise ValueError('No skills detected in the job description. Add recognizable skills and try again.')
    coverages = [keyword_coverage(text, requirements, skills) for text in resume_texts]
    tfidf = calc_tfidf([clean_for_scoring(text) for text in resume_texts], clean_for_scoring(jd_text))
    semantic = calc_semantic([light_clean(text) for text in resume_texts], light_clean(jd_text))
    results = []
    for i, text in enumerate(resume_texts):
        raw_tfidf, raw_semantic = float(tfidf[i]), float(semantic[i])
        if config['use_scaling']:
            used_tfidf = _scale_match_component(raw_tfidf, config['tfidf_low'], config['tfidf_high'])
            used_semantic = _scale_match_component(raw_semantic, config['semantic_low'], config['semantic_high'])
        else:
            used_tfidf, used_semantic = raw_tfidf, raw_semantic
        coverage = coverages[i]['coverage']
        combined = (weights[0] * coverage + weights[1] * used_semantic + weights[2] * used_tfidf)
        experience = _experience_factor(jd_text, text)
        score = combined * experience['experience_factor']
        results.append({
            'coverage_score': round(coverage * 100, 1),
            'semantic_score': round(used_semantic * 100, 1),
            'tfidf_score': round(used_tfidf * 100, 1),
            'semantic_raw': round(raw_semantic * 100, 1),
            'tfidf_raw': round(raw_tfidf * 100, 1),
            'combined_before_experience': round(combined * 100, 1),
            'match_index': round(score * 100, 1),
            'scaling_used': bool(config['use_scaling']),
            'coverage_details': coverages[i],
            **experience,
        })
    return results
