import os
import re
import html
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from bs4 import BeautifulSoup

BASE_PATH = r'D:\Hans_Glanzmann\436167332711627763-1789638831\7897076666aabb50cb36a1'
INDEX_FILE = os.path.join(BASE_PATH, 'index.html')
OUTPUT_EXCEL_PRIMARY = r'D:\Hans_Glanzmann\Website_Dokumentation_Hans_Glanzmann.xlsx'
OUTPUT_EXCEL_LOCAL = r'C:\Users\edlei\PycharmProjects\DocumentWebsite\Website_Dokumentation_Hans_Glanzmann.xlsx'

def read_html_file(fpath):
    with open(fpath, 'r', encoding='utf-8', errors='replace') as f:
        return f.read()

def clean_text(text):
    if not text:
        return ''
    t = html.unescape(text)
    t = t.replace('\xa0', ' ').replace('\u200b', '').replace('\ufeff', '')
    lines = [re.sub(r'[ \t]+', ' ', line).strip() for line in t.split('\n')]
    clean_lines = []
    for line in lines:
        if line or (clean_lines and clean_lines[-1]):
            clean_lines.append(line)
    return '\n'.join(clean_lines).strip()

def get_site_structure():
    html_content = read_html_file(INDEX_FILE)
    soup = BeautifulSoup(html_content, 'lxml')

    menu_ul = soup.find('ul', class_='wsite-menu-default')
    pages_ordered = []
    seen_files = set()

    def traverse(ul, level=0, parent_title='', parent_file='', breadcrumb=[]):
        for li in ul.find_all('li', recursive=False):
            a = li.find('a', recursive=False)
            if not a:
                continue
            menu_title = clean_text(a.get_text(separator=' ', strip=True))
            if menu_title.endswith('>'):
                menu_title = menu_title[:-1].strip()
            href = a.get('href', '').strip()
            fname = os.path.basename(href)
            
            cur_breadcrumb = breadcrumb + [menu_title]
            
            page_info = {
                'level': level,
                'menu_title': menu_title,
                'file_name': fname,
                'parent_title': parent_title,
                'parent_file': parent_file,
                'breadcrumb': ' > '.join(cur_breadcrumb),
                'in_nav': True
            }
            pages_ordered.append(page_info)
            seen_files.add(fname)
            
            subwrap = li.find('div', class_='wsite-menu-wrap', recursive=False)
            if subwrap:
                subul = subwrap.find('ul', class_='wsite-menu', recursive=False)
                if subul:
                    traverse(subul, level + 1, menu_title, fname, cur_breadcrumb)

    if menu_ul:
        traverse(menu_ul)

    all_html = [f for f in os.listdir(BASE_PATH) if f.endswith('.html')]
    for f in sorted(all_html):
        if f not in seen_files:
            pages_ordered.append({
                'level': -1,
                'menu_title': f.replace('.html', '').replace('-', ' ').title(),
                'file_name': f,
                'parent_title': '(Direktlink / Nicht im Hauptmenü)',
                'parent_file': '',
                'breadcrumb': f'(Direktlink) > {f}',
                'in_nav': False
            })

    return pages_ordered

def parse_artwork_metadata(raw_title):
    if not raw_title:
        return {'title': '', 'technique': '', 'dimensions': '', 'year': '', 'number': ''}
    
    t = clean_text(raw_title)
    
    # Dimensions e.g. 100 x 140cm, 28.5 x 38.5cm, Ø 50cm, 17x17x2cm
    dim_match = re.search(r'(\d+(?:[.,]\d+)?\s*(?:x|X|/)\s*\d+(?:[.,]\d+)?(?:\s*(?:x|X)\s*\d+(?:[.,]\d+)?)?\s*cm|\b\d+(?:[.,]\d+)?\s*cm\b|\bØ\s*\d+\s*cm\b)', t)
    dims = dim_match.group(1).strip() if dim_match else ''
    
    # Year e.g. 1986, 1989-2017, 2003/04, 2018/25, 2010
    year_match = re.search(r'\b((?:19|20)\d{2}(?:\s*[-/–]\s*(?:\d{4}|\d{2}))?)\b', t)
    year = year_match.group(1).strip() if year_match else ''
    
    # Number e.g. - 01, Nr. 1, Serie 1, #01, 1a
    num_match = re.search(r'(?:[-–]\s*|Nr\.?\s*|#\s*)(\d{1,3})\s*$', t)
    number = num_match.group(1).strip() if num_match else ''
    
    # Technique
    tech = ''
    tech_patterns = [
        r'(\bÖl\s+auf\s+(?:Leinwand|LW|Holz|Holzplatte|MDF|Hartfaser|Papier|Karton)\b)',
        r'(\bAcryl\s+auf\s+(?:Leinwand|LW|Holz|Holzplatte|MDF|Hartfaser|Papier|Karton)\b)',
        r'(\bÖlfarbe\s+auf\s+(?:Leinwand|LW|Holz|Pizzakarton|Papier|Karton)\b)',
        r'(\bIndustrielack\s+auf\s+(?:Hartfaser|Holz|LW)\b)',
        r'(\bRadierung(?:\s+auf\s+[A-Za-zäöüÄÖÜ]+)?\b)',
        r'(\bFarbradierung\b)',
        r'(\bAquatinta\b)',
        r'(\bKaltnadel\b)',
        r'(\bStrichätzung\b)',
        r'(\bGouache\b)',
        r'(\bAquarell\b)',
        r'(\bMischtechnik\b)',
        r'(\bZeichnung\b)',
        r'(\bBleistift\b)',
        r'(\bKreide\b)',
        r'(\bPastell\b)',
        r'(\bCollage\b)',
        r'(\bRelief\b)',
        r'(\bBozzetti\b)',
        r'(\bGips(?:\s+oder\s+[A-Za-z]+)?\b)',
        r'(\bBronze\b)',
        r'(\bKeramik\b)'
    ]
    for pat in tech_patterns:
        m = re.search(pat, t, re.IGNORECASE)
        if m:
            tech = m.group(1).strip()
            break

    title_clean = t.strip(' "\'“”„')
    
    return {
        'title': title_clean,
        'technique': tech,
        'dimensions': dims,
        'year': year,
        'number': number
    }

def generate_excel_documentation():
    print("Parsing website structure...")
    pages_meta = get_site_structure()
    print(f"Total pages discovered: {len(pages_meta)}")

    pages_data = []
    all_texts = []
    all_artworks = []
    all_links = []

    text_counter = 0
    artwork_counter = 0
    link_counter = 0

    for page_idx, pm in enumerate(pages_meta, 1):
        fname = pm['file_name']
        fpath = os.path.join(BASE_PATH, fname)
        file_url = f"file:///{fpath.replace('\\', '/')}"

        if not os.path.exists(fpath):
            print(f"Warning: File not found {fpath}")
            continue

        html_raw = read_html_file(fpath)
        soup = BeautifulSoup(html_raw, 'lxml')

        # Meta information
        html_title = clean_text(soup.title.string) if soup.title and soup.title.string else ''
        meta_desc_tag = soup.find('meta', attrs={'name': lambda n: n and n.lower() == 'description'})
        meta_desc = clean_text(meta_desc_tag.get('content', '')) if meta_desc_tag else ''
        meta_kw_tag = soup.find('meta', attrs={'name': lambda n: n and n.lower() == 'keywords'})
        meta_kw = clean_text(meta_kw_tag.get('content', '')) if meta_kw_tag else ''

        content = soup.find(id='wsite-content') or soup.find(class_='wsite-elements') or soup.find('body')

        page_headings = []
        if content:
            for h in content.find_all(['h1', 'h2', 'h3', 'h4', 'font'], class_=lambda c: c and 'title' in c):
                htxt = clean_text(h.get_text(separator=' ', strip=True))
                if htxt and htxt not in page_headings and len(htxt) > 1 and not htxt.startswith('<<<') and not htxt.endswith('>>>'):
                    page_headings.append(htxt)
            for h in content.find_all(['h1', 'h2']):
                htxt = clean_text(h.get_text(separator=' ', strip=True))
                if htxt and htxt not in page_headings and len(htxt) > 1 and not htxt.startswith('<<<') and not htxt.endswith('>>>'):
                    page_headings.append(htxt)

        current_section = pm['menu_title']
        page_text_count = 0
        page_artwork_count = 0
        page_link_count = 0

        if content:
            # 1. Text elements
            text_containers = content.find_all(['h1', 'h2', 'h3', 'h4', 'div', 'p', 'blockquote'], class_=True)
            seen_texts_on_page = set()

            for elem in text_containers:
                classes = elem.get('class', [])
                elem_type = 'Text'
                if any('title' in c for c in classes) or elem.name in ['h1', 'h2', 'h3']:
                    elem_type = 'Überschrift'
                elif 'paragraph' in classes or any('text' in c for c in classes) or elem.name == 'p':
                    elem_type = 'Absatz'
                elif elem.name == 'blockquote':
                    elem_type = 'Zitat / Blockquote'
                else:
                    continue

                txt = clean_text(elem.get_text(separator=' ', strip=True))
                if not txt or txt in seen_texts_on_page:
                    continue

                if re.match(r'^[<>\s]{2,}.*|.*[<>\s]{2,}$', txt) and len(txt) < 40:
                    continue

                seen_texts_on_page.add(txt)

                if elem_type == 'Überschrift' and len(txt) < 150:
                    current_section = txt

                text_counter += 1
                page_text_count += 1
                word_count = len(txt.split())

                all_texts.append({
                    'text_id': f"T-{text_counter:04d}",
                    'page_idx': page_idx,
                    'file_name': fname,
                    'menu_title': pm['menu_title'],
                    'breadcrumb': pm['breadcrumb'],
                    'section': current_section,
                    'type': elem_type,
                    'text': txt,
                    'word_count': word_count
                })

            # 2. Artworks / Images
            seen_img_srcs = set()
            galleries = content.find_all(class_=lambda c: c and 'galleryImageHolder' in c)
            for g in galleries:
                img = g.find('img')
                a = g.find('a')
                caption_elem = g.find(class_=lambda c: c and ('galleryCaptionHolder' in c or 'galleryImageCaption' in c))

                img_src = img.get('src', '').strip() if img else ''
                full_href = a.get('href', '').strip() if a else ''
                raw_caption = a.get('title', '') if a and a.get('title') else (caption_elem.get_text(separator=' ', strip=True) if caption_elem else '')
                if not raw_caption and img:
                    raw_caption = img.get('alt', '')

                raw_caption = clean_text(raw_caption)

                local_thumb = os.path.join(BASE_PATH, img_src.replace('/', '\\')) if img_src else ''
                local_full = os.path.join(BASE_PATH, full_href.replace('/', '\\')) if full_href else ''

                exists_thumb = os.path.exists(local_thumb) if local_thumb else False
                exists_full = os.path.exists(local_full) if local_full else False
                file_size_kb = round(os.path.getsize(local_full if exists_full else local_thumb) / 1024, 1) if (exists_full or exists_thumb) else 0

                meta = parse_artwork_metadata(raw_caption)

                artwork_counter += 1
                page_artwork_count += 1
                if img_src:
                    seen_img_srcs.add(img_src)

                all_artworks.append({
                    'artwork_id': f"W-{artwork_counter:04d}",
                    'page_idx': page_idx,
                    'file_name': fname,
                    'menu_title': pm['menu_title'],
                    'breadcrumb': pm['breadcrumb'],
                    'section': current_section,
                    'raw_caption': raw_caption,
                    'title': meta['title'],
                    'technique': meta['technique'],
                    'dimensions': meta['dimensions'],
                    'year': meta['year'],
                    'number': meta['number'],
                    'image_type': 'Galerie-Kunstwerk',
                    'thumb_path': img_src,
                    'full_path': full_href,
                    'exists_local': 'Ja' if (exists_thumb or exists_full) else 'Nein',
                    'file_size_kb': file_size_kb
                })

            # Standalone images
            standalone_divs = content.find_all('div', class_=lambda c: c and ('wsite-image' in c or 'image' in c))
            for sdiv in standalone_divs:
                img = sdiv.find('img')
                if not img:
                    continue
                img_src = img.get('src', '').strip()
                if not img_src or img_src in seen_img_srcs:
                    continue
                seen_img_srcs.add(img_src)

                a = sdiv.find('a')
                full_href = a.get('href', '').strip() if a else ''
                caption_elem = sdiv.find(class_=lambda c: c and 'caption' in c)
                raw_caption = caption_elem.get_text(separator=' ', strip=True) if caption_elem else (img.get('alt', '') or (a.get('title', '') if a else ''))
                raw_caption = clean_text(raw_caption)

                local_thumb = os.path.join(BASE_PATH, img_src.replace('/', '\\'))
                local_full = os.path.join(BASE_PATH, full_href.replace('/', '\\')) if full_href else ''

                exists_thumb = os.path.exists(local_thumb)
                exists_full = os.path.exists(local_full) if local_full else False
                file_size_kb = round(os.path.getsize(local_full if exists_full else local_thumb) / 1024, 1) if (exists_full or exists_thumb) else 0

                meta = parse_artwork_metadata(raw_caption)

                artwork_counter += 1
                page_artwork_count += 1

                all_artworks.append({
                    'artwork_id': f"W-{artwork_counter:04d}",
                    'page_idx': page_idx,
                    'file_name': fname,
                    'menu_title': pm['menu_title'],
                    'breadcrumb': pm['breadcrumb'],
                    'section': current_section,
                    'raw_caption': raw_caption,
                    'title': meta['title'] if meta['title'] else os.path.basename(img_src),
                    'technique': meta['technique'],
                    'dimensions': meta['dimensions'],
                    'year': meta['year'],
                    'number': meta['number'],
                    'image_type': 'Einzelabbildung / Foto',
                    'thumb_path': img_src,
                    'full_path': full_href if full_href else img_src,
                    'exists_local': 'Ja' if (exists_thumb or exists_full) else 'Nein',
                    'file_size_kb': file_size_kb
                })

            # 3. Links & Media
            for a in content.find_all('a', href=True):
                href = a.get('href', '').strip()
                link_text = clean_text(a.get_text(separator=' ', strip=True))
                if not href or href.startswith('#') or href.startswith('javascript:'):
                    continue

                link_type = 'Interner Seitenlink'
                if href.startswith('http://') or href.startswith('https://'):
                    link_type = 'Externer Link'
                elif href.startswith('mailto:'):
                    link_type = 'E-Mail-Kontakt'
                elif href.lower().endswith(('.mp3', '.wav', '.ogg', '.m4a', '.swf')):
                    link_type = 'Audio / Mediendatei'
                elif href.lower().endswith(('.pdf', '.doc', '.docx', '.zip')):
                    link_type = 'Dokument / Download'
                elif href.lower().endswith(('.jpg', '.jpeg', '.png', '.gif')):
                    link_type = 'Bildlink'

                loc_path = os.path.join(BASE_PATH, href.replace('/', '\\'))
                loc_exists = 'Ja' if os.path.exists(loc_path) else ('-' if link_type == 'Externer Link' else 'Nein')

                link_counter += 1
                page_link_count += 1

                all_links.append({
                    'link_id': f"L-{link_counter:04d}",
                    'page_idx': page_idx,
                    'file_name': fname,
                    'menu_title': pm['menu_title'],
                    'type': link_type,
                    'text': link_text if link_text else (a.get('title', '') or href),
                    'target_url': href,
                    'exists_local': loc_exists
                })

        text_snippets = [t['text'] for t in all_texts if t['page_idx'] == page_idx and t['type'] != 'Überschrift']
        summary = text_snippets[0][:250] + '...' if (text_snippets and len(text_snippets[0]) > 250) else (text_snippets[0] if text_snippets else '')

        level_name = {
            0: '0 - Hauptmenü',
            1: '1 - Unterseite (1. Ebene)',
            2: '2 - Detailseite (2. Ebene)',
            -1: 'Direktlink / Nicht im Menü'
        }.get(pm['level'], str(pm['level']))

        pages_data.append({
            'page_idx': page_idx,
            'level': level_name,
            'breadcrumb': pm['breadcrumb'],
            'parent_title': pm['parent_title'],
            'menu_title': pm['menu_title'],
            'file_name': fname,
            'local_path': fpath,
            'file_url': file_url,
            'html_title': html_title,
            'meta_description': meta_desc,
            'meta_keywords': meta_kw,
            'headings_summary': ' | '.join(page_headings),
            'text_count': page_text_count,
            'artwork_count': page_artwork_count,
            'link_count': page_link_count,
            'summary': summary
        })

    # Build Excel Workbook
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Styles
    font_header = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    fill_header = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
    fill_sub_header = PatternFill(start_color='2C5282', end_color='2C5282', fill_type='solid')

    font_title = Font(name='Segoe UI', size=14, bold=True, color='1F497D')
    font_bold = Font(name='Segoe UI', size=10, bold=True)
    font_regular = Font(name='Segoe UI', size=10)
    font_small = Font(name='Segoe UI', size=9, color='555555')

    fill_zebra_1 = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    fill_zebra_2 = PatternFill(start_color='F7FAFC', end_color='F7FAFC', fill_type='solid')

    border_thin_gray = Side(style='thin', color='D2D6DC')
    border_cell = Border(left=border_thin_gray, right=border_thin_gray, top=border_thin_gray, bottom=border_thin_gray)
    border_header = Border(left=border_thin_gray, right=border_thin_gray, top=border_thin_gray, bottom=Side(style='medium', color='1F497D'))

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    align_right = Alignment(horizontal='right', vertical='center', wrap_text=True)

    # -------------------------------------------------------------------------
    # Sheet 1: Seitenübersicht & Struktur
    # -------------------------------------------------------------------------
    ws_pages = wb.create_sheet(title='Seitenübersicht & Struktur')
    ws_pages.views.sheetView[0].showGridLines = True

    headers_pages = [
        ('Nr.', 6, align_center),
        ('Hierarchie-Ebene', 18, align_center),
        ('Menüpfad / Breadcrumb', 40, align_left),
        ('Übergeordnete Seite', 22, align_left),
        ('Menütitel', 25, align_left),
        ('HTML-Dateiname', 28, align_left),
        ('Browser-Seitentitel (<title>)', 35, align_left),
        ('Meta-Beschreibung', 35, align_left),
        ('Meta-Schlagwörter', 30, align_left),
        ('Hauptüberschriften auf Seite', 35, align_left),
        ('Anzahl Texte', 14, align_center),
        ('Anzahl Werke / Bilder', 18, align_center),
        ('Anzahl Links', 14, align_center),
        ('Inhaltsauszug / Zusammenfassung', 45, align_left),
        ('Lokaler Dateipfad', 45, align_left),
    ]

    for col_idx, (hname, width, al) in enumerate(headers_pages, 1):
        cell = ws_pages.cell(row=1, column=col_idx, value=hname)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_header
        ws_pages.column_dimensions[get_column_letter(col_idx)].width = width
    ws_pages.row_dimensions[1].height = 28

    for row_idx, p in enumerate(pages_data, 2):
        fill = fill_zebra_1 if row_idx % 2 == 0 else fill_zebra_2
        row_vals = [
            p['page_idx'],
            p['level'],
            p['breadcrumb'],
            p['parent_title'],
            p['menu_title'],
            p['file_name'],
            p['html_title'],
            p['meta_description'],
            p['meta_keywords'],
            p['headings_summary'],
            p['text_count'],
            p['artwork_count'],
            p['link_count'],
            p['summary'],
            p['local_path']
        ]
        for col_idx, val in enumerate(row_vals, 1):
            cell = ws_pages.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_regular
            cell.fill = fill
            cell.border = border_cell
            cell.alignment = headers_pages[col_idx-1][2]
        ws_pages.row_dimensions[row_idx].height = 22

    ws_pages.freeze_panes = 'A2'
    ws_pages.auto_filter.ref = ws_pages.dimensions

    # -------------------------------------------------------------------------
    # Sheet 2: Inhalte & Texte
    # -------------------------------------------------------------------------
    ws_texts = wb.create_sheet(title='Inhalte & Texte')
    ws_texts.views.sheetView[0].showGridLines = True

    headers_texts = [
        ('Text-ID', 10, align_center),
        ('Seite Nr.', 10, align_center),
        ('Dateiname', 26, align_left),
        ('Rubrik / Menütitel', 25, align_left),
        ('Hierarchie-Pfad', 35, align_left),
        ('Sektions-Überschrift', 30, align_left),
        ('Inhaltstyp', 16, align_center),
        ('Vollständiger Textinhalt', 65, align_left),
        ('Wortanzahl', 12, align_center)
    ]

    for col_idx, (hname, width, al) in enumerate(headers_texts, 1):
        cell = ws_texts.cell(row=1, column=col_idx, value=hname)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_header
        ws_texts.column_dimensions[get_column_letter(col_idx)].width = width
    ws_texts.row_dimensions[1].height = 28

    for row_idx, t in enumerate(all_texts, 2):
        fill = fill_zebra_1 if row_idx % 2 == 0 else fill_zebra_2
        row_vals = [
            t['text_id'],
            t['page_idx'],
            t['file_name'],
            t['menu_title'],
            t['breadcrumb'],
            t['section'],
            t['type'],
            t['text'],
            t['word_count']
        ]
        for col_idx, val in enumerate(row_vals, 1):
            cell = ws_texts.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_bold if t['type'] == 'Überschrift' else font_regular
            cell.fill = fill
            cell.border = border_cell
            cell.alignment = headers_texts[col_idx-1][2]
        ws_texts.row_dimensions[row_idx].height = 20 if len(t['text']) < 80 else 35

    ws_texts.freeze_panes = 'A2'
    ws_texts.auto_filter.ref = ws_texts.dimensions

    # -------------------------------------------------------------------------
    # Sheet 3: Kunstwerke & Bildkatalog
    # -------------------------------------------------------------------------
    ws_art = wb.create_sheet(title='Kunstwerke & Bildkatalog')
    ws_art.views.sheetView[0].showGridLines = True

    headers_art = [
        ('Werk-ID', 10, align_center),
        ('Seite Nr.', 10, align_center),
        ('HTML-Dateiname', 26, align_left),
        ('Rubrik / Menütitel', 25, align_left),
        ('Hierarchie-Pfad', 35, align_left),
        ('Sektion / Werkgruppe', 30, align_left),
        ('Original-Bildtitel / Beschriftung', 40, align_left),
        ('Bereinigter Werkname', 30, align_left),
        ('Technik / Material', 25, align_left),
        ('Abmessungen (Maße)', 20, align_left),
        ('Entstehungsjahr', 16, align_center),
        ('Werk-Nr.', 12, align_center),
        ('Bildtyp', 18, align_center),
        ('Vorschaubild-Pfad', 35, align_left),
        ('Original-Vollbildpfad', 35, align_left),
        ('Lokal vorhanden', 15, align_center),
        ('Dateigröße (KB)', 15, align_right)
    ]

    for col_idx, (hname, width, al) in enumerate(headers_art, 1):
        cell = ws_art.cell(row=1, column=col_idx, value=hname)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_header
        ws_art.column_dimensions[get_column_letter(col_idx)].width = width
    ws_art.row_dimensions[1].height = 28

    for row_idx, a in enumerate(all_artworks, 2):
        fill = fill_zebra_1 if row_idx % 2 == 0 else fill_zebra_2
        row_vals = [
            a['artwork_id'],
            a['page_idx'],
            a['file_name'],
            a['menu_title'],
            a['breadcrumb'],
            a['section'],
            a['raw_caption'],
            a['title'],
            a['technique'],
            a['dimensions'],
            a['year'],
            a['number'],
            a['image_type'],
            a['thumb_path'],
            a['full_path'],
            a['exists_local'],
            a['file_size_kb']
        ]
        for col_idx, val in enumerate(row_vals, 1):
            cell = ws_art.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_regular
            cell.fill = fill
            cell.border = border_cell
            cell.alignment = headers_art[col_idx-1][2]
        ws_art.row_dimensions[row_idx].height = 20

    ws_art.freeze_panes = 'A2'
    ws_art.auto_filter.ref = ws_art.dimensions

    # -------------------------------------------------------------------------
    # Sheet 4: Medien & Verlinkungen
    # -------------------------------------------------------------------------
    ws_links = wb.create_sheet(title='Medien & Verlinkungen')
    ws_links.views.sheetView[0].showGridLines = True

    headers_links = [
        ('Link-ID', 10, align_center),
        ('Seite Nr.', 10, align_center),
        ('Quellseite (Dateiname)', 26, align_left),
        ('Quell-Rubrik', 25, align_left),
        ('Linktyp / Medientyp', 20, align_center),
        ('Link-Beschriftung (Text)', 35, align_left),
        ('Ziel-URL / Dateipfad', 45, align_left),
        ('Lokal vorhanden', 15, align_center)
    ]

    for col_idx, (hname, width, al) in enumerate(headers_links, 1):
        cell = ws_links.cell(row=1, column=col_idx, value=hname)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_header
        ws_links.column_dimensions[get_column_letter(col_idx)].width = width
    ws_links.row_dimensions[1].height = 28

    for row_idx, l in enumerate(all_links, 2):
        fill = fill_zebra_1 if row_idx % 2 == 0 else fill_zebra_2
        row_vals = [
            l['link_id'],
            l['page_idx'],
            l['file_name'],
            l['menu_title'],
            l['type'],
            l['text'],
            l['target_url'],
            l['exists_local']
        ]
        for col_idx, val in enumerate(row_vals, 1):
            cell = ws_links.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_regular
            cell.fill = fill
            cell.border = border_cell
            cell.alignment = headers_links[col_idx-1][2]
        ws_links.row_dimensions[row_idx].height = 20

    ws_links.freeze_panes = 'A2'
    ws_links.auto_filter.ref = ws_links.dimensions

    # -------------------------------------------------------------------------
    # Sheet 5: Zusammenfassung & Statistik
    # -------------------------------------------------------------------------
    ws_stats = wb.create_sheet(title='Zusammenfassung & Statistik')
    ws_stats.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_stats.merge_cells('A1:F1')
    title_cell = ws_stats['A1']
    title_cell.value = "Webseiten-Dokumentation: Hans Glanzmann (Weebly Backup)"
    title_cell.font = Font(name='Segoe UI', size=16, bold=True, color='1F497D')
    title_cell.alignment = Alignment(horizontal='left', vertical='center')
    ws_stats.row_dimensions[1].height = 35

    ws_stats['A2'] = f"Ausgangsdatei: file:///D:/Hans_Glanzmann/436167332711627763-1789638831/7897076666aabb50cb36a1/index.html"
    ws_stats['A2'].font = font_small
    ws_stats.row_dimensions[2].height = 18

    # KPI Summary Table
    ws_stats.cell(row=4, column=1, value="Kennzahl").font = font_header
    ws_stats.cell(row=4, column=1).fill = fill_header
    ws_stats.cell(row=4, column=1).border = border_header
    ws_stats.cell(row=4, column=1).alignment = align_left
    ws_stats.column_dimensions['A'].width = 35

    ws_stats.cell(row=4, column=2, value="Wert").font = font_header
    ws_stats.cell(row=4, column=2).fill = fill_header
    ws_stats.cell(row=4, column=2).border = border_header
    ws_stats.cell(row=4, column=2).alignment = align_center
    ws_stats.column_dimensions['B'].width = 20

    ws_stats.cell(row=4, column=3, value="Beschreibung / Details").font = font_header
    ws_stats.cell(row=4, column=3).fill = fill_header
    ws_stats.cell(row=4, column=3).border = border_header
    ws_stats.cell(row=4, column=3).alignment = align_left
    ws_stats.column_dimensions['C'].width = 50

    total_words = sum(t['word_count'] for t in all_texts)
    local_art_count = sum(1 for a in all_artworks if a['exists_local'] == 'Ja')

    kpis = [
        ("Gesamtanzahl HTML-Seiten", len(pages_data), "Alle Webseiten inklusive Haupt- und Unterseiten"),
        ("Hauptmenü-Seiten (Ebene 0)", sum(1 for p in pages_meta if p['level'] == 0), "Übergeordnete Hauptkategorien"),
        ("Unterseiten Ebene 1", sum(1 for p in pages_meta if p['level'] == 1), "Erste Unterebene der Navigation"),
        ("Unterseiten Ebene 2", sum(1 for p in pages_meta if p['level'] == 2), "Detaillierte Werkzyklen und Werkserien"),
        ("Direktlink-Seiten (Nicht im Menü)", sum(1 for p in pages_meta if p['level'] == -1), "Verlinkte Sonderseiten (Werkverzeichnis, Testseiten)"),
        ("Erfasste Kunstwerke & Abbildungen", len(all_artworks), "Galerie-Bilder, Serien, Einzelabbildungen und Skizzen"),
        ("Bilder lokal auf Datenträger vorhanden", local_art_count, f"{round(local_art_count/len(all_artworks)*100, 1) if all_artworks else 0}% Verfügbarkeit der Mediendateien"),
        ("Erfasste Textblöcke & Absätze", len(all_texts), "Überschriften, Werkkommentare, einführende Texte"),
        ("Gesamte Wortanzahl", total_words, "Wortanzahl über die gesamte Website"),
        ("Erfasste Links & Medienreferenzen", len(all_links), "Interne Querverweise, Audio-Player, Downloads & externe Links")
    ]

    for k_idx, (k_name, k_val, k_desc) in enumerate(kpis, 5):
        fill = fill_zebra_1 if k_idx % 2 == 1 else fill_zebra_2
        ws_stats.cell(row=k_idx, column=1, value=k_name).font = font_bold
        ws_stats.cell(row=k_idx, column=1).fill = fill
        ws_stats.cell(row=k_idx, column=1).border = border_cell
        ws_stats.cell(row=k_idx, column=1).alignment = align_left

        ws_stats.cell(row=k_idx, column=2, value=k_val).font = font_bold
        ws_stats.cell(row=k_idx, column=2).fill = fill
        ws_stats.cell(row=k_idx, column=2).border = border_cell
        ws_stats.cell(row=k_idx, column=2).alignment = align_center

        ws_stats.cell(row=k_idx, column=3, value=k_desc).font = font_regular
        ws_stats.cell(row=k_idx, column=3).fill = fill
        ws_stats.cell(row=k_idx, column=3).border = border_cell
        ws_stats.cell(row=k_idx, column=3).alignment = align_left
        ws_stats.row_dimensions[k_idx].height = 20

    # Rubriken Summary Table
    start_rubrik_row = len(kpis) + 7
    ws_stats.cell(row=start_rubrik_row, column=1, value="Übersicht nach Rubriken / Werkgruppen").font = font_title
    ws_stats.row_dimensions[start_rubrik_row].height = 25

    rubrik_headers = [
        ("Rubrik / Menübereich", 35, align_left),
        ("HTML-Datei", 25, align_left),
        ("Hierarchie", 18, align_center),
        ("Bilder / Werke", 15, align_center),
        ("Texte", 12, align_center),
        ("Links", 12, align_center)
    ]
    for col_idx, (hname, width, al) in enumerate(rubrik_headers, 1):
        cell = ws_stats.cell(row=start_rubrik_row+1, column=col_idx, value=hname)
        cell.font = font_header
        cell.fill = fill_sub_header
        cell.alignment = align_center
        cell.border = border_header
        col_letter = get_column_letter(col_idx)
        if ws_stats.column_dimensions[col_letter].width < width:
            ws_stats.column_dimensions[col_letter].width = width
    ws_stats.row_dimensions[start_rubrik_row+1].height = 25

    for r_idx, p in enumerate(pages_data, start_rubrik_row+2):
        fill = fill_zebra_1 if r_idx % 2 == 1 else fill_zebra_2
        row_vals = [
            p['menu_title'],
            p['file_name'],
            p['level'],
            p['artwork_count'],
            p['text_count'],
            p['link_count']
        ]
        for col_idx, val in enumerate(row_vals, 1):
            cell = ws_stats.cell(row=r_idx, column=col_idx, value=val)
            cell.font = font_regular
            cell.fill = fill
            cell.border = border_cell
            cell.alignment = rubrik_headers[col_idx-1][2]
        ws_stats.row_dimensions[r_idx].height = 20

    # Save Excel
    os.makedirs(os.path.dirname(OUTPUT_EXCEL_PRIMARY), exist_ok=True)
    wb.save(OUTPUT_EXCEL_PRIMARY)
    print(f"Excel successfully saved to: {OUTPUT_EXCEL_PRIMARY}")

    wb.save(OUTPUT_EXCEL_LOCAL)
    print(f"Excel copy saved to: {OUTPUT_EXCEL_LOCAL}")

if __name__ == '__main__':
    generate_excel_documentation()
