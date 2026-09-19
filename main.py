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

OUTPUT_HTML_PRIMARY = r'D:\Hans_Glanzmann\Website_Dokumentation_Hans_Glanzmann.html'
OUTPUT_HTML_BASEDIR = os.path.join(BASE_PATH, 'Website_Dokumentation_Hans_Glanzmann.html')
OUTPUT_HTML_LOCAL = r'C:\Users\edlei\PycharmProjects\DocumentWebsite\Website_Dokumentation_Hans_Glanzmann.html'

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

    all_html = [f for f in os.listdir(BASE_PATH) if f.endswith('.html') and not f.startswith('Website_Dokumentation_')]
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

def extract_website_data():
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

    return pages_data, all_texts, all_artworks, all_links, kpis, pages_meta

def generate_excel_documentation(pages_data, all_texts, all_artworks, all_links, kpis, pages_meta):
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

    # Sheet 1: Seitenübersicht & Struktur
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

    # Sheet 2: Inhalte & Texte
    ws_texts = wb.create_sheet(title='Inhalte & Texte')
    ws_texts.views.sheetView[0].showGridLines = True

    headers_texts = [
        ('Text-ID', 10, align_center),
        ('Seite Nr.', 10, align_center),
        ('HTML-Dateiname', 26, align_left),
        ('Rubrik / Menütitel', 25, align_left),
        ('Hierarchie-Pfad', 35, align_left),
        ('Abschnitts-Überschrift', 30, align_left),
        ('Inhaltstyp', 20, align_center),
        ('Textinhalt', 70, align_left),
        ('Wortanzahl', 12, align_right)
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
            cell.font = font_regular
            cell.fill = fill
            cell.border = border_cell
            cell.alignment = headers_texts[col_idx-1][2]
        ws_texts.row_dimensions[row_idx].height = 20 if len(t['text']) < 80 else 35

    ws_texts.freeze_panes = 'A2'
    ws_texts.auto_filter.ref = ws_texts.dimensions

    # Sheet 3: Kunstwerke & Bildkatalog
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

    # Sheet 4: Medien & Verlinkungen
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

    # Sheet 5: Zusammenfassung & Statistik
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
    try:
        os.makedirs(os.path.dirname(OUTPUT_EXCEL_PRIMARY), exist_ok=True)
        wb.save(OUTPUT_EXCEL_PRIMARY)
        print(f"Excel successfully saved to: {OUTPUT_EXCEL_PRIMARY}")
    except PermissionError:
        print(f"Hinweis: {OUTPUT_EXCEL_PRIMARY} ist derzeit geöffnet und konnte nicht überschrieben werden.")

    try:
        wb.save(OUTPUT_EXCEL_LOCAL)
        print(f"Excel copy saved to: {OUTPUT_EXCEL_LOCAL}")
    except Exception as e:
        print(f"Hinweis bei lokalem Excel-Speichern: {e}")

def generate_html_documentation(pages_data, all_texts, all_artworks, all_links, kpis, pages_meta):
    print("Generating HTML documentation...")
    
    # Pre-calculate filter unique lists
    all_techniques = sorted(list(set(a['technique'] for a in all_artworks if a['technique'])))
    all_link_types = sorted(list(set(l['type'] for l in all_links if l['type'])))
    all_text_types = sorted(list(set(t['type'] for t in all_texts if t['type'])))
    all_levels = sorted(list(set(p['level'] for p in pages_data)))

    # Relative base prefix for images and html files
    rel_site_prefix = "436167332711627763-1789638831/7897076666aabb50cb36a1/"
    abs_site_prefix = "file:///D:/Hans_Glanzmann/436167332711627763-1789638831/7897076666aabb50cb36a1/"

    # Build KPI Summary cards
    total_pages = len(pages_data)
    total_art = len(all_artworks)
    local_art = sum(1 for a in all_artworks if a['exists_local'] == 'Ja')
    total_txt = len(all_texts)
    total_w = sum(t['word_count'] for t in all_texts)
    total_lnk = len(all_links)

    parts = []

    parts.append(f"""<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hans Glanzmann – Website-Dokumentation & Werkverzeichnis</title>
    <style>
        :root {{
            --primary: #1F497D;
            --primary-dark: #143257;
            --primary-light: #2C5282;
            --accent: #2B6CB0;
            --bg-page: #F8FAFC;
            --bg-card: #FFFFFF;
            --border-color: #D2D6DC;
            --border-subtle: #E2E8F0;
            --text-main: #2D3748;
            --text-muted: #718096;
            --zebra-row: #F7FAFC;
            --shadow-sm: 0 1px 3px rgba(0,0,0,0.06);
            --shadow-md: 0 4px 6px -1px rgba(0,0,0,0.1);
        }}

        * {{
            box-sizing: border-box;
        }}

        body {{
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', Arial, sans-serif;
            background-color: var(--bg-page);
            color: var(--text-main);
            margin: 0;
            padding: 0;
            line-height: 1.5;
            -webkit-font-smoothing: antialiased;
        }}

        /* Header Banner */
        .header-banner {{
            background: linear-gradient(135deg, #1F497D 0%, #17375E 100%);
            color: #FFFFFF;
            padding: 24px 36px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.12);
        }}

        .header-title-row {{
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            flex-wrap: wrap;
            gap: 16px;
        }}

        .header-banner h1 {{
            margin: 0 0 6px 0;
            font-size: 24px;
            font-weight: 700;
            letter-spacing: -0.3px;
        }}

        .header-banner .subtitle {{
            font-size: 14px;
            color: #CBD5E1;
            margin: 0 0 16px 0;
        }}

        .header-meta {{
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
            font-size: 13px;
        }}

        .header-meta-chip {{
            display: inline-flex;
            align-items: center;
            gap: 6px;
            background: rgba(255,255,255,0.12);
            padding: 6px 14px;
            border-radius: 6px;
            border: 1px solid rgba(255,255,255,0.18);
        }}

        /* Excel-Style Tab Bar */
        .tab-bar {{
            background: #FFFFFF;
            border-bottom: 2px solid #CBD5E1;
            padding: 0 36px;
            display: flex;
            gap: 6px;
            overflow-x: auto;
            position: sticky;
            top: 0;
            z-index: 100;
            box-shadow: 0 2px 4px rgba(0,0,0,0.04);
        }}

        .tab-button {{
            padding: 13px 20px;
            background: transparent;
            border: none;
            border-bottom: 3px solid transparent;
            font-size: 14px;
            font-weight: 600;
            color: #4A5568;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            transition: all 0.2s ease;
            white-space: nowrap;
            font-family: inherit;
        }}

        .tab-button:hover {{
            color: #1F497D;
            background: #F1F5F9;
        }}

        .tab-button.active {{
            color: #1F497D;
            border-bottom-color: #1F497D;
            background: #F8FAFC;
        }}

        .tab-badge {{
            background: #E2E8F0;
            color: #4A5568;
            font-size: 11px;
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 12px;
        }}

        .tab-button.active .tab-badge {{
            background: #1F497D;
            color: #FFFFFF;
        }}

        /* Main Container */
        .container {{
            max-width: 1800px;
            margin: 0 auto;
            padding: 24px 36px;
        }}

        .tab-pane {{
            display: none;
        }}

        .tab-pane.active {{
            display: block;
        }}

        /* KPI Cards Grid */
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }}

        .kpi-card {{
            background: #FFFFFF;
            border-radius: 8px;
            border: 1px solid var(--border-color);
            padding: 18px 20px;
            box-shadow: var(--shadow-sm);
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            border-top: 4px solid var(--primary);
        }}

        .kpi-label {{
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: var(--text-muted);
            margin-bottom: 6px;
        }}

        .kpi-value {{
            font-size: 28px;
            font-weight: 700;
            color: var(--primary);
            margin-bottom: 4px;
        }}

        .kpi-subtext {{
            font-size: 12px;
            color: #718096;
        }}

        /* Toolbar & Controls */
        .toolbar {{
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 16px;
            background: #FFFFFF;
            padding: 12px 18px;
            border-radius: 8px;
            border: 1px solid var(--border-color);
            box-shadow: var(--shadow-sm);
        }}

        .toolbar-group {{
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            align-items: center;
        }}

        .search-input {{
            padding: 8px 14px;
            border: 1px solid #CBD5E1;
            border-radius: 6px;
            font-size: 13px;
            min-width: 280px;
            outline: none;
            font-family: inherit;
        }}

        .search-input:focus {{
            border-color: var(--primary);
            box-shadow: 0 0 0 2px rgba(31,73,125,0.15);
        }}

        .filter-select {{
            padding: 8px 12px;
            border: 1px solid #CBD5E1;
            border-radius: 6px;
            font-size: 13px;
            background-color: #FFFFFF;
            outline: none;
            cursor: pointer;
            font-family: inherit;
            color: var(--text-main);
        }}

        .filter-select:focus {{
            border-color: var(--primary);
        }}

        .btn {{
            padding: 8px 14px;
            border-radius: 6px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            border: 1px solid #CBD5E1;
            background: #FFFFFF;
            color: var(--text-main);
            display: inline-flex;
            align-items: center;
            gap: 6px;
            transition: all 0.15s ease;
            font-family: inherit;
        }}

        .btn:hover {{
            background: #F1F5F9;
            border-color: #94A3B8;
        }}

        .btn-primary {{
            background: var(--primary);
            color: #FFFFFF;
            border-color: var(--primary-dark);
        }}

        .btn-primary:hover {{
            background: var(--primary-light);
        }}

        .btn-active {{
            background: #E2E8F0;
            border-color: #94A3B8;
            font-weight: 700;
        }}

        .count-indicator {{
            font-size: 13px;
            color: var(--text-muted);
            font-weight: 500;
        }}

        /* Table Design */
        .table-wrap {{
            background: #FFFFFF;
            border-radius: 8px;
            border: 1px solid var(--border-color);
            overflow: auto;
            max-height: 72vh;
            box-shadow: var(--shadow-sm);
        }}

        table.excel-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
            text-align: left;
        }}

        table.excel-table th {{
            background-color: #1F497D;
            color: #FFFFFF;
            padding: 10px 14px;
            font-weight: 600;
            font-size: 13px;
            position: sticky;
            top: 0;
            z-index: 10;
            border-right: 1px solid #2C5282;
            border-bottom: 2px solid #143257;
            user-select: none;
            white-space: nowrap;
        }}

        table.excel-table th.sortable {{
            cursor: pointer;
        }}

        table.excel-table th.sortable:hover {{
            background-color: #2C5282;
        }}

        table.excel-table th .sort-arrow {{
            display: inline-block;
            margin-left: 4px;
            font-size: 10px;
            color: #CBD5E1;
        }}

        table.excel-table td {{
            padding: 8px 14px;
            border-bottom: 1px solid var(--border-subtle);
            border-right: 1px solid var(--border-subtle);
            vertical-align: middle;
        }}

        table.excel-table tr:nth-child(even) {{
            background-color: #F8FAFC;
        }}

        table.excel-table tr:hover {{
            background-color: #EDF2F7;
        }}

        .text-center {{ text-align: center; }}
        .text-right {{ text-align: right; }}

        /* Badges */
        .badge {{
            display: inline-block;
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 600;
            white-space: nowrap;
        }}

        .badge-yes {{ background: #DEF7EC; color: #03543F; border: 1px solid #BCF0DA; }}
        .badge-no {{ background: #FDE8E8; color: #9B1C1C; border: 1px solid #FBD5D5; }}
        .badge-neutral {{ background: #EDF2F7; color: #4A5568; }}
        .badge-lvl0 {{ background: #E1EFFE; color: #1E429F; border: 1px solid #BEE3F8; font-weight: 700; }}
        .badge-lvl1 {{ background: #EBF8FF; color: #2B6CB0; border: 1px solid #CBD5E1; }}
        .badge-lvl2 {{ background: #FAF5FF; color: #6B46C1; border: 1px solid #E9D8FD; }}
        .badge-direct {{ background: #F7FAFC; color: #718096; border: 1px solid #E2E8F0; }}
        .badge-tech {{ background: #F1F5F9; color: #1F497D; border: 1px solid #CBD5E1; font-weight: 500; }}
        .badge-type {{ background: #EDF2F7; color: #2D3748; border: 1px solid #E2E8F0; }}
        .badge-id {{ font-family: Consolas, monospace; font-size: 12px; font-weight: 600; color: #1F497D; background: #EBF8FF; padding: 2px 6px; border-radius: 4px; }}

        /* Artwork Thumbnails */
        .thumb-wrap {{
            width: 48px;
            height: 48px;
            border-radius: 4px;
            overflow: hidden;
            background: #EDF2F7;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            border: 1px solid #CBD5E1;
            flex-shrink: 0;
        }}

        .thumb-img {{
            width: 100%;
            height: 100%;
            object-fit: cover;
            transition: transform 0.2s ease;
        }}

        .thumb-wrap:hover .thumb-img {{
            transform: scale(1.12);
        }}

        .thumb-placeholder {{
            font-size: 18px;
            color: #A0AEC0;
        }}

        /* Gallery Grid View */
        .gallery-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
            gap: 18px;
        }}

        .gallery-card {{
            background: #FFFFFF;
            border: 1px solid var(--border-color);
            border-radius: 8px;
            overflow: hidden;
            box-shadow: var(--shadow-sm);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
            display: flex;
            flex-direction: column;
        }}

        .gallery-card:hover {{
            transform: translateY(-3px);
            box-shadow: var(--shadow-md);
        }}

        .gallery-card-img-wrap {{
            width: 100%;
            height: 200px;
            background: #EDF2F7;
            display: flex;
            align-items: center;
            justify-content: center;
            overflow: hidden;
            cursor: pointer;
            position: relative;
        }}

        .gallery-card-img {{
            width: 100%;
            height: 100%;
            object-fit: cover;
            transition: transform 0.3s ease;
        }}

        .gallery-card-img:hover {{
            transform: scale(1.04);
        }}

        .gallery-card-badge {{
            position: absolute;
            top: 8px;
            left: 8px;
            background: rgba(31,73,125,0.85);
            color: #FFFFFF;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 700;
            backdrop-filter: blur(2px);
        }}

        .gallery-card-body {{
            padding: 14px;
            flex: 1;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }}

        .gallery-card-title {{
            font-size: 14px;
            font-weight: 600;
            color: #1F497D;
            margin-bottom: 6px;
            line-height: 1.3;
        }}

        .gallery-card-meta {{
            font-size: 12px;
            color: #718096;
            margin-bottom: 4px;
        }}

        .gallery-card-footer {{
            margin-top: 10px;
            padding-top: 8px;
            border-top: 1px solid var(--border-subtle);
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 11px;
        }}

        /* Links */
        a.table-link {{
            color: #2B6CB0;
            text-decoration: none;
            font-weight: 500;
        }}

        a.table-link:hover {{
            text-decoration: underline;
            color: #1F497D;
        }}

        /* Lightbox Modal */
        .modal {{
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(15, 23, 42, 0.85);
            z-index: 1000;
            align-items: center;
            justify-content: center;
            padding: 24px;
            backdrop-filter: blur(4px);
        }}

        .modal.active {{
            display: flex;
        }}

        .modal-content {{
            background: #FFFFFF;
            border-radius: 12px;
            max-width: 960px;
            max-height: 90vh;
            width: 100%;
            overflow: hidden;
            display: flex;
            flex-direction: column;
            box-shadow: 0 20px 25px -5px rgba(0,0,0,0.3);
            position: relative;
        }}

        .modal-header {{
            padding: 16px 20px;
            background: #1F497D;
            color: #FFFFFF;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        .modal-title {{
            font-size: 16px;
            font-weight: 700;
            margin: 0;
        }}

        .modal-close {{
            background: none;
            border: none;
            color: #FFFFFF;
            font-size: 24px;
            cursor: pointer;
            line-height: 1;
            padding: 0 4px;
        }}

        .modal-body {{
            display: flex;
            flex-direction: row;
            overflow: auto;
            max-height: calc(90vh - 65px);
        }}

        @media (max-width: 768px) {{
            .modal-body {{
                flex-direction: column;
            }}
        }}

        .modal-img-area {{
            flex: 1.2;
            background: #1A202C;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 16px;
            min-height: 300px;
        }}

        .modal-img {{
            max-width: 100%;
            max-height: 70vh;
            object-fit: contain;
            border-radius: 4px;
        }}

        .modal-details {{
            flex: 1;
            padding: 24px;
            overflow-y: auto;
            font-size: 13px;
        }}

        .modal-details-table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 12px;
        }}

        .modal-details-table td {{
            padding: 8px 10px;
            border-bottom: 1px solid var(--border-subtle);
        }}

        .modal-details-table td:first-child {{
            font-weight: 600;
            color: #4A5568;
            width: 130px;
        }}

        /* Section Title */
        .section-heading {{
            font-size: 18px;
            font-weight: 700;
            color: #1F497D;
            margin: 24px 0 12px 0;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
    </style>
</head>
<body>

    <!-- Header Banner -->
    <header class="header-banner">
        <div class="header-title-row">
            <div>
                <h1>Hans Glanzmann – Website-Dokumentation & Werkverzeichnis</h1>
                <div class="subtitle">Vollständige Archivierung und Katalogisierung aller Webseiten, Texte, Kunstwerke und Medien (Weebly-Backup)</div>
            </div>
            <div>
                <button class="btn btn-primary" onclick="window.print()">🖨️ Drucken / PDF</button>
            </div>
        </div>
        <div class="header-meta">
            <div class="header-meta-chip">📁 <strong>Quelle:</strong> D:\\Hans_Glanzmann</div>
            <div class="header-meta-chip">📑 <strong>Seiten:</strong> {total_pages}</div>
            <div class="header-meta-chip">🎨 <strong>Kunstwerke:</strong> {total_art}</div>
            <div class="header-meta-chip">📝 <strong>Texte:</strong> {total_txt} ({total_w:,} Wörter)</div>
            <div class="header-meta-chip">🔗 <strong>Links:</strong> {total_lnk}</div>
            <div class="header-meta-chip">✔️ <strong>Bilder lokal:</strong> {local_art} / {total_art} ({round(local_art/total_art*100, 1)}%)</div>
        </div>
    </header>

    <!-- Excel-Style Tab Bar -->
    <nav class="tab-bar">
        <button class="tab-button active" onclick="switchTab('stats', this)">
            📊 Zusammenfassung & Statistik
        </button>
        <button class="tab-button" onclick="switchTab('pages', this)">
            📑 Seitenübersicht & Struktur <span class="tab-badge">{total_pages}</span>
        </button>
        <button class="tab-button" onclick="switchTab('texts', this)">
            📝 Inhalte & Texte <span class="tab-badge">{total_txt}</span>
        </button>
        <button class="tab-button" onclick="switchTab('artworks', this)">
            🎨 Kunstwerke & Bildkatalog <span class="tab-badge">{total_art}</span>
        </button>
        <button class="tab-button" onclick="switchTab('links', this)">
            🔗 Medien & Verlinkungen <span class="tab-badge">{total_lnk}</span>
        </button>
    </nav>

    <main class="container">

        <!-- TAB 1: Zusammenfassung & Statistik -->
        <section id="tab-stats" class="tab-pane active">
            <div class="kpi-grid">
                <div class="kpi-card">
                    <div class="kpi-label">Erfasste HTML-Seiten</div>
                    <div class="kpi-value">{total_pages}</div>
                    <div class="kpi-subtext">35 Haupt- und Unterseiten</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Kunstwerke & Abbildungen</div>
                    <div class="kpi-value">{total_art}</div>
                    <div class="kpi-subtext">1'617 von 1'618 Bildern lokal vorhanden (99.9%)</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Redaktionelle Textblöcke</div>
                    <div class="kpi-value">{total_txt}</div>
                    <div class="kpi-subtext">Gesamtumfang: {total_w:,} Wörter</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Medien & Verlinkungen</div>
                    <div class="kpi-value">{total_lnk}</div>
                    <div class="kpi-subtext">Interne Seitenlinks, Mediendateien, E-Mails</div>
                </div>
            </div>

            <div class="section-heading">📈 Kennzahlen & Strukturübersicht</div>
            <div class="table-wrap" style="max-height: none; margin-bottom: 24px;">
                <table class="excel-table">
                    <thead>
                        <tr>
                            <th style="width: 300px;">Kennzahl</th>
                            <th style="width: 140px;" class="text-center">Wert</th>
                            <th>Beschreibung / Details</th>
                        </tr>
                    </thead>
                    <tbody>
    """)

    for k_name, k_val, k_desc in kpis:
        val_str = f"{k_val:,}" if isinstance(k_val, int) else str(k_val)
        parts.append(f"""
                        <tr>
                            <td><strong>{html.escape(k_name)}</strong></td>
                            <td class="text-center"><strong>{html.escape(val_str)}</strong></td>
                            <td>{html.escape(k_desc)}</td>
                        </tr>
        """)

    parts.append("""
                    </tbody>
                </table>
            </div>

            <div class="section-heading">📑 Übersicht nach Rubriken / Werkgruppen</div>
            <div class="table-wrap" style="max-height: none;">
                <table class="excel-table">
                    <thead>
                        <tr>
                            <th>Rubrik / Menübereich</th>
                            <th>HTML-Datei</th>
                            <th class="text-center">Hierarchie</th>
                            <th class="text-center">Bilder / Werke</th>
                            <th class="text-center">Texte</th>
                            <th class="text-center">Links</th>
                        </tr>
                    </thead>
                    <tbody>
    """)

    for p in pages_data:
        lvl = p['level']
        badge_cls = "badge-lvl0" if "0 -" in lvl else ("badge-lvl1" if "1 -" in lvl else ("badge-lvl2" if "2 -" in lvl else "badge-direct"))
        parts.append(f"""
                        <tr>
                            <td><strong>{html.escape(p['menu_title'])}</strong></td>
                            <td><a href="{rel_site_prefix}{p['file_name']}" class="table-link" target="_blank">{html.escape(p['file_name'])} ↗</a></td>
                            <td class="text-center"><span class="badge {badge_cls}">{html.escape(lvl)}</span></td>
                            <td class="text-center"><strong>{p['artwork_count']}</strong></td>
                            <td class="text-center">{p['text_count']}</td>
                            <td class="text-center">{p['link_count']}</td>
                        </tr>
        """)

    parts.append(f"""
                    </tbody>
                </table>
            </div>
        </section>

        <!-- TAB 2: Seitenübersicht & Struktur -->
        <section id="tab-pages" class="tab-pane">
            <div class="toolbar">
                <div class="toolbar-group">
                    <input type="text" id="pages-search" class="search-input" placeholder="🔍 Seiten filtern (Titel, Datei, Überschrift...)" oninput="filterPagesTable()">
                    <select id="pages-level-filter" class="filter-select" onchange="filterPagesTable()">
                        <option value="">Alle Ebenen</option>
    """)

    for lvl in all_levels:
        parts.append(f"""<option value="{html.escape(lvl)}">{html.escape(lvl)}</option>""")

    parts.append(f"""
                    </select>
                </div>
                <div class="count-indicator" id="pages-count">Zeige {len(pages_data)} von {len(pages_data)} Seiten</div>
            </div>

            <div class="table-wrap">
                <table class="excel-table" id="pages-table">
                    <thead>
                        <tr>
                            <th class="sortable text-center" style="width: 50px;" onclick="sortTable('pages-table', 0)">Nr. <span class="sort-arrow">▼</span></th>
                            <th class="sortable text-center" style="width: 130px;" onclick="sortTable('pages-table', 1)">Ebene <span class="sort-arrow">▼</span></th>
                            <th class="sortable" style="width: 220px;" onclick="sortTable('pages-table', 2)">Menüpfad / Breadcrumb <span class="sort-arrow">▼</span></th>
                            <th class="sortable" style="width: 180px;" onclick="sortTable('pages-table', 3)">Menütitel <span class="sort-arrow">▼</span></th>
                            <th class="sortable" style="width: 180px;" onclick="sortTable('pages-table', 4)">HTML-Dateiname <span class="sort-arrow">▼</span></th>
                            <th class="sortable" style="width: 200px;" onclick="sortTable('pages-table', 5)">Browser-Seitentitel <span class="sort-arrow">▼</span></th>
                            <th style="width: 200px;">Hauptüberschriften</th>
                            <th class="sortable text-center" style="width: 80px;" onclick="sortTable('pages-table', 7)">Texte <span class="sort-arrow">▼</span></th>
                            <th class="sortable text-center" style="width: 80px;" onclick="sortTable('pages-table', 8)">Werke <span class="sort-arrow">▼</span></th>
                            <th class="sortable text-center" style="width: 80px;" onclick="sortTable('pages-table', 9)">Links <span class="sort-arrow">▼</span></th>
                            <th style="min-width: 250px;">Inhaltsauszug</th>
                        </tr>
                    </thead>
                    <tbody>
    """)

    for p in pages_data:
        lvl = p['level']
        badge_cls = "badge-lvl0" if "0 -" in lvl else ("badge-lvl1" if "1 -" in lvl else ("badge-lvl2" if "2 -" in lvl else "badge-direct"))
        parts.append(f"""
                        <tr data-level="{html.escape(lvl)}">
                            <td class="text-center">{p['page_idx']}</td>
                            <td class="text-center"><span class="badge {badge_cls}">{html.escape(lvl)}</span></td>
                            <td>{html.escape(p['breadcrumb'])}</td>
                            <td><strong>{html.escape(p['menu_title'])}</strong></td>
                            <td><a href="{rel_site_prefix}{p['file_name']}" class="table-link" target="_blank">{html.escape(p['file_name'])} ↗</a></td>
                            <td>{html.escape(p['html_title'])}</td>
                            <td><small>{html.escape(p['headings_summary'])}</small></td>
                            <td class="text-center">{p['text_count']}</td>
                            <td class="text-center"><strong>{p['artwork_count']}</strong></td>
                            <td class="text-center">{p['link_count']}</td>
                            <td><small style="color: #4A5568;">{html.escape(p['summary'])}</small></td>
                        </tr>
        """)

    parts.append(f"""
                    </tbody>
                </table>
            </div>
        </section>

        <!-- TAB 3: Inhalte & Texte -->
        <section id="tab-texts" class="tab-pane">
            <div class="toolbar">
                <div class="toolbar-group">
                    <input type="text" id="texts-search" class="search-input" placeholder="🔍 Texte durchsuchen (Stichwort, ID, Rubrik...)" oninput="filterTextsTable()">
                    <select id="texts-type-filter" class="filter-select" onchange="filterTextsTable()">
                        <option value="">Alle Texttypen</option>
    """)

    for tt in all_text_types:
        parts.append(f"""<option value="{html.escape(tt)}">{html.escape(tt)}</option>""")

    parts.append(f"""
                    </select>
                </div>
                <div class="count-indicator" id="texts-count">Zeige {len(all_texts)} von {len(all_texts)} Texte</div>
            </div>

            <div class="table-wrap">
                <table class="excel-table" id="texts-table">
                    <thead>
                        <tr>
                            <th class="sortable text-center" style="width: 80px;" onclick="sortTable('texts-table', 0)">Text-ID <span class="sort-arrow">▼</span></th>
                            <th class="sortable text-center" style="width: 60px;" onclick="sortTable('texts-table', 1)">Seite <span class="sort-arrow">▼</span></th>
                            <th class="sortable" style="width: 180px;" onclick="sortTable('texts-table', 2)">Quellseite <span class="sort-arrow">▼</span></th>
                            <th class="sortable" style="width: 180px;" onclick="sortTable('texts-table', 3)">Rubrik / Menü <span class="sort-arrow">▼</span></th>
                            <th class="sortable" style="width: 180px;" onclick="sortTable('texts-table', 4)">Abschnitt <span class="sort-arrow">▼</span></th>
                            <th class="sortable text-center" style="width: 120px;" onclick="sortTable('texts-table', 5)">Inhaltstyp <span class="sort-arrow">▼</span></th>
                            <th>Textinhalt</th>
                            <th class="sortable text-right" style="width: 90px;" onclick="sortTable('texts-table', 7)">Wörter <span class="sort-arrow">▼</span></th>
                        </tr>
                    </thead>
                    <tbody>
    """)

    for t in all_texts:
        tt = t['type']
        badge_cls = "badge-lvl0" if "Überschrift" in tt else ("badge-lvl2" if "Zitat" in tt else "badge-type")
        parts.append(f"""
                        <tr data-type="{html.escape(tt)}">
                            <td class="text-center"><span class="badge-id">{html.escape(t['text_id'])}</span></td>
                            <td class="text-center">{t['page_idx']}</td>
                            <td><a href="{rel_site_prefix}{t['file_name']}" class="table-link" target="_blank">{html.escape(t['file_name'])} ↗</a></td>
                            <td>{html.escape(t['menu_title'])}</td>
                            <td><small>{html.escape(t['section'])}</small></td>
                            <td class="text-center"><span class="badge {badge_cls}">{html.escape(tt)}</span></td>
                            <td style="white-space: pre-line; line-height: 1.4;">{html.escape(t['text'])}</td>
                            <td class="text-right">{t['word_count']}</td>
                        </tr>
        """)

    parts.append(f"""
                    </tbody>
                </table>
            </div>
        </section>

        <!-- TAB 4: Kunstwerke & Bildkatalog -->
        <section id="tab-artworks" class="tab-pane">
            <div class="toolbar">
                <div class="toolbar-group">
                    <input type="text" id="art-search" class="search-input" placeholder="🔍 Werke filtern (Titel, Technik, Maße, Jahr, ID...)" oninput="filterArtworks()">
                    <select id="art-tech-filter" class="filter-select" onchange="filterArtworks()">
                        <option value="">Alle Techniken</option>
    """)

    for tech in all_techniques:
        parts.append(f"""<option value="{html.escape(tech)}">{html.escape(tech)}</option>""")

    parts.append(f"""
                    </select>
                    <select id="art-exist-filter" class="filter-select" onchange="filterArtworks()">
                        <option value="">Alle Status</option>
                        <option value="Ja">Lokal vorhanden (Ja)</option>
                        <option value="Nein">Fehlt lokal (Nein)</option>
                    </select>
                </div>
                <div class="toolbar-group">
                    <button class="btn btn-active" id="btn-view-table" onclick="setArtworkView('table')">📋 Tabellenansicht</button>
                    <button class="btn" id="btn-view-grid" onclick="setArtworkView('grid')">🖼️ Galerieansicht</button>
                    <div class="count-indicator" id="art-count">Zeige {len(all_artworks)} von {len(all_artworks)} Werke</div>
                </div>
            </div>

            <!-- Table View -->
            <div class="table-wrap" id="art-table-wrap">
                <table class="excel-table" id="art-table">
                    <thead>
                        <tr>
                            <th class="sortable text-center" style="width: 75px;" onclick="sortTable('art-table', 0)">ID <span class="sort-arrow">▼</span></th>
                            <th class="text-center" style="width: 60px;">Bild</th>
                            <th class="sortable text-center" style="width: 50px;" onclick="sortTable('art-table', 2)">S. <span class="sort-arrow">▼</span></th>
                            <th class="sortable" style="width: 160px;" onclick="sortTable('art-table', 3)">Rubrik / Menü <span class="sort-arrow">▼</span></th>
                            <th class="sortable" style="width: 220px;" onclick="sortTable('art-table', 4)">Werkname / Titel <span class="sort-arrow">▼</span></th>
                            <th class="sortable" style="width: 160px;" onclick="sortTable('art-table', 5)">Technik / Material <span class="sort-arrow">▼</span></th>
                            <th class="sortable" style="width: 120px;" onclick="sortTable('art-table', 6)">Maße <span class="sort-arrow">▼</span></th>
                            <th class="sortable text-center" style="width: 80px;" onclick="sortTable('art-table', 7)">Jahr <span class="sort-arrow">▼</span></th>
                            <th class="sortable text-center" style="width: 70px;" onclick="sortTable('art-table', 8)">Werk-Nr. <span class="sort-arrow">▼</span></th>
                            <th class="sortable text-center" style="width: 110px;" onclick="sortTable('art-table', 9)">Bildtyp <span class="sort-arrow">▼</span></th>
                            <th class="sortable text-center" style="width: 70px;" onclick="sortTable('art-table', 10)">Lokal <span class="sort-arrow">▼</span></th>
                            <th class="sortable text-right" style="width: 75px;" onclick="sortTable('art-table', 11)">KB <span class="sort-arrow">▼</span></th>
                        </tr>
                    </thead>
                    <tbody>
    """)

    for a in all_artworks:
        exists = a['exists_local'] == 'Ja'
        badge_exists = "badge-yes" if exists else "badge-no"
        img_src_rel = f"{rel_site_prefix}{a['thumb_path']}" if a['thumb_path'] else ""
        img_full_rel = f"{rel_site_prefix}{a['full_path']}" if a['full_path'] else img_src_rel
        img_abs = f"{abs_site_prefix}{a['thumb_path']}" if a['thumb_path'] else ""
        
        safe_title = html.escape(a['title'] if a['title'] else a['raw_caption'])
        safe_tech = html.escape(a['technique'])
        safe_dims = html.escape(a['dimensions'])
        safe_year = html.escape(a['year'])
        safe_id = html.escape(a['artwork_id'])
        safe_num = html.escape(a['number'])
        safe_rubrik = html.escape(a['menu_title'])
        safe_raw = html.escape(a['raw_caption'])
        safe_file = html.escape(a['file_name'])

        parts.append(f"""
                        <tr data-tech="{html.escape(a['technique'])}" data-exists="{a['exists_local']}">
                            <td class="text-center"><span class="badge-id">{safe_id}</span></td>
                            <td class="text-center">
                                <div class="thumb-wrap" onclick="openLightbox('{img_full_rel}', '{safe_id}', '{safe_title}', '{safe_tech}', '{safe_dims}', '{safe_year}', '{safe_num}', '{safe_rubrik}', '{safe_file}', '{a['file_size_kb']}')">
                                    <img src="{img_src_rel}" data-fallback="{img_abs}" class="thumb-img" loading="lazy" alt="{safe_title}" onerror="handleImgError(this)">
                                </div>
                            </td>
                            <td class="text-center">{a['page_idx']}</td>
                            <td>{safe_rubrik}</td>
                            <td>
                                <strong>{safe_title}</strong>
                                {f'<div style="font-size: 11px; color: #718096; margin-top: 2px;">Orig: {safe_raw}</div>' if safe_raw and safe_raw != safe_title else ''}
                            </td>
                            <td>{f'<span class="badge badge-tech">{safe_tech}</span>' if safe_tech else '-'}</td>
                            <td>{safe_dims or '-'}</td>
                            <td class="text-center">{safe_year or '-'}</td>
                            <td class="text-center"><strong>{safe_num or '-'}</strong></td>
                            <td class="text-center"><small>{html.escape(a['image_type'])}</small></td>
                            <td class="text-center"><span class="badge {badge_exists}">{a['exists_local']}</span></td>
                            <td class="text-right">{a['file_size_kb']}</td>
                        </tr>
        """)

    parts.append("""
                    </tbody>
                </table>
            </div>

            <!-- Grid View (Hidden initially) -->
            <div class="gallery-grid" id="art-grid-wrap" style="display: none;">
    """)

    for a in all_artworks:
        exists = a['exists_local'] == 'Ja'
        badge_exists = "badge-yes" if exists else "badge-no"
        img_src_rel = f"{rel_site_prefix}{a['thumb_path']}" if a['thumb_path'] else ""
        img_full_rel = f"{rel_site_prefix}{a['full_path']}" if a['full_path'] else img_src_rel
        img_abs = f"{abs_site_prefix}{a['thumb_path']}" if a['thumb_path'] else ""

        safe_title = html.escape(a['title'] if a['title'] else a['raw_caption'])
        safe_tech = html.escape(a['technique'])
        safe_dims = html.escape(a['dimensions'])
        safe_year = html.escape(a['year'])
        safe_id = html.escape(a['artwork_id'])
        safe_num = html.escape(a['number'])
        safe_rubrik = html.escape(a['menu_title'])
        safe_raw = html.escape(a['raw_caption'])
        safe_file = html.escape(a['file_name'])

        parts.append(f"""
                <div class="gallery-card" data-tech="{html.escape(a['technique'])}" data-exists="{a['exists_local']}">
                    <div class="gallery-card-img-wrap" onclick="openLightbox('{img_full_rel}', '{safe_id}', '{safe_title}', '{safe_tech}', '{safe_dims}', '{safe_year}', '{safe_num}', '{safe_rubrik}', '{safe_file}', '{a['file_size_kb']}')">
                        <span class="gallery-card-badge">{safe_id}</span>
                        <img src="{img_src_rel}" data-fallback="{img_abs}" class="gallery-card-img" loading="lazy" alt="{safe_title}" onerror="handleImgError(this)">
                    </div>
                    <div class="gallery-card-body">
                        <div>
                            <div class="gallery-card-title">{safe_title}</div>
                            <div class="gallery-card-meta">{safe_tech}</div>
                            <div class="gallery-card-meta">{f'{safe_dims} • ' if safe_dims else ''}{safe_year}</div>
                        </div>
                        <div class="gallery-card-footer">
                            <span class="badge {badge_exists}">Lokal: {a['exists_local']}</span>
                            <span style="color: #718096;">{safe_rubrik}</span>
                        </div>
                    </div>
                </div>
        """)

    parts.append(f"""
            </div>
        </section>

        <!-- TAB 5: Medien & Verlinkungen -->
        <section id="tab-links" class="tab-pane">
            <div class="toolbar">
                <div class="toolbar-group">
                    <input type="text" id="links-search" class="search-input" placeholder="🔍 Verlinkungen filtern (URL, Text, Rubrik...)" oninput="filterLinksTable()">
                    <select id="links-type-filter" class="filter-select" onchange="filterLinksTable()">
                        <option value="">Alle Linktypen</option>
    """)

    for lt in all_link_types:
        parts.append(f"""<option value="{html.escape(lt)}">{html.escape(lt)}</option>""")

    parts.append(f"""
                    </select>
                </div>
                <div class="count-indicator" id="links-count">Zeige {len(all_links)} von {len(all_links)} Verlinkungen</div>
            </div>

            <div class="table-wrap">
                <table class="excel-table" id="links-table">
                    <thead>
                        <tr>
                            <th class="sortable text-center" style="width: 80px;" onclick="sortTable('links-table', 0)">Link-ID <span class="sort-arrow">▼</span></th>
                            <th class="sortable text-center" style="width: 60px;" onclick="sortTable('links-table', 1)">Seite <span class="sort-arrow">▼</span></th>
                            <th class="sortable" style="width: 200px;" onclick="sortTable('links-table', 2)">Quellseite <span class="sort-arrow">▼</span></th>
                            <th class="sortable" style="width: 200px;" onclick="sortTable('links-table', 3)">Rubrik <span class="sort-arrow">▼</span></th>
                            <th class="sortable text-center" style="width: 150px;" onclick="sortTable('links-table', 4)">Linktyp / Medientyp <span class="sort-arrow">▼</span></th>
                            <th class="sortable" style="width: 250px;" onclick="sortTable('links-table', 5)">Link-Beschriftung (Text) <span class="sort-arrow">▼</span></th>
                            <th>Ziel-URL / Dateipfad</th>
                            <th class="sortable text-center" style="width: 110px;" onclick="sortTable('links-table', 7)">Lokal <span class="sort-arrow">▼</span></th>
                        </tr>
                    </thead>
                    <tbody>
    """)

    for l in all_links:
        t_cls = "badge-lvl0" if "Intern" in l['type'] else ("badge-lvl2" if "Audio" in l['type'] or "Dokument" in l['type'] else "badge-type")
        stat_cls = "badge-yes" if l['exists_local'] == 'Ja' else ("badge-no" if l['exists_local'] == 'Nein' else "badge-neutral")
        
        target_href = l['target_url']
        if not (target_href.startswith('http://') or target_href.startswith('https://') or target_href.startswith('mailto:')):
            target_href = f"{rel_site_prefix}{target_href}"

        parts.append(f"""
                        <tr data-type="{html.escape(l['type'])}">
                            <td class="text-center"><span class="badge-id">{html.escape(l['link_id'])}</span></td>
                            <td class="text-center">{l['page_idx']}</td>
                            <td><a href="{rel_site_prefix}{l['file_name']}" class="table-link" target="_blank">{html.escape(l['file_name'])} ↗</a></td>
                            <td>{html.escape(l['menu_title'])}</td>
                            <td class="text-center"><span class="badge {t_cls}">{html.escape(l['type'])}</span></td>
                            <td>{html.escape(l['text'])}</td>
                            <td><a href="{html.escape(target_href)}" class="table-link" target="_blank" style="word-break: break-all;">{html.escape(l['target_url'])} ↗</a></td>
                            <td class="text-center"><span class="badge {stat_cls}">{html.escape(l['exists_local'])}</span></td>
                        </tr>
        """)

    parts.append("""
                    </tbody>
                </table>
            </div>
        </section>

    </main>

    <!-- Modal Lightbox -->
    <div class="modal" id="lightbox-modal" onclick="if(event.target===this) closeLightbox()">
        <div class="modal-content">
            <div class="modal-header">
                <h3 class="modal-title" id="modal-title">Werkdetails</h3>
                <button class="modal-close" onclick="closeLightbox()">&times;</button>
            </div>
            <div class="modal-body">
                <div class="modal-img-area">
                    <img src="" id="modal-image" class="modal-img" alt="Kunstwerk">
                </div>
                <div class="modal-details">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                        <span class="badge-id" id="modal-id" style="font-size: 14px; padding: 4px 10px;">W-0000</span>
                        <a href="" id="modal-img-link" target="_blank" class="table-link">Originaldatei öffnen ↗</a>
                    </div>
                    <h2 id="modal-work-title" style="margin: 0 0 16px 0; font-size: 18px; color: #1F497D;">Titel</h2>
                    
                    <table class="modal-details-table">
                        <tr><td>Technik:</td><td id="modal-tech">-</td></tr>
                        <tr><td>Abmessungen:</td><td id="modal-dims">-</td></tr>
                        <tr><td>Entstehungsjahr:</td><td id="modal-year">-</td></tr>
                        <tr><td>Werk-Nr.:</td><td id="modal-num">-</td></tr>
                        <tr><td>Rubrik / Seite:</td><td id="modal-rubrik">-</td></tr>
                        <tr><td>Quellseite:</td><td id="modal-file">-</td></tr>
                        <tr><td>Dateigröße:</td><td id="modal-size">-</td></tr>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <script>
        // Tab switching
        function switchTab(tabId, btn) {
            document.querySelectorAll('.tab-pane').forEach(el => el.classList.remove('active'));
            document.querySelectorAll('.tab-button').forEach(el => el.classList.remove('active'));
            
            const target = document.getElementById('tab-' + tabId);
            if (target) target.classList.add('active');
            if (btn) btn.classList.add('active');
            
            window.location.hash = tabId;
        }

        // Handle URL Hash
        window.addEventListener('DOMContentLoaded', () => {
            const hash = window.location.hash.replace('#', '');
            if (hash) {
                const btn = document.querySelector(`button[onclick*="'${hash}'"]`);
                if (btn) switchTab(hash, btn);
            }
        });

        // Image error fallback
        function handleImgError(img) {
            if (!img.dataset.tried) {
                img.dataset.tried = '1';
                img.src = img.dataset.fallback;
            } else {
                img.style.display = 'none';
                if (img.parentElement) {
                    img.parentElement.innerHTML = '<span class="thumb-placeholder">🖼️</span>';
                }
            }
        }

        // Lightbox
        function openLightbox(imgSrc, id, title, tech, dims, year, num, rubrik, file, size) {
            document.getElementById('modal-image').src = imgSrc;
            document.getElementById('modal-id').textContent = id;
            document.getElementById('modal-title').textContent = id + ' – ' + title;
            document.getElementById('modal-work-title').textContent = title;
            document.getElementById('modal-tech').textContent = tech || '-';
            document.getElementById('modal-dims').textContent = dims || '-';
            document.getElementById('modal-year').textContent = year || '-';
            document.getElementById('modal-num').textContent = num || '-';
            document.getElementById('modal-rubrik').textContent = rubrik || '-';
            document.getElementById('modal-file').innerHTML = `<a href="436167332711627763-1789638831/7897076666aabb50cb36a1/${file}" target="_blank" class="table-link">${file} ↗</a>`;
            document.getElementById('modal-size').textContent = size ? size + ' KB' : '-';
            document.getElementById('modal-img-link').href = imgSrc;

            document.getElementById('lightbox-modal').classList.add('active');
        }

        function closeLightbox() {
            document.getElementById('lightbox-modal').classList.remove('active');
        }

        window.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') closeLightbox();
        });

        // Artwork View Toggle (Table vs Grid)
        function setArtworkView(mode) {
            const tableWrap = document.getElementById('art-table-wrap');
            const gridWrap = document.getElementById('art-grid-wrap');
            const btnTable = document.getElementById('btn-view-table');
            const btnGrid = document.getElementById('btn-view-grid');

            if (mode === 'grid') {
                tableWrap.style.display = 'none';
                gridWrap.style.display = 'grid';
                btnTable.classList.remove('btn-active');
                btnGrid.classList.add('btn-active');
            } else {
                tableWrap.style.display = 'block';
                gridWrap.style.display = 'none';
                btnTable.classList.add('btn-active');
                btnGrid.classList.remove('btn-active');
            }
        }

        // Filtering Functions
        function filterPagesTable() {
            const query = document.getElementById('pages-search').value.toLowerCase().trim();
            const level = document.getElementById('pages-level-filter').value;
            const rows = document.querySelectorAll('#pages-table tbody tr');
            let visible = 0;

            rows.forEach(row => {
                const rowLevel = row.getAttribute('data-level');
                const rowText = row.textContent.toLowerCase();
                const matchQuery = !query || rowText.includes(query);
                const matchLevel = !level || rowLevel === level;

                if (matchQuery && matchLevel) {
                    row.style.display = '';
                    visible++;
                } else {
                    row.style.display = 'none';
                }
            });

            document.getElementById('pages-count').textContent = `Zeige ${visible} von ${rows.length} Seiten`;
        }

        function filterTextsTable() {
            const query = document.getElementById('texts-search').value.toLowerCase().trim();
            const type = document.getElementById('texts-type-filter').value;
            const rows = document.querySelectorAll('#texts-table tbody tr');
            let visible = 0;

            rows.forEach(row => {
                const rowType = row.getAttribute('data-type');
                const rowText = row.textContent.toLowerCase();
                const matchQuery = !query || rowText.includes(query);
                const matchType = !type || rowType === type;

                if (matchQuery && matchType) {
                    row.style.display = '';
                    visible++;
                } else {
                    row.style.display = 'none';
                }
            });

            document.getElementById('texts-count').textContent = `Zeige ${visible} von ${rows.length} Texte`;
        }

        function filterArtworks() {
            const query = document.getElementById('art-search').value.toLowerCase().trim();
            const tech = document.getElementById('art-tech-filter').value;
            const exist = document.getElementById('art-exist-filter').value;

            // Table Rows
            const tableRows = document.querySelectorAll('#art-table tbody tr');
            let visible = 0;

            tableRows.forEach(row => {
                const rowTech = row.getAttribute('data-tech');
                const rowExist = row.getAttribute('data-exists');
                const rowText = row.textContent.toLowerCase();
                const matchQuery = !query || rowText.includes(query);
                const matchTech = !tech || rowTech === tech;
                const matchExist = !exist || rowExist === exist;

                if (matchQuery && matchTech && matchExist) {
                    row.style.display = '';
                    visible++;
                } else {
                    row.style.display = 'none';
                }
            });

            // Grid Cards
            const gridCards = document.querySelectorAll('#art-grid-wrap .gallery-card');
            gridCards.forEach(card => {
                const cardTech = card.getAttribute('data-tech');
                const cardExist = card.getAttribute('data-exists');
                const cardText = card.textContent.toLowerCase();
                const matchQuery = !query || cardText.includes(query);
                const matchTech = !tech || cardTech === tech;
                const matchExist = !exist || cardExist === exist;

                if (matchQuery && matchTech && matchExist) {
                    card.style.display = 'flex';
                } else {
                    card.style.display = 'none';
                }
            });

            document.getElementById('art-count').textContent = `Zeige ${visible} von ${tableRows.length} Werke`;
        }

        function filterLinksTable() {
            const query = document.getElementById('links-search').value.toLowerCase().trim();
            const type = document.getElementById('links-type-filter').value;
            const rows = document.querySelectorAll('#links-table tbody tr');
            let visible = 0;

            rows.forEach(row => {
                const rowType = row.getAttribute('data-type');
                const rowText = row.textContent.toLowerCase();
                const matchQuery = !query || rowText.includes(query);
                const matchType = !type || rowType === type;

                if (matchQuery && matchType) {
                    row.style.display = '';
                    visible++;
                } else {
                    row.style.display = 'none';
                }
            });

            document.getElementById('links-count').textContent = `Zeige ${visible} von ${rows.length} Verlinkungen`;
        }

        // Table Column Sorting
        const sortDirections = {};
        function sortTable(tableId, colIndex) {
            const table = document.getElementById(tableId);
            const tbody = table.querySelector('tbody');
            const rows = Array.from(tbody.querySelectorAll('tr'));
            
            const key = tableId + '_' + colIndex;
            const asc = !sortDirections[key];
            sortDirections[key] = asc;

            // Reset arrow indicators
            table.querySelectorAll('th .sort-arrow').forEach(a => a.textContent = '▼');
            const currentTh = table.querySelectorAll('th')[colIndex];
            if (currentTh && currentTh.querySelector('.sort-arrow')) {
                currentTh.querySelector('.sort-arrow').textContent = asc ? '▲' : '▼';
            }

            rows.sort((a, b) => {
                const aCell = a.cells[colIndex] ? a.cells[colIndex].textContent.trim() : '';
                const bCell = b.cells[colIndex] ? b.cells[colIndex].textContent.trim() : '';
                
                const aNum = parseFloat(aCell.replace(/[^0-9.-]/g, ''));
                const bNum = parseFloat(bCell.replace(/[^0-9.-]/g, ''));

                if (!isNaN(aNum) && !isNaN(bNum) && /^[0-9.,\\s-]+$/.test(aCell) && /^[0-9.,\\s-]+$/.test(bCell)) {
                    return asc ? aNum - bNum : bNum - aNum;
                }
                return asc ? aCell.localeCompare(bCell, 'de', { numeric: true }) : bCell.localeCompare(aCell, 'de', { numeric: true });
            });

            rows.forEach(row => tbody.appendChild(row));
        }
    </script>
</body>
</html>
    """)

    doc_html = "".join(parts)

    # Save HTML files
    try:
        os.makedirs(os.path.dirname(OUTPUT_HTML_PRIMARY), exist_ok=True)
        with open(OUTPUT_HTML_PRIMARY, 'w', encoding='utf-8') as f:
            f.write(doc_html)
        print(f"HTML documentation successfully saved to: {OUTPUT_HTML_PRIMARY}")
    except Exception as e:
        print(f"Hinweis beim Speichern von {OUTPUT_HTML_PRIMARY}: {e}")

    try:
        with open(OUTPUT_HTML_BASEDIR, 'w', encoding='utf-8') as f:
            f.write(doc_html)
        print(f"HTML documentation copy saved to base dir: {OUTPUT_HTML_BASEDIR}")
    except Exception as e:
        print(f"Hinweis beim Speichern von {OUTPUT_HTML_BASEDIR}: {e}")

    try:
        with open(OUTPUT_HTML_LOCAL, 'w', encoding='utf-8') as f:
            f.write(doc_html)
        print(f"HTML documentation copy saved to local project: {OUTPUT_HTML_LOCAL}")
    except Exception as e:
        print(f"Hinweis beim Speichern von {OUTPUT_HTML_LOCAL}: {e}")

def main():
    pages_data, all_texts, all_artworks, all_links, kpis, pages_meta = extract_website_data()
    generate_excel_documentation(pages_data, all_texts, all_artworks, all_links, kpis, pages_meta)
    generate_html_documentation(pages_data, all_texts, all_artworks, all_links, kpis, pages_meta)
    print("\nDocumentation generation complete!")

if __name__ == '__main__':
    main()
