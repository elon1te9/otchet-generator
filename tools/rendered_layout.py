"""Measure exported PDF content without counting whitespace or pending figures.

This is a layout gate, not proof of prose quality or screenshot authenticity.
All coordinates are PDF points, relative to the top-left of the page.
"""
from report_plan import normalized

PT_PER_MM = 72 / 25.4


def merged_height(intervals, top, bottom):
    intervals = sorted((max(top, a), min(bottom, b)) for a, b in intervals if b > top and a < bottom)
    total = 0.0
    end = top
    for a, b in intervals:
        total += max(0, b - max(a, end))
        end = max(end, b)
    return total


def content_lines(page, box):
    left, top, right, bottom = box
    chars = [c for c in page.chars if c.get('text', '').strip()
             and left <= c['x0'] < right and top <= c['top']
             and c['bottom'] <= bottom + .1]
    groups = []
    for char in sorted(chars, key=lambda c: (c['top'], c['x0'])):
        if not groups or abs(char['top'] - groups[-1][0]['top']) > 3:
            groups.append([char])
        else:
            groups[-1].append(char)
    lines = []
    for group in groups:
        words = page.within_bbox((left, max(top, min(c['top'] for c in group) - .1),
                                  right, min(bottom, max(c['bottom'] for c in group) + .1))).extract_text() or ''
        if not normalized(words):
            continue
        lines.append({'text': normalized(words), 'top': min(c['top'] for c in group),
                      'bottom': max(c['bottom'] for c in group),
                      'size': max(c.get('size', 14) for c in group)})
    return lines


def heading_location(lines, title):
    title = normalized(title)
    for start in range(len(lines)):
        text = ''
        for end in range(start, min(start + 8, len(lines))):
            text = normalized(text + ' ' + lines[end]['text'])
            if text == title:
                return start, end
            if not title.startswith(text):
                break
    return None


def page_occupancy(page, box, lines):
    left, top, right, bottom = box
    intervals = []
    # Text gets its normal TNR line box (1.15 font height * line spacing),
    # capped at the required size. Gaps and oversized font do not add credit.
    tables = page.find_tables()
    for line in lines:
        in_table = any(t.bbox[1] <= line['top'] < t.bbox[3] for t in tables)
        size = min(line['size'], 12 if in_table else 14)
        pitch = size * 1.15 * (1 if in_table else 1.5)
        middle = (line['top'] + line['bottom']) / 2
        intervals.append((middle - pitch / 2, middle + pitch / 2))
    images = [im for im in page.images if im['x1'] > left and im['x0'] < right
              and im['bottom'] > top and im['top'] < bottom]
    intervals.extend((im['top'], im['bottom']) for im in images)
    occupied = merged_height(intervals, top, bottom)
    return {'occupied_height_pt': round(occupied, 2), 'work_height_pt': round(bottom - top, 2),
            'occupancy': round(occupied / (bottom - top), 4),
            'image_count': len(images), 'table_count': len(tables)}


def inspect_rendered(pdf, headings):
    import pdfplumber
    issues, pages, locations, line_cache = [], [], {}, []
    with pdfplumber.open(pdf) as document:
        for number, page in enumerate(document.pages, 1):
            box = (30 * PT_PER_MM, 20 * PT_PER_MM,
                   page.width - 10 * PT_PER_MM, page.height - 20 * PT_PER_MM)
            lines = content_lines(page, box)
            line_cache.append(lines)
            pages.append({'page': number, 'last_line': lines[-1]['text'] if lines else '',
                          'sentence_end_preference_met': bool(lines and lines[-1]['text'].endswith('.')),
                          'headings': [], **page_occupancy(page, box, lines)})
            for heading in headings:
                position = heading_location(lines, heading['text'])
                if position:
                    locations.setdefault(heading['text'], []).append((number, position))
        # The TOC is before the first body heading. Prefer the final exact match
        # to avoid a plain cached entry without a page number being mistaken for body.
        matches = locations.get(headings[0]['text'], [])
        first_body = matches[-1][0] if matches else None
        if first_body is None:
            issues.append({'code': 'pdf-body-start', 'message': 'Cannot locate the first body heading in PDF.'})
        for page in pages:
            page['body'] = first_body is not None and page['page'] >= first_body
            if page['body'] and page['occupancy'] < .70:
                issues.append({'code': 'page-occupancy', 'message':
                    f"Page {page['page']}: {page['occupancy']:.1%} occupied; minimum 70%, including section ends."})
            if page['body']:
                physical = document.pages[page['page'] - 1]
                lines = line_cache[page['page'] - 1]
                last_bottom = lines[-1]['bottom'] if lines else 0
                if page['last_line'].startswith('Рисунок '):
                    issues.append({'code':'pdf-figure-page-end', 'message':
                                   f"Page {page['page']}: figure caption has no following text."})
                if any(im['top'] < physical.height - 20 * PT_PER_MM
                       and im['bottom'] >= last_bottom for im in physical.images):
                    issues.append({'code':'pdf-figure-page-end', 'message':
                                   f"Page {page['page']}: image has no following text."})
        for heading in headings:
            found = [loc for loc in locations.get(heading['text'], []) if first_body and loc[0] >= first_body]
            if len(found) != 1:
                issues.append({'code': 'pdf-heading', 'message': f"Expected one rendered heading: {heading['text']}; found {len(found)}."})
                continue
            number, (start, end) = found[0]
            page = document.pages[number - 1]
            box = (30 * PT_PER_MM, 20 * PT_PER_MM, page.width - 10 * PT_PER_MM, page.height - 20 * PT_PER_MM)
            lines = line_cache[number - 1]
            pages[number - 1]['headings'].append({'text':heading['text'], 'level':heading['level']})
            if heading['level'] == 1 and (start != 0 or lines[start]['top'] > box[1] + 28):
                issues.append({'code': 'pdf-section-page-start', 'message': f"Page {number}: section does not start at the top: {heading['text']}."})
            # A section immediately followed by a subsection must keep both and
            # the first body line together, not only the two headings.
            next_line = end + 1
            while next_line < len(lines):
                next_position = next((position for h in headings
                                      if (position := heading_location(lines[next_line:], h['text'])) and position[0] == 0), None)
                if next_position is None:
                    break
                next_line += next_position[1] + 1
            if next_line >= len(lines):
                issues.append({'code': 'pdf-orphan-heading', 'message': f"Page {number}: heading has no following text: {heading['text']}."})
    return pages, issues, first_body
