"""Structural pagination policy shared by both DOCX validators."""
from docx.enum.text import WD_ALIGN_PARAGRAPH
from report_plan import outline_level


def effective(paragraph, attribute):
    value = getattr(paragraph.paragraph_format, attribute)
    style = paragraph.style
    while value is None and style is not None:
        value = getattr(style.paragraph_format, attribute)
        style = style.base_style
    return WD_ALIGN_PARAGRAPH.LEFT if attribute == 'alignment' and value is None else value


def pagination_issues(doc, begin):
    issues = []
    paragraphs = doc.paragraphs
    for index in range(begin, len(paragraphs)):
        paragraph = paragraphs[index]
        level = outline_level(paragraph)
        previous = paragraphs[index - 1] if index else None
        following = paragraphs[index + 1] if index + 1 < len(paragraphs) else None
        label = f'Paragraph {index}: {paragraph.text[:70]}'
        def issue(code, message):
            issues.append({'code': code, 'message': label + ': ' + message})
        breaks = paragraph._p.xpath('.//w:br[@w:type="page"]')
        if breaks and (len(breaks) != 1 or level in (0, 1)
                       or following is None or outline_level(following) != 0):
            issue('body-page-break', 'Explicit breaks are allowed only at a level-1 boundary.')
        if level in (0, 1):
            if not effective(paragraph, 'keep_with_next') or not effective(paragraph, 'keep_together'):
                issue('heading-pagination', 'Keep the entire heading with following text.')
        if level == 0:
            boundary = previous is not None and bool(previous._p.xpath(
                './/w:br[@w:type="page"] | ./w:pPr/w:sectPr[not(w:type) or w:type/@w:val="nextPage"]'))
            if not effective(paragraph, 'page_break_before') and not boundary:
                issue('section-page-start', 'A level-1 heading must start a new page.')
            if boundary and effective(paragraph, 'page_break_before'):
                issue('duplicate-section-break', 'Do not combine an explicit break with page_break_before.')
        if level == 1:
            if effective(paragraph, 'page_break_before'):
                issue('subsection-page-break', 'Subsections continue on the current page.')
            if previous is None or outline_level(previous) != 0:
                if previous is None or previous.text.strip() or previous._p.xpath('.//w:drawing | .//w:br'):
                    issue('subsection-gap', 'Exactly one empty paragraph must precede a subsection.')
                elif index > begin + 1 and not paragraphs[index - 2].text.strip():
                    issue('subsection-gap', 'More than one empty paragraph precedes the subsection.')
                elif index > begin + 1 and outline_level(paragraphs[index - 2]) == 0:
                    issue('section-subsection-gap', 'No empty paragraph between a section and its first subsection.')
    return issues
