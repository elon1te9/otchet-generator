"""Check that each practice-report subsection ends with its own summary paragraph.

This is a structural gate. A human must still read the result for relevance,
logic and consistency with actual evidence.
"""
import argparse,json,re
from pathlib import Path
from docx import Document
from report_plan import load_plan, body_start, body_headings, outline_level, structure_issues

SUMMARY = re.compile(r'^(?:В результате\b|В итоге\b|Таким образом\b)')
HEADING = re.compile(r'^\d+(?:\.\d+)?\s|^ЗАКЛЮЧЕНИЕ$')

def check(doc, plan=None):
    paragraphs=doc.paragraphs
    begin=body_start(doc, plan)
    if begin is None:
        return {'status':'fail','issues':[{'code':'body-start'}],'conclusions':[]}
    issues=structure_issues(doc, plan) if plan else []
    conclusions=[]
    titles = [h['text'] for h in plan['headings'] if h['level']==2] if plan else [p.text for _,p in body_headings(doc,begin) if outline_level(p)==1]
    for title in titles:
        index=next((i for i,p in enumerate(paragraphs) if i>=begin and p.text.strip()==title),None)
        if index is None:
            issues.append({'code':'missing-subsection','subsection':title});continue
        end=next((i for i,p in body_headings(doc,index+1)),len(paragraphs))
        content=[p.text.strip() for p in paragraphs[index+1:end] if p.text.strip()]
        last=content[-1] if content else ''
        conclusions.append({'subsection':title,'last_paragraph':last})
        if not SUMMARY.match(last):issues.append({'code':'missing-conclusion','subsection':title})
        elif len(last.split())<6 or not last.endswith('.'):
            issues.append({'code':'incomplete-conclusion','subsection':title})
    return {'status':'fail' if issues else 'pass','issues':issues,'conclusions':conclusions,'manual_review_required':'Read each summary for meaning and evidence; introductory words are not sufficient.'}

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('docx',type=Path);parser.add_argument('--plan',required=True,type=Path);parser.add_argument('--json',type=Path);args=parser.parse_args()
    result=check(Document(args.docx),load_plan(args.plan));payload=json.dumps(result,ensure_ascii=False,indent=2)
    if args.json:args.json.parent.mkdir(parents=True,exist_ok=True);args.json.write_text(payload,encoding='utf-8')
    print(payload);raise SystemExit(result['status']!='pass')
