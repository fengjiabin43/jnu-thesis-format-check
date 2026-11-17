"""Read-only DOCX formatting evidence. Python 3 standard library only."""
import argparse
import hashlib
import json
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
W = '{' + NS['w'] + '}'
A = '{' + NS['a'] + '}'

def attr(node, name, default=None):
    return node.get(W + name, default) if node is not None else default

def text(node):
    return ''.join(x.text or '' for x in node.iter(W + 't'))

def number(value, divisor=1):
    try:
        return round(float(value) / divisor, 4)
    except (ValueError, TypeError):
        return None

def run_properties(rpr):
    if rpr is None:
        return {}
    font = rpr.find('w:rFonts', NS)
    size = rpr.find('w:sz', NS)
    color = rpr.find('w:color', NS)
    bold = rpr.find('w:b', NS)
    return {'fonts': {k.split('}')[-1]: v for k,v in font.attrib.items()} if font is not None else {},
            'size_pt': number(attr(size, 'val'), 2),
            'bold': attr(bold, 'val', '1') not in ('0','false','off') if bold is not None else None,
            'color': {k.split('}')[-1]:v for k,v in color.attrib.items()} if color is not None else {}}

def audit(path):
    path = Path(path)
    with ZipFile(path) as z:
        damaged = z.testzip()
        if damaged:
            raise ValueError('Damaged DOCX part: ' + damaged)
        parts = {n:ET.fromstring(z.read(n)) for n in z.namelist()
                 if n.startswith('word/') and n.endswith('.xml')}
        doc = parts['word/document.xml']
        style_root = parts.get('word/styles.xml')
        styles = {}
        if style_root is not None:
            for s in style_root.findall('w:style', NS):
                sid = attr(s, 'styleId')
                ppr = s.find('w:pPr', NS)
                styles[sid] = {'name':attr(s.find('w:name',NS),'val'),
                    'based_on':attr(s.find('w:basedOn',NS),'val'),
                    'linked_style':attr(s.find('w:link',NS),'val'),
                    'outline_level':attr(ppr.find('w:outlineLvl',NS),'val') if ppr is not None else None,
                    'run_properties':run_properties(s.find('w:rPr',NS))}
        nonblack, themes, automatic, drawing, errors, headings, captions, sections = [],[],[],[],[],[],[],[]
        for name, root in parts.items():
            for node in root.iter(W+'color'):
                info = {'part':name,'attributes':{k.split('}')[-1]:v for k,v in node.attrib.items()}}
                value = attr(node,'val','').upper()
                if value == 'AUTO':
                    automatic.append(info)
                elif value != '000000':
                    nonblack.append(info)
                if any(attr(node,key) is not None for key in ('themeColor','themeTint','themeShade')):
                    themes.append(info)
            for tag in ('rPr','defRPr','endParaRPr'):
                for rp in root.iter(A+tag):
                    fill = rp.find('a:solidFill',NS)
                    if fill is not None:
                        for c in fill:
                            if c.tag != A+'srgbClr' or c.get('val','').upper()!='000000':
                                drawing.append({'part':name,'kind':c.tag.split('}')[-1],'value':c.get('val')})
            for node in root.iter(W+'t'):
                if re.search(r'Error!|错误!|错误！|未定义书签|找不到引用源',node.text or '',re.I):
                    errors.append({'part':name,'text':node.text})
        for paragraph in doc.iter(W+'p'):
            ppr = paragraph.find('w:pPr',NS)
            sid = attr(ppr.find('w:pStyle',NS),'val') if ppr is not None else None
            style = styles.get(sid,{})
            outline = attr(ppr.find('w:outlineLvl',NS),'val') if ppr is not None else None
            level = outline if outline is not None else style.get('outline_level')
            title = text(paragraph)
            if level is not None and level != '9':
                headings.append({'text':title,'level':int(level)+1,'style_id':sid})
            if re.match(r'^\s*[图表]\s*\d',title):
                captions.append({'text':title,'style_id':sid})
        for section in doc.iter(W+'sectPr'):
            size = section.find('w:pgSz',NS)
            margin = section.find('w:pgMar',NS)
            page_num = section.find('w:pgNumType',NS)
            sections.append({'width_cm':number(attr(size,'w'),1440/2.54),
                'height_cm':number(attr(size,'h'),1440/2.54),
                'margins_cm':{key:number(attr(margin,key),1440/2.54) for key in
                    ('top','bottom','left','right','header','footer','gutter')},
                'page_number_format':attr(page_num,'fmt'),
                'page_number_start':attr(page_num,'start')})
        headers = {name:[text(p) for p in root.iter(W+'p') if text(p)]
                   for name,root in parts.items() if re.match(r'word/(header|footer)\d+\.xml$',name)}
        media = {n:hashlib.sha256(z.read(n)).hexdigest() for n in z.namelist() if n.startswith('word/media/')}
        instructions = [n.text or '' for n in doc.iter(W+'instrText')]
        instructions.extend(attr(n,'instr','') for n in doc.iter(W+'fldSimple'))
        return {'file':str(path.resolve()),'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),
            'counts':{'paragraphs':len(list(doc.iter(W+'p'))),'tables':len(list(doc.iter(W+'tbl'))),
                'media':len(media),'sections':len(sections)},
            'nonblack_font_colors':nonblack,'theme_font_overrides':themes,
            'automatic_font_colors':automatic,'drawing_text_colors_to_check':drawing,
            'field_errors':errors,'styles':styles,'headings':headings,'captions':captions,
            'sections':sections,'headers_and_footers':headers,
            'toc_fields':[i for i in instructions if re.search(r'\bTOC\b',i)],
            'media_sha256':media,
            'visual_checks':['实际分页','图题同页','续表题与重复列标题','图片内文字颜色','缺字与截断'],
            'supplementary_checks':['封面字段','摘要字数和关键词数量','参考文献排序及字段','实体打印安排']}

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('docx')
    parser.add_argument('--out',type=Path)
    args = parser.parse_args()
    result = audit(args.docx)
    content = json.dumps(result,ensure_ascii=False,indent=2)
    if args.out:
        args.out.write_text(content,encoding='utf-8')
        print(json.dumps({'report':str(args.out),'counts':result['counts'],
            'nonblack_font_colors':len(result['nonblack_font_colors']),
            'theme_font_overrides':len(result['theme_font_overrides']),
            'automatic_font_colors':len(result['automatic_font_colors']),
            'drawing_text_colors_to_check':len(result['drawing_text_colors_to_check']),
            'field_errors':len(result['field_errors'])},ensure_ascii=False))
    else:
        print(content)
