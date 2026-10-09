#!/usr/bin/env python3
"""Margin's bounded DOCX bridge. All inputs/outputs must be in a caller-owned job.

CLI: word_engine.py import|export REQUEST.json
Uses the trusted MARGIN_PANDOC/FOLIO_PANDOC executable, or pandoc on PATH.
The caller must provide OS filesystem/network isolation for this entire process.
"""
import copy, datetime, difflib, hashlib, io, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path
from lxml import etree as ET

PANDOC = os.environ.get('MARGIN_PANDOC') or os.environ.get('FOLIO_PANDOC') or shutil.which('pandoc') or 'pandoc'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS={'w':W}; q=lambda name:'{'+W+'}'+name
MAX_ZIP=15_000_000; MAX_EXPANDED=80_000_000; MAX_PART=20_000_000
NOW=lambda:datetime.datetime.now(datetime.timezone.utc).isoformat(timespec='seconds').replace('+00:00','Z')
PREAMBLE=r'''\documentclass[11pt,a4paper]{article}
\usepackage[margin=28mm]{geometry}
\usepackage{fontspec}
\setmainfont{texgyrepagella}[Extension=.otf,UprightFont=*-regular,BoldFont=*-bold,ItalicFont=*-italic,BoldItalicFont=*-bolditalic]
\usepackage{parskip}
\setcounter{secnumdepth}{0}
\usepackage{longtable,booktabs,array,calc}
\IfFileExists{multirow.sty}{\usepackage{multirow}}{}
\usepackage[hidelinks]{hyperref}
\providecommand{\tightlist}{\setlength{\itemsep}{0pt}\setlength{\parskip}{0pt}}
\providecommand{\pandocbounded}[1]{#1}
\hypersetup{pdfauthor={},pdftitle={},pdfsubject={},pdfkeywords={}}
'''

def pd(args, data=None):
    args=[PANDOC,'--sandbox',*args]
    p=subprocess.run(args,input=data,stdout=subprocess.PIPE,stderr=subprocess.PIPE,timeout=45,check=False)
    if p.returncode: raise ValueError('Word conversion failed: '+p.stderr.decode('utf8','replace')[:1600])
    return p.stdout,p.stderr.decode('utf8','replace')

def ast(source,fmt):
    if fmt=='latex':
        # Pandoc's reader sees setcounter only before a label redefinition.
        source=re.sub(r'(\\begin\{enumerate\}\s*)(\\def\\labelenum[^\n]*\n)(\\setcounter\{enum[^}]*\}\{\d+\}\s*)',r'\1\3\2',source)
    out,log=pd(['-f','latex+raw_tex' if fmt=='latex' else fmt,'-t','json'],source.encode('utf8'))
    doc=json.loads(out)
    if fmt=='latex' and doc.get('meta',{}).get('title'):
        doc['blocks']=[b for b in doc['blocks'] if not (b.get('t')=='RawBlock' and b.get('c')==['latex',r'\maketitle'])]
    return doc,log

def write_ast(doc,fmt,extra=()):
    return pd(['-f','json','-t',fmt,*extra],json.dumps(doc,ensure_ascii=False).encode('utf8'))

def warn(warnings,code,message,count=1):
    for item in warnings:
        if item['code']==code: item['count']+=count;return
    warnings.append({'code':code,'message':message,'count':count})

def norm(text): return re.sub(r'\s+',' ',text).strip()

def span_class(node):
    return node.get('c',[['',[],[]]])[0][1] if isinstance(node,dict) and node.get('t')=='Span' else []

def text_of(node,view='after'):
    if isinstance(node,list):return ''.join(text_of(n,view) for n in node)
    if not isinstance(node,dict):return ''
    typ=node.get('t');c=node.get('c')
    if typ in ('Str','Code'):return c if typ=='Str' else c[1]
    if typ in ('Space','SoftBreak','LineBreak'):return ' '
    if typ=='Span':
        classes=span_class(node)
        if 'comment-start' in classes or 'comment-end' in classes:return ''
        if ('insertion' in classes and view=='before') or ('deletion' in classes and view=='after'):return ''
        return text_of(c[1],view)
    if typ in ('Para','Plain'):return text_of(c,view)+'\n'
    if typ=='Header':return text_of(c[2],view)+'\n'
    if typ=='OrderedList':return '\n'.join(text_of(x,view) for x in c[1])
    if typ=='BulletList':return '\n'.join(text_of(x,view) for x in c)
    if typ=='Note':return ' [Footnote: '+norm(text_of(c,view))+'] '
    if typ=='Div':return text_of(c[1],view)
    if typ in ('Link','Image'):return text_of(c[1],view)
    if typ=='Quoted':return text_of(c[1],view)
    if typ in ('RawInline','RawBlock'):return ''
    if isinstance(c,list):return text_of(c,view)
    return ''

def walk(node):
    if isinstance(node,dict):
        yield node
        for value in node.values():yield from walk(value)
    elif isinstance(node,list):
        for value in node:yield from walk(value)

def xml(data):
    if b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper():raise ValueError('DOCX XML entities and doctypes are not supported.')
    return ET.fromstring(data,ET.XMLParser(resolve_entities=False,no_network=True,remove_blank_text=False))

def settle_docx(source,destination,view,warnings=None):
    """Resolve native revisions in a disposable copy, including paragraph marks."""
    with zipfile.ZipFile(source) as z:parts={i.filename:z.read(i.filename) for i in z.infolist()}
    for filename,data in list(parts.items()):
        if not re.match(r'word/(document|footnotes|endnotes|header\d+|footer\d+)\.xml$',filename):continue
        tree=xml(data);merges={}
        for p in tree.iter(q('p')):
            rp=p.find('w:pPr/w:rPr',NS)
            merges[p]=rp is not None and rp.find('w:'+('ins' if view=='before' else 'del'),NS) is not None
        def resolve(parent):
            for child in list(parent):
                name=child.tag.rsplit('}',1)[-1]
                if name in ('ins','del','moveFrom','moveTo'):
                    keep=(name in ('ins','moveTo'))==(view=='after')
                    if keep:
                        resolve(child);at=parent.index(child)
                        for item in list(child):parent.insert(at,item);at+=1
                    parent.remove(child)
                elif name=='delText':child.tag=q('t')
                elif name.endswith('PrChange'):parent.remove(child)
                else:resolve(child)
        resolve(tree)
        parents=[p for p in tree.iter() if any(x.tag==q('p') for x in p)]
        for parent in parents:
            i=0
            while i<len(parent):
                p=parent[i]
                if p.tag!=q('p') or not merges.get(p):i+=1;continue
                visible=any(x.tag in (q('t'),q('footnoteReference'),q('drawing'),q('object')) and (x.tag!=q('t') or x.text) for x in p.iter())
                if not visible:parent.remove(p);continue
                if i+1<len(parent) and parent[i+1].tag==q('p'):
                    following=parent[i+1]
                    for item in list(following):
                        if item.tag!=q('pPr'):p.append(item)
                    parent.remove(following);merges[p]=merges.get(following,False)
                    if not merges[p]:i+=1
                else:
                    if warnings is not None:warn(warnings,'paragraph_boundary','A tracked paragraph boundary next to a non-paragraph requires manual checking.')
                    i+=1
        parts[filename]=ET.tostring(tree,encoding='UTF-8',xml_declaration=True)
    with zipfile.ZipFile(destination,'w',zipfile.ZIP_DEFLATED) as z:
        for name,data in parts.items():z.writestr(name,data)

def xml_text(element,view='after'):
    name=element.tag.rsplit('}',1)[-1]
    if (view=='before' and name in ('ins','moveTo')) or (view=='after' and name in ('del','moveFrom')):return ''
    if name in ('t','delText'):return element.text or ''
    if name in ('tab','br','cr'):return ' '
    if name=='footnoteReference':return ' [Footnote '+element.get(q('id'),'')+'] '
    return ''.join(xml_text(x,view) for x in element)

def inspect_docx(path):
    if Path(path).stat().st_size>MAX_ZIP:raise ValueError('The Word file exceeds the 15 MB import limit.')
    warnings=[];comments=[];revisions=[]
    with zipfile.ZipFile(path) as z:
        files=z.infolist()
        if len(files)>2000 or sum(f.file_size for f in files)>MAX_EXPANDED:raise ValueError('The expanded Word file is too large.')
        if len({f.filename for f in files})!=len(files):raise ValueError('The Word file has duplicate ZIP entries.')
        for f in files:
            if f.file_size>MAX_PART or (f.file_size>1_000_000 and f.file_size/max(1,f.compress_size)>500):raise ValueError('The Word file has an unsafe compressed part.')
            if f.filename.startswith('/') or '..' in Path(f.filename).parts or f.flag_bits&1:raise ValueError('Unsupported Word archive entry.')
            if f.filename.lower().endswith(('vbaproject.bin','.exe','.js')):raise ValueError('Macros or executable content are not accepted.')
        names={f.filename for f in files}
        if 'word/document.xml' not in names:raise ValueError('This is not a supported DOCX document.')
        doc=xml(z.read('word/document.xml'))
        for filename in names:
            if filename.endswith('.rels'):
                for r in xml(z.read(filename)):
                    if r.get('TargetMode')=='External':
                        typ=r.get('Type','').rsplit('/',1)[-1]
                        if typ!='hyperlink':raise ValueError('External Word resources are not accepted ('+typ+').')
                        warn(warnings,'external_hyperlink','External hyperlinks are retained as text links; no resource is fetched.')
        tags=[x.tag.rsplit('}',1)[-1] for x in doc.iter()]
        reports={
          'drawing':('images','Images are not embedded in this standalone LaTeX conversion.'),
          'pict':('drawings','Drawings and text boxes may not be represented in the converted body.'),
          'txbxContent':('text_boxes','Text boxes require manual checking against the original Word file.'),
          'sdt':('content_controls','Content control values become ordinary text; controls are not retained.'),
          'fldChar':('fields','Word fields become their cached visible text; field updating is not retained.'),
          'fldSimple':('fields','Word fields become their cached visible text; field updating is not retained.'),
          'altChunk':('alt_chunks','Embedded alternative document content is unsupported.'),
          'object':('embedded_objects','Embedded objects are unsupported.'),
          'moveFrom':('tracked_moves','Tracked moves are represented as textual changes, not native moves.'),
          'pPrChange':('format_revisions','Tracked paragraph formatting changes are not reviewable in Margin.'),
          'rPrChange':('format_revisions','Tracked text formatting changes are not reviewable in Margin.'),
          'tblPrChange':('table_revisions','Tracked table formatting changes are unsupported.'),
        }
        for tag,(code,message) in reports.items():
            n=tags.count(tag)
            if n:warn(warnings,code,message,n)
        tracked_rows=sum(1 for p in doc.iter(q('trPr')) if p.find('w:ins',NS) is not None or p.find('w:del',NS) is not None)
        if tracked_rows:warn(warnings,'tracked_table_rows','Tracked table-row insertions/deletions need manual review; conversion may not preserve their decision semantics.',tracked_rows)
        for filename in sorted(names):
            if re.match(r'word/(header|footer)\d+\.xml$',filename):
                if norm(xml_text(xml(z.read(filename)))):warn(warnings,'headers_footers','Headers and footers are not carried into the clean legal layout.')
        if 'word/comments.xml' in names:
            for c in xml(z.read('word/comments.xml')).findall('w:comment',NS):
                comments.append({'id':c.get(q('id')),'author':c.get(q('author'),'Unknown reviewer'),'at':c.get(q('date'),'') or None,'initials':c.get(q('initials'),''),'text':'\n'.join(xml_text(p) for p in c.findall('w:p',NS))})
        if 'word/commentsExtended.xml' in names:warn(warnings,'comment_threads','Reply/resolution metadata is retained only in the original Word file; comment text is imported.')
        for part in ('word/document.xml','word/footnotes.xml','word/endnotes.xml'):
            if part not in names:continue
            tree=doc if part=='word/document.xml' else xml(z.read(part))
            for item in tree.iter():
                typ=item.tag.rsplit('}',1)[-1]
                if typ in ('ins','del','moveFrom','moveTo'):
                    revisions.append({'kind':typ,'id':item.get(q('id')),'author':item.get(q('author'),'Unknown reviewer'),'at':item.get(q('date')),'text':''.join(item.itertext()),'part':part})
        body=doc.find('w:body',NS)
        source={view:'\n'.join(xml_text(p,view) for p in body.iter(q('p'))) for view in ('before','after')}
        counts={'comments':len(comments),'revisions':len(revisions),'footnotes':0,'numberedParagraphs':len(doc.findall('.//w:numPr',NS)),'tables':tags.count('tbl')}
        if 'word/footnotes.xml' in names:counts['footnotes']=sum(1 for p in xml(z.read('word/footnotes.xml')).findall('w:footnote',NS) if int(p.get(q('id'),'0'))>0)
    warn(warnings,'restyled_layout','Imported wording uses Margin’s clean legal layout. Original Word styling, pagination and section settings are retained in the original file.')
    return {'comments':comments,'revisions':revisions,'warnings':warnings,'counts':counts,'source':source}

def transform_images(node,warnings):
    if isinstance(node,list):return [transform_images(x,warnings) for x in node]
    if not isinstance(node,dict):return node
    if node.get('t')=='Image':
        return {'t':'Str','c':'[Image: '+(norm(text_of(node)) or 'see original Word document')+']'}
    return {k:transform_images(v,warnings) for k,v in node.items()}

def tex_escape(text):
    return ''.join({'\\':'\\textbackslash{}','%':'\\%','&':'\\&','#':'\\#','$':'\\$','_':'\\_','{':'\\{','}':'\\}','^':'\\textasciicircum{}','~':'\\textasciitilde{}'}.get(c,c) for c in text)

def to_tex(doc):
    blocks=[]
    for i,b in enumerate(doc['blocks']):
        blocks.extend([{'t':'RawBlock','c':['latex',f'%FOLIOBLOCK{i}START']},b,{'t':'RawBlock','c':['latex',f'%FOLIOBLOCK{i}END']}])
    converted,_=write_ast({**doc,'blocks':blocks},'latex',['--wrap=none'])
    converted=converted.decode('utf8');fragments=[]
    for i in range(len(doc['blocks'])):
        m=re.search(r'%FOLIOBLOCK'+str(i)+r'START\s*\n([\s\S]*?)\n%FOLIOBLOCK'+str(i)+r'END',converted)
        if not m:raise ValueError('Could not identify the converted Word paragraph boundaries.')
        fragment=m.group(1).strip('\n')
        fragment=re.sub(r'(\\begin\{enumerate\}\s*)(\\def\\labelenum[^\n]*\n)(\\setcounter\{enum[^}]*\}\{\d+\}\s*)',r'\1\3\2',fragment)
        fragments.append(fragment)
    title=text_of(doc.get('meta',{}).get('title',{}).get('c',[])).strip()
    source=PREAMBLE
    if title:source+='\\title{'+tex_escape(title)+'}\n\\author{}\n\\date{}\n'
    source+='\\begin{document}\n'
    if title:source+='\\maketitle\n'
    source+='\n';anchors=[]
    for i,fragment in enumerate(fragments):
        start=len(source.encode('utf-16-le'))//2;source+=fragment+'\n\n';anchors.append({'start':start,'end':len(source.encode('utf-16-le'))//2,'source':fragment+'\n\n','blockIndex':i,'text':norm(text_of(doc['blocks'][i]))})
    source+='\\end{document}\n'
    return source,anchors

def inline_groups(node):
    if isinstance(node,list):
        for value in node:yield from inline_groups(value)
    elif isinstance(node,dict):
        typ=node.get('t')
        if typ in ('Para','Plain'):yield node['c']
        elif typ=='Header':yield node['c'][2]
        else:
            for value in node.values():yield from inline_groups(value)

def split_numbered_clauses(doc):
    result=copy.deepcopy(doc);blocks=[]
    for block in result['blocks']:
        if block['t']!='OrderedList':blocks.append(block);continue
        attrs,items=block['c']
        for offset,item in enumerate(items):blocks.append({'t':'OrderedList','c':[[attrs[0]+offset,*copy.deepcopy(attrs[1:])],[item]]})
    result['blocks']=blocks
    return result

def import_word(req):
    path=req['inputPath'];record=inspect_docx(path);docs={}
    with tempfile.TemporaryDirectory(prefix='settled-',dir=req.get('jobDir') or tempfile.gettempdir()) as work:
        for view,mode in [('before','reject'),('after','accept'),('all','all')]:
            incoming=path
            if view!='all':
                incoming=str(Path(work)/(view+'.docx'));settle_docx(path,incoming,view,record['warnings'])
            data,log=pd(['-f','docx','--track-changes='+mode,'-t','json',incoming]);docs[view]=split_numbered_clauses(json.loads(data))
            if log.strip():warn(record['warnings'],'pandoc_reader',log.strip())
    before,banchors=to_tex(transform_images(docs['before'],record['warnings']))
    after,aanchors=to_tex(transform_images(docs['after'],record['warnings']))
    if record['counts']['numberedParagraphs']:warn(record['warnings'],'clause_numbering','Numbered clauses are independently reviewable and retain their displayed start numbers. Check numbering continuity after mixed review decisions.')
    mappings={}
    for view,anchors in [('before',banchors),('after',aanchors)]:
        cursor=0;mapping={}
        for i,block in enumerate(docs['all']['blocks']):
            key=norm(text_of(block,view))
            if not key:continue
            candidates=[j for j in range(cursor,len(anchors)) if anchors[j]['text']==key]
            if candidates:
                j=candidates[0];mapping[i]=anchors[j];cursor=j+1
        mappings[view]=mapping
    byid={c['id']:c for c in record['comments']};revisions=[]
    for block_index,block in enumerate(docs['all']['blocks']):
        for n in walk(block):
            if n.get('t')!='Span':continue
            attrs,content=n['c'];classes=attrs[1];kv=dict(attrs[2]);cid=kv.get('comment-id',kv.get('id',attrs[0]))
            if 'comment-start' in classes and cid in byid:
                c=byid[cid];c['beforeAnchor']=mappings['before'].get(block_index);c['afterAnchor']=mappings['after'].get(block_index)
                # Preserve the exact range's wording separately from the safe block anchor.
                for group in inline_groups(block):
                    recording=False;range_nodes=[]
                    for item in group:
                        if item is n:recording=True;continue
                        if recording and item.get('t')=='Span' and 'comment-end' in span_class(item) and dict(item['c'][0][2]).get('id',item['c'][0][0])==cid:break
                        if recording:range_nodes.append(item)
                    if recording:
                        c['anchorTextBefore']=norm(text_of(range_nodes,'before'));c['anchorTextAfter']=norm(text_of(range_nodes,'after'));break
            for cls in ('insertion','deletion','paragraph-insertion','paragraph-deletion'):
                if cls in classes:revisions.append({'kind':cls,'author':kv.get('author','Unknown reviewer'),'at':kv.get('date'),'text':text_of(content),'beforeAnchor':mappings['before'].get(block_index),'afterAnchor':mappings['after'].get(block_index)})
    for c in record['comments']:
        if not c.get('beforeAnchor') and not c.get('afterAnchor'):warn(record['warnings'],'unanchored_comment','A comment could not be mapped to a unique converted passage; its text and author are retained.')
    title=norm(text_of(docs['after'].get('meta',{}).get('title',{}).get('c',[]))) or req.get('title') or Path(path).stem
    return {'before':before,'after':after,'title':title,'comments':record['comments'],'revisions':revisions,'originalRevisions':record['revisions'],'warnings':record['warnings'],'counts':record['counts'],'sourceText':record['source']['after'],'beforeSourceText':record['source']['before'],'convertedText':text_of(docs['after']['blocks']),'beforeConvertedText':text_of(docs['before']['blocks']),'blocks':{'before':banchors,'after':aanchors}}

def change_span(kind,content,date):return {'t':'Span','c':[['',[kind],[['author','Margin'],['date',date]]],content]}
def key(node):return json.dumps(node,sort_keys=True,ensure_ascii=False,separators=(',',':'))
def tokenise(inlines):
    out=[]
    for item in inlines:
        if item.get('t')=='Str':out.extend({'t':'Str','c':part} for part in re.findall(r'\w+|[^\w]',item['c'],re.UNICODE))
        else:out.append(copy.deepcopy(item))
    return out

def diff_inlines(before,after,date):
    # A substantially rewritten paragraph reads better as two coherent runs.
    if before and after and difflib.SequenceMatcher(None,norm(text_of(before)),norm(text_of(after)),autojunk=False).ratio()<0.5:
        return [change_span('deletion',copy.deepcopy(before),date),change_span('insertion',copy.deepcopy(after),date)]
    before=tokenise(before);after=tokenise(after);out=[]
    for op,a,b,c,d in difflib.SequenceMatcher(None,[key(x) for x in before],[key(x) for x in after],autojunk=False).get_opcodes():
        if op=='equal':out.extend(copy.deepcopy(after[c:d]));continue
        if a<b:out.append(change_span('deletion',copy.deepcopy(before[a:b]),date))
        if c<d:out.append(change_span('insertion',copy.deepcopy(after[c:d]),date))
    return out

def whole_block(block,kind,date):
    result=copy.deepcopy(block);typ=result['t']
    if typ in ('Para','Plain'):result['c']=[change_span(kind,result['c'],date)]
    elif typ=='Header':result['c'][2]=[change_span(kind,result['c'][2],date)]
    elif typ in ('OrderedList','BulletList'):
        items=result['c'][1] if typ=='OrderedList' else result['c']
        for i,item in enumerate(items):items[i]=[whole_block(x,kind,date) for x in item]
    elif typ=='BlockQuote':result['c']=[whole_block(x,kind,date) for x in result['c']]
    else:raise ValueError('Native tracked export does not support adding/removing '+typ+' structures. Export a clean Word document or revise the structure in Word.')
    return result

def diff_pair(before,after,date):
    typ=after['t']
    if before['t']!=typ:return [whole_block(before,'deletion',date),whole_block(after,'insertion',date)]
    result=copy.deepcopy(after)
    if typ in ('Para','Plain'):result['c']=diff_inlines(before['c'],after['c'],date)
    elif typ=='Header':
        if before['c'][0]!=after['c'][0]:raise ValueError('Native tracked export cannot represent a changed heading level. Export clean Word or change the heading structure in Word.')
        result['c'][2]=diff_inlines(before['c'][2],after['c'][2],date)
    elif typ in ('OrderedList','BulletList'):
        if typ=='OrderedList' and before['c'][0]!=after['c'][0]:raise ValueError('Native tracked export cannot safely represent a change to numbering style/start. Export clean Word or review numbering in Word.')
        old=before['c'][1] if typ=='OrderedList' else before['c'];new=after['c'][1] if typ=='OrderedList' else after['c']
        items=[]
        for op,a,b,c,d in difflib.SequenceMatcher(None,[key(x) for x in old],[key(x) for x in new],autojunk=False).get_opcodes():
            if op=='equal':items.extend(copy.deepcopy(new[c:d]))
            elif op=='replace' and b-a==d-c:items.extend(diff_blocks(x,y,date) for x,y in zip(old[a:b],new[c:d]))
            else:
                items.extend([[whole_block(x,'deletion',date) for x in item] for item in old[a:b]])
                items.extend([[whole_block(x,'insertion',date) for x in item] for item in new[c:d]])
        if typ=='OrderedList':result['c'][1]=items
        else:result['c']=items
    elif typ=='BlockQuote':result['c']=diff_blocks(before['c'],after['c'],date)
    else:raise ValueError('Native tracked export cannot safely represent changes inside '+typ+'. Export clean Word or make those changes in Word.')
    return [result]

def diff_blocks(before,after,date):
    out=[]
    for op,a,b,c,d in difflib.SequenceMatcher(None,[key(x) for x in before],[key(x) for x in after],autojunk=False).get_opcodes():
        if op=='equal':out.extend(copy.deepcopy(after[c:d]))
        elif op=='replace' and b-a==d-c:
            for old,new in zip(before[a:b],after[c:d]):out.extend(diff_pair(old,new,date))
        else:
            out.extend(whole_block(x,'deletion',date) for x in before[a:b]);out.extend(whole_block(x,'insertion',date) for x in after[c:d])
    return out

def add_comments(doc,comments,warnings):
    groups=list(inline_groups(doc['blocks']));unplaced=[];placed=0
    for index,comment in enumerate(comments):
        anchor=comment.get('sourceAnchor','');anchor=anchor.get('source','') if isinstance(anchor,dict) else anchor
        anchor_text=comment.get('anchorText') or comment.get('anchorTextAfter')
        context=''
        if anchor:
            parsed,_=ast(anchor,'latex');context=norm(text_of(parsed['blocks']))
            if not anchor_text:anchor_text=context
        target=norm(anchor_text or '')
        matches=[g for g in groups if target and target in norm(text_of(g,'after')) and (not context or norm(text_of(g,'after')) in context or context in norm(text_of(g,'after')))]
        if len(matches)!=1:
            # Whole balanced source blocks (e.g. an enumerate environment) can span
            # several paragraphs. Require exactly one contiguous matching range.
            matches=[]
            if target:
                for a in range(len(groups)):
                    joined=''
                    for b in range(a,len(groups)):
                        joined=norm(joined+' '+text_of(groups[b],'after'))
                        if joined==target:matches.append((groups[a],groups[b]));break
                        if len(joined)>len(target):break
            if len(matches)!=1:unplaced.append({'id':comment.get('id',str(index)),'text':comment.get('text',''),'reason':'The anchor is absent or ambiguous in the reviewed wording.'});continue
            first,last=matches[0]
        else:first=last=matches[0]
        cid=str(index);attrs=[['id',cid],['author',comment.get('author') or 'Margin'],['date',comment.get('at') or NOW()]]
        start={'t':'Span','c':[['',['comment-start'],attrs],[{'t':'Str','c':comment.get('text','')}]]}
        end={'t':'Span','c':[['',['comment-end'],[['id',cid]]],[]]}
        first.insert(0,start);last.append(end);placed+=1
        if norm(text_of(first,'after'))!=target:warn(warnings,'comment_paragraph_anchor','A comment is attached to its containing paragraph rather than its original character range.')
    return placed,unplaced

def style_and_revision_marks(path,tracked):
    with zipfile.ZipFile(path) as z:parts={i.filename:z.read(i.filename) for i in z.infolist()}
    doc=xml(parts['word/document.xml'])
    if tracked:
        for p in doc.iter(q('p')):
            children=[x for x in p if x.tag!=q('pPr') and x.tag not in (q('bookmarkStart'),q('bookmarkEnd'),q('commentRangeStart'),q('commentRangeEnd')) and not (x.tag==q('r') and x.find('w:commentReference',NS) is not None)]
            if children and all(x.tag==q('ins') for x in children):kind='ins'
            elif children and all(x.tag==q('del') for x in children):kind='del'
            else:continue
            pp=p.find('w:pPr',NS)
            if pp is None:pp=ET.Element(q('pPr'));p.insert(0,pp)
            rp=pp.find('w:rPr',NS)
            if rp is None:rp=ET.SubElement(pp,q('rPr'))
            ET.SubElement(rp,q(kind),{q('id'):children[0].get(q('id'),'0'),q('author'):'Margin',q('date'):children[0].get(q('date'),NOW())})
        settings=xml(parts['word/settings.xml'])
        if settings.find('w:trackRevisions',NS) is None:ET.SubElement(settings,q('trackRevisions'))
        parts['word/settings.xml']=ET.tostring(settings,encoding='UTF-8',xml_declaration=True)
    for sect in doc.iter(q('sectPr')):
        size=sect.find('w:pgSz',NS)
        if size is None:size=ET.SubElement(sect,q('pgSz'))
        size.set(q('w'),'11906');size.set(q('h'),'16838')
        margin=sect.find('w:pgMar',NS)
        if margin is None:margin=ET.SubElement(sect,q('pgMar'))
        for side in ('top','right','bottom','left'):margin.set(q(side),'1587')
    parts['word/document.xml']=ET.tostring(doc,encoding='UTF-8',xml_declaration=True)
    # Pandoc maintains separate insertion/deletion counters; Word IDs must be
    # unique across these types, including paragraph marks and note stories.
    revision_id=0
    for filename in ('word/document.xml','word/footnotes.xml','word/endnotes.xml'):
        if filename not in parts:continue
        tree=xml(parts[filename])
        for element in tree.iter():
            if element.tag in (q('ins'),q('del'),q('moveFrom'),q('moveTo')):
                element.set(q('id'),str(revision_id));revision_id+=1
        parts[filename]=ET.tostring(tree,encoding='UTF-8',xml_declaration=True)
    styles=xml(parts['word/styles.xml'])
    for style in styles.iter(q('style')):
        if style.get(q('type'))!='paragraph':continue
        rp=style.find('w:rPr',NS)
        if rp is None:rp=ET.SubElement(style,q('rPr'))
        fonts=rp.find('w:rFonts',NS)
        if fonts is None:fonts=ET.SubElement(rp,q('rFonts'))
        for att in ('ascii','hAnsi','cs'):fonts.set(q(att),'Georgia')
        colour=rp.find('w:color',NS)
        if colour is None:colour=ET.SubElement(rp,q('color'))
        colour.attrib.clear();colour.set(q('val'),'000000')
    parts['word/styles.xml']=ET.tostring(styles,encoding='UTF-8',xml_declaration=True)
    # Preserve visible title content but clear outgoing document properties.
    for filename in ('docProps/core.xml','docProps/custom.xml'):
        if filename in parts:
            tree=xml(parts[filename])
            for child in list(tree):tree.remove(child)
            parts[filename]=ET.tostring(tree,encoding='UTF-8',xml_declaration=True)
    if 'docProps/app.xml' in parts:
        tree=xml(parts['docProps/app.xml'])
        for child in list(tree):
            if child.tag.rsplit('}',1)[-1] in ('Company','Manager','HyperlinkBase','Template'):tree.remove(child)
        parts['docProps/app.xml']=ET.tostring(tree,encoding='UTF-8',xml_declaration=True)
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
        for name,data in parts.items():z.writestr(name,data)

def semantic(doc):
    # Ignore writer-specific heading IDs and style attributes; retain text,
    # paragraph/list boundaries, numbering starts/styles and note content.
    result=[]
    for b in doc['blocks']:
        t=b['t']
        if t in ('Para','Plain'):result.append(('paragraph',norm(text_of(b))))
        elif t=='Header':result.append(('heading',b['c'][0],norm(text_of(b))))
        elif t in ('OrderedList','BulletList'):
            items=b['c'][1] if t=='OrderedList' else b['c']
            attrs=copy.deepcopy(b['c'][0]) if t=='OrderedList' else None
            if attrs:
                if attrs[1]['t']=='DefaultStyle':attrs[1]={'t':'Decimal'}
                if attrs[2]['t']=='DefaultDelim':attrs[2]={'t':'Period'}
            result.append((t,attrs,[semantic({'blocks':item}) for item in items]))
        else:result.append((t,norm(text_of(b))))
    return result

def export_word(req):
    mode=req.get('mode','clean')
    if mode not in ('clean','tracked'):raise ValueError('Unknown Word export mode.')
    warnings=[];after,log=ast(req['after'],'latex');before,before_log=ast(req.get('before',req['after']),'latex')
    unsupported=[n for n in walk([after['blocks'],before['blocks'] if mode=='tracked' else []]) if n.get('t') in ('RawInline','RawBlock','Image')]
    if unsupported:raise ValueError('Word export would omit unsupported LaTeX commands or images. Export the source/PDF or simplify those structures first.')
    if log.strip() or (mode=='tracked' and before_log.strip()):raise ValueError('Word export needs a source conversion check: '+(log or before_log).strip()[:1200])
    doc=copy.deepcopy(after)
    if mode=='tracked':
        if before.get('meta')!=after.get('meta'):raise ValueError('Native tracked export cannot represent changes in document metadata. Export clean Word or keep the title metadata unchanged.')
        doc['blocks']=diff_blocks(before['blocks'],after['blocks'],req.get('date') or NOW())
        warn(warnings,'generated_comparison','Tracked changes are a new Margin comparison against the original. They are attributed to Margin; original reviewer authors/dates remain in the saved review.')
    placed,unplaced=add_comments(doc,req.get('comments',[]),warnings)
    if unplaced:return {'written':False,'unplacedComments':unplaced,'warnings':warnings,'counts':{'commentsPlaced':placed,'commentsUnplaced':len(unplaced)}}
    output=req['outputPath'];data,write_log=write_ast(doc,'docx');Path(output).write_bytes(data);style_and_revision_marks(output,mode=='tracked')
    # Endpoints must survive the actual DOCX writer and our OOXML paragraph marks.
    for label,expected,track in [('after',after,'accept')]+([('before',before,'reject')] if mode=='tracked' else []):
        actual,_=pd(['-f','docx','--track-changes='+track,'-t','json',output]);actual=json.loads(actual)
        if semantic(actual)!=semantic(expected):
            Path(output).unlink(missing_ok=True)
            raise ValueError('The generated Word '+label+' endpoint did not preserve paragraph/numbering/footnote semantics. Use clean Word or PDF for this document.')
    with zipfile.ZipFile(output) as z:
        body=xml(z.read('word/document.xml'));revision_count=len(list(body.iter(q('ins'))))+len(list(body.iter(q('del'))))
        comment_count=len(xml(z.read('word/comments.xml')).findall('w:comment',NS)) if 'word/comments.xml' in z.namelist() else 0
    if comment_count!=placed:Path(output).unlink(missing_ok=True);raise ValueError('Not all review comments survived Word export.')
    warn(warnings,'restyled_layout','Word export uses Margin’s clean legal layout; LaTeX pagination and custom typography are not reproduced exactly.')
    if write_log.strip():warn(warnings,'pandoc_writer',write_log.strip())
    return {'written':True,'outputPath':output,'warnings':warnings,'counts':{'commentsPlaced':placed,'commentsUnplaced':0,'nativeRevisionElements':revision_count},'unplacedComments':[]}

if __name__=='__main__':
    try:
        mode=sys.argv[1];request=json.loads(Path(sys.argv[2]).read_text('utf8'))
        result=import_word(request) if mode=='import' else export_word(request) if mode=='export' else None
        if result is None:raise ValueError('Unknown operation.')
        print(json.dumps(result,ensure_ascii=False))
    except Exception as error:
        print(json.dumps({'error':str(error)},ensure_ascii=False));sys.exit(1)
