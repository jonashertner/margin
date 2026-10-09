"""Run with bundled Python; fixtures are fictional and created by QA agent."""
import copy, hashlib, importlib.util, json, os, tempfile, unittest, zipfile
from pathlib import Path
from lxml import etree as ET
SPEC=importlib.util.spec_from_file_location('word_engine',Path(__file__).resolve().parents[1]/'lib'/'word_engine.py')
e=importlib.util.module_from_spec(SPEC);SPEC.loader.exec_module(e)
FIXTURES=Path(os.environ.get('FOLIO_WORD_FIXTURES',str(Path(__file__).resolve().parent/'fixtures'/'word')))

class WordBridge(unittest.TestCase):
 @classmethod
 def setUpClass(cls):
  cls.work=tempfile.TemporaryDirectory(prefix='folio-word-tests-')
  cls.root=Path(cls.work.name)
  cls.manifest=json.loads((FIXTURES/'expected-manifest.json').read_text())
  cls.native=e.import_word({'inputPath':str(FIXTURES/'03-native-tracked.docx'),'jobDir':str(cls.root)})
 @classmethod
 def tearDownClass(cls):cls.work.cleanup()
 def output(self,name):return str(self.root/name)
 def export(self,before,after,**extra):return e.export_word({'before':before,'after':after,'mode':'tracked','outputPath':self.output(self._testMethodName+'.docx'),'date':'2026-10-09T12:00:00Z',**extra})
 def xmlparts(self,path):
  with zipfile.ZipFile(path) as z:return {n:e.xml(z.read(n)) for n in z.namelist() if n.endswith('.xml')}
 def comments(self):
  return [{**c,'sourceAnchor':c['afterAnchor'],'anchorText':c['anchorTextAfter']} for c in self.native['comments']]
 def test_native_endpoints_wording_clauses_and_footnotes(self):
  expected=self.manifest['files']['03-native-tracked.docx']
  for view,key in [('before','beforeClauseParagraphs'),('after','afterClauseParagraphs')]:
   parsed,_=e.ast(self.native[view],'latex');clauses=[b for b in parsed['blocks'] if b['t']=='OrderedList']
   self.assertEqual(len(clauses),6)
   self.assertEqual([b['c'][0][0] for b in clauses],list(range(1,7)))
   for b,text in zip(clauses,expected[key]):self.assertIn(text,e.norm(e.text_of(b)))
   self.assertIn(self.manifest['sharedFootnotes'][0]['text'],e.text_of(parsed['blocks']))
  self.assertEqual(len(self.native['originalRevisions']),9)
  self.assertEqual(self.native['counts']['revisions'],9)
 def test_comments_authors_dates_and_unchanged_scope_anchor(self):
  for c,expected in zip(self.native['comments'],self.manifest['sharedComments']):
   self.assertEqual(c['text'],expected['text']);self.assertEqual(c['author'],expected['author']);self.assertEqual(c['at'],expected['date'])
   for view in ('before','after'):
    anchor=c[view+'Anchor'];self.assertTrue(anchor)
    self.assertEqual(self.native[view].encode('utf-16-le')[anchor['start']*2:anchor['end']*2].decode('utf-16-le'),anchor['source'])
  self.assertIn('Neither party admits liability.',self.native['comments'][0]['afterAnchor']['source'])
  self.assertEqual(self.native['comments'][1]['anchorTextBefore'],'48,000')
  self.assertEqual(self.native['comments'][1]['anchorTextAfter'],'52,000')
 def test_source_bytes_unchanged(self):
  for name,expected in self.manifest['files'].items():self.assertEqual(hashlib.sha256((FIXTURES/name).read_bytes()).hexdigest(),expected['sha256'])
 def test_clean_export_retains_comments_and_zero_revisions(self):
  result=self.export(self.native['before'],self.native['after'],mode='clean',comments=self.comments());self.assertTrue(result['written'])
  self.assertEqual(result['counts']['commentsPlaced'],2);self.assertEqual(result['counts']['nativeRevisionElements'],0)
  parts=self.xmlparts(result['outputPath']);body=parts['word/document.xml'];self.assertEqual(len(list(body.iter(e.q('numPr')))),6)
  self.assertEqual(len(list(body.iter(e.q('footnoteReference')))),1)
  self.assertEqual([c.get(e.q('author')) for c in parts['word/comments.xml']],['Elias Brunner','Mara Keller'])
 def test_tracked_export_native_unique_revision_ids_and_folio_author(self):
  result=self.export(self.native['before'],self.native['after'],comments=self.comments());self.assertTrue(result['written'])
  parts=self.xmlparts(result['outputPath']);nodes=[n for tree in parts.values() for n in tree.iter() if n.tag in (e.q('ins'),e.q('del'))]
  self.assertGreater(len(nodes),0);ids=[n.get(e.q('id')) for n in nodes];self.assertEqual(len(ids),len(set(ids)))
  self.assertEqual({n.get(e.q('author')) for n in nodes},{'Margin'})
  self.assertTrue(all(n.get(e.q('date'))=='2026-10-09T12:00:00Z' for n in nodes))
  for deletion in parts['word/document.xml'].iter(e.q('del')):self.assertFalse(list(deletion.iter(e.q('t'))))
 def test_paragraph_mark_only_import(self):
  record=e.import_word({'inputPath':str(FIXTURES/'04-paragraph-boundaries.docx'),'jobDir':str(self.root)})
  expected=self.manifest['files']['04-paragraph-boundaries.docx']
  self.assertNotEqual(record['before'],record['after'])
  for view in ('before','after'):
   parsed,_=e.ast(record[view],'latex');texts=[e.norm(e.text_of(b)) for b in parsed['blocks'] if b['t']=='Para']
   self.assertEqual(texts[-3:],[e.norm(x) for x in expected[view+'Paragraphs']])
  self.assertTrue(self.export(record['before'],record['after'])['written'])
 def test_whole_paragraph_insertion_and_deletion_have_mark_revisions(self):
  for before,after,kind in [('Stable.','Stable.\n\nAdded.','ins'),('Stable.\n\nDeleted.','Stable.','del')]:
   result=self.export(before,after);self.assertTrue(result['written']);body=self.xmlparts(result['outputPath'])['word/document.xml']
   self.assertEqual(len(body.findall('.//w:pPr/w:rPr/w:'+kind,e.NS)),1)
 def test_changed_footnote_endpoints(self):
  self.assertTrue(self.export(r'Payment is due.\footnote{Old explanatory note.}',r'Payment is due.\footnote{Revised explanatory note.}')['written'])
 def test_unsupported_structures_are_reported(self):
  record=e.import_word({'inputPath':str(FIXTURES/'05-special-constructs.docx'),'jobDir':str(self.root)})
  codes={w['code'] for w in record['warnings']}
  self.assertTrue({'headers_footers','content_controls','fields','tracked_table_rows','text_boxes'}<=codes)
  self.assertIn('Delivery is at the north loading bay.',record['convertedText'])
 def test_missing_or_ambiguous_comments_block_export(self):
  for anchor,source in [('Missing text.','Present text.'),('Repeated.','Repeated.\n\nRepeated.')]:
   path=self.output('blocked.docx');result=self.export(source,source,mode='clean',comments=[{'text':'Important note','sourceAnchor':anchor}],outputPath=path)
   self.assertFalse(result['written']);self.assertEqual(result['counts']['commentsUnplaced'],1);self.assertFalse(Path(path).exists())
 def test_unsupported_before_and_after_macros_fail_closed(self):
  for before,after in [(r'\unknowncommand{Important old text}','New text.'),('Old text.',r'\unknowncommand{Important new text}')]:
   with self.assertRaisesRegex(ValueError,'unsupported LaTeX'):self.export(before,after)
 def test_metadata_change_blocks_tracked_export(self):
  with self.assertRaisesRegex(ValueError,'metadata'):self.export('\\title{Old}\n\\begin{document}Body.\\end{document}','\\title{New}\n\\begin{document}Body.\\end{document}')
 def test_utf16_anchor_offsets(self):
  doc,_=e.ast('A 😀 passage.\n\nSecond paragraph.','latex');source,anchors=e.to_tex(doc)
  for a in anchors:self.assertEqual(source.encode('utf-16-le')[a['start']*2:a['end']*2].decode('utf-16-le'),a['source'])
 def altered_zip(self,name,modify):
  with zipfile.ZipFile(FIXTURES/'01-baseline.docx') as z:parts={i.filename:z.read(i.filename) for i in z.infolist()}
  modify(parts);path=self.output(name)
  with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
   for n,data in parts.items():z.writestr(n,data)
  return path
 def test_external_resource_is_blocked(self):
  def modify(parts):parts['word/_rels/evil.xml.rels']=b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" TargetMode="External" Target="https://example.invalid/private"/></Relationships>'
  with self.assertRaisesRegex(ValueError,'External Word resources'):e.inspect_docx(self.altered_zip('external.docx',modify))
 def test_xml_entities_are_blocked(self):
  def modify(parts):parts['word/document.xml']=b'<!DOCTYPE x [<!ENTITY x SYSTEM "file:///tmp/canary">]><x>&x;</x>'
  with self.assertRaisesRegex(ValueError,'entities and doctypes'):e.inspect_docx(self.altered_zip('entity.docx',modify))
 def test_archive_traversal_and_macros_are_blocked(self):
  for entry in ('../escape','word/vbaProject.bin'):
   with self.assertRaises(ValueError):e.inspect_docx(self.altered_zip('bad.docx',lambda parts:parts.update({entry:b'not-executed'})))

if __name__=='__main__':unittest.main(verbosity=2)
