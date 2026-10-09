"""Remove outgoing document metadata and automatic PDF actions from a derivative."""
import sys
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject
reader = PdfReader(sys.argv[1])
writer = PdfWriter()
writer.clone_document_from_reader(reader)
writer.metadata = None
for key in ('/Metadata', '/OpenAction', '/AA'):
    writer._root_object.pop(NameObject(key), None)
names = writer._root_object.get('/Names')
if names:
    names = names.get_object()
    for key in ('/JavaScript', '/EmbeddedFiles'):
        names.pop(NameObject(key), None)
for page in writer.pages:
    page.pop(NameObject('/AA'), None)
    for ref in page.get('/Annots', []):
        annotation = ref.get_object()
        annotation.pop(NameObject('/AA'), None)
        action = annotation.get('/A')
        if action and action.get_object().get('/S') in ('/JavaScript', '/Launch', '/SubmitForm', '/ImportData'):
            annotation.pop(NameObject('/A'), None)
with open(sys.argv[2], 'wb') as stream:
    writer.write(stream)
