"""Offline pre-deploy gate: fill each bundled form and export it with production PDF code."""
import asyncio,sys,tempfile
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from pypdf import PdfReader
from fixed_word_templates import CATALOG,render
from word_quality import convert_word_pdf
async def main():
 with tempfile.TemporaryDirectory(prefix='fixed-word-check-') as td:
  for kind,spec in CATALOG.items():
   root=Path(td);docx=root/(kind+'.docx');pdf=docx.with_suffix('.pdf')
   render(docx,kind,spec['sample']);await convert_word_pdf(docx,pdf)
   reader=PdfReader(pdf);text=' '.join(' '.join(p.extract_text().split()) for p in reader.pages)
   assert 1<=len(reader.pages)<=2,(kind,len(reader.pages))
   assert all(not list(p.images) for p in reader.pages)
   for value in spec['sample'].values():assert ' '.join(value.rstrip('.').split()) in text,(kind,value)
   for label in spec['fields'].values():assert '['+label+']' not in text
   for brand in ('auto.ru','авто.ру','dogovor.expert','яндекс','yandex'):assert brand not in text.casefold()
   print('fixed_word runtime_check=passed kind='+kind+' pages='+str(len(reader.pages)),flush=True)
if __name__=='__main__':asyncio.run(main())
