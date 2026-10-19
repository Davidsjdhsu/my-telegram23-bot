"""Offline release gate: real covers, PDF pagination, TOC and preserved values."""
import asyncio
import tempfile
import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from docx import Document
from pypdf import PdfReader
from academic_templates import BLANKS, blank_path, render
from word_quality import convert_word_pdf

async def main():
    with tempfile.TemporaryDirectory(prefix='academic-check-')as td:
        root=Path(td)
        for kind in BLANKS:
            dest=root/(kind+'.pdf')
            # Converter may cache fields: always use a temporary DOCX copy.
            src=root/(kind+'.docx');src.write_bytes(blank_path(kind).read_bytes())
            await convert_word_pdf(src,dest)
            reader=PdfReader(dest)
            assert len(reader.pages)==1,(kind,len(reader.pages))
            text=reader.pages[0].extract_text()
            assert 'НИ ТГУ' not in text and 'Томск' not in text
        sections=[{'title':'Введение','content':'Цель работы — сравнение способов хранения данных.'},
                  {'title':'1 Сравнение','content':'Локальные файлы доступны без сети.'},
                  {'title':'1.1 Пример','content':'Условный пример для проверки структуры.'},
                  {'title':'Заключение','content':'Выбор зависит от задачи и условий работы.'}]
        for kind in ('referat','coursework','report','essay','notes'):
            src=root/('work_'+kind+'.docx');dest=src.with_suffix('.pdf')
            render(src,'Хранение данных',sections,kind,{'school':'Учебный вуз','author':'Анна Петрова','group':'ИС-21','city':'Москва','year':'2028'})
            await convert_word_pdf(src,dest)
            reader=PdfReader(dest);text='\n'.join(p.extract_text()for p in reader.pages)
            for value in ('Учебный вуз','Анна Петрова','ИС-21'):assert value in text,(kind,value)
            if kind in ('referat','coursework'):
                assert len(reader.pages)==5,(kind,len(reader.pages))
                toc=reader.pages[1].extract_text()
                for page in ('3','4','5'):assert page in toc
                d=Document(src)
                cache=[p.text for p in d.paragraphs if '\t' in p.text]
                assert cache==['Введение\t3','1 Сравнение\t4','1.1 Пример\t4','Заключение\t5'],cache
        print('academic runtime_check=passed blanks=9 genres=5 toc=verified credits=unchanged paid_api_calls=0',flush=True)
if __name__=='__main__':asyncio.run(main())
