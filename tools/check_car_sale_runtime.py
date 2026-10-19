"""No API/Telegram calls: verify the bundled fixed template and PDF export."""
import asyncio
from pathlib import Path
import sys
import tempfile
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from pypdf import PdfReader
from car_sale_template import render
from word_quality import convert_word_pdf

async def main():
    with tempfile.TemporaryDirectory(prefix='car-sale-check-') as td:
        root=Path(td);docx=root/'contract.docx';pdf=root/'contract.pdf'
        render(docx,{'seller_name':'Проверочный Продавец','price':'1 200 000'})
        await convert_word_pdf(docx,pdf)
        result=PdfReader(pdf)
        assert len(result.pages)==2
        text='\n'.join(p.extract_text() for p in result.pages)
        assert 'Проверочный Продавец' in text and '1 200 000' in text
        assert 'Подтверждение исполнения' in text
        assert 'auto.ru' not in text.lower() and 'авто.ру' not in text.lower()
        assert all(not list(p.images) for p in result.pages)
        print('car_sale runtime_check=passed pages=2 branding=none')
if __name__=='__main__':asyncio.run(main())
