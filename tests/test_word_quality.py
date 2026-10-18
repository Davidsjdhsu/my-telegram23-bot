"""Regressions for delivered Word defects and bounded page-aware recovery."""
import asyncio
import copy
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import AsyncMock
from docx import Document
from test_office import office_env
import test_office
from word_quality import (clean_word_meta, merge_word_expansion, prepare_word_files,
                          word_generation_schema, WORD_EDITORIAL_GUIDANCE)


class WordMetadataTests(unittest.TestCase):
    def test_optional_placeholders_removed_but_values_and_legal_fields_preserved(self):
        meta = {'org': '[учебное заведение]', 'author': 'Анна Петрова', 'school': '[название учебного заведения]'}
        self.assertEqual(clean_word_meta(meta, 'doc'), {'org':'', 'author':'Анна Петрова', 'school':''})
        self.assertEqual(clean_word_meta(meta, 'statement'), meta)
        self.assertEqual(meta['org'], '[учебное заведение]')
        self.assertEqual(clean_word_meta({'org':'[Лаборатория Солнце]'}, 'doc')['org'], '[Лаборатория Солнце]')

    def test_generated_schema_has_no_fake_institution_or_date(self):
        schema = '{"org":"Министерство образования","school":"[учебное заведение]","year":"2026"}'
        self.assertEqual(json.loads(word_generation_schema('referat',schema)), {'org':'','school':'','year':''})
        self.assertEqual(word_generation_schema('statement',schema), schema)

    def test_plain_document_omits_placeholder_and_retains_real_organization(self):
        env = office_env()
        with tempfile.TemporaryDirectory() as td:
            path = str(Path(td)/'sample.docx')
            for org in ('[учебное заведение]', 'Мастерская Север'):
                env['build_word'](path,'Хранение файлов',[{'title':'Сравнение','content':'Текст'}],meta={'org':org})
                paragraphs = Document(path).paragraphs
                text='\n'.join(p.text for p in paragraphs)
                self.assertNotIn('[учебное заведение]',text)
                if org=='Мастерская Север': self.assertIn(org,text)
                title=next(p for p in paragraphs if p.text=='Хранение файлов')
                self.assertEqual(title.style.name,'Title')
                self.assertNotIn('pBdr',title.style.element.xml)

    def test_study_cover_omits_unknown_details_and_template_keeps_blanks(self):
        env=office_env()
        with tempfile.TemporaryDirectory() as td:
            path=str(Path(td)/'sample.docx')
            for kind in ('referat','report','essay','coursework'):
                env['build_word'](path,'Хранение файлов',[{'title':'Введение','content':'Текст'}],kind,
                                  {'school':'[учебное заведение]','author':'Анна Петрова'})
                text='\n'.join(p.text for p in Document(path).paragraphs)
                self.assertIn('Анна Петрова',text)
                self.assertNotIn('[',text)
                self.assertNotIn('Министерство',text)
            env['build_word'](path,'Шаблон',[], 'referat',{},extra_data={'template':True})
            self.assertIn('[Название учебного заведения]','\n'.join(p.text for p in Document(path).paragraphs))

    def test_invalid_revision_is_not_accepted(self):
        original={'sections':[{'title':'А','content':'Текст','table':{'headers':['X'],'rows':[['42']]}}]}
        for expanded in (None, [], {'sections':None}, {'sections':[{'title':'Б','content':'Другой'}]}):
            self.assertIs(merge_word_expansion(original,expanded), original)


class WordPaginationTests(unittest.IsolatedAsyncioTestCase):
    async def run_case(self, candidate_pages=2, failure=None, requested=2, initial_pages=1):
        with tempfile.TemporaryDirectory() as td:
            docx,pdf=Path(td)/'original.docx',Path(td)/'original.pdf'
            original={'title':'Тема','meta':{'author':'Анна'},'sections':[{'title':'А','content':'initial','table':{'headers':['X'],'rows':[['42']]}}]}
            built=[]
            async def build(value,path):
                built.append(copy.deepcopy(value))
                pages=initial_pages if value['sections'][0]['content']=='initial' else candidate_pages
                Path(path).write_text(str(pages))
            async def convert(source,target):
                if failure=='conversion' and 'revised' in source: raise RuntimeError('fixture')
                Path(target).write_text(Path(source).read_text())
            async def expand(value,actual,wanted):
                if failure=='model': raise RuntimeError('fixture')
                if failure=='cancel': raise asyncio.CancelledError()
                if failure=='malformed': return {'sections':None}
                return {'title':'Malicious replacement','meta':{'author':'Other'},'sections':[{'title':'А','content':'expanded','table':{'rows':[['99']]}}]}
            revision=AsyncMock(side_effect=expand)
            content,count=await prepare_word_files(original,requested,str(docx),str(pdf),build=build,convert=convert,
                                                  expand=revision,measure=lambda p:int(Path(p).read_text()))
            return dict(content=content,count=count,docx=docx.read_text(),pdf=pdf.read_text(),calls=revision.await_count,built=built)

    async def test_two_page_order_really_gets_one_bounded_revision(self):
        r=await self.run_case()
        self.assertEqual((r['count'],r['calls'],r['docx'],r['pdf']),(2,1,'2','2'))
        self.assertEqual(r['content']['title'],'Тема')
        self.assertEqual(r['content']['meta'],{'author':'Анна'})
        self.assertEqual(r['content']['sections'][0]['table']['rows'],[['42']])

    async def test_no_extra_model_request_when_length_is_met(self):
        r=await self.run_case(initial_pages=2)
        self.assertEqual(r['calls'],0)

    async def test_optional_model_conversion_and_json_failures_keep_valid_pair(self):
        for failure in ('model','conversion','malformed'):
            r=await self.run_case(failure=failure)
            self.assertEqual((r['count'],r['docx'],r['pdf']),(1,'1','1'))
            self.assertEqual(r['calls'],1)

    async def test_shorter_or_oversized_revision_cannot_replace_original(self):
        for pages in (1,10):
            r=await self.run_case(candidate_pages=pages)
            self.assertEqual((r['count'],r['docx'],r['pdf']),(1,'1','1'))

    async def test_cancellation_propagates(self):
        with self.assertRaises(asyncio.CancelledError): await self.run_case(failure='cancel')


class WordDeliveryPageTests(unittest.IsolatedAsyncioTestCase):
    def setup_env(self, actual):
        e,m,s=test_office.WordWorkflowTests().env()
        async def prepare(*args, **kwargs):
            return await prepare_word_files(*args, **kwargs, measure=lambda _:actual)
        e['prepare_word_files']=prepare
        return e,m,s

    async def test_short_result_disclosed_and_only_delivered_pages_charged(self):
        e,m,s=self.setup_env(1)
        await e['word_build'](m,s)
        self.assertEqual(m.answer_document.await_count,2)
        e['spend_credits'].assert_called_once_with(-9901,'word',amount=10)
        self.assertTrue(any('1 стр.' in str(c) and '10 кредитов' in str(c) for c in m.answer.await_args_list))
        self.assertEqual(e['ask_grok'].await_count,2)

    async def test_extra_pages_do_not_increase_charge(self):
        e,m,s=self.setup_env(3)
        await e['word_build'](m,s)
        e['spend_credits'].assert_called_once_with(-9901,'word',amount=20)
        self.assertEqual(e['ask_grok'].await_count,1)

    async def test_plain_document_prompt_has_editorial_structure_without_legal_language(self):
        e,m,s=self.setup_env(2)
        await e['word_build'](m,s)
        prompt=e['ask_grok'].await_args.args[0]
        self.assertIn(WORD_EDITORIAL_GUIDANCE,prompt)
        self.assertNotIn('точно по формулировкам ГК РФ',prompt)
        self.assertNotIn('детекторы ИИ-текста',prompt)
