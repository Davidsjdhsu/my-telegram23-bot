import ast, json, os, random, re, tempfile, unittest
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, AsyncMock
import word_quality
ROOT=Path(__file__).resolve().parents[1]
COLORS={'bg':(255,255,255),'ink':(30,40,50),'mute':(110,120,130),'line':(45,90,130)}
def office_env():
    tree=ast.parse((ROOT/'main.py').read_text());ns=dict(asyncio=__import__("asyncio"),datetime=datetime,re=re,os=os,json=json,random=random,convert_word_pdf=word_quality.convert_word_pdf,merge_word_expansion=word_quality.merge_word_expansion)
    from service_pricing import VERSION
    ns.update(PRICE_VERSION=VERSION,billing_saved=AsyncMock(),send_no_credits_notice=AsyncMock())
    nodes=[ast.ImportFrom(module='__future__',names=[ast.alias(name='annotations')],level=0)]
    for n in tree.body:
        if isinstance(n,ast.ImportFrom) and (n.module or '').split('.')[0] in ('docx','openpyxl','word_quality','service_pricing','wallet'):nodes.append(n)
        elif isinstance(n,ast.Import) and any(x.name=='openpyxl' for x in n.names):nodes.append(n)
        elif isinstance(n,(ast.FunctionDef,ast.AsyncFunctionDef)) and (n.name.startswith('_xl_') or n.name in ('_run','_p','_vspace','_set_cell_borders','_set_cell_margins','_shade_cell','_cell_p','_student_title_page','_tabbed_line','_add_table','_parse_markdown_table','_body','build_word','build_excel_items','build_excel_startup','word_build','build_word_from_upload','safe_filename')):
            n.decorator_list=[];nodes.append(n)
        elif isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id in ('CONTRACT_FAMILY','ACT_FAMILY','STATEMENT_FAMILY','UNILATERAL_FAMILY','DECLARATION_FAMILY') for t in n.targets):nodes.append(n)
    exec(compile(ast.fix_missing_locations(ast.Module(body=nodes,type_ignores=[])),str(ROOT/'main.py'),'exec'),ns);return ns
class OfficeTests(unittest.TestCase):
    def test_expansion_preserves_tables_meta_and_order(self):
        a={'title':'Тема','meta':{'author':'Иван'},'sections':[{'title':'А','content':'Коротко','table':{'headers':['X'],'rows':[['42']]}}]};b={'sections':[{'title':'А','content':'Подробный текст'}]}
        c=word_quality.merge_word_expansion(a,b);self.assertEqual(c['meta'],a['meta']);self.assertEqual(c['sections'][0]['table'],a['sections'][0]['table']);b['sections'][0]['title']='Б';self.assertEqual(word_quality.merge_word_expansion(a,b),a)
    def test_word_layout(self):
        e=office_env()
        with tempfile.TemporaryDirectory() as td:
            p=str(Path(td)/'test.docx');e['build_word'](p,'Документ',[{'title':'Раздел','content':'Текст','table':{'headers':['Имя','Количество'],'rows':[['А','12']]}}]);d=e['Document'](p)
            self.assertAlmostEqual(d.sections[0].page_width.cm,21,places=2);h=next(p for p in d.paragraphs if p.text=='Раздел');self.assertEqual(h.style.name,'Heading 1');self.assertTrue(h.paragraph_format.keep_with_next);self.assertIn('tblHeader',d.tables[0].rows[0]._tr.xml)
    def test_budget_type_and_labels(self):
        e=office_env()
        with tempfile.TemporaryDirectory() as td:
            p=str(Path(td)/'budget.xlsx');e['build_excel_items'](p,'Бюджет','family_budget',COLORS,[{'name':'Аренда','type':'expense','amount':200}]);w=e['openpyxl'].load_workbook(p).active;self.assertEqual(w['C5'].value,'Расход');self.assertIn('A6:C6',str(w.merged_cells))
            with self.assertRaises(ValueError):e['build_excel_items'](p,'Бюджет','family_budget',COLORS,[{'name':'Аренда','type':'unknown','amount':200}])
class WordWorkflowTests(unittest.IsolatedAsyncioTestCase):
    def env(self):
        e=office_env();e.update(user_lang=lambda _:'ru',get_user=Mock(return_value={'generations':0,'history':[]}),start_job=Mock(return_value=(True,'')),finish_job=Mock(),ensure_pdf_converter=Mock(),tr=lambda key,*a,**k:key,normalize_word_pages=lambda p:p,word_target_words=lambda p:p*250,grok_json_lang_instruction=lambda _:'',content_gen_lang=lambda *a:'ru',WORD_SIZE_KINDS=set(),STUDY_KINDS=set(),META_SCHEMAS={'doc':'{}'},WORD_KIND_DESC={'doc':'Документ'},ask_grok=AsyncMock(return_value='{"title":"Test","sections":[{"title":"Part","content":"Text"}]}'),extract_json=json.loads,run_heavy=AsyncMock(),convert_word_pdf=AsyncMock(),spend_credits=Mock(),word_cost=lambda p, kind='doc':p*10,note_success=Mock(),FSInputFile=lambda path,**kw:path,main_kb=lambda *a,**k:None,word_confirm_kb=lambda *a:None,signature_line=AsyncMock(return_value=''),can_afford=lambda *a:True,extract_page_count=lambda _:2,save_users=Mock())
        async def prepare(*args, **kwargs):
            # Only page measurement is stubbed; run the real orchestration and failure handling.
            return await word_quality.prepare_word_files(*args, **kwargs, measure=lambda _: 2)
        e['prepare_word_files'] = prepare
        e.update(ReplyKeyboardRemove=lambda:None, ReplyKeyboardMarkup=lambda **kw:SimpleNamespace(**kw),
                 KeyboardButton=lambda **kw:SimpleNamespace(**kw),
                 Form=SimpleNamespace(waiting_word_quality='Form:waiting_word_quality', waiting_academic_meta='Form:waiting_academic_meta',
                                      waiting_word_confirm='Form:waiting_word_confirm'))
        m=SimpleNamespace(from_user=SimpleNamespace(id=-9901),answer=AsyncMock(),answer_document=AsyncMock());s=SimpleNamespace(get_data=AsyncMock(return_value={'word_kind':'doc','word_pages':2,'topic':'Test','price_version':e['PRICE_VERSION']}),clear=AsyncMock(),update_data=AsyncMock(),set_state=AsyncMock());return e,m,s
    async def test_model_failure_unlocks(self):
        e,m,s=self.env();e['ask_grok'].side_effect=RuntimeError('fixture');await e['word_build'](m,s);e['finish_job'].assert_called_once();e['spend_credits'].assert_not_called();m.answer_document.assert_not_awaited();s.clear.assert_not_awaited()
    async def test_upload_failure_unlocks(self):
        e,m,s=self.env();e['ask_grok'].side_effect=RuntimeError('fixture');await e['build_word_from_upload'](m,s,'source','instruction');e['finish_job'].assert_called_once();e['spend_credits'].assert_not_called()
    async def test_pdf_failure_no_delivery_or_charge(self):
        e,m,s=self.env();e['convert_word_pdf'].side_effect=RuntimeError('fixture');await e['word_build'](m,s);m.answer_document.assert_not_awaited();e['spend_credits'].assert_not_called();s.clear.assert_not_awaited();e['finish_job'].assert_called_once()
    async def test_second_delivery_failure_no_charge(self):
        e,m,s=self.env();m.answer_document.side_effect=[None,RuntimeError('fixture')];await e['word_build'](m,s);self.assertEqual(m.answer_document.await_count,2);e['spend_credits'].assert_not_called()
    async def test_charge_after_pair(self):
        e,m,s=self.env();order=[]
        async def converted(*a):order.append('convert')
        async def sent(*a,**kw):order.append('send')
        e['convert_word_pdf'].side_effect=converted;m.answer_document.side_effect=sent;e['spend_credits'].side_effect=lambda *a,**kw:order.append('charge');await e['word_build'](m,s);self.assertEqual(order,['convert','send','send','charge']);e['finish_job'].assert_called_once()
