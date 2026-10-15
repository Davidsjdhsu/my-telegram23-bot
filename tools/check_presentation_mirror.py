"""Run the actual bot delivery function with public fixtures, without Telegram/storage.

Only delivery, account state, locks and billing are replaced. Production plan,
content, editorial review, image providers, layouts, serialization and PDF writer
are executed unmodified. Requires the service's existing provider environment.
"""
import argparse
import ast
import asyncio
import copy
import hashlib
import json
import os
from pathlib import Path
import sys
import time
from types import SimpleNamespace

ROOT=Path(__file__).resolve().parents[1]
sys.path[:0]=[str(ROOT),str(ROOT/'tests')]
import test_presentation_support as scaffolding
import presentation_quality as pq
from check_presentation_quality import adapters as text_adapters
from check_presentation_images import adapters as image_adapters

FIXTURES={
 'hotel':('Бутик-отель в Тоскане: отдых среди виноградников', 'Для людей, выбирающих отпуск. Номера, интерьеры, местная кухня, прогулки, отдых и ограничения такого формата. Это концепция вымышленного бутик-отеля, без реальных цен, отзывов, услуг и наград. Журнальный стиль, кремовый фон, оливковые и терракотовые акценты.'),
 'technology':('Роботы на складе: как автоматизация помогает людям', 'Для широкой аудитории. Реальные мобильные складские роботы, перевозка грузов, комплектация, задачи человека, безопасность и ограничения. Добавь схему работы и сравнение человека с роботом по одинаковым критериям. Без вымышленных цифр.'),
 'education':('Круговорот воды в природе', 'Для учеников 7 класса. Испарение, конденсация, осадки, сток и роль Солнца. Добавь понятную схему последовательности. Объясни, почему круговорот не создаёт новую воду. Без вымышленных данных.'),
 'business':('План открытия книжного магазина', 'Для начинающего предпринимателя. Концепция вымышленного проекта: ассортимент, место, поставщики, категории расходов, первые покупатели. Сравни офлайн-магазин и интернет-магазин по трём одинаковым критериям. Не выдумывай цены, продажи, источники и гарантии успеха. Деловой стиль.'),
 'nature':('Океан и его удивительные обитатели', 'Для учеников 7–9 классов. Киты, акулы, осьминоги, коралловые рифы, глубоководные животные и защита океана. Объясни приспособления на понятных примерах. Без вымышленных фактов и цифр.'),
 'product':('Настольная лампа для домашнего рабочего места', 'Презентация концепции вымышленного продукта. Освещение стола, направленный и рассеянный свет, регулировка положения, материалы, удобство использования и ограничения. Сравни два способа освещения по одинаковым критериям. Журнальный стиль, тёплый светлый фон. Без выдуманных измерений, медицинских обещаний, цены и отзывов.')}


def production_environment(client):
    # This helper supplies isolated account/delivery state, not an alternate generator.
    env,state,_=scaffolding.BuildTests().build_env(8)
    text=text_adapters(client); images=image_adapters(client)
    env.update({k:text[k] for k in ('ask_grok','extract_json','grok_failed','_create_presentation_plan')})
    env.update({k:images[k] for k in ('generate_image','ask_presentation_vision')})
    env.update(local_issues=pq.local_issues,review_content=pq.review_content,DeckQualityError=pq.DeckQualityError,
               REVIEW_SYSTEM=pq.REVIEW_SYSTEM,BASE_DIR=str(ROOT))
    from reportlab.pdfbase.ttfonts import TTFont
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE,XL_LEGEND_POSITION
    env.update(TTFont=TTFont,CategoryChartData=CategoryChartData,XL_CHART_TYPE=XL_CHART_TYPE,XL_LEGEND_POSITION=XL_LEGEND_POSITION)
    tree=ast.parse((ROOT/'main.py').read_text())
    constants={'THEMES','FONT_CANDIDATES','GROK_JSON_LANG_HINT'}
    functions={'register_pdf_fonts','pick_theme','add_chart','content_gen_lang','grok_json_lang_instruction'}
    nodes=[]
    for n in tree.body:
        if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id in constants for t in n.targets):nodes.append(n)
        if isinstance(n,(ast.FunctionDef,ast.AsyncFunctionDef)) and n.name in functions:nodes.append(n)
    exec(compile(ast.fix_missing_locations(ast.Module(body=[ast.ImportFrom(module='__future__',names=[ast.alias(name='annotations')],level=0)]+nodes,type_ignores=[])),'mirror-production-config','exec'),env)
    # _create_presentation_plan's globals must use the same actual language functions.
    text.update({k:env[k] for k in ('content_gen_lang','grok_json_lang_instruction')})
    async def run_heavy(fn,*args):return await asyncio.to_thread(fn,*args)
    env['run_heavy']=run_heavy
    return env,state


async def run_case(name,client,out,ordinary=False):
    env,state=production_environment(client)
    topic,extra=FIXTURES[name]
    data={'topic':topic,'extra':extra,'slides':8,'theme_name':'minimal' if ordinary else 'designer_soft',
          'mode':'ai','content_lang':'ru','photo_mode':'ai'}
    state.data=copy.deepcopy(data)
    out.mkdir(parents=True,exist_ok=True)
    record={'case':name,'request':data,'source_sha256':{},'deliveries':[],'messages':[], 'reviews':[]}
    for p in ('main.py','presentation_support.py','presentation_profiles.py','presentation_designer.py','presentation_visuals.py','presentation_quality.py'):
        record['source_sha256'][p]=hashlib.sha256((ROOT/p).read_bytes()).hexdigest()
    real_plan=env['_create_presentation_plan'];real_review=env['review_content']
    real_generate=env['ask_grok'];real_validate=env['validate_content'];real_scene=env['build_scene'];real_designer=env['validate_designer']
    record['model_responses']=[];record['validation_errors']=[]
    async def generate(*args,**kwargs):
        result=await real_generate(*args,**kwargs)
        record['model_responses'].append(result)
        return result
    def validate(*args):
        try:return real_validate(*args)
        except Exception as error:
            record['validation_errors'].append(type(error).__name__+': '+str(error));raise
    def scene(*args):
        try:return real_scene(*args)
        except Exception as error:
            record['validation_errors'].append(type(error).__name__+': '+str(error));raise
    def validate_design(*args):
        try:return real_designer(*args)
        except Exception as error:
            record['validation_errors'].append(type(error).__name__+': '+str(error));raise
    env.update(ask_grok=generate,validate_content=validate,build_scene=scene,validate_designer=validate_design)

    async def plan(*args):
        result=await real_plan(*args);record['plan']=copy.deepcopy(result)
        print('MIRROR_PLAN '+json.dumps({'case':name,'family':result['design'].get('family'),'titles':[s['title'] for s in result['slides']]},ensure_ascii=False),flush=True)
        return result
    async def review(*args):
        result=await real_review(*args);record['reviews'].append(copy.deepcopy(result));record['content']=copy.deepcopy(args[2]);return result
    async def send_message(uid,text,**kwargs):record['messages'].append(text)
    async def send_document(uid,file,**kwargs):
        target=out/Path(file.filename).name
        payload=Path(file.path).read_bytes();target.write_bytes(payload)
        record['deliveries'].append({'filename':target.name,'bytes':len(payload),'sha256':hashlib.sha256(payload).hexdigest()})
    env.update(_create_presentation_plan=plan,review_content=review,
               bot=SimpleNamespace(send_message=send_message,send_document=send_document))
    message=SimpleNamespace(from_user=SimpleNamespace(id=-987654321),chat=SimpleNamespace(id=-987654321))
    start=time.monotonic()
    await env['_build_presentation'](message,state)
    record['seconds']=round(time.monotonic()-start);record['complete']=len(record['deliveries'])==2
    record['simulated_billing_calls']=env['spend_credits'].call_count
    (out/'record.json').write_text(json.dumps(record,ensure_ascii=False,indent=2))
    if record['complete']:
        assert record['simulated_billing_calls']==1
        print('MIRROR_PASS '+json.dumps({'case':name,'seconds':record['seconds'],'files':record['deliveries']},ensure_ascii=False),flush=True)
    else:
        assert record['simulated_billing_calls']==0
        print('MIRROR_FAIL '+json.dumps({'case':name,'seconds':record['seconds'],'messages':record['messages']},ensure_ascii=False),flush=True)
    return record['complete']


async def main(args):
    if not all(os.environ.get(k) for k in ('XAI_API_KEY','REPLICATE_API_TOKEN')):
        print('MIRROR_BLOCKED missing_provider_credentials');return 2
    from openai import AsyncOpenAI
    async with AsyncOpenAI(api_key=os.environ['XAI_API_KEY'],base_url='https://api.x.ai/v1',max_retries=0) as client:
        results=[]
        for case in args.cases:
            results.append(await run_case(case,client,Path(args.output)/case,args.ordinary))
        return 0 if all(results) else 1

if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--cases',nargs='+',choices=FIXTURES,default=list(FIXTURES))
    parser.add_argument('--output',required=True)
    parser.add_argument('--ordinary',action='store_true')
    raise SystemExit(asyncio.run(main(parser.parse_args())))
