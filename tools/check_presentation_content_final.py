"""Three bounded live text checks using bot authoring/editor/renderers; no Telegram/billing.

Images are fixtures: this checks content, native diagrams and actual PDF conversion,
not live image generation. No customer data or credentials enter output.
"""
import asyncio
import json
import os
from pathlib import Path
import sys
from types import SimpleNamespace
ROOT=Path(__file__).resolve().parents[1]
sys.path[:0]=[str(ROOT),str(ROOT/'tests'),str(ROOT/'tools')]
import test_presentation_support as fixtures
import presentation_quality as pq
import presentation_pdf as pdf
from check_presentation_semantics import adapters

CASES=[('fridge','Как работает холодильник','designer_strict','Для 7–9 классов. Объясни перенос тепла, хладагент, компрессор и нагрев задней стенки. Добавь схему полного цикла работы. Без вымышленных фактов.'),
       ('rain','Почему идёт дождь','minimal','Для 7 класса. Объясни испарение, образование облаков и осадков. Покажи схему круговорота воды и объясни, почему не каждое облако даёт дождь.'),
       ('coffee','Как готовят кофе в кофейне','designer_soft','Для новичков. От зерна до готового напитка: выбор, помол, приготовление. Добавь схему приготовления эспрессо. Не путай разные способы приготовления и не придумывай числа.')]

async def main():
    from openai import AsyncOpenAI
    output=Path('/tmp/alix-content-final-results');output.mkdir(exist_ok=True)
    async with AsyncOpenAI(api_key=os.environ['XAI_API_KEY'],base_url='https://api.x.ai/v1',max_retries=0) as client:
        api=adapters(client)
        for name,topic,style,extra in CASES:
            if len(sys.argv)>1 and name not in sys.argv[1:]:continue
            env,state,sent=fixtures.BuildTests().build_env(6)
            env.update({k:api[k] for k in ('ask_grok','extract_json','grok_failed','_create_presentation_plan')})
            env.update(local_issues=pq.local_issues,DeckQualityError=pq.DeckQualityError,
                       ensure_pdf_converter=pdf.ensure_pdf_converter,convert_presentation_pdf=pdf.convert_presentation_pdf)
            record={'case':name,'reviews':[],'content':None,'images':'fixtures, not real image-provider calls'}
            async def review(*args):
                result=await pq.review_content(*args)
                record['reviews'].append(result);record['content']=args[2]
                return result
            env['review_content']=review
            state.data={'topic':topic,'extra':extra,'slides':6,'theme_name':style,'mode':'ai','content_lang':'ru'}
            await env['_build_presentation'](SimpleNamespace(from_user=SimpleNamespace(id=-987654321)),state)
            record['complete']=len(sent)==2
            (output/(name+'.json')).write_text(json.dumps(record,ensure_ascii=False,indent=2))
            for filename,payload in sent:(output/(name+Path(filename).suffix)).write_bytes(payload)
            print('CONTENT_FINAL '+json.dumps({'case':name,'complete':record['complete'],'reviews':record['reviews']},ensure_ascii=False),flush=True)
            if record['content']:
                for slide in record['content']['slides']:
                    print('CONTENT_TEXT '+json.dumps({'case':name,'slide':slide['number'],'diagram':slide.get('diagram'),'text':pq.visible_copy(slide)},ensure_ascii=False),flush=True)
            if not record['complete']:raise RuntimeError('Case failed: '+name)
    print('CONTENT_FINAL_CHECK_PASSED',flush=True)

if __name__=='__main__':asyncio.run(main())
