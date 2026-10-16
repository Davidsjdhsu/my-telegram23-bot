"""Bounded live text check through the actual bot builder. No Telegram or billing.

Only public synthetic fixtures are used. Images and vision are fixture adapters;
authoring, plans, editorial review, validation and PPTX/PDF delivery use bot code.
Run on Render with existing XAI_API_KEY. No keys or user storage are printed.
"""
import ast
import asyncio
import json
import os
from pathlib import Path
import sys
from types import SimpleNamespace
ROOT=Path(__file__).resolve().parents[1]
sys.path[:0]=[str(ROOT),str(ROOT/'tests')]
import test_presentation_support as fixtures
import presentation_support as ps
import presentation_quality as pq


def adapters(client):
    names={'ask_grok','_create_presentation_plan','extract_json','_balanced_json_spans','grok_failed'}
    nodes=[n for n in ast.parse((ROOT/'main.py').read_text()).body
           if isinstance(n,(ast.FunctionDef,ast.AsyncFunctionDef)) and n.name in names]
    assert {n.name for n in nodes}==names
    ns=dict(vars(ps),REVIEW_SYSTEM=pq.REVIEW_SYSTEM,asyncio=asyncio,json=json,client=client,GROK_ERROR_PREFIX='MODEL_ERROR:',
            content_gen_lang=lambda *a:'ru',grok_json_lang_instruction=lambda *a:'\nВсе тексты на русском.')
    exec(compile(ast.fix_missing_locations(ast.Module(body=[ast.ImportFrom(module='__future__',names=[ast.alias(name='annotations')],level=0)]+nodes,type_ignores=[])),'text-adapters','exec'),ns)
    return ns

CASES=[('rain','Почему идёт дождь','designer_soft','Для 7 класса. Испарение, подъём и охлаждение воздуха, образование облаков, рост капель, выпадение дождя и возвращение воды. Почему не каждое облако приносит дождь. На последнем слайде кратко покажи весь процесс. Короткий естественный текст, без повторов и вымышленных цифр.'),
       ('plant','Как прорастает фасоль','minimal','Для 7 класса. Прорастание семени, роль воды, температуры и кислорода. Как развивается проросток и когда ему нужен свет. Объясни причины простыми словами, без абсолютных правил для всех видов растений.'),
       ('workshop','Мастерская по ремонту велосипедов','designer_strict','Концепция вымышленной мастерской. Приём, осмотр, согласование ремонта, ремонт и выдача. Сравни самостоятельный ремонт и мастерскую по двум содержательным критериям. Без вымышленных цен, адресов и обещаний. В конце добавь схему процесса ремонта.')]

async def main():
    from openai import AsyncOpenAI
    if not os.environ.get('XAI_API_KEY'):raise RuntimeError('Missing text provider access')
    output=Path('/tmp/alix-semantic-check-results');output.mkdir(exist_ok=True)
    async with AsyncOpenAI(api_key=os.environ['XAI_API_KEY'],base_url='https://api.x.ai/v1',max_retries=0) as client:
        api=adapters(client)
        if len(sys.argv)==1 or 'negative' in sys.argv[1:]:
            blocks=[{'heading':'Наличные','text':'Покупатель передаёт купюры кассиру.'},
                    {'heading':'Карта','text':'Покупатель прикладывает банковскую карту к терминалу.'},
                    {'heading':'Перевод','text':'Покупатель переводит деньги через приложение.'}]
            bad={'title':'Способы оплаты','slides':[{'number':2,'title':'Как можно заплатить',
                 'content':'Разные способы оплаты.','image_prompt':'Checkout counter.',
                 'designer':{'layout':'process','blocks':blocks}}]}
            data={'topic':'Способы оплаты','extra':'Наличные, карта или перевод как альтернативы.'}
            plan={'title':'Способы оплаты','slides':[{'number':1,'title':'Способы оплаты','point':'Обзор'},
                  {'number':2,'title':'Как можно заплатить','point':'Три альтернативы'}]}
            rejected=await pq.review_content(data,plan,bad,api['ask_grok'],api['extract_json'])
            assert any(x['problem']=='logic' for x in rejected), rejected
            bad['slides'][0]['designer']['layout']='sections'
            approved=await pq.review_content(data,plan,bad,api['ask_grok'],api['extract_json'])
            assert not [x for x in approved if x['problem'] in ('logic','unsupported','missing')], approved
            (output/'negative.json').write_text(json.dumps({'rejected':rejected,'corrected':approved},ensure_ascii=False))
            print('SEMANTIC_CHECK_NEGATIVE_PASSED '+json.dumps(rejected,ensure_ascii=False),flush=True)
        for name,topic,style,extra in CASES:
            if len(sys.argv)>1 and name not in sys.argv[1:]:continue
            count=8 if name=='rain' else 6
            env,state,sent=fixtures.BuildTests().build_env(count)
            env.update({k:api[k] for k in ('ask_grok','extract_json','grok_failed','_create_presentation_plan')})
            env.update(local_issues=pq.local_issues,DeckQualityError=pq.DeckQualityError)
            record={'case':name,'reviews':[],'content':None,'image_provider':'fixture; no paid image calls','vision':'fixture descriptions'}
            async def review(*args):
                result=await pq.review_content(*args);record['reviews'].append(result);record['content']=args[2];return result
            env['review_content']=review
            state.data={'topic':topic,'extra':extra,'slides':count,'theme_name':style,'mode':'ai','content_lang':'ru'}
            await env['_build_presentation'](SimpleNamespace(from_user=SimpleNamespace(id=-987654321)),state)
            record['plan']=state.data.get('pres_plan');record['complete']=len(sent)==2;record['billing_simulated']=env['spend_credits'].call_count
            (output/f'{name}.json').write_text(json.dumps(record,ensure_ascii=False,indent=2))
            for filename,payload in sent:
                (output/f'{name}-{filename}').write_bytes(payload)
            print('SEMANTIC_CHECK '+json.dumps({'case':name,'complete':record['complete'],'reviews':record['reviews']},ensure_ascii=False),flush=True)
            if record['content']:
                for slide in record['content']['slides']:
                    print('SEMANTIC_TEXT '+json.dumps({'case':name,'slide':slide['number'],'text':pq.visible_copy(slide)},ensure_ascii=False),flush=True)
            if not record['complete']:raise RuntimeError('Copy check did not deliver '+name)
    print('SEMANTIC_CHECK_ALL_PASSED',flush=True)

if __name__=='__main__':asyncio.run(main())
