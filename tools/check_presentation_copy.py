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
import tempfile
from types import SimpleNamespace
from unittest.mock import AsyncMock
ROOT=Path(__file__).resolve().parents[1]
sys.path[:0]=[str(ROOT),str(ROOT/'tests')]
import test_presentation_support as fixtures
import presentation_support as ps
import presentation_quality as pq
from PIL import Image


def adapters(client):
    names={'ask_grok','_create_presentation_plan','extract_json','_balanced_json_spans','grok_failed'}
    nodes=[n for n in ast.parse((ROOT/'main.py').read_text()).body
           if isinstance(n,(ast.FunctionDef,ast.AsyncFunctionDef)) and n.name in names]
    assert {n.name for n in nodes}==names
    ns=dict(vars(ps),REVIEW_SYSTEM=pq.REVIEW_SYSTEM,asyncio=asyncio,json=json,client=client,GROK_ERROR_PREFIX='MODEL_ERROR:',
            content_gen_lang=lambda *a:'ru',grok_json_lang_instruction=lambda *a:'\nВсе тексты на русском.')
    exec(compile(ast.fix_missing_locations(ast.Module(body=[ast.ImportFrom(module='__future__',names=[ast.alias(name='annotations')],level=0)]+nodes,type_ignores=[])),'text-adapters','exec'),ns)
    return ns

CASES=[('album','Мой фотоальбом','designer_soft','Короткие подписи, без вымышленных событий.'),
       ('education','Почему идёт дождь','minimal','Для 7 класса. Испарение, охлаждение воздуха, конденсация, капли и осадки. Без вымышленных чисел.'),
       ('business','Мастерская по ремонту велосипедов','designer_strict','Концепция вымышленной мастерской. Приём велосипеда, осмотр, согласование ремонта и выдача. Без вымышленных цен, адресов, отзывов и обещаний.')]

async def main():
    from openai import AsyncOpenAI
    if not os.environ.get('XAI_API_KEY'):raise RuntimeError('Missing text provider access')
    output=Path('/tmp/alix-copy-check-results');output.mkdir(exist_ok=True)
    async with AsyncOpenAI(api_key=os.environ['XAI_API_KEY'],base_url='https://api.x.ai/v1',max_retries=0) as client:
        api=adapters(client)
        for name,topic,style,extra in CASES:
            if len(sys.argv)>1 and name not in sys.argv[1:]:continue
            env,state,sent=fixtures.BuildTests().build_env(6)
            env.update({k:api[k] for k in ('ask_grok','extract_json','grok_failed','_create_presentation_plan')})
            env.update(local_issues=pq.local_issues,DeckQualityError=pq.DeckQualityError)
            record={'case':name,'reviews':[],'content':None,'image_provider':'fixture; no paid image calls','vision':'fixture descriptions'}
            async def review(*args):
                result=await pq.review_content(*args);record['reviews'].append(result);record['content']=args[2];return result
            env['review_content']=review
            state.data={'topic':topic,'extra':extra,'slides':6,'theme_name':style,'mode':'ai','content_lang':'ru'}
            with tempfile.TemporaryDirectory() as td:
                if name=='album':
                    descriptions=['Нарисованный персонаж на белом фоне.','Синяя стойка для обслуживания автомобиля.',
                                  'Ребёнок сидит в игрушечной машине.','Взрослый и ребёнок смотрят в камеру.',
                                  'Люди стоят рядом с автомобилем в помещении.','Белый автомобиль припаркован у дома.']
                    photos=[]
                    for i in range(6):
                        p=Path(td)/f'photo-{i}.png';Image.new('RGB',(640,480),(20+i*25,80,160)).save(p);photos.append(str(p))
                    async def vision(path,*args,**kw):
                        index=int(Path(path).stem.split('-')[-1])
                        return json.dumps({'slides':[index+1],'description':descriptions[index], 'preserve_full_frame':True,'safe_frames':[]})
                    state.data.update(user_photos=photos,pres_photo_captions={},pres_photo_context='',photo_album_only=True)
                    env['ask_presentation_vision']=AsyncMock(side_effect=vision)
                await env['_build_presentation'](SimpleNamespace(from_user=SimpleNamespace(id=-987654321)),state)
            record['complete']=len(sent)==2;record['billing_simulated']=env['spend_credits'].call_count
            (output/f'{name}.json').write_text(json.dumps(record,ensure_ascii=False,indent=2))
            for filename,payload in sent:
                (output/f'{name}-{filename}').write_bytes(payload)
            print('COPY_CHECK '+json.dumps({'case':name,'complete':record['complete'],'reviews':record['reviews']},ensure_ascii=False),flush=True)
            if record['content']:
                for slide in record['content']['slides']:
                    print('COPY_TEXT '+json.dumps({'case':name,'slide':slide['number'],'text':pq.visible_copy(slide)},ensure_ascii=False),flush=True)
            if not record['complete']:raise RuntimeError('Copy check did not deliver '+name)
    print('COPY_CHECK_ALL_PASSED',flush=True)

if __name__=='__main__':asyncio.run(main())
