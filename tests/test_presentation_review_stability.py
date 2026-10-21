import copy
import json
import unittest
from types import SimpleNamespace
from unittest.mock import AsyncMock

import presentation_quality as quality
import test_presentation_support as fixtures


def issue(number, problem='logic', fix='Remove the false causal connection.'):
    return {'slide':number,'problem':problem,'fix':fix}


class ReviewStabilityTests(unittest.IsolatedAsyncioTestCase):
    async def test_unchanged_accepted_slides_are_not_reopened(self):
        original=fixtures.make_content(4)
        revised=copy.deepcopy(original)
        revised['slides'][0]['content']='Исправленное объяснение действия.'
        data={'_editor_previous':{'content':original,'issues':[issue(2)]}}
        model=AsyncMock(return_value='{"issues":[]}')
        self.assertEqual(await quality.review_content(data,{},revised,model,json.loads),[])
        self.assertEqual(model.await_count,1)
        self.assertIn('Область проверки: [2]',model.await_args.args[0])
        self.assertIn('остальными слайдами',model.await_args.args[0])

    async def test_same_content_keeps_unresolved_errors_without_new_review(self):
        original=fixtures.make_content(4);pending=[issue(2,'unsupported')]
        model=AsyncMock()
        data={'_editor_previous':{'content':original,'issues':pending}}
        self.assertEqual(await quality.review_content(data,{},original,model,json.loads),pending)
        model.assert_not_awaited()

    async def test_unfixed_error_is_not_lost_when_other_slide_changes(self):
        original=fixtures.make_content(4);pending=[issue(2),issue(3,'unsupported')]
        revised=copy.deepcopy(original);revised['slides'][0]['content']='Исправленная связь.'
        data={'_editor_previous':{'content':original,'issues':pending}}
        model=AsyncMock(return_value='{"issues":[]}')
        self.assertEqual(await quality.review_content(data,{},revised,model,json.loads),[pending[1]])

    async def test_changed_diagram_and_image_are_checked_again(self):
        original=fixtures.make_content(4);revised=copy.deepcopy(original)
        revised['slides'][0]['diagram']={'kind':'sequence','nodes':[
            {'label':'Нагрев','text':'Вода получает тепло.'},{'label':'Пар','text':'Часть воды испаряется.'}]}
        revised['slides'][1].update(visual_kind='photo',image_prompt='A new photo subject')
        model=AsyncMock(return_value='{"issues":[]}')
        data={'_editor_previous':{'content':original,'issues':[]}}
        self.assertEqual(await quality.review_content(data,{},revised,model,json.loads),[])
        self.assertEqual(model.await_count,3)

    async def test_all_distinct_findings_reach_author_without_six_item_truncation(self):
        content=fixtures.make_content(5)
        for slide in content['slides']:slide['visual_kind']='photo'
        visual=[issue(n,'image_mismatch',f'Replace invented component photo on slide {n}.') for n in range(2,6)]
        editorial=[issue(n,'logic',f'Remove false connection on slide {n}.') for n in range(2,6)]
        editorial.append(issue(2,'logic','Correct a different independent false conclusion.'))
        model=AsyncMock(side_effect=[json.dumps({'issues':visual}),json.dumps({'issues':editorial})])
        result=await quality.review_content({}, {},content,model,json.loads)
        self.assertEqual(len(result),9)
        self.assertCountEqual(result,editorial+visual)

    async def test_new_factual_error_on_changed_slide_still_blocks(self):
        original=fixtures.make_content(4);revised=copy.deepcopy(original)
        revised['slides'][0]['content']='Неверное новое утверждение.'
        data={'_editor_previous':{'content':original,'issues':[]}}
        pending=[issue(2,'unsupported')]
        model=AsyncMock(return_value=json.dumps({'issues':pending}))
        self.assertEqual(await quality.review_content(data,{},revised,model,json.loads),pending)

    async def test_out_of_scope_review_is_invalid_not_automatic_approval(self):
        original=fixtures.make_content(4);revised=copy.deepcopy(original)
        revised['slides'][0]['content']='Исправленное объяснение.'
        model=AsyncMock(return_value=json.dumps({'issues':[issue(3)]}))
        with self.assertRaises(quality.DeckQualityError):
            await quality.review_content({'_editor_previous':{'content':original,'issues':[issue(2)]}}, {},revised,model,json.loads)
        self.assertEqual(model.await_count,2)

    async def test_actual_builder_passes_exact_previous_review_and_still_delivers(self):
        env,state,sent=fixtures.BuildTests().build_env(3)
        before=fixtures.make_content(3);after=copy.deepcopy(before)
        after['slides'][0]['content']='Уточнённое объяснение преобразования света.'
        env['ask_grok']=AsyncMock(side_effect=[json.dumps(before),json.dumps(after)])
        env['review_content']=AsyncMock(side_effect=[[issue(2)],[]])
        await env['_build_presentation'](SimpleNamespace(from_user=SimpleNamespace(id=-1)),state)
        previous=env['review_content'].await_args_list[1].args[0]['_editor_previous']
        self.assertEqual(previous['content']['slides'][0]['content'],before['slides'][0]['content'])
        self.assertEqual(previous['issues'],[issue(2)])
        self.assertEqual(len(sent),2);env['spend_credits'].assert_called_once()
