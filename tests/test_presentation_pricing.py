import copy
import json
import unittest
from types import SimpleNamespace
from unittest.mock import AsyncMock, Mock

import test_presentation_support as f
from presentation_pricing import presentation_cost, presentation_price_notice


class PresentationPricingTests(unittest.IsolatedAsyncioTestCase):
    def priced_order(self, style, balance=1000):
        env, state, sent = f.BuildTests().build_env(8)
        state.data['theme_name'] = style
        state.data['pres_plan_key'] = f.ps.plan_key(state.data)
        if style.startswith('designer_'):
            content = f.make_content(8)
            for slide in content['slides']:
                slide['designer'] = {'layout': 'auto', 'blocks': []}
            env['ask_grok'] = AsyncMock(return_value=json.dumps(content))
        user = {'credits': balance, 'generations': 0, 'history': []}
        ledger = f.load_functions(['can_afford', 'spend_credits'], {
            'ADMIN_IDS': set(), 'get_user': lambda uid: user,
            'save_users': Mock(), 'CREDIT_COSTS': {'presentation': 10}})
        env.update(get_user=lambda uid: user, can_afford=ledger['can_afford'],
                   spend_credits=Mock(wraps=ledger['spend_credits']))
        return env, state, sent, user, ledger

    def test_rates_follow_actual_renderer_including_custom_colors(self):
        for style in ('default', 'minimal', 'tech', 'business', 'unknown', None):
            self.assertEqual(presentation_cost(8, style), 128)
        for style in ('designer_soft', 'designer_strict'):
            self.assertEqual(presentation_cost(8, style), 256)
            self.assertEqual(presentation_cost('bad', style), 96)
            # Copied custom colours select the ordinary renderer.
            self.assertEqual(presentation_cost(8, {'theme_name': style, 'custom_colors': {'bg': 'white'}}), 128)

    async def test_successful_delivery_debits_actual_balance_at_selected_rate(self):
        for style, amount in [('minimal', 128), ('designer_soft', 256), ('designer_strict', 256)]:
            with self.subTest(style=style):
                env, state, sent, user, ledger = self.priced_order(style)
                original = env['bot'].send_document
                async def send(*a, **kw):
                    self.assertEqual(user['credits'], 1000)
                    return await original(*a, **kw)
                env['bot'].send_document = send
                result = await env['_build_presentation'](SimpleNamespace(from_user=SimpleNamespace(id=-987654321)), state)
                self.assertEqual(result, {'delivered': True})
                self.assertEqual(len(sent), 2)
                self.assertEqual(user['credits'], 1000-amount)
                env['spend_credits'].assert_called_once_with(-987654321, 'presentation', amount=amount)
                ledger['save_users'].assert_called_once()
                self.assertIn(str(amount)+' кредитов', env['bot'].send_message.await_args_list[0].args[1])

    async def test_second_file_delivery_failure_has_no_success_signal_or_charge(self):
        env, state, sent, user, ledger = self.priced_order('minimal')
        original = env['bot'].send_document
        async def send(*a, **kw):
            if sent:
                raise RuntimeError('test PDF send failure')
            return await original(*a, **kw)
        env['bot'].send_document = send
        result = await env['_build_presentation'](SimpleNamespace(from_user=SimpleNamespace(id=-987654321)), state)
        self.assertIsNone(result)
        self.assertEqual(len(sent), 1)
        self.assertEqual(user['credits'], 1000)
        env['spend_credits'].assert_not_called()

    async def test_designer_insufficient_balance_stops_before_generation(self):
        for style in ('designer_soft', 'designer_strict'):
            env, state, sent, user, ledger = self.priced_order(style, balance=255)
            await env['_build_presentation'](SimpleNamespace(from_user=SimpleNamespace(id=-987654321)), state)
            self.assertEqual(sent, [])
            env['start_job'].assert_not_called()
            env['ask_grok'].assert_not_awaited()
            ledger['save_users'].assert_not_called()
            self.assertEqual(user['credits'], 255)

    async def test_missing_images_partial_consent_keeps_designer_price(self):
        env, state, sent, user, ledger = self.priced_order('designer_strict')
        env['generate_image'] = AsyncMock(return_value=False)
        message = SimpleNamespace(from_user=SimpleNamespace(id=-987654321))
        await env['_build_presentation'](message, state)
        self.assertEqual(user['credits'], 1000)
        self.assertIn('256', env['bot'].send_message.await_args.args[1])
        await env['_build_presentation'](message, state, allow_missing_images=True)
        self.assertEqual(len(sent), 2)
        self.assertEqual(user['credits'], 744)
        ledger['save_users'].assert_called_once()

    async def test_style_switch_updates_quote_and_rejects_insufficient_balance(self):
        env = f.WorkflowTests().env(_create_presentation_plan=AsyncMock(return_value=f.make_plan(8)),
            _send_presentation_draft=AsyncMock(), STYLE_BY_LABEL={'Designer':'designer_strict','Standard':'minimal'},
            format_plan=lambda *a:'Plan', style_kb=lambda **kw:None)
        env = f.load_functions(['process_style'], env)
        state = f.State({'slides':8,'theme_name':'minimal'})
        m = f.WorkflowTests().message('Designer')
        env['can_afford'] = lambda uid, cost: cost <= 255
        before = copy.deepcopy(state.data)
        await env['process_style'](m, state)
        self.assertEqual(state.data, before)
        env['_create_presentation_plan'].assert_not_awaited()
        env['can_afford'] = lambda uid, cost: cost <= 256
        await env['process_style'](m, state)
        self.assertIn('256 кредитов', env['_send_presentation_draft'].await_args.args[1])
        m.text = 'Standard'
        await env['process_style'](m, state)
        self.assertIn('128 кредитов', env['_send_presentation_draft'].await_args.args[1])

    def test_quotes_include_total_and_formats_in_all_languages(self):
        for lang in ('ru','en','de','ar','zh','es','fr'):
            for style, amount in [('minimal',128),('designer_soft',256),('designer_strict',256)]:
                quote = presentation_price_notice(8, style, lang)
                for part in (str(amount),'PPTX','PDF'):
                    self.assertIn(part, quote)
