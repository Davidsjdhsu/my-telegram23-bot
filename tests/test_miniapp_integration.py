import re
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, Mock
from test_presentation_support import load_functions
from presentation_pricing import presentation_cost


class MiniAppIntegrationTests(unittest.IsolatedAsyncioTestCase):
    async def test_designer_payload_reaches_bot_and_checks_actual_price(self):
        for style, amount in [('default', 40), ('designer_soft', 80), ('designer_strict', 80)]:
            for affordable in [True, False]:
                with self.subTest(style=style, affordable=affordable):
                    start, rejected = AsyncMock(), AsyncMock()
                    balance = Mock(return_value=affordable)
                    fn = load_functions(['_handle_miniapp_payload'], {
                        'ensure_answerable': lambda m: m, 'user_lang': lambda uid: 'ru',
                        'get_user': lambda uid: {}, '_clip_text': lambda value, n: (value or '')[:n],
                        'can_afford': balance, 'presentation_cost': presentation_cost,
                        '_start_presentation_request': start, 'send_no_credits_notice': rejected,
                    })['_handle_miniapp_payload']
                    message = SimpleNamespace(from_user=SimpleNamespace(id=123))
                    state = object()
                    await fn(message, state, {'topic': 'Облака', 'style': style, 'slides': 8, 'photo_mode': 'own'}, 'gen_presentation')
                    balance.assert_called_once_with(123, amount)
                    if affordable:
                        self.assertEqual(start.call_args.kwargs['settings']['theme_name'], style)
                        self.assertEqual(start.call_args.kwargs['settings']['photo_mode'], 'own')
                        rejected.assert_not_awaited()
                    else:
                        start.assert_not_awaited()
                        rejected.assert_awaited_once()

    def test_external_preview_paths_are_served_but_private_paths_are_not(self):
        import ast
        root = Path(__file__).resolve().parents[1]
        tree = ast.parse((root / 'main.py').read_text())
        assignment = next(n for n in tree.body if isinstance(n, ast.Assign) and any(isinstance(t, ast.Name) and t.id == '_ROOT_STATIC_ALLOWED' for t in n.targets))
        allowed = re.compile(ast.literal_eval(assignment.value.args[0]))
        html = (root / 'index.html').read_text()
        paths = re.findall(r'src="(assets/[^"]+)"', html)
        self.assertEqual(len(paths), 16)
        self.assertEqual(len(set(paths)), 8)
        for path in paths:
            self.assertTrue(allowed.fullmatch(path))
            self.assertTrue((root / path).is_file())
        for path in ['users.json', 'assets/../users.json', 'assets/secrets.webp', 'main.py']:
            self.assertFalse(allowed.fullmatch(path))
        self.assertFalse(re.search(r'data:image/[^;]+;base64,', html))
