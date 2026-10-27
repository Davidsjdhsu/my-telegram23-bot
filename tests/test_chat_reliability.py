import asyncio
import json
import tempfile
import time
import unittest
from pathlib import Path
from types import SimpleNamespace as NS
from unittest.mock import AsyncMock, Mock

from test_chat_support import main_namespace
from chat_support import is_retry_request


def response(text='Ответ: 1000 шариков.', finish='stop'):
    return NS(choices=[NS(message=NS(content=json.dumps({"status":"answer","text":text}, ensure_ascii=False) if text else ""), finish_reason=finish)])


class PhotoTests(unittest.IsolatedAsyncioTestCase):
    def env(self, outcomes):
        create = AsyncMock(side_effect=outcomes)
        options = Mock(return_value=NS(chat=NS(completions=NS(create=create))))
        async def heavy(fn, *args): return fn(*args)
        ns = main_namespace('ask_grok_vision')
        ns.update(time=time, VISION_MODEL='grok-4.6', GROK_ERROR_PREFIX='ERROR:',
                  VISION_SOLVE_PROMPT='Read accurately; do not invent geometry.',
                  run_heavy=heavy, _prepare_vision_image=lambda *a: ('fixture', 'image/jpeg'),
                  client=NS(with_options=options))
        return ns, create, options

    async def test_timeout_falls_back_with_same_image_and_bounded_retries(self):
        ns, create, options = self.env([TimeoutError(), response()])
        self.assertIn('1000', await ns['ask_grok_vision']('photo.jpg', 'Реши'))
        self.assertEqual(create.await_count, 2)
        first, second = [c.kwargs for c in create.await_args_list]
        self.assertNotEqual(first['model'], second['model'])
        self.assertEqual(first['messages'], second['messages'])
        self.assertEqual(first['extra_body'], {'reasoning_effort': 'low'})
        self.assertTrue(all(c.kwargs['max_retries'] == 0 for c in options.call_args_list))

    async def test_empty_or_truncated_response_is_not_delivered_as_success(self):
        for initial in (response(''), response('Незаконченное вычисление', 'length')):
            ns, create, _ = self.env([initial, response()])
            self.assertEqual(await ns['ask_grok_vision']('photo.jpg', 'Реши'), 'Ответ: 1000 шариков.')
            self.assertEqual(create.await_count, 2)

    async def test_two_failures_stop_and_do_not_return_raw_provider_error(self):
        ns, create, _ = self.env([RuntimeError('secret request data'), TimeoutError()])
        value = await ns['ask_grok_vision']('photo.jpg', 'Реши')
        self.assertEqual(value, 'ERROR:TimeoutError')
        self.assertEqual(create.await_count, 2)

    async def test_auth_error_does_not_retry(self):
        error = RuntimeError('private details'); error.status_code = 401
        ns, create, _ = self.env([error])
        self.assertTrue((await ns['ask_grok_vision']('photo.jpg', 'Реши')).startswith('ERROR:'))
        self.assertEqual(create.await_count, 1)

    async def test_cancellation_never_starts_fallback(self):
        ns, create, _ = self.env([asyncio.CancelledError()])
        with self.assertRaises(asyncio.CancelledError):
            await ns['ask_grok_vision']('photo.jpg', 'Реши')
        self.assertEqual(create.await_count, 1)

    async def test_failed_photo_has_no_charge_or_false_blur_diagnosis(self):
        from contextlib import asynccontextmanager
        @asynccontextmanager
        async def typing(uid): yield
        ns = main_namespace('_handle_vision_upload_impl')
        user = {}; charges = Mock(); answer = AsyncMock()
        ns.update(user_lang=lambda uid: 'ru', get_user=lambda uid: user,
                  can_afford=lambda *a: True, CREDIT_COSTS={'vision': 3},
                  start_job=lambda uid,**kw: (True, None), finish_job=Mock(), chat_typing=typing,
                  analyze_user_image=AsyncMock(return_value='ERROR:TimeoutError'),
                  grok_failed=lambda s: s.startswith('ERROR:'), spend_credits=charges)
        await ns['_handle_vision_upload_impl'](NS(from_user=NS(id=1), answer=answer), None, 'photo.jpg', '')
        self.assertIn('Кредиты не списаны', answer.await_args.args[0])
        self.assertNotIn('размыти', answer.await_args.args[0])
        charges.assert_not_called()

    def test_exif_orientation_is_applied_before_encoding(self):
        import base64, io, os
        from PIL import Image
        ns = main_namespace('_prepare_vision_image'); ns.update(base64=base64, os=os)
        with tempfile.TemporaryDirectory() as td:
            path = str(Path(td) / 'rotated.jpg')
            exif = Image.Exif(); exif[274] = 6
            Image.new('RGB', (80, 40)).save(path, exif=exif)
            encoded, mime = ns['_prepare_vision_image'](path)
            with Image.open(io.BytesIO(base64.b64decode(encoded))) as im:
                self.assertEqual(im.size, (40, 80))
            self.assertEqual(mime, 'image/jpeg')
            self.assertFalse(Path(path + '.vis.jpg').exists())


class RetryTests(unittest.IsolatedAsyncioTestCase):
    def test_only_explicit_retry_commands(self):
        for text in ('повтори', 'Повтори, пожалуйста!', 'давай попробуй ещё раз',
                     'продолжи сборку', 'retry', 'please try again'):
            self.assertTrue(is_retry_request(text), text)
        for text in ('не повторяй', 'повтори цену', 'почему надо повторить',
                     'повтори без картинок', 'нет', '', None):
            self.assertFalse(is_retry_request(text), text)

    async def test_resumes_existing_order_without_omitting_images(self):
        ns = main_namespace('waiting_confirm_fallback')
        build = AsyncMock(); choice = AsyncMock()
        ns.update(user_lang=lambda uid: 'ru', _build_presentation=build, ask_photo_source=choice)
        state = NS(get_data=AsyncMock(return_value={'pres_plan': {'title': 'Example'}, 'photo_mode': 'own'}))
        message = NS(text='повтори', from_user=NS(id=1), answer=AsyncMock())
        await ns['waiting_confirm_fallback'](message, state)
        build.assert_awaited_once_with(message, state)
        choice.assert_not_awaited()

    async def test_unknown_text_does_not_start_or_charge_a_job(self):
        ns = main_namespace('waiting_confirm_fallback'); build = AsyncMock()
        ns.update(user_lang=lambda uid: 'ru', _build_presentation=build, confirm_kb=lambda lang: 'buttons')
        message = NS(text='не надо повторять', from_user=NS(id=1), answer=AsyncMock())
        await ns['waiting_confirm_fallback'](message, None)
        build.assert_not_awaited()
        self.assertIn('План сохранён', message.answer.await_args.args[0])


class ChatTests(unittest.IsolatedAsyncioTestCase):
    async def test_chat_fallback_keeps_history_and_task(self):
        ns = main_namespace('ask_grok_chat')
        create = AsyncMock(side_effect=[response(''), response('Проще: это рабочая память.')])
        ns.update(client=NS(with_options=Mock(return_value=NS(chat=NS(completions=NS(create=create))))))
        result = await ns['ask_grok_chat']('объясни проще', history=[{'role': 'user', 'content': 'Что такое ОЗУ?'}])
        self.assertIn('рабочая память', result)
        first, second = [c.kwargs for c in create.await_args_list]
        self.assertEqual(first['messages'], second['messages'])
        self.assertIn('Что такое ОЗУ?', str(second['messages']))


if __name__ == '__main__': unittest.main()
