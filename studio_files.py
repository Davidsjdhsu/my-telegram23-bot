"""Private Telegram file references. No generation or payment on retrieval."""
import asyncio
import hashlib
import json
import time
from pathlib import Path
from aiohttp import web
from studio_word import atomic


class FileLibrary:
    def __init__(self, root):
        self.root = Path(root)
        self.root.mkdir(parents=True, exist_ok=True, mode=0o700)
        self.busy = set()
        self.last_open = {}

    def load(self, uid):
        try:
            rows = json.loads((self.root / (str(uid) + '.json')).read_text())
            if not isinstance(rows, list):
                return []
            return [r for r in rows if isinstance(r, dict)
                    and isinstance(r.get('id'), str) and isinstance(r.get('title'), str)
                    and isinstance(r.get('created'), (int, float))
                    and isinstance(r.get('files'), dict)
                    and all(isinstance(k, str) and isinstance(v, str) for k, v in r['files'].items())]
        except (OSError, ValueError):
            return []

    def record(self, uid, message, title, ext, group):
        document = getattr(message, 'document', None)
        if not document or not getattr(document, 'file_id', None):
            return
        token = hashlib.sha256(str(group).encode()).hexdigest()[:32]
        rows = self.load(uid)
        row = next((r for r in rows if r['id'] == token), None)
        if row is None:
            row = {'id': token, 'title': title[:200], 'created': time.time(), 'files': {}}
            rows.append(row)
        row['files'][ext] = document.file_id
        try:
            atomic(self.root / (str(uid) + '.json'), rows[-100:])
        except OSError:
            # The file has already reached Telegram; do not turn an index error
            # into an ambiguous delivery or retry the send.
            pass

    def public(self, uid):
        return [{'id': r['id'], 'title': r['title'], 'formats': [e.upper() for e in r['files']],
                 'created': r['created']} for r in reversed(self.load(uid))]

    def attach(self, studio, app, bot):
        async def open_file(request):
            uid = studio.authorized(request)['user']['id']
            studio.check_write(request)
            row = next((r for r in self.load(uid) if r['id'] == request.match_info['file_id']), None)
            if row is None:
                raise web.HTTPNotFound()
            if uid in self.busy or time.monotonic() - self.last_open.get(uid, -100) < 5:
                raise web.HTTPTooManyRequests(text='Подождите несколько секунд.')
            self.busy.add(uid)
            self.last_open[uid] = time.monotonic()
            try:
                for ext, file_id in row['files'].items():
                    await bot.send_document(uid, file_id, caption=row['title'] + ' · ' + ext.upper())
                return web.json_response({'ok': True})
            finally:
                self.busy.discard(uid)
        app.router.add_post('/studio/files/{file_id}/open', open_file)
