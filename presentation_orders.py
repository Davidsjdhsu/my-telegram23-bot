from response_outcomes import GenerationStopped, ensure_generation, ensure_pptx, generation_notice
"""Durable presentation delivery journal shared by chat and Mini App.

Telegram has no idempotency key for sendDocument. An interrupted send is
explicitly unresolved; never resend or debit automatically in that state.
"""
import asyncio
import copy
import hashlib
import json
import shutil
import time
from pathlib import Path
from studio_word import atomic


class PresentationOrder:
    def __init__(self, service, row):
        self.service, self.row = service, row
        self.directory = service.root / row['id']

    def save(self):
        atomic(self.directory / 'order.json', self.row)

    def stage(self, text):
        self.row['stage'] = text
        self.save()

    def save_pptx(self, path, title, slides):
        pending = self.directory / 'pending.pptx'
        shutil.copyfile(path, pending)
        pending.replace(self.directory / 'document.pptx')
        self.row.update(pptx_ready=True, title=title, slide_count=slides)
        self.save()

    def input_data(self, data):
        data = copy.deepcopy(data)
        mapping = self.row.get('input_photos', {})
        data['user_photos'] = [mapping.get(p, p) for p in (data.get('user_photos') or [])]
        data['pres_photo_captions'] = {mapping.get(p, p): v for p, v in (data.get('pres_photo_captions') or {}).items()}
        return data

    async def deliver(self, pptx, pdf, cost, title):
        if cost != self.row['cost']:
            raise ValueError('Стоимость заказа изменилась; требуется новое подтверждение.')
        for ext, path in [('pptx', pptx), ('pdf', pdf)]:
            pending = self.directory / ('pending.' + ext)
            destination = self.directory / ('document.' + ext)
            if Path(path).resolve() != destination.resolve():
                shutil.copyfile(path, pending)
                pending.replace(destination)
        self.row.update(title=title, files_ready=True)
        self.save()
        return await self.service.deliver(self)


class PresentationOrders:
    def __init__(self, root, send, charge, afford, notify):
        self.root = Path(root)
        self.root.mkdir(parents=True, exist_ok=True, mode=0o700)
        self.send, self.charge, self.afford, self.notify = send, charge, afford, notify
        self.running = set()
        self.cleanup()
        for path in self.root.glob('*/order.json'):
            try:
                row = json.loads(path.read_text())
                if row['status'] in ('building', 'delivering'):
                    row.update(status='delivery_unknown' if row.get('sending') else 'interrupted',
                               stage='Сборка прервалась при перезапуске сервера.', restart_notice=True)
                    atomic(path, row)
            except (OSError, ValueError, KeyError):
                continue

    def cleanup(self):
        # Keep receipts, discard expired large files and personal request text.
        for path in self.root.glob('*/order.json'):
            try:
                row = json.loads(path.read_text())
                if time.time() - row.get('created', time.time()) <= 86400 or row['status'] in ('building', 'delivering', 'delivery_unknown') or row.get('billing_pending'):
                    continue
                for child in path.parent.iterdir():
                    if child.is_file() and child != path:
                        child.unlink()
                row.pop('data', None)
                row.pop('input_photos', None)
                if row['status'] != 'completed':
                    row.update(status='expired', stage='Срок хранения заказа истёк. Создайте новый план.')
                atomic(path, row)
            except (OSError, ValueError, KeyError):
                continue

    @staticmethod
    def key(uid, data):
        # Full request, not topic alone: different settings are different orders.
        fields = {k: v for k, v in data.items() if not k.startswith('_')}
        return hashlib.sha256(json.dumps([uid, fields], sort_keys=True, ensure_ascii=False).encode()).hexdigest()

    def get(self, uid, data, cost):
        self.cleanup()
        token = self.key(uid, data)
        directory = self.root / token
        path = directory / 'order.json'
        if path.exists():
            row = json.loads(path.read_text())
            if row['uid'] != uid or row['cost'] != cost:
                raise ValueError('Заказ не совпадает с подтверждённым.')
        else:
            directory.mkdir(mode=0o700)
            row = dict(id=token, uid=uid, data=copy.deepcopy(data), cost=cost,
                       status='ready', stage='Заказ сохранён', created=time.time(), sent=[])
            photos = {}
            for i, source in enumerate(data.get('user_photos') or []):
                target = directory / ('input-' + str(i) + Path(source).suffix)
                shutil.copyfile(source, target)
                photos[source] = str(target)
            row['input_photos'] = photos
            atomic(path, row)
        return PresentationOrder(self, row)

    def latest(self, uid):
        rows = []
        for path in self.root.glob('*/order.json'):
            try:
                row = json.loads(path.read_text())
                if row['uid'] == uid and row['status'] not in ('completed', 'expired', 'rejected'):
                    rows.append(row)
            except (OSError, ValueError, KeyError):
                continue
        return max(rows, key=lambda r: r['created'], default=None)

    async def notify_restarts(self):
        for path in self.root.glob('*/order.json'):
            try:
                row = json.loads(path.read_text())
                if not row.get('restart_notice'):
                    continue
                message = ('Сервер перезапустился, сборка презентации прервалась. Заказ сохранён. '
                           'Откройте меню или отправьте /continue_presentation, чтобы продолжить.')
                if row['status'] == 'delivery_unknown':
                    message = ('Связь прервалась при отправке презентации. Проверьте файлы в чате и обратитесь '
                               'в поддержку. Автоматической повторной отправки и списания не будет.')
                await self.notify(row['uid'], message)
                row.pop('restart_notice', None)
                atomic(path, row)
            except Exception:
                # Retry on next start; never log the user's request or credentials.
                continue

    async def deliver(self, job):
        row = job.row
        if row['status'] == 'completed':
            return {'delivered': True, 'charged': row.get('charged', 0)}
        if row.get('sending') or row['status'] == 'delivery_unknown':
            row.update(status='delivery_unknown', stage='Нужно уточнить отправку в поддержке.')
            job.save()
            return {'delivered': False, 'status': 'delivery_unknown'}
        ensure_generation(row.get('data', {}).get('pres_plan'))
        ensure_pptx(job.directory / 'document.pptx')
        if not row.get('billing_pending') and not self.afford(row['uid'], row['cost']):
            raise ValueError('Недостаточно кредитов. Пополните баланс и продолжите этот заказ.')
        for ext in ('pptx', 'pdf'):
            if ext in row['sent']:
                continue
            row.update(status='delivering', sending=ext, stage='Отправляем ' + ext.upper())
            job.save()
            receipt = await self.send(row['uid'], job.directory / ('document.' + ext), row.get('title', 'Презентация'), ext)
            row.setdefault('messages', {})[ext] = getattr(receipt, 'message_id', None)
            row['sent'].append(ext)
            row.pop('sending', None)
            job.save()
        row['billing_pending'] = True
        job.save()
        row['charged'] = await self.charge(row['uid'], row['id'], row['cost'], row.get('title', 'Презентация'))
        row.pop('billing_pending', None)
        row.update(status='completed', stage='Презентация готова', restart_notice=False)
        job.save()
        return {'delivered': True, 'charged': row['charged']}

    async def run(self, uid, data, cost, build):
        job = self.get(uid, data, cost)
        if job.row['status'] == 'completed':
            return {'delivered': True, 'charged': job.row.get('charged', 0)}
        if job.row['status'] in ('expired', 'delivery_unknown', 'rejected'):
            return {'delivered': False, 'status': job.row['status']}
        if uid in self.running:
            return {'delivered': False, 'status': 'building'}
        self.running.add(uid)
        try:
            if job.row.get('files_ready'):
                return await self.deliver(job)
            job.row.update(status='building', stage='Подготавливаем содержание', restart_notice=False)
            job.save()
            async with asyncio.timeout(1800):
                result = await build(job)
            if job.row['status'] != 'completed':
                job.row.update(status='delivery_unknown' if job.row.get('sending') else 'interrupted',
                               stage='Сборка не завершена. Готовые этапы сохранены.')
                job.save()
            return result if job.row['status'] == 'completed' else {'delivered': False, 'status': job.row['status']}
        except GenerationStopped as error:
            row = job.row
            row.update(status='rejected', stage=generation_notice(error), restart_notice=False)
            job.save()
            await self.notify(uid, generation_notice(error))
            return {'delivered': False, 'status': 'rejected'}
        except (Exception, asyncio.CancelledError):
            job.row.update(status='delivery_unknown' if job.row.get('sending') else 'interrupted',
                           stage='Сборка прервалась. Заказ сохранён.')
            job.save()
            try:
                await self.notify(uid, 'Сборка презентации прервалась. Откройте меню или отправьте /continue_presentation. Заказ сохранён; повторное списание защищено.')
            except Exception:
                job.row['restart_notice'] = True
                job.save()
            return {'delivered': False, 'status': job.row['status']}
        finally:
            self.running.discard(uid)
