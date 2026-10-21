"""Private, bounded checkpoints for an interrupted illustrated presentation.

Only accepted content and accepted pictures are retained. The input fingerprint
includes the confirmed plan and photo selections; a changed order cannot reuse
an earlier order's content or consent. No provider credentials are stored.
"""
import hashlib
import json
import logging
import os
from pathlib import Path
import shutil
import time

log = logging.getLogger(__name__)
VERSION = 1
TTL = 24 * 3600
MAX_BYTES = 128 * 1024 * 1024
PARTIAL_LABELS = {'ru': '📄 Собрать без этих картинок', 'en': '📄 Build without these images'}
RESUME_LABELS = {'ru': '🔄 Продолжить с картинками', 'en': '🔄 Continue with images'}


def missing_images(slides, assets):
    return [i + 1 for i, s in enumerate(slides)
            if not assets[i] and not s.get('chart') and not s.get('_skip_ai_image')]


def recovery_notice(missing, lang, cost, slides=None):
    numbers = ', '.join(map(str, missing))
    details = []
    labels = {
        'waiting_timeout': ('Не дождались готового изображения', 'Image wait timed out'),
        'generation_unavailable': ('Не удалось получить изображение от сервиса', 'Image service unavailable'),
        'generation_failed': ('Сервис сообщил об остановке создания изображения', 'Image generation stopped'),
        'expired': ('Сохранённый результат сервиса уже недоступен', 'Provider result expired'),
        'invalid_image': ('Сервис вернул повреждённое изображение', 'Invalid image received'),
        'review_unavailable': ('Не удалось завершить проверку изображения', 'Image review unavailable'),
        'requires_verified_source': ('Нужна достоверная иллюстрация; созданное изображение не подходит', 'A verified illustration is required'),
        'quality_rejected': ('Изображение не прошло проверку качества', 'Image did not pass quality review'),
    }
    groups = {}
    for number in missing:
        reason = slides[number - 1].get('_image_issue') if slides and number <= len(slides) else None
        if reason in labels:
            groups.setdefault(reason, []).append(str(number))
    for reason, group in groups.items():
        details.append(labels[reason][0 if lang == 'ru' else 1] + ': ' + ', '.join(group) + '.')
    detail = ('\n' + '\n'.join(details)) if details else ''
    if lang == 'ru':
        return (f'Пока не готовы изображения для слайдов: {numbers}. Кредиты не списаны.{detail}\n'
                'Текст и удачные картинки сохранены на срок до 24 часов. '
                '«Продолжить с картинками» продолжит только недостающие изображения: '
                'сначала проверит сохранённые задания сервиса.\n\n'
                'Или выбери «Собрать без этих картинок»: все слайды и текст сохранятся, '
                'на перечисленных слайдах будет оформление без фото. '
                f'Стоимость этой версии — {cost} кредитов, как в подтверждённом плане. '
                'Списание — после отправки PPTX и PDF.')
    return (f'Images for slides {numbers} are not ready. No credits charged.{detail}\n'
            'Text and accepted images are saved for up to 24 hours. Continue with images '
            'resumes only missing pictures, checking saved provider jobs first. '
            'Or choose Build without these images: all slides and text remain, '
            'with text layouts for the listed slides. '
            f'The price remains {cost} credits, charged after both PPTX and PDF are sent.')


class PresentationRecovery:
    def __init__(self, uid, data, root):
        fields = {k: data.get(k) for k in (
            'topic', 'pres_topic', 'user_text', 'extra', 'slides', 'mode', 'content_lang',
            'theme_name', 'custom_colors', 'photo_mode', 'user_photos', 'pres_photo_captions',
            'pres_photo_context', 'photo_album_only', 'pres_plan', 'pres_request_id')}
        payload = json.dumps([VERSION, str(uid), fields], ensure_ascii=False, sort_keys=True)
        self.key = hashlib.sha256(payload.encode()).hexdigest()
        self.root = Path(root)
        self.directory = self.root / self.key

    def load(self, destination):
        try:
            path = self.directory / 'checkpoint.json'
            if not path.is_file() or time.time() - path.stat().st_mtime > TTL:
                return None
            if path.stat().st_size > 2_000_000:
                return None
            value = json.loads(path.read_text())
            if value.get('version') != VERSION or value.get('key') != self.key:
                return None
            slides, assets = value['slides'], value['assets']
            if not 3 <= len(slides) <= 30 or len(assets) != len(slides):
                return None
            for i, asset in enumerate(assets):
                if not asset:
                    continue
                source = self.directory / f'asset-{i}.jpg'
                raw = source.read_bytes()
                if hashlib.sha256(raw).hexdigest() != asset.pop('file_hash'):
                    return None
                target = Path(destination) / f'resumed-{i}.jpg'
                target.write_bytes(raw)
                asset['path'] = str(target)
            value['missing'] = missing_images(slides, assets)
            return value
        except (OSError, ValueError, TypeError, KeyError):
            log.warning('presentation_recovery reason=checkpoint_unavailable')
            return None

    def save(self, plan, content, slides, assets, warnings):
        try:
            self.root.mkdir(parents=True, exist_ok=True, mode=0o700)
            now = time.time()
            entries = []
            for path in self.root.glob('*/checkpoint.json'):
                if path.parent == self.directory:
                    continue
                if now - path.stat().st_mtime > TTL:
                    shutil.rmtree(path.parent, ignore_errors=True)
                else:
                    entries.append((path.stat().st_mtime, path.parent,
                                    sum(f.stat().st_size for f in path.parent.iterdir() if f.is_file())))
            incoming = sum(Path(a['path']).stat().st_size for a in assets if a)
            total = incoming + sum(x[2] for x in entries)
            for _, directory, size in sorted(entries):
                if total <= MAX_BYTES:
                    break
                shutil.rmtree(directory, ignore_errors=True)
                total -= size
            if incoming > MAX_BYTES:
                return False
            self.directory.mkdir(exist_ok=True, mode=0o700)
            stored = []
            for i, asset in enumerate(assets):
                if asset is None:
                    stored.append(None)
                    continue
                raw = Path(asset['path']).read_bytes()
                target = self.directory / f'asset-{i}.jpg'
                target.write_bytes(raw)
                stored.append({**asset, 'path': target.name, 'file_hash': hashlib.sha256(raw).hexdigest()})
            value = dict(version=VERSION, key=self.key, plan=plan, content=content,
                         slides=slides, assets=stored, warnings=warnings)
            temporary = self.directory / 'checkpoint.tmp'
            with temporary.open('w') as out:
                json.dump(value, out, ensure_ascii=False)
                out.flush()
                os.fsync(out.fileno())
            os.replace(temporary, self.directory / 'checkpoint.json')
            return True
        except (OSError, ValueError, TypeError, KeyError):
            log.warning('presentation_recovery reason=checkpoint_save_failed')
            return False

    def clear(self):
        shutil.rmtree(self.directory, ignore_errors=True)
