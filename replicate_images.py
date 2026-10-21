"""Bounded waits that preserve a prediction across presentation retries.

Only an ID, request fingerprint and coarse status enter the checkpoint; never
credentials, provider responses or signed output URLs. Creation POSTs are not
retried automatically because an interrupted response can hide a paid job.
"""
import asyncio
import hashlib
import json
import logging
from pathlib import Path
import re
import time

import httpx

log = logging.getLogger(__name__)
WAIT_SECONDS = 180
POLL_SECONDS = 2
API = 'https://api.replicate.com/v1'
TERMINAL = {'failed', 'canceled', 'aborted'}


async def generate_prediction(prompt, path, aspect_ratio, token, job=None):
    job = job if job is not None else {}
    fingerprint = hashlib.sha256(json.dumps([prompt, aspect_ratio]).encode()).hexdigest()
    if job.get('fingerprint') != fingerprint:
        job.clear()
    job['fingerprint'] = fingerprint
    started = time.monotonic()
    resumed = bool(job.get('id'))
    headers = {'Authorization': f'Bearer {token}'}

    async def get_existing(http, url, **kwargs):
        for attempt in range(3):
            try:
                result = await http.get(url, timeout=20, **kwargs)
                if result.status_code == 429 or result.status_code >= 500:
                    if attempt < 2:
                        await asyncio.sleep(attempt + 1)
                        continue
                result.raise_for_status()
                return result
            except httpx.TransportError:
                if attempt == 2:
                    raise
                await asyncio.sleep(attempt + 1)

    try:
        async with asyncio.timeout(WAIT_SECONDS):
            async with httpx.AsyncClient(timeout=30) as http:
                if not job.get('id'):
                    response = await http.post(
                        API + '/models/black-forest-labs/flux-schnell/predictions',
                        headers=headers, json={'input': {'prompt': prompt,
                            'aspect_ratio': aspect_ratio, 'output_format': 'png'}})
                    response.raise_for_status()
                    info = response.json()
                    prediction_id = info.get('id')
                    if not isinstance(prediction_id, str) or not re.fullmatch(r'[A-Za-z0-9_-]{1,128}', prediction_id):
                        raise ValueError('Invalid prediction ID')
                    job['id'] = prediction_id
                else:
                    prediction_id = job['id']
                    if not isinstance(prediction_id, str) or not re.fullmatch(r'[A-Za-z0-9_-]{1,128}', prediction_id):
                        job.pop('id', None)
                        raise ValueError('Invalid saved prediction ID')
                    info = None
                # Build the authenticated URL ourselves, never from model/user data.
                status_url = API + '/predictions/' + prediction_id
                while True:
                    if info is None:
                        try:
                            info = (await get_existing(http, status_url, headers=headers)).json()
                        except httpx.HTTPStatusError as error:
                            if error.response.status_code in (404, 410):
                                job.pop('id', None)
                                job['reason'] = 'expired'
                                return False
                            raise
                    status = info.get('status')
                    job['status'] = status if status in TERMINAL | {'starting', 'processing', 'succeeded'} else 'unknown'
                    if status in TERMINAL:
                        job.pop('id', None)
                        job['reason'] = 'generation_failed'
                        return False
                    if status == 'succeeded':
                        output = info.get('output')
                        url = output[0] if isinstance(output, list) and output else output
                        if not isinstance(url, str) or not url.startswith('https://'):
                            job.pop('id', None)
                            job['reason'] = 'expired'
                            return False
                        # Do not send the API credential to the output host.
                        try:
                            raw = await get_existing(http, url)
                        except httpx.HTTPStatusError as error:
                            if error.response.status_code in (404, 410):
                                job.pop('id', None)
                                job['reason'] = 'expired'
                                return False
                            raise
                        Path(path).write_bytes(raw.content)
                        job['reason'] = 'ready'
                        return True
                    await asyncio.sleep(POLL_SECONDS)
                    info = None
    except TimeoutError:
        job['reason'] = 'waiting_timeout'
    except asyncio.CancelledError:
        job['reason'] = 'waiting_timeout'
        raise
    except Exception:
        job['reason'] = 'generation_unavailable'
    finally:
        log.info('image_provider seconds=%.2f resumed=%s status=%s reason=%s',
                 time.monotonic() - started, resumed, job.get('status', 'unknown'),
                 job.get('reason', 'unknown'))
    return False
