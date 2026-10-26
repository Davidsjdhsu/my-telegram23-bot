"""Password for worksheet editing (not encryption). Never persist plaintext."""
import os
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from openpyxl.utils.protection import hash_password


def protection(body, *, draft=False):
    enabled = body.get('excel_protect') is True
    if not enabled:
        return {'excel_protect': False, 'excel_password_hash': None}
    password = body.get('excel_password', '')
    if not isinstance(password, str) or len(password) > 64 or any(ord(c) < 32 for c in password):
        raise ValueError('Пароль: от 4 до 64 символов, без переносов строк.')
    value = body.get('excel_password_hash')
    if password:
        if len(password) < 4:
            if draft:
                return {'excel_protect': True, 'excel_password_hash': None}
            raise ValueError('Добавьте пароль: от 4 до 64 символов.')
        value = hash_password(password)
    if value is not None and (not isinstance(value, str) or not re.fullmatch(r'[0-9A-F]{1,4}', value)):
        raise ValueError('Укажите пароль защиты заново.')
    if not value and not draft:
        raise ValueError('Добавьте пароль для защиты формул.')
    return {'excel_protect': True, 'excel_password_hash': value}


def protect_file(path, password_hash):
    """Patch only sheetProtection: preserve cached formula values and styles."""
    if not password_hash:
        return
    if not re.fullmatch(r'[0-9A-F]{1,4}', password_hash):
        raise ValueError('Неверная настройка защиты.')
    path = Path(path)
    pending = path.with_suffix('.protection.tmp')
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    try:
        with zipfile.ZipFile(path) as source, zipfile.ZipFile(pending, 'w', zipfile.ZIP_DEFLATED) as target:
            for item in source.infolist():
                raw = source.read(item.filename)
                if re.fullmatch(r'xl/worksheets/sheet\d+\.xml', item.filename):
                    root = ET.fromstring(raw)
                    node = root.find(ns + 'sheetProtection')
                    if node is None or node.get('sheet') not in ('1', 'true'):
                        raise ValueError('Защита листа не включена.')
                    node.set('password', password_hash)
                    raw = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                target.writestr(item, raw)
        os.replace(pending, path)
    finally:
        pending.unlink(missing_ok=True)
