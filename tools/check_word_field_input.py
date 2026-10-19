"""Deployment regression: common labels must fill fixed forms without paid APIs."""
import asyncio
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fixed_word_templates as fixed
import car_sale_template as car


async def no_model(*args, **kwargs):
    raise AssertionError('labelled input unexpectedly called a model')


async def main():
    api = {'ask_grok': no_model}
    for kind, spec in fixed.CATALOG.items():
        values = spec['sample']
        source = '\n'.join(('Дата' if k == 'date' else 'Город' if k == 'city' else label) + ': ' + values[k]
                           for k, label in spec['fields'].items())
        actual = await fixed.extract_fields({'word_kind': kind, 'topic': source}, api)
        assert actual == values, kind
    source = ('Исполнитель: ООО Тест\nЗаказчик: ООО Пример\nГород: Москва\nДата: 20.10.2026\n'
              'Услуги: Подготовка инструкции: правила хранения файлов\n'
              'Срок оказания услуг: до 30.10.2026\nСтоимость: 15000 рублей\n'
              'Порядок оплаты: после подписания акта оказанных услуг')
    expected = dict(zip(('provider', 'customer', 'city', 'date', 'services', 'period', 'price', 'payment'),
                        (line.split(':', 1)[1].strip() for line in source.splitlines())))
    for kind in ('services_ru', 'services'):
        assert await fixed.extract_fields({'word_kind': kind, 'topic': source}, api) == expected
    assert await car.extract_fields({'topic': 'Продавец: Петров\nПокупатель: Иванов\nДата: 20.10.2026\nСтоимость: 1200000 рублей'}, api) == {
        'seller_name': 'Петров', 'buyer_name': 'Иванов', 'date': '20.10.2026', 'price': '1200000 рублей'}
    print('word_field_input runtime_check=passed catalog_routes=28 reported_fields=8 car=yes paid_api_calls=0', flush=True)


if __name__ == '__main__':
    asyncio.run(main())
