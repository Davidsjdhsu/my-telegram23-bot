"""Server-authoritative manual constructor tariff; no AI generation surcharge."""
VERSION = 'constructor-10-3-2-cap20-v2'
BASE = 10
EXTRA_PHOTO = 3
PATTERN = 2


def quote(project):
    if not isinstance(project, dict) or project.get('version') != 1:
        raise ValueError('Неверный проект.')
    slides = project.get('slides')
    if not isinstance(slides, list) or not 1 <= len(slides) <= 20:
        raise ValueError('Нужно от 1 до 20 слайдов.')
    prices = []
    for i, slide in enumerate(slides, 1):
        elements = slide.get('elements') if isinstance(slide, dict) else None
        if not isinstance(elements, list) or len(elements) > 50:
            raise ValueError('Неверные элементы слайда.')
        counts = dict.fromkeys(('photo', 'pattern', 'text', 'diagram'), 0)
        for element in elements:
            kind = element.get('type') if isinstance(element, dict) else None
            if kind not in counts:
                raise ValueError('Неизвестный элемент.')
            counts[kind] += 1
        if counts['photo'] > 3 or counts['pattern'] > 3:
            raise ValueError(f'Слайд {i}: максимум 3 фотографии и 3 узора.')
        if counts['text'] > 6 or counts['diagram'] > 1:
            raise ValueError(f'Слайд {i}: максимум 6 текстовых блоков и 1 схема.')
        prices.append(min(20, BASE + EXTRA_PHOTO * max(0, counts['photo'] - 1) + PATTERN * counts['pattern']))
    return {'version': VERSION, 'slides': prices, 'cost': sum(prices)}
