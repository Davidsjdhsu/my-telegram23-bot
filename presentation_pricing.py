"""Presentation prices shared by balance checks, quotes and final billing."""
from presentation_designer import is_designer

PRESENTATION_CREDITS_PER_SLIDE = 5
DESIGNER_CREDITS_PER_SLIDE = 10
PRESENTATION_MIN_SLIDES = 3


def presentation_rate(data=None):
    context = data if isinstance(data, dict) else {'theme_name': data}
    return DESIGNER_CREDITS_PER_SLIDE if is_designer(context) else PRESENTATION_CREDITS_PER_SLIDE


def presentation_cost(slides, data=None):
    try:
        count = max(PRESENTATION_MIN_SLIDES, int(slides))
    except (TypeError, ValueError):
        count = PRESENTATION_MIN_SLIDES
    return count * presentation_rate(data)


def presentation_price_notice(slides, data=None, lang='ru'):
    rate = presentation_rate(data)
    cost = presentation_cost(slides, data)
    templates = {
        'ru': 'Стоимость: {cost} кредитов ({rate} за слайд, включая обложку). PPTX и PDF входят в стоимость.',
        'en': 'Price: {cost} credits ({rate} per slide, including the cover). PPTX and PDF are included.',
        'de': 'Preis: {cost} Credits ({rate} pro Folie, inklusive Titelfolie). PPTX und PDF sind enthalten.',
        'ar': 'السعر: {cost} رصيد ({rate} لكل شريحة، بما فيها الغلاف). يشمل PPTX وPDF.',
        'zh': '费用：{cost}积分（每页{rate}积分，含封面）。包含PPTX和PDF。',
        'es': 'Precio: {cost} créditos ({rate} por diapositiva, incluida la portada). Incluye PPTX y PDF.',
        'fr': 'Prix : {cost} crédits ({rate} par diapositive, couverture comprise). PPTX et PDF inclus.',
    }
    return templates.get(lang, templates['en']).format(cost=cost, rate=rate)
