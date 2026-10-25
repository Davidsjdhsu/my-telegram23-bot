"""Fail the image build if deployed preview assets cannot produce real scenes."""
import sys
from pathlib import Path
from xml.etree import ElementTree
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from studio_palette_preview import render_preview, preview_photo
from PIL import Image
with Image.open(preview_photo()) as photo:
    photo.verify()
for style in ('strict','soft'):
    for page in range(4):
        svg=render_preview(style,'d42','wine',page)
        ElementTree.fromstring(svg)
        if page in (0,2):
            assert '/studio/constructor/coffee.jpg' in svg
        assert '/preview-pattern/'+style+'-d42.jpg' in svg
print('Designer palette runtime: 8 scenes and bundled photograph OK')

for layout in ('classic','cards','steps'):
    svg=render_preview('soft','d42','wine',3,'gradient',layout)
    ElementTree.fromstring(svg)
    assert 'linearGradient' in svg
print('Designer choices: diagram variants and gradient OK')
