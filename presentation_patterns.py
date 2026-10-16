"""Topic-matched, bundled artwork: fifty distinct motifs in each designer style.

Images are generated once during development, never through a paid runtime call.
The confirmed design kit is stable across all slides and retries of a deck.
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parent
STYLES = {'designer_soft': 'soft', 'designer_strict': 'strict'}


def pattern_node(design, style, number, *, custom=False):
    prefix = STYLES.get(style)
    key = design.get('template_id', '')
    if not prefix or key not in {f'd{i:02}' for i in range(1, 51)}:
        return None
    path = ROOT / 'assets' / 'patterns' / f'{prefix}-{key}.jpg'
    if not path.is_file():
        return None
    # Keep artwork upright. Horizontal mirroring is safe for these text-free motifs.
    # A stable deck seed changes placement without changing the chosen subject.
    flip = (int(design.get('variant', 0)) + (number - 1) // 2) % 2 == 1
    node = dict(type='background_image', path=str(path), box=[0, 0, 1280, 720],
                origin='decoration', radius=0, crop=False, flip_h=flip,
                pattern_id=f'{prefix}-{key}', opacity=.48)
    if custom:
        # Respect an explicitly requested palette: pattern is a restrained overlay.
        node['opacity'] = .18
    return node


def add_pattern(scene, design, style, number, *, custom=False):
    node = pattern_node(design, style, number, custom=custom)
    if node is None:
        return False
    scene['nodes'][:] = [n for n in scene['nodes'] if n['type'] != 'background_image']
    scene['nodes'].insert(0, node)
    scene['pattern_id'] = node['pattern_id']
    return True
