"""Bounded designer choices shared by previews and the real PPTX renderer."""
DIAGRAM_LAYOUTS = ('classic', 'cards', 'steps')
BACKGROUND_MODES = ('solid', 'gradient')

def options(body):
    background = body.get('background_mode', 'solid')
    diagram = body.get('diagram_layout', 'classic')
    if background not in BACKGROUND_MODES or diagram not in DIAGRAM_LAYOUTS:
        raise ValueError('Unknown designer option')
    return dict(background_mode=background, diagram_layout=diagram)

def finish_scene(scene, design):
    if design.get('background_mode') != 'gradient':
        return scene
    from presentation_profiles import relative_luminance
    bg = scene['background']
    # Shade toward the same light/dark pole: keep existing text contrast.
    end = '#000000' if relative_luminance(bg) < .3 else '#FFFFFF'
    # Keep all original artwork. Apply a translucent gradient only over the
    # background stack, before photos, diagrams and any readable text.
    at = 0
    for i, n in enumerate(scene['nodes']):
        if n['type'] == 'background_image' or (n['type']=='panel' and list(n['box'])==[0,0,1280,720]):
            at = i+1
        else:
            break
    scene['nodes'].insert(at, {'type':'panel','box':[0,0,1280,720], 'fill':bg, 'radius':0,
        'gradient':{'angle':135,'stops':[(0,end,0),(1,end,.42)]}, 'role':'background_gradient'})
    return scene
