# -*- coding: utf-8 -*-
"""Persisted View HUD settings (chips, placement, auto-start).

Stored in pyRevit user_config section ``ViewHUD``. Toggle and Settings
read the same keys so a saved layout survives sessions. ``lib/`` is
cached — reload the extension after editing this module.
"""

from pyrevit.userconfig import user_config

from revit.view_hud import HudStyle

CONFIG_SECTION = 'ViewHUD'

CHIP_PHASE = 'phase'
CHIP_WORKSET = 'workset'
CHIP_PHASE_FILTER = 'phase_filter'
CHIP_DETAIL_LEVEL = 'detail_level'
CHIP_DISPLAY_STYLE = 'display_style'
CHIP_SCOPE_BOX = 'scope_box'
CHIP_SECTION_BOX = 'section_box'

# Ribbon order left-to-right on the overlay bar.
CHIP_ORDER = (
    CHIP_PHASE,
    CHIP_WORKSET,
    CHIP_PHASE_FILTER,
    CHIP_DETAIL_LEVEL,
    CHIP_DISPLAY_STYLE,
    CHIP_SCOPE_BOX,
    CHIP_SECTION_BOX,
)

CHIP_LABELS = {
    CHIP_PHASE: 'Phase',
    CHIP_WORKSET: 'Active workset',
    CHIP_PHASE_FILTER: 'Phase filter',
    CHIP_DETAIL_LEVEL: 'Detail level',
    CHIP_DISPLAY_STYLE: 'Visual style',
    CHIP_SCOPE_BOX: 'Scope box',
    CHIP_SECTION_BOX: 'Section box',
}

CHIP_HINTS = {
    CHIP_PHASE: 'Click the name to cycle. Caret opens the phase list.',
    CHIP_WORKSET: 'Active workset for new elements (workshared models).',
    CHIP_PHASE_FILTER: 'View phase filter (Show All, Show Previous + New, …).',
    CHIP_DETAIL_LEVEL: 'Coarse / Medium / Fine.',
    CHIP_DISPLAY_STYLE: 'Wireframe, hidden line, shaded, realistic, …',
    CHIP_SCOPE_BOX: 'View scope box. None means no scope box assigned.',
    CHIP_SECTION_BOX: '3D section box on/off.',
}

_CHIP_DEFAULTS = {
    CHIP_PHASE: True,
    CHIP_WORKSET: True,
    CHIP_PHASE_FILTER: False,
    CHIP_DETAIL_LEVEL: False,
    CHIP_DISPLAY_STYLE: False,
    CHIP_SCOPE_BOX: True,
    CHIP_SECTION_BOX: False,
}

ANCHORS = ('top-left', 'top-center', 'top-right')
ANCHOR_LABELS = {
    'top-left': 'Top left',
    'top-center': 'Top middle',
    'top-right': 'Top right',
}

DEFAULT_ANCHOR = 'top-center'
DEFAULT_IDLE_OPACITY = 0.5
OPACITY_CHOICES = (0.35, 0.5, 0.75, 1.0)


def _as_bool(value, default=False):
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in ('1', 'true', 'yes', 'on'):
        return True
    if text in ('0', 'false', 'no', 'off', ''):
        return False
    return default


def _as_float(value, default):
    try:
        return float(value)
    except Exception:
        return default


def _section():
    if not hasattr(user_config, CONFIG_SECTION):
        user_config.add_section(CONFIG_SECTION)
    return getattr(user_config, CONFIG_SECTION)


def load_hud_config():
    """Dict: chips {id: bool}, anchor, idle_opacity, auto_start."""
    section = _section()
    chips = {}
    for chip_id in CHIP_ORDER:
        key = 'show_{}'.format(chip_id)
        chips[chip_id] = _as_bool(
            section.get_option(key, default_value=_CHIP_DEFAULTS[chip_id]),
            _CHIP_DEFAULTS[chip_id])
    anchor = section.get_option('anchor', default_value=DEFAULT_ANCHOR)
    if anchor not in ANCHORS:
        anchor = DEFAULT_ANCHOR
    idle = _as_float(
        section.get_option('idle_opacity', default_value=DEFAULT_IDLE_OPACITY),
        DEFAULT_IDLE_OPACITY)
    if idle < 0.15:
        idle = 0.15
    if idle > 1.0:
        idle = 1.0
    auto_start = _as_bool(
        section.get_option('auto_start', default_value=False), False)
    return {
        'chips': chips,
        'anchor': anchor,
        'idle_opacity': idle,
        'auto_start': auto_start,
    }


def save_hud_config(cfg):
    section = _section()
    chips = cfg.get('chips') or {}
    for chip_id in CHIP_ORDER:
        section.set_option(
            'show_{}'.format(chip_id),
            bool(chips.get(chip_id, _CHIP_DEFAULTS[chip_id])))
    anchor = cfg.get('anchor', DEFAULT_ANCHOR)
    if anchor not in ANCHORS:
        anchor = DEFAULT_ANCHOR
    section.set_option('anchor', anchor)
    idle = _as_float(cfg.get('idle_opacity'), DEFAULT_IDLE_OPACITY)
    section.set_option('idle_opacity', idle)
    section.set_option('auto_start', bool(cfg.get('auto_start', False)))
    user_config.save_changes()


def enabled_chip_ids(cfg=None):
    cfg = cfg or load_hud_config()
    chips = cfg.get('chips') or {}
    return [chip_id for chip_id in CHIP_ORDER if chips.get(chip_id)]


def hud_style_from_config(cfg=None):
    cfg = cfg or load_hud_config()
    return HudStyle(
        anchor=cfg.get('anchor', DEFAULT_ANCHOR),
        idle_opacity=_as_float(
            cfg.get('idle_opacity'), DEFAULT_IDLE_OPACITY),
        hover_opacity=1.0,
    )


def is_auto_start_enabled():
    return bool(load_hud_config().get('auto_start'))
