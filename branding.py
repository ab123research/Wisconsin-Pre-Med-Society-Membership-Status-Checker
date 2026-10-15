"""Self-contained Wisconsin Pre-Medical Society presentation assets."""

from base64 import b64encode
from functools import lru_cache
from html import escape
from pathlib import Path

ASSETS = Path(__file__).resolve().parent / "assets"


@lru_cache(maxsize=1)
def logo_uri() -> str:
    """Use the supplied society logo exactly as provided."""
    data = b64encode((ASSETS / "wps-logo.png").read_bytes()).decode("ascii")
    return "data:image/png;base64," + data


@lru_cache(maxsize=1)
def brand_css() -> str:
    css = (ASSETS / "theme.css").read_text(encoding="utf-8")
    for token, filename, mime in (
        ("__CONTOURS__", "contours.svg", "image/svg+xml"),
        ("__DISPLAY_FONT__", "fonts/RedHatDisplay.ttf", "font/ttf"),
        ("__TEXT_FONT__", "fonts/RedHatText.ttf", "font/ttf"),
    ):
        data = b64encode((ASSETS / filename).read_bytes()).decode("ascii")
        css = css.replace(token, f"data:{mime};base64,{data}")
    return "<style>" + css + "</style>"


def hero_html(term: str) -> str:
    art = (ASSETS / "medical-art.svg").read_text(encoding="utf-8")
    return (
        '<header class="wps-brandbar">'
        f'<div class="wps-brand"><img class="wps-logo" src="{logo_uri()}" '
        'alt="Wisconsin Pre-Medical Society logo" width="104" height="104">'
        '<div><div class="wps-brand-name">Wisconsin Pre-Medical Society</div>'
        '<div class="wps-university">University of Wisconsin–Madison</div></div></div>'
        '<span class="wps-portal-label">Member portal</span></header>'
        '<section class="wps-hero" aria-labelledby="wps-title">'
        '<div class="wps-hero-copy">'
        f'<div class="wps-term"><span aria-hidden="true"></span>{escape(term)}</div>'
        '<h1 id="wps-title">Your path to<br>active membership.</h1>'
        '<p>Your status, your involvement, and your next steps.<br class="wps-desktop-break"> '
        'All in one place.</p></div>'
        f'<div class="wps-hero-art" aria-hidden="true">{art}</div></section>'
        '<div class="wps-requirement-strip" aria-label="Active membership requirements">'
        '<div><span class="wps-strip-number">03</span><span>General meetings</span></div>'
        '<div><span class="wps-strip-number">01</span><span>Event per color</span></div>'
        '<div><span class="wps-strip-check" aria-hidden="true">✓</span><span>Semester dues paid</span></div>'
        '</div>'
    )


def form_heading_html() -> str:
    return ('<div class="wps-form-heading"><span class="wps-eyebrow">Member lookup</span>'
            '<h2>Check your membership</h2>'
            '<p>Enter your full name as it appears on the membership sheet.</p></div>')


def footer_html(term: str) -> str:
    return (f'<footer class="wps-footer"><img class="wps-footer-logo" src="{logo_uri()}" '
            'alt="" width="64" height="64">'
            '<div><strong>Wisconsin Pre-Medical Society</strong>'
            f'<span>UW–Madison · {escape(term)}</span></div>'
            '<span class="wps-footer-end">Member portal</span></footer>')
