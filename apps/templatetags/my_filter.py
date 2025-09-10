import os

from django import template
from datetime import timedelta


register = template.Library()


@register.filter
def filename(value):
    return os.path.basename(value.file.name)


@register.filter
def to_space(value):
    return value.replace('%20', ' ').replace('25', '')


@register.filter
def format_duration_hhmm(duration):
    if not isinstance(duration, timedelta):
        return ""
    total_seconds = int(duration.total_seconds())
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    return f"{hours}j {minutes:02d}m"
