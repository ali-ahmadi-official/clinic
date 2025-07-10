from django import template

register = template.Library()

@register.filter
def split(value, key):
    return value.split(key)

@register.filter
def trim(value):
    if isinstance(value, str):
        return value.strip()
    return value
