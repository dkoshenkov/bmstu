from django import template
import os

register = template.Library()


@register.filter(name='endswith')
def endswith(value, arg):
    """Check if value ends with the specified argument (arg)."""
    return value.lower().endswith(arg)


@register.filter(name='basename')
def basename(value):
    """Returns the base name of a file path."""
    return os.path.basename(value)