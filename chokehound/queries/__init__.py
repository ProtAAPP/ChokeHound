"""
Query modules for choke points analysis.

This package contains query definitions for choke points analysis.
"""

from .registry import QueryRegistry, SecurityQuery

__all__ = ['QueryRegistry', 'SecurityQuery']

# Import choke points query module to register it
from . import choke_points

