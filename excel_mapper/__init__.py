"""
excel_mapper - A Python library for mapping Excel columns to object attributes.

Install with: pip install excel-mapper
Import with: from excel_mapper import ExcelMapper
"""

from .core import ExcelMapper
from .version import __version__

__all__ = ["ExcelMapper", "__version__"]
