"""
dashboard
---------
Dashboard builder for the deployment tracking workbook.
Builds and refreshes the Dashboard 2 sheet with live stats from all tool sheets.
"""

from .builder import build_dashboard, refresh_ca_block

__version__ = "1.0.0"
__all__ = ["build_dashboard", "refresh_ca_block"]
