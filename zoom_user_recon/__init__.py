"""
zoom_user_recon
---------------
Zoom User Status Audit engine.
Exposes two entry points for xlwings RunPython calls from Excel.
"""

from .recon import run_zoom_user_audit, clear_zoom_results

__version__ = "1.0.0"
__all__ = ["run_zoom_user_audit", "clear_zoom_results"]
