"""
zp_user_recon
-------------
Zoom Phone User Reconciliation engine.
Exposes one entry point for xlwings RunPython calls from Excel.
"""

from .recon import run_zp_reconciliation

__version__ = "1.0.0"
__all__ = ["run_zp_reconciliation"]
