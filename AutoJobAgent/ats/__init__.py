"""
ATS package initialization.
Provides functions to detect and get submitters for various ATS systems.
"""

import re
from typing import Dict, Optional, Union

from playwright.sync_api import BrowserContext

# Import ATS submitters
from .base_submitter import BaseSubmitter
from .workday import WorkdaySubmitter
from .icims import ICIMSSubmitter
from .greenhouse import GreenhouseSubmitter

# Try to import Lever if available (marked as future stub)
try:
    from .lever import LeverSubmitter
    LEVER_AVAILABLE = True
except ImportError:
    LEVER_AVAILABLE = False

# URL patterns for ATS detection
ATS_PATTERNS = {
    "workday": [
        r"myworkdayjobs\.com",
        r"workday\.com",
        r"wd3\.myworkdayjobs\.com",
        r"workdayjobs\.com"
    ],
    "icims": [
        r"icims\.com",
        r"jobs\.icims\.com",
        r"careers\.icims\.com"
    ],
    "greenhouse": [
        r"greenhouse\.io",
        r"boards\.greenhouse\.io",
        r"app\.greenhouse\.io"
    ],
    "lever": [
        r"lever\.co",
        r"jobs\.lever\.co",
        r"careers\.lever\.co"
    ]
}

# Registry of ATS submitters
ATS_SUBMITTERS = {
    "workday": WorkdaySubmitter,
    "icims": ICIMSSubmitter,
    "greenhouse": GreenhouseSubmitter
}

# Register Lever if available
if LEVER_AVAILABLE:
    ATS_SUBMITTERS["lever"] = LeverSubmitter


def detect(url: str) -> str:
    """
    Detect the ATS system from a job URL.
    
    Args:
        url: Job posting URL
        
    Returns:
        ATS system name or "unknown"
    """
    if not url:
        return "unknown"
    
    url = url.lower()
    
    for ats_name, patterns in ATS_PATTERNS.items():
        for pattern in patterns:
            if re.search(pattern, url):
                return ats_name
    
    return "unknown"


def get_submitter(ats_name: str, browser_context: BrowserContext) -> BaseSubmitter:
    """
    Get a submitter instance for the specified ATS.
    
    Args:
        ats_name: ATS system name
        browser_context: Playwright browser context
        
    Returns:
        ATS submitter instance
        
    Raises:
        ValueError: If ATS is unknown or not supported
    """
    if ats_name == "unknown":
        raise ValueError("Unknown ATS system")
    
    if ats_name == "lever" and not LEVER_AVAILABLE:
        raise ValueError("Lever ATS support is not available")
    
    if ats_name not in ATS_SUBMITTERS:
        raise ValueError(f"Unsupported ATS system: {ats_name}")
    
    # Create and return the submitter instance
    return ATS_SUBMITTERS[ats_name](browser_context)


# Additional helper functions

def register_submitter(ats_name: str, submitter_class) -> None:
    """
    Register a new ATS submitter.
    
    Args:
        ats_name: ATS system name
        submitter_class: Submitter class
    """
    ATS_SUBMITTERS[ats_name] = submitter_class


def get_supported_ats() -> list:
    """
    Get a list of supported ATS systems.
    
    Returns:
        List of supported ATS system names
    """
    return list(ATS_SUBMITTERS.keys())
