"""
Base submitter class for ATS integrations.
All ATS-specific submitters should inherit from this class.
"""

from abc import ABC, abstractmethod
from typing import Dict, Optional, Union

from playwright.sync_api import BrowserContext, Page, TimeoutError as PlaywrightTimeoutError
from rich.console import Console

# Import utils for common functionality
import utils

console = Console()

class BaseSubmitter(ABC):
    """
    Base class for all ATS submitters.
    Defines the common interface and utility methods for ATS automation.
    """
    
    def __init__(self, browser_context: BrowserContext):
        """
        Initialize the submitter with a browser context.
        
        Args:
            browser_context: Playwright browser context
        """
        self.ctx = browser_context
    
    @abstractmethod
    def submit(self, job: Dict, profile: Dict, resume_path: str, cover_letter_path: str) -> str:
        """
        Submit an application to an ATS.
        Must be implemented by subclasses.
        
        Args:
            job: Job dictionary with details
            profile: User profile dictionary
            resume_path: Path to the resume file
            cover_letter_path: Path to the cover letter file
            
        Returns:
            Status string (e.g., "Applied", "Failed", "Manual")
        """
        pass
    
    def wait_for_navigation(self, page: Page, timeout: int = 30000) -> bool:
        """
        Wait for page navigation to complete.
        
        Args:
            page: Playwright page
            timeout: Timeout in milliseconds
            
        Returns:
            True if navigation completed, False otherwise
        """
        try:
            page.wait_for_load_state("networkidle", timeout=timeout)
            return True
        except PlaywrightTimeoutError:
            console.print("[yellow]Navigation timeout, continuing anyway[/yellow]")
            return False
    
    def check_for_captcha(self, page: Page) -> bool:
        """
        Check if a CAPTCHA is present on the page.
        
        Args:
            page: Playwright page
            
        Returns:
            True if CAPTCHA detected, False otherwise
            
        Raises:
            NeedsHumanException: If CAPTCHA is detected
        """
        captcha_selectors = [
            "iframe[src*=captcha]",
            "iframe[src*=recaptcha]",
            "div.g-recaptcha",
            "div[class*=captcha]",
            "input[name*=captcha]"
        ]
        
        for selector in captcha_selectors:
            if page.is_visible(selector):
                console.print("[bold red]CAPTCHA detected![/bold red]")
                raise utils.NeedsHumanException("CAPTCHA detected, human intervention required")
        
        return False
    
    def fill_form_fields(self, page: Page, field_mappings: Dict[str, str]) -> int:
        """
        Fill form fields based on mappings.
        
        Args:
            page: Playwright page
            field_mappings: Dictionary mapping selectors to values
            
        Returns:
            Number of fields successfully filled
        """
        fields_filled = 0
        
        for selector, value in field_mappings.items():
            try:
                if utils.fill_if_empty(page, selector, value):
                    fields_filled += 1
            except Exception as e:
                console.print(f"[yellow]Failed to fill field {selector}: {e}[/yellow]")
        
        return fields_filled
    
    def upload_resume(self, page: Page, resume_path: str) -> bool:
        """
        Upload a resume to the ATS.
        
        Args:
            page: Playwright page
            resume_path: Path to the resume file
            
        Returns:
            True if successful, False otherwise
            
        Raises:
            NeedsHumanException: If upload fails
        """
        # Common resume upload selectors
        resume_selectors = [
            "input[type=file][name*=resume]",
            "input[type=file][name*=cv]",
            "input[type=file]:below(:text('Resume'))",
            "input[type=file]:below(:text('CV'))",
            "input[type=file]:below(:text('Upload'))",
            "input[type=file][accept*=pdf]",
            "input[type=file][accept*=docx]"
        ]
        
        return utils.smart_attach(page, resume_selectors, resume_path)
    
    def upload_cover_letter(self, page: Page, cover_letter_path: str) -> bool:
        """
        Upload a cover letter to the ATS.
        
        Args:
            page: Playwright page
            cover_letter_path: Path to the cover letter file
            
        Returns:
            True if successful, False otherwise
        """
        # Common cover letter upload selectors
        cover_letter_selectors = [
            "input[type=file][name*=cover]",
            "input[type=file][name*=letter]",
            "input[type=file]:below(:text('Cover'))",
            "input[type=file]:below(:text('Letter'))",
            "input[type=file]:nth-of-type(2)"
        ]
        
        try:
            return utils.smart_attach(page, cover_letter_selectors, cover_letter_path)
        except utils.NeedsHumanException:
            # Cover letter is often optional, so don't raise an exception
            console.print("[yellow]Could not upload cover letter automatically[/yellow]")
            return False
    
    def fill_personal_info(self, page: Page, profile: Dict) -> int:
        """
        Fill common personal information fields.
        
        Args:
            page: Playwright page
            profile: User profile dictionary
            
        Returns:
            Number of fields successfully filled
        """
        # Map profile fields to common form field selectors
        field_mappings = {
            "input[name*=firstName]": profile.get("name", "").split()[0] if profile.get("name") else "",
            "input[name*=lastName]": profile.get("name", "").split()[-1] if profile.get("name") else "",
            "input[name*=first_name]": profile.get("name", "").split()[0] if profile.get("name") else "",
            "input[name*=last_name]": profile.get("name", "").split()[-1] if profile.get("name") else "",
            "input[type=email]": profile.get("email", ""),
            "input[name*=email]": profile.get("email", ""),
            "input[name*=phone]": profile.get("phone", ""),
            "input[name*=telephone]": profile.get("phone", ""),
            "input[name*=location]": profile.get("location", ""),
            "input[name*=address]": profile.get("location", "")
        }
        
        return self.fill_form_fields(page, field_mappings)
    
    def click_next_button(self, page: Page) -> bool:
        """
        Click the next/continue button.
        
        Args:
            page: Playwright page
            
        Returns:
            True if button was clicked, False otherwise
        """
        next_button_selectors = [
            "button:has-text('Next')",
            "button:has-text('Continue')",
            "button:has-text('Save & Continue')",
            "input[type=submit][value*=Next]",
            "input[type=submit][value*=Continue]",
            "a:has-text('Next')",
            "a:has-text('Continue')"
        ]
        
        for selector in next_button_selectors:
            try:
                if page.is_visible(selector):
                    page.click(selector)
                    self.wait_for_navigation(page)
                    return True
            except Exception:
                continue
        
        return False
    
    def click_submit_button(self, page: Page) -> bool:
        """
        Click the final submit button.
        
        Args:
            page: Playwright page
            
        Returns:
            True if button was clicked, False otherwise
        """
        submit_button_selectors = [
            "button:has-text('Submit')",
            "button:has-text('Apply')",
            "button:has-text('Send Application')",
            "input[type=submit][value*=Submit]",
            "input[type=submit][value*=Apply]",
            "a:has-text('Submit')",
            "a:has-text('Apply')"
        ]
        
        for selector in submit_button_selectors:
            try:
                if page.is_visible(selector):
                    page.click(selector)
                    self.wait_for_navigation(page)
                    return True
            except Exception:
                continue
        
        return False
    
    def wait_for_human(self, page: Page, message: str) -> None:
        """
        Wait for human intervention.
        
        Args:
            page: Playwright page
            message: Message to display
        """
        console.print(f"[bold yellow]Human intervention required: {message}[/bold yellow]")
        console.print("[bold yellow]Press Enter when ready to continue...[/bold yellow]")
        input()
