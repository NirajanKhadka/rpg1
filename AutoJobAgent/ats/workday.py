"""
Workday ATS submitter implementation.
Handles job application automation for Workday-based job portals.
"""

from typing import Dict

from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError
from rich.console import Console

from .base_submitter import BaseSubmitter
import utils

console = Console()

class WorkdaySubmitter(BaseSubmitter):
    """
    Submitter for Workday ATS.
    Handles automation of job applications on Workday-based portals.
    """
    
    def submit(self, job: Dict, profile: Dict, resume_path: str, cover_letter_path: str) -> str:
        """
        Submit an application to Workday ATS.
        
        Args:
            job: Job dictionary with details
            profile: User profile dictionary
            resume_path: Path to the resume file
            cover_letter_path: Path to the cover letter file
            
        Returns:
            Status string (e.g., "Applied", "Failed", "Manual")
        """
        # Create a new page
        page = self.ctx.new_page()
        
        try:
            # Navigate to the job URL
            console.print(f"[green]Navigating to job URL: {job['url']}[/green]")
            page.goto(job["url"], timeout=30000)
            self.wait_for_navigation(page)
            
            # Check for CAPTCHA
            self.check_for_captcha(page)
            
            # Sign in if needed (using stored cookies)
            if page.is_visible("text='Sign In'"):
                console.print("[yellow]Sign-in required[/yellow]")
                page.click("text='Sign In'")
                self.wait_for_human(page, "Complete SSO login then press Enter")
                self.wait_for_navigation(page)
            
            # Check if we need to click Apply button
            apply_button_selectors = [
                "button:has-text('Apply')",
                "a:has-text('Apply')",
                "button:has-text('Apply Now')",
                "a:has-text('Apply Now')"
            ]
            
            for selector in apply_button_selectors:
                if page.is_visible(selector):
                    console.print("[green]Clicking Apply button[/green]")
                    page.click(selector)
                    self.wait_for_navigation(page)
                    break
            
            # Upload resume
            console.print("[green]Uploading resume[/green]")
            resume_selectors = [
                "input[type=file][name*=resume]",
                "input[type=file]:below(:text('Resume'))",
                "input[type=file][accept*='.pdf']",
                "input[type=file][accept*='.docx']"
            ]
            
            try:
                utils.smart_attach(page, resume_selectors, resume_path)
            except utils.NeedsHumanException as e:
                console.print(f"[yellow]{str(e)}[/yellow]")
                self.wait_for_human(page, "Please upload resume manually then press Enter")
            
            # Upload cover letter (if field exists)
            console.print("[green]Checking for cover letter upload[/green]")
            cover_letter_selectors = [
                "input[type=file][name*=cover]",
                "input[type=file]:below(:text('Cover'))",
                "input[type=file]:nth-of-type(2)"
            ]
            
            try:
                cl_element = None
                for selector in cover_letter_selectors:
                    cl_element = page.query_selector(selector)
                    if cl_element:
                        console.print("[green]Uploading cover letter[/green]")
                        cl_element.set_input_files(cover_letter_path)
                        break
            except Exception as e:
                console.print(f"[yellow]Cover letter upload failed: {e}[/yellow]")
                # Cover letter is often optional, so continue
            
            # Fill personal information
            console.print("[green]Filling personal information[/green]")
            fields_filled = self.fill_personal_info(page, profile)
            console.print(f"[green]Filled {fields_filled} personal information fields[/green]")
            
            # Handle additional common fields
            additional_fields = {
                "input[name*=linkedin]": profile.get("linkedin", ""),
                "input[name*=website]": profile.get("github", ""),
                "input[name*=github]": profile.get("github", ""),
                "input[name*=portfolio]": profile.get("github", "")
            }
            
            self.fill_form_fields(page, additional_fields)
            
            # Click through Next buttons until we reach the end or run out of Next buttons
            console.print("[green]Navigating through application steps[/green]")
            next_clicks = utils.loop_click(page, "button:has-text('Next')", max_clicks=10)
            console.print(f"[green]Clicked Next {next_clicks} times[/green]")
            
            # Check for review page and submit
            if page.is_visible("button:has-text('Submit')"):
                console.print("[green]Submitting application[/green]")
                page.click("button:has-text('Submit')")
                self.wait_for_navigation(page, timeout=10000)
                
                # Wait a moment to ensure submission completes
                page.wait_for_timeout(2000)
                
                # Check for confirmation indicators
                confirmation_selectors = [
                    "text='Thank you for your application'",
                    "text='Application submitted'",
                    "text='successfully submitted'",
                    "text='has been received'"
                ]
                
                for selector in confirmation_selectors:
                    if page.is_visible(selector):
                        console.print("[bold green]Application successfully submitted![/bold green]")
                        return "Applied"
                
                # If no confirmation found but no errors either, assume success
                return "Applied"
            else:
                # If we can't find the Submit button, human intervention needed
                console.print("[yellow]Could not find Submit button[/yellow]")
                self.wait_for_human(page, "Please complete and submit the application manually")
                return "Manual"
            
        except utils.NeedsHumanException:
            console.print("[yellow]Human intervention required[/yellow]")
            self.wait_for_human(page, "Please complete the application manually")
            return "Manual"
        
        except Exception as e:
            console.print(f"[bold red]Workday submission error: {e}[/bold red]")
            return "Failed"
        
        finally:
            # Close the page
            page.close()
