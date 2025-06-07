"""
iCIMS ATS submitter implementation.
Handles job application automation for iCIMS-based job portals.
"""

from typing import Dict

from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError
from rich.console import Console

from .base_submitter import BaseSubmitter
import utils

console = Console()

class ICIMSSubmitter(BaseSubmitter):
    """
    Submitter for iCIMS ATS.
    Handles automation of job applications on iCIMS-based portals.
    """
    
    def submit(self, job: Dict, profile: Dict, resume_path: str, cover_letter_path: str) -> str:
        """
        Submit an application to iCIMS ATS.
        
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
            
            # Handle login/account creation if needed
            if self._check_for_login_requirement(page):
                console.print("[yellow]Login required[/yellow]")
                self.wait_for_human(page, "Please login or create account then press Enter")
                self.wait_for_navigation(page)
            
            # Click Apply button (iCIMS often has a prominent apply button)
            apply_button_selectors = [
                "button:has-text('Apply')",
                "a:has-text('Apply')",
                "button:has-text('Apply Now')",
                "a:has-text('Apply Now')",
                "input[type=button][value*='Apply']",
                "a.iCIMS_Button:has-text('Apply')"
            ]
            
            for selector in apply_button_selectors:
                if page.is_visible(selector):
                    console.print("[green]Clicking Apply button[/green]")
                    page.click(selector)
                    self.wait_for_navigation(page)
                    break
            
            # iCIMS often has a multi-step application process with tabs/sections
            # Check if we need to navigate to a specific section first
            self._navigate_to_upload_section(page)
            
            # Upload resume (iCIMS specific selectors)
            console.print("[green]Uploading resume[/green]")
            resume_selectors = [
                "input[type=file][name*='resume']",
                "input[type=file][id*='resume']",
                "input[type=file][name*='Resume']",
                "input[type=file][id*='Resume']",
                "input[type=file][name*='upload']",
                "input[type=file].iCIMS_FileUpload",
                "input[type=file]:below(label:has-text('Resume'))"
            ]
            
            try:
                utils.smart_attach(page, resume_selectors, resume_path)
            except utils.NeedsHumanException as e:
                console.print(f"[yellow]{str(e)}[/yellow]")
                self.wait_for_human(page, "Please upload resume manually then press Enter")
            
            # Upload cover letter if field exists
            console.print("[green]Checking for cover letter upload[/green]")
            cover_letter_selectors = [
                "input[type=file][name*='cover']",
                "input[type=file][id*='cover']",
                "input[type=file][name*='Cover']",
                "input[type=file][id*='Cover']",
                "input[type=file]:below(label:has-text('Cover'))",
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
            
            # Fill personal information (iCIMS often has specific field IDs)
            console.print("[green]Filling personal information[/green]")
            
            # iCIMS-specific field mappings
            icims_fields = {
                "input[name*='firstName']": profile.get("name", "").split()[0] if profile.get("name") else "",
                "input[name*='lastName']": profile.get("name", "").split()[-1] if profile.get("name") else "",
                "input[id*='firstName']": profile.get("name", "").split()[0] if profile.get("name") else "",
                "input[id*='lastName']": profile.get("name", "").split()[-1] if profile.get("name") else "",
                "input[name*='email']": profile.get("email", ""),
                "input[id*='email']": profile.get("email", ""),
                "input[name*='phoneNumber']": profile.get("phone", ""),
                "input[id*='phoneNumber']": profile.get("phone", ""),
                "input[name*='phone']": profile.get("phone", ""),
                "input[id*='phone']": profile.get("phone", ""),
                "input[name*='address']": profile.get("location", ""),
                "input[id*='address']": profile.get("location", ""),
                "input[name*='city']": profile.get("location", "").split(",")[0] if "," in profile.get("location", "") else "",
                "input[name*='linkedin']": profile.get("linkedin", ""),
                "input[id*='linkedin']": profile.get("linkedin", ""),
                "input[name*='github']": profile.get("github", ""),
                "input[id*='github']": profile.get("github", "")
            }
            
            fields_filled = self.fill_form_fields(page, icims_fields)
            console.print(f"[green]Filled {fields_filled} personal information fields[/green]")
            
            # Handle iCIMS-specific sections and navigation
            self._navigate_application_sections(page)
            
            # Handle screening questions (common in iCIMS)
            if self._handle_screening_questions(page):
                console.print("[green]Completed screening questions[/green]")
            
            # Check for review page and submit
            if page.is_visible("button:has-text('Submit')") or page.is_visible("input[type=submit]"):
                console.print("[green]Submitting application[/green]")
                
                # Try different submit button selectors
                submit_selectors = [
                    "button:has-text('Submit')",
                    "input[type=submit]",
                    "button.iCIMS_Button:has-text('Submit')",
                    "button:has-text('Submit Application')"
                ]
                
                for selector in submit_selectors:
                    if page.is_visible(selector):
                        page.click(selector)
                        self.wait_for_navigation(page, timeout=10000)
                        break
                
                # Wait a moment to ensure submission completes
                page.wait_for_timeout(2000)
                
                # Check for confirmation indicators
                confirmation_selectors = [
                    "text='Thank you for your application'",
                    "text='Application submitted'",
                    "text='successfully submitted'",
                    "text='has been received'",
                    "div.iCIMS_SuccessMessage"
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
            console.print(f"[bold red]iCIMS submission error: {e}[/bold red]")
            return "Failed"
        
        finally:
            # Close the page
            page.close()
    
    def _check_for_login_requirement(self, page: Page) -> bool:
        """
        Check if login is required before applying.
        
        Args:
            page: Playwright page
            
        Returns:
            True if login is required, False otherwise
        """
        login_indicators = [
            "text='Sign In'",
            "text='Log In'",
            "text='Create an Account'",
            "form[id*='login']",
            "input[name*='password']"
        ]
        
        for indicator in login_indicators:
            if page.is_visible(indicator):
                return True
        
        return False
    
    def _navigate_to_upload_section(self, page: Page) -> bool:
        """
        Navigate to the resume/document upload section if needed.
        
        Args:
            page: Playwright page
            
        Returns:
            True if navigation was successful, False otherwise
        """
        upload_section_selectors = [
            "a:has-text('Upload Resume')",
            "a:has-text('Documents')",
            "a:has-text('Resume')",
            "button:has-text('Upload')"
        ]
        
        for selector in upload_section_selectors:
            if page.is_visible(selector):
                try:
                    page.click(selector)
                    self.wait_for_navigation(page)
                    return True
                except Exception:
                    continue
        
        return False
    
    def _navigate_application_sections(self, page: Page) -> int:
        """
        Navigate through the various sections of an iCIMS application.
        
        Args:
            page: Playwright page
            
        Returns:
            Number of sections navigated
        """
        sections_navigated = 0
        
        # iCIMS often has Next/Continue buttons to navigate between sections
        next_button_selectors = [
            "button:has-text('Next')",
            "button:has-text('Continue')",
            "button:has-text('Save & Continue')",
            "input[type=button][value*='Next']",
            "input[type=button][value*='Continue']",
            "a.iCIMS_Button:has-text('Next')",
            "a.iCIMS_Button:has-text('Continue')"
        ]
        
        # Keep clicking Next buttons until we can't find any more
        max_sections = 10  # Limit to avoid infinite loops
        for _ in range(max_sections):
            next_clicked = False
            
            for selector in next_button_selectors:
                if page.is_visible(selector):
                    try:
                        page.click(selector)
                        self.wait_for_navigation(page)
                        sections_navigated += 1
                        next_clicked = True
                        break
                    except Exception:
                        continue
            
            # If we didn't click any Next button, we're done
            if not next_clicked:
                break
        
        return sections_navigated
    
    def _handle_screening_questions(self, page: Page) -> bool:
        """
        Handle common screening questions in iCIMS applications.
        
        Args:
            page: Playwright page
            
        Returns:
            True if questions were handled, False otherwise
        """
        # Check if we're on a page with screening questions
        question_indicators = [
            "div.iCIMS_Question",
            "div.screening-question",
            "label:has-text('Yes')",
            "label:has-text('No')"
        ]
        
        has_questions = False
        for indicator in question_indicators:
            if page.is_visible(indicator):
                has_questions = True
                break
        
        if not has_questions:
            return False
        
        # Try to answer common questions favorably
        try:
            # Handle Yes/No questions (usually best to answer "Yes" for positive questions)
            yes_options = page.query_selector_all("input[type=radio][value='Yes']")
            for option in yes_options:
                try:
                    option.check()
                except Exception:
                    pass
            
            # Handle work authorization questions
            auth_options = page.query_selector_all("input[type=radio][value*='authorized']")
            for option in auth_options:
                try:
                    option.check()
                except Exception:
                    pass
            
            # Handle dropdown selects (often experience level questions)
            # Try to select the first option for each dropdown
            dropdowns = page.query_selector_all("select")
            for dropdown in dropdowns:
                try:
                    # Get all options
                    options = dropdown.query_selector_all("option")
                    # Select the second option (first is often "Please select")
                    if len(options) > 1:
                        value = options[1].get_attribute("value")
                        if value:
                            dropdown.select_option(value=value)
                except Exception:
                    pass
            
            return True
            
        except Exception as e:
            console.print(f"[yellow]Error handling screening questions: {e}[/yellow]")
            return False
