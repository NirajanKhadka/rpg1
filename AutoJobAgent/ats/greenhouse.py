"""
Greenhouse ATS submitter implementation.
Handles job application automation for Greenhouse-based job portals.
"""

from typing import Dict, List, Optional

from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError
from rich.console import Console

from .base_submitter import BaseSubmitter
import utils

console = Console()

class GreenhouseSubmitter(BaseSubmitter):
    """
    Submitter for Greenhouse ATS.
    Handles automation of job applications on Greenhouse-based portals.
    """
    
    def submit(self, job: Dict, profile: Dict, resume_path: str, cover_letter_path: str) -> str:
        """
        Submit an application to Greenhouse ATS.
        
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
            
            # Greenhouse often has an "Apply for this job" button
            apply_button_selectors = [
                "a:has-text('Apply for this job')",
                "a:has-text('Apply Now')",
                "button:has-text('Apply')",
                "button:has-text('Apply for this job')",
                "button:has-text('Apply Now')"
            ]
            
            for selector in apply_button_selectors:
                if page.is_visible(selector):
                    console.print("[green]Clicking Apply button[/green]")
                    page.click(selector)
                    self.wait_for_navigation(page)
                    break
            
            # Greenhouse application form usually has a specific structure
            # First, check if we need to fill out personal information
            console.print("[green]Filling personal information[/green]")
            
            # Fill name fields
            self._fill_name_fields(page, profile)
            
            # Fill email
            email_selectors = [
                "input#email",
                "input[name='email']",
                "input[name='job_application[email]']"
            ]
            for selector in email_selectors:
                if utils.fill_if_empty(page, selector, profile.get("email", "")):
                    break
            
            # Fill phone
            phone_selectors = [
                "input#phone",
                "input[name='phone']",
                "input[name='job_application[phone]']"
            ]
            for selector in phone_selectors:
                if utils.fill_if_empty(page, selector, profile.get("phone", "")):
                    break
            
            # Upload resume
            console.print("[green]Uploading resume[/green]")
            resume_selectors = [
                "input#resume",
                "input[name='resume']",
                "input[name='job_application[resume]']",
                "input[type=file]:below(label:has-text('Resume'))",
                "input[type=file][accept='.pdf,.doc,.docx']"
            ]
            
            try:
                utils.smart_attach(page, resume_selectors, resume_path)
            except utils.NeedsHumanException as e:
                console.print(f"[yellow]{str(e)}[/yellow]")
                self.wait_for_human(page, "Please upload resume manually then press Enter")
            
            # Upload cover letter if field exists
            console.print("[green]Checking for cover letter upload[/green]")
            cover_letter_selectors = [
                "input#cover_letter",
                "input[name='cover_letter']",
                "input[name='job_application[cover_letter]']",
                "input[type=file]:below(label:has-text('Cover Letter'))"
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
            
            # Fill additional fields
            self._fill_additional_fields(page, profile)
            
            # Handle custom questions
            if self._handle_custom_questions(page):
                console.print("[green]Answered custom questions[/green]")
            
            # Handle EEO questions (Equal Employment Opportunity)
            if self._handle_eeo_questions(page):
                console.print("[green]Completed EEO questions[/green]")
            
            # Check for terms agreement checkbox
            self._check_terms_agreement(page)
            
            # Submit the application
            console.print("[green]Submitting application[/green]")
            submit_button_selectors = [
                "input[type=submit][value='Submit Application']",
                "button:has-text('Submit Application')",
                "button:has-text('Submit')",
                "input[type=submit]"
            ]
            
            for selector in submit_button_selectors:
                if page.is_visible(selector):
                    page.click(selector)
                    self.wait_for_navigation(page, timeout=10000)
                    break
            
            # Wait a moment to ensure submission completes
            page.wait_for_timeout(2000)
            
            # Check for confirmation indicators
            confirmation_selectors = [
                "text='Thank you for applying'",
                "text='Your application has been submitted'",
                "text='Thank you for your interest'",
                "div.application-confirmation"
            ]
            
            for selector in confirmation_selectors:
                if page.is_visible(selector):
                    console.print("[bold green]Application successfully submitted![/bold green]")
                    return "Applied"
            
            # If no confirmation found but no errors either, assume success
            return "Applied"
            
        except utils.NeedsHumanException:
            console.print("[yellow]Human intervention required[/yellow]")
            self.wait_for_human(page, "Please complete the application manually")
            return "Manual"
        
        except Exception as e:
            console.print(f"[bold red]Greenhouse submission error: {e}[/bold red]")
            return "Failed"
        
        finally:
            # Close the page
            page.close()
    
    def _fill_name_fields(self, page: Page, profile: Dict) -> bool:
        """
        Fill name fields in Greenhouse application form.
        
        Args:
            page: Playwright page
            profile: User profile dictionary
            
        Returns:
            True if fields were filled, False otherwise
        """
        # Get first and last name
        full_name = profile.get("name", "")
        if not full_name:
            return False
        
        name_parts = full_name.split()
        first_name = name_parts[0] if name_parts else ""
        last_name = name_parts[-1] if len(name_parts) > 1 else ""
        
        # Try different selectors for first name
        first_name_selectors = [
            "input#first_name",
            "input[name='first_name']",
            "input[name='job_application[first_name]']"
        ]
        
        first_filled = False
        for selector in first_name_selectors:
            if utils.fill_if_empty(page, selector, first_name):
                first_filled = True
                break
        
        # Try different selectors for last name
        last_name_selectors = [
            "input#last_name",
            "input[name='last_name']",
            "input[name='job_application[last_name]']"
        ]
        
        last_filled = False
        for selector in last_name_selectors:
            if utils.fill_if_empty(page, selector, last_name):
                last_filled = True
                break
        
        # If separate fields didn't work, try full name field
        if not (first_filled and last_filled):
            full_name_selectors = [
                "input#name",
                "input[name='name']",
                "input[name='job_application[name]']"
            ]
            
            for selector in full_name_selectors:
                if utils.fill_if_empty(page, selector, full_name):
                    return True
        
        return first_filled or last_filled
    
    def _fill_additional_fields(self, page: Page, profile: Dict) -> int:
        """
        Fill additional fields in Greenhouse application form.
        
        Args:
            page: Playwright page
            profile: User profile dictionary
            
        Returns:
            Number of fields filled
        """
        fields_filled = 0
        
        # Fill location/address
        location_selectors = [
            "input#location",
            "input[name='location']",
            "input[name='job_application[location]']",
            "input#address",
            "input[name='address']"
        ]
        
        for selector in location_selectors:
            if utils.fill_if_empty(page, selector, profile.get("location", "")):
                fields_filled += 1
                break
        
        # Fill LinkedIn
        linkedin_selectors = [
            "input#linkedin",
            "input[name='linkedin']",
            "input[name='job_application[linkedin_url]']",
            "input[placeholder*='LinkedIn']"
        ]
        
        for selector in linkedin_selectors:
            if utils.fill_if_empty(page, selector, profile.get("linkedin", "")):
                fields_filled += 1
                break
        
        # Fill GitHub
        github_selectors = [
            "input#github",
            "input[name='github']",
            "input[name='job_application[website]']",
            "input[placeholder*='GitHub']"
        ]
        
        for selector in github_selectors:
            if utils.fill_if_empty(page, selector, profile.get("github", "")):
                fields_filled += 1
                break
        
        # Fill website (use GitHub if no website specified)
        website_selectors = [
            "input#website",
            "input[name='website']",
            "input[name='job_application[website]']"
        ]
        
        for selector in website_selectors:
            if utils.fill_if_empty(page, selector, profile.get("github", "")):
                fields_filled += 1
                break
        
        return fields_filled
    
    def _handle_custom_questions(self, page: Page) -> bool:
        """
        Handle custom questions in Greenhouse application form.
        
        Args:
            page: Playwright page
            
        Returns:
            True if questions were handled, False otherwise
        """
        # Check if there are custom questions
        question_containers = page.query_selector_all("div.custom-question")
        if not question_containers:
            return False
        
        questions_handled = 0
        
        for container in question_containers:
            try:
                # Try to identify the question type
                
                # Check for text inputs
                text_input = container.query_selector("input[type=text]")
                if text_input:
                    text_input.fill("Yes")
                    questions_handled += 1
                    continue
                
                # Check for textareas
                textarea = container.query_selector("textarea")
                if textarea:
                    # For textareas, provide a more detailed response
                    textarea.fill("Yes, I am eligible and available to start immediately.")
                    questions_handled += 1
                    continue
                
                # Check for radio buttons (usually yes/no questions)
                radio_yes = container.query_selector("input[type=radio][value='Yes']")
                if radio_yes:
                    radio_yes.check()
                    questions_handled += 1
                    continue
                
                # Check for dropdowns
                select = container.query_selector("select")
                if select:
                    # Get all options
                    options = select.query_selector_all("option")
                    # Select the second option (first is often "Please select")
                    if len(options) > 1:
                        value = options[1].get_attribute("value")
                        if value:
                            select.select_option(value=value)
                            questions_handled += 1
                            continue
            
            except Exception as e:
                console.print(f"[yellow]Error handling custom question: {e}[/yellow]")
        
        return questions_handled > 0
    
    def _handle_eeo_questions(self, page: Page) -> bool:
        """
        Handle EEO (Equal Employment Opportunity) questions in Greenhouse.
        
        Args:
            page: Playwright page
            
        Returns:
            True if questions were handled, False otherwise
        """
        # Check if there's an EEO section
        eeo_section = page.query_selector("div#equal_opportunity_employer")
        if not eeo_section:
            return False
        
        try:
            # Look for "decline to self identify" options
            decline_options = page.query_selector_all("input[type=radio][value*='decline']")
            for option in decline_options:
                option.check()
            
            # If no decline options found, try to find "I don't wish to answer" options
            if not decline_options:
                dont_answer_options = page.query_selector_all("input[type=radio][value*='don']")
                for option in dont_answer_options:
                    option.check()
            
            return True
            
        except Exception as e:
            console.print(f"[yellow]Error handling EEO questions: {e}[/yellow]")
            return False
    
    def _check_terms_agreement(self, page: Page) -> bool:
        """
        Check terms agreement checkbox if present.
        
        Args:
            page: Playwright page
            
        Returns:
            True if checkbox was checked, False otherwise
        """
        try:
            # Look for terms agreement checkboxes
            terms_selectors = [
                "input[type=checkbox][id*='term']",
                "input[type=checkbox][name*='term']",
                "input[type=checkbox][id*='agree']",
                "input[type=checkbox][name*='agree']",
                "input[type=checkbox]:below(label:has-text('agree'))"
            ]
            
            for selector in terms_selectors:
                checkbox = page.query_selector(selector)
                if checkbox:
                    checkbox.check()
                    return True
            
            return False
            
        except Exception as e:
            console.print(f"[yellow]Error checking terms agreement: {e}[/yellow]")
            return False
