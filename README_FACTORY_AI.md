"""
Workday ATS Submitter Module
----------------------------
Handles automated job application submissions to Workday sites using Playwright.
"""
import logging
import re
from pathlib import Path
from typing import Dict, Tuple, Any, List, Optional

from playwright.async_api import Page, FrameLocator, TimeoutError as PlaywrightTimeoutError

from .base_submitter import ApplicationSubmitter

logger = logging.getLogger(__name__)

class WorkdaySubmitter(ApplicationSubmitter):
    """Submitter for Workday job applications using Playwright."""

    async def _perform_submission(self, job_data: Dict[str, Any], resume_path: Path, cover_letter_path: Path) -> Tuple[str, str]:
        job_title_for_prompts = job_data.get('title', 'this Workday job')
        logger.info(f"Starting Workday submission for '{job_data.get('title')}' at '{job_data.get('company')}'")
        
        headless_mode = self._ask_user_choice("Run browser in visible mode for Workday submission?", ["Yes (Visible)", "No (Headless)"]) == 2
        await self._initialize_playwright(headless=headless_mode) # Browser type from profile or default
        page = self.page
        if not page: return "Failed", "Playwright page not initialized."

        try:
            await page.goto(job_data['url'], wait_until="networkidle", timeout=60000)
            logger.info(f"Successfully navigated to Workday job page: {job_data['url']}")
            await self._wait_for_natural_flow(2)
            if await self._check_for_captcha(page): await self._handle_captcha_interaction(page)

            # Common Workday "Apply" button patterns
            apply_button_selectors = [
                "button[data-automation-id='adventureButton']", # Often the main one
                "button[data-automation-id='applyButton']",
                "a[data-automation-id='adventureButtonLink']",
                "button:has-text('Apply Now')", "button:has-text('Apply')",
                "a:has-text('Apply Now')", "a:has-text('Apply')",
            ]
            apply_clicked = await self._click_button_with_retry(page, apply_button_selectors, "Workday Apply Button", delay_after_click=3)

            if not apply_clicked:
                logger.warning("Could not find or click the main Apply button on the Workday job page.")
                await self._take_screenshot("workday_apply_button_not_found")
                user_choice = self._ask_user_choice("Could not find Apply button. What to do?", ["Try to continue manually (if you see the form)", "Skip this job"])
                if user_choice == 2: return "Failed", "Could not find Apply button."
                # If user wants to continue manually, they might be on the form already or will click it.
                logger.info("User will attempt to proceed manually from job description page.")
                # Fall through to form filling, assuming user gets to the form.

            await page.wait_for_load_state("networkidle", timeout=45000) # Wait for potential page transition
            await self._wait_for_natural_flow(2)
            logger.info("On first page after clicking Apply (or attempting to).")
            await self._take_screenshot("workday_after_apply_click")
            if await self._check_for_captcha(page): await self._handle_captcha_interaction(page)

            # Workday often has options like "Autofill with Resume", "Apply Manually", or directly goes to a form.
            autofill_selectors = [
                "button[data-automation-id='file-upload-prompt-upload-button']", # Common for "Autofill with resume"
                "button:has-text('Autofill with Resume')", "button:has-text('Upload Resume')",
            ]
            manual_or_next_selectors = [
                "button[data-automation-id='manuallyEnterDataButton']", # "Apply Manually"
                "button:has-text('Apply Manually')", "button:has-text('Fill Out Manually')",
                "button:textmatches('Next', 'i')", "button:textmatches('Continue', 'i')",
                "button[data-automation-id='nextButton']", "button[data-automation-id='saveAndContinueButton']"
            ]
            
            autofill_button_found = any(await page.query_selector(s) for s in autofill_selectors)
            
            if autofill_button_found and self._ask_user_choice("Option to 'Autofill with Resume' found. Attempt this?", ["Yes", "No, fill manually/next"]) == 1:
                if await self._click_button_with_retry(page, autofill_selectors, "Autofill with Resume Button", delay_after_click=1):
                    # Workday's resume upload for autofill is often a direct input[type=file] triggered by the button.
                    # The actual input might be hidden or revealed.
                    upload_result = await self._upload_file_with_retry(page, 
                                                                    ["input[type='file'][data-automation-id='file-upload-input-ref']", "input[type='file']"], # Common Workday file input
                                                                    [], # No separate button click needed after autofill button
                                                                    resume_path, "Resume for Autofill", job_title_for_prompts, required=True)
                    if upload_result == "skip_job": return "Skipped (Autofill Upload Failed)", "Failed to upload resume for autofill."
                    await page.wait_for_load_state("networkidle", timeout=60000) # Give ample time for autofill
                    logger.info("Attempted autofill with resume. Subsequent fields may be pre-filled.")
                    if await self._check_for_captcha(page): await self._handle_captcha_interaction(page)
            elif not await self._click_button_with_retry(page, manual_or_next_selectors, "Apply Manually / Next Button", delay_after_click=2):
                 logger.warning("Could not find 'Autofill with Resume' or 'Apply Manually/Next' button after initial apply click.")
                 await self._take_screenshot("workday_autofill_or_manual_not_found")
                 stuck_action = await self._handle_stuck_automation("initial form navigation (Autofill/Manual/Next)")
                 if stuck_action == "skip_job": return "Failed", "Could not find autofill/manual apply option."
                 elif stuck_action == "manual_takeover":
                     logger.info("User taking over for initial form navigation.")
                     # After manual interaction, the script will proceed to the form filling loop.
                     if await self._check_for_captcha(page): await self._handle_captcha_interaction(page)


            # --- Multi-Step Form Filling Loop ---
            max_steps = 15 # Workday forms can be long
            for step in range(max_steps):
                await page.wait_for_load_state("networkidle", timeout=45000)
                await self._wait_for_natural_flow(1.5) # Slightly longer for Workday pages
                current_url = page.url
                logger.info(f"On Workday form page (step {step + 1}): {current_url}")
                await self._take_screenshot(f"workday_step_{step+1}")
                if await self._check_for_captcha(page): await self._handle_captcha_interaction(page)

                profile = self.profile_data
                # Personal Information (common Workday data-automation-ids)
                fields_to_fill = [
                    ("First Name", ["input[data-automation-id='legalNameSection_firstName']", "input[data-automation-id='firstName']"], profile.get('first_name', profile.get('name', '').split(' ')[0]), True),
                    ("Last Name", ["input[data-automation-id='legalNameSection_lastName']", "input[data-automation-id='lastName']"], profile.get('last_name', profile.get('name', '').split(' ')[-1] if ' ' in profile.get('name','') else ''), True),
                    ("Email Address", ["input[data-automation-id='email']"], profile.get('email'), True),
                    ("Phone Number", ["input[data-automation-id='phone-number']", "input[type='tel']"], profile.get('phone'), False), # Often optional
                    ("Address Line 1", ["input[data-automation-id='addressSection_addressLine1']"], profile.get('address', {}).get('street'), False),
                    ("City", ["input[data-automation-id='addressSection_city']"], profile.get('address', {}).get('city'), False),
                    # State/Province might be a dropdown or text input
                    ("State/Province", ["input[data-automation-id='addressSection_region']", "input[data-automation-id='province']"], profile.get('address', {}).get('state'), False),
                    ("Postal Code", ["input[data-automation-id='addressSection_postalCode']"], profile.get('address', {}).get('zip'), False),
                    # Country is often a dropdown
                    ("Country", ["select[data-automation-id='country']"], profile.get('address', {}).get('country'), False), # Assuming select by label/partial label match
                    ("LinkedIn Profile URL", ["input[data-automation-id='linkedinQuestion']"], profile.get('linkedin_url'), False),
                    ("Website/Portfolio", ["input[data-automation-id='website']", "input[aria-labelledby*='website']"], profile.get('portfolio_url', profile.get('website_url')), False),
                ]

                for field_name, selectors, value, required in fields_to_fill:
                    if "Country" in field_name and value: # Special handling for country dropdown
                        res = await self._select_dropdown_option(page, selectors, value, field_name, job_title_for_prompts, by_value=False, required=required) # Try by label first
                        if res != "success" and res != "skip_job": # Try by value if label fails (e.g. "US" for "United States")
                             res = await self._select_dropdown_option(page, selectors, profile.get('address', {}).get('country_code', value), field_name, job_title_for_prompts, by_value=True, required=required)
                    else:
                        res = await self._fill_field_with_retry(page, selectors, value, field_name, required=required)
                    if res == "skip_job": return "Skipped (Missing Profile Data)", f"User skipped due to missing '{field_name}' for {job_title_for_prompts}"
                    if res == "failed_fill" and required:
                        logger.error(f"Failed to fill required Workday field: {field_name}")
                        # Potentially ask user to fill manually or skip
                        stuck_action = await self._handle_stuck_automation(f"filling required field '{field_name}' on Workday")
                        if stuck_action == "skip_job": return "Skipped (Field Fill Fail)", f"Skipped due to failure filling '{field_name}'."
                        # If manual_takeover, we assume user fixed it and loop continues.

                # Resume Upload (if not autofilled or if fields reappear)
                # Workday often has separate sections for "My Experience" where resume is primary
                # and "Application Questions" or "Additional Documents" for cover letter
                resume_upload_selectors = ["input[type='file'][data-automation-id='file-upload-input-ref']", "button[data-automation-id='resumeUpload']"]
                resume_section_header = await page.query_selector("h3:has-text('Resume'), h2:has-text('Resume')") # Check if resume section is visible
                if resume_section_header and await resume_section_header.is_visible():
                    if not await page.query_selector("div[data-automation-id*='resume'] span[title*='.pdf'], div[data-automation-id*='resume'] span[title*='.docx']"): # If no resume seems uploaded
                        logger.info("Attempting to upload resume on Workday form.")
                        res_upload = await self._upload_file_with_retry(page, resume_upload_selectors, ["button[data-automation-id='resumeSection_uploadResume']"], resume_path, "Resume", job_title_for_prompts, required=True)
                        if res_upload == "skip_job": return "Skipped (Resume Upload Failed)", f"Resume upload failed for {job_title_for_prompts}"
