import asyncio
import logging
import re
from pathlib import Path
from typing import Dict, Tuple, Any, List

from playwright.async_api import Page, FrameLocator, TimeoutError as PlaywrightTimeoutError
import docx # For reading cover letter text if needed

from .base_submitter import ApplicationSubmitter # Assuming base_submitter.py is in the same directory

logger = logging.getLogger(__name__)

class GreenhouseSubmitter(ApplicationSubmitter):
    """Submitter for Greenhouse job applications using Playwright."""

    async def _perform_submission(self, job_data: Dict[str, Any], resume_path: Path, cover_letter_path: Path) -> Tuple[str, str]:
        self.current_job_data = job_data
        job_title_for_prompts = job_data.get('title', 'this Greenhouse job')
        logger.info(f"Starting Greenhouse submission for '{job_data.get('title')}' at '{job_data.get('company')}'")

        headless_mode = self._ask_user_choice("Run browser in visible mode for Greenhouse submission?", ["Yes (Visible)", "No (Headless)"]) == 2
        await self._initialize_playwright(headless=headless_mode, browser_type=self.profile_data.get("browser", "chromium"))
        page = self.page
        if not page: return "Failed", "Playwright page not initialized."

        try:
            await page.goto(job_data['url'], wait_until="networkidle", timeout=60000)
            logger.info(f"Navigated to Greenhouse job URL: {job_data['url']}")
            await self._wait_for_natural_flow(1)
            await self._take_screenshot("greenhouse_job_page_start")
            if await self._check_for_captcha(): await self._handle_captcha_interaction()
            
            form_page: Page | FrameLocator = page 
            iframe_selector = "iframe#grnhse_iframe"
            iframe_element = await page.query_selector(iframe_selector)

            if iframe_element:
                logger.info("Greenhouse iframe detected. Switching context.")
                try:
                    await page.wait_for_function(f"() => document.querySelector('{iframe_selector}').contentDocument && document.querySelector('{iframe_selector}').contentDocument.readyState === 'complete'", timeout=15000)
                    form_page = page.frame_locator(iframe_selector)
                    logger.info("Switched to Greenhouse iframe context.")
                    await self._wait_for_natural_flow(1)
                    if await self._check_for_captcha(): await self._handle_captcha_interaction() 
                except PlaywrightTimeoutError:
                    logger.warning(f"Timeout waiting for Greenhouse iframe '{iframe_selector}' to load. Proceeding with main page.")
                except Exception as e_iframe:
                    logger.warning(f"Error accessing Greenhouse iframe '{iframe_selector}': {e_iframe}. Proceeding with main page.")
            else: 
                apply_button_selectors_main = [
                    "button:has-text('Apply Now')", "a:has-text('Apply Now')",
                    "button:has-text('Apply for this Job')", "a:has-text('Apply for this Job')",
                    "#apply_button" 
                ]
                clicked_main_apply = False
                for sel in apply_button_selectors_main:
                    main_page_button = await page.query_selector(sel)
                    if main_page_button and await main_page_button.is_visible():
                        if await self._click_button(page, [sel], "Greenhouse Apply Button (Main Page)", delay=2):
                            clicked_main_apply = True
                            await page.wait_for_load_state("networkidle", timeout=30000) 
                            if await self._check_for_captcha(): await self._handle_captcha_interaction()
                            iframe_element = await page.query_selector(iframe_selector) # Re-check for iframe
                            if iframe_element:
                                form_page = page.frame_locator(iframe_selector)
                                logger.info("Switched to Greenhouse iframe context after clicking apply on main page.")
                            break 
                if not clicked_main_apply:
                    logger.info("No immediate apply button on main page, or form is directly embedded.")
            
            await self._wait_for_natural_flow(2) 
            await self._take_screenshot("greenhouse_form_page")

            # --- Fill Greenhouse Form ---
            profile = self.profile_data
            
            res = await self._fill_text_field(form_page, ["#first_name", "input[name='job_application[first_name]']"], profile.get('first_name', profile.get('name', '').split(' ')[0]), "First Name", job_title_for_prompts, required=True)
            if res == "skip_job": return "Skipped (Missing Profile Data)", "Missing First Name"
            res = await self._fill_text_field(form_page, ["#last_name", "input[name='job_application[last_name]']"], profile.get('last_name', profile.get('name', '').split(' ')[-1] if ' ' in profile.get('name','') else ''), "Last Name", job_title_for_prompts, required=True)
            if res == "skip_job": return "Skipped (Missing Profile Data)", "Missing Last Name"
            res = await self._fill_text_field(form_page, ["#email", "input[name='job_application[email]']"], profile.get('email'), "Email", job_title_for_prompts, required=True)
            if res == "skip_job": return "Skipped (Missing Profile Data)", "Missing Email"
            res = await self._fill_text_field(form_page, ["#phone", "input[name='job_application[phone]']"], profile.get('phone'), "Phone", job_title_for_prompts, required=False)
            if res == "skip_job": return "Skipped (Missing Profile Data)", f"User skipped due to Phone for {job_title_for_prompts}"

            resume_input_selectors = ["input#resume", "input[name='job_application[resume]']", "input[type='file'][aria-describedby*='resume_filename']"]
            resume_button_selectors = ["button:has-text('Attach Resume')", "button:has-text('Upload Resume')", "button[data-qa='resume-upload']"]
            
            uploaded_resume = False
            if await self._click_button(form_page, resume_button_selectors, "Attach Resume Button", delay=0.5):
                await self._wait_for_natural_flow(0.5) 
            
            for sel in resume_input_selectors: 
                res_upload = await self._upload_file_to_input(form_page, sel, resume_path, "Resume", job_title_for_prompts, required=True)
                if res_upload == "skip_job": return "Skipped (Upload Failed)", f"Resume upload failed for {job_title_for_prompts}"
                if res_upload: 
                    uploaded_resume = True
                    break
            
            if not uploaded_resume:
                logger.warning("Failed to upload resume on Greenhouse.")
                if self._ask_user_choice("Failed to upload resume. Continue without it or skip job?", ["Continue without resume", "Skip job"]) == 2:
                    return "Skipped (Resume Upload Fail)", "User skipped due to resume upload failure."

            cl_input_selectors = ["input#cover_letter", "input[name='job_application[cover_letter]']", "input[type='file'][aria-describedby*='cover_letter_filename']"]
            cl_button_selectors = ["button:has-text('Attach Cover Letter')", "button:has-text('Upload Cover Letter')", "button[data-qa='cover-letter-upload']"]
            cl_textarea_selector = "textarea[name*='cover_letter']" # Greenhouse sometimes uses a textarea for CL
            cl_section_exists = any(await form_page.query_selector(s) for s in cl_input_selectors + cl_button_selectors + [cl_textarea_selector])

            if cl_section_exists and cover_letter_path.exists():
                logger.info("Attempting to handle cover letter on Greenhouse form.")
                if await self._click_button(form_page, cl_button_selectors, "Attach Cover Letter Button", delay=0.5):
                    await self._wait_for_natural_flow(0.5)
                
                uploaded_cl = False
                for sel in cl_input_selectors:
                    res_cl_upload = await self._upload_file_to_input(form_page, sel, cover_letter_path, "Cover Letter", job_title_for_prompts, required=False)
                    if res_cl_upload == "skip_job": return "Skipped (Upload Failed)", f"User skipped due to Cover Letter upload for {job_title_for_prompts}"
                    if res_cl_upload: uploaded_cl = True; break
                
                if not uploaded_cl: 
                    cl_text_content = ""
                    try:
                        doc = docx.Document(cover_letter_path)
                        cl_text_content = "\n".join([p.text for p in doc.paragraphs[:5]]) 
                    except Exception as e_read_cl:
                        logger.warning(f"Could not read cover letter content for pasting: {e_read_cl}")

                    if cl_text_content:
                        if await self._fill_text_field(form_page, [cl_textarea_selector], cl_text_content, "Cover Letter Text", job_title_for_prompts, required=False):
                            logger.info("Pasted cover letter content into textarea.")
                        else:
                            logger.warning("Could not automatically upload or paste cover letter on Greenhouse form.")
            elif cl_section_exists:
                 logger.info("Cover letter field exists but no cover letter file provided or found.")


            await self._fill_text_field(form_page, ["input[name*='linkedin_url']", "input[aria-label*='LinkedIn']"], profile.get('linkedin_url'), "LinkedIn URL", job_title_for_prompts, required=False)
            await self._fill_text_field(form_page, ["input[name*='github_url']", "input[name*='portfolio_url']", "input[aria-label*='Portfolio']", "input[aria-label*='Website']"], profile.get('portfolio_url', profile.get('github_url', profile.get('website_url'))), "Portfolio/Website URL", job_title_for_prompts, required=False)

            custom_question_sections = await form_page.query_selector_all("div.field, div.application-question")
            if custom_question_sections:
                print("\n[yellow]This Greenhouse application may have custom questions. Please review and answer them in the browser if automation doesn't cover them.[/yellow]")
                for i, q_element in enumerate(custom_question_sections):
                    label_el = await q_element.query_selector("label")
                    textarea_el = await q_element.query_selector("textarea")
                    if label_el and textarea_el and await textarea_el.is_visible():
                        q_text = (await label_el.inner_text()).strip().lower()
                        if "tell us more" in q_text or "anything else" in q_text:
                            if self._ask_user_choice(f"Found question: '{q_text}'. Add a generic response or fill manually?", ["Add generic", "Fill manually"]) == 1:
                                await self._fill_text_field(form_page, [f"#{await textarea_el.get_attribute('id')}"], "I am very interested in this opportunity and believe my skills align well.", "Generic Custom Question", job_title_for_prompts)
                
                if self._ask_user_choice("Do you want to pause and fill any remaining custom questions now?", ["Yes, I'll fill them", "No, try to submit as is"]) == 1:
                    input("Press Enter in the console after you have filled the custom questions...")
                    if await self._check_for_captcha(): await self._handle_captcha_interaction()
                else:
                    logger.info("User chose to proceed without explicitly filling custom questions.")


            submit_selectors = ["button[type='submit']", "button:has-text('Submit Application')", "#submit_app", "button[data-qa='submit-button']"]
            if self._ask_user_choice("Form filling attempt complete. Submit application to Greenhouse?", ["Yes", "No, review manually"]) == 1:
                if await self._click_button(form_page, submit_selectors, "Submit Application", delay=5): 
                    await page.wait_for_load_state("networkidle", timeout=45000)
                    await self._take_screenshot("greenhouse_after_submit")
                    page_content = (await page.content()).lower()
                    if "thank you for applying" in page_content or "application submitted" in page_content or "we received your application" in page_content or "application has been sent" in page_content:
                        return "Applied", "Application submitted successfully via Greenhouse automation."
                    else:
                        error_messages = await form_page.query_selector_all(".field-error-message, .error-message, [class*='error-message'], [class*='error_message']")
                        if error_messages:
                            errors = [await e.inner_text() for e in error_messages if await e.is_visible() and (await e.inner_text()).strip()]
                            if errors:
                                error_str = "; ".join(errors)
                                logger.warning(f"Greenhouse submission might have failed. Errors found: {error_str}")
                                return "Failed", f"Submission errors: {error_str[:150]}"
                        return "Applied (Confirmation Unclear)", "Submitted, but confirmation message not detected clearly."
                else:
                    return "Failed", "Could not click final Submit button on Greenhouse."
            else:
                return "Pending (User Review)", "User chose to review Greenhouse application manually."

        except Exception as e:
            if str(e).startswith("UserSkippedJob"): 
                 return "Skipped (User Decision)", str(e)
            logger.error(f"Error during Greenhouse submission: {e}", exc_info=True)
            await self._take_screenshot("greenhouse_submission_error")
            return "Failed", f"Greenhouse submission error: {str(e)[:100]}"

