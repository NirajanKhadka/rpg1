#!/usr/bin/env python3
"""
Document Generator Module
------------------------
Provides functionality to generate customized .docx resumes and cover letters
for job applications in .docx format. Integrates with Ollama for intelligent
content customization and converts DOCX to PDF on Windows.
"""

import os
import re
import logging
import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple

try:
    import docx
    from docx import Document
except ImportError:
    print("python-docx library not found. Please install it with 'pip install python-docx'")
    raise ImportError("python-docx library not found. Please install it with 'pip install python-docx'")

try:
    import ollama
except ImportError:
    print("ollama library not found. Please install it with 'pip install ollama'")
    ollama = None

# Assuming utils.py is in the same directory or accessible via PYTHONPATH
try:
    from utils import convert_doc_to_pdf_windows, sanitize_filename, ask_user_choice, ensure_file_exists
except ImportError:
    logging.error("Failed to import utility functions. Ensure utils.py is accessible.")
    # Define fallbacks if utils is not available for some reason during generation, though it should be.
    def convert_doc_to_pdf_windows(doc_path: str, pdf_path: str) -> bool: return False
    def sanitize_filename(name: str, max_length: int = 200) -> str: return re.sub(r'[^\w\s-]', '', name).strip()[:max_length]
    def ask_user_choice(question: str, choices: List[str]) -> Optional[int]:
        print(f"\n{question}")
        for i, choice in enumerate(choices): print(f"  {i+1}) {choice}")
        try: return int(input(f"Enter your choice (1-{len(choices)}): "))
        except: return None
    def ensure_file_exists(file_path_str: str, purpose: str = "file") -> bool: return Path(file_path_str).is_file()


logger = logging.getLogger(__name__)

class DocumentGenerator:
    """Generates customized .docx resumes and cover letters using Ollama and converts to PDF."""

    def __init__(self, profile_data: Dict[str, Any], profile_folder_path: Path, output_dir_for_profile_docs: Path, ollama_model: str = "mistral"):
        self.profile_data = profile_data
        self.profile_folder_path = profile_folder_path # e.g., profiles/Nirajan/
        self.output_dir = output_dir_for_profile_docs # e.g., output/Nirajan/documents/
        self.ollama_model = profile_data.get("ollama_model", ollama_model) # Allow profile override
        self.ollama_available = self._check_ollama_availability()

        self.base_resume_docx_path = self.profile_folder_path / self.profile_data["resume_docx"]
        self.base_cover_letter_docx_path = self.profile_folder_path / self.profile_data["cover_letter_docx"]

        self._verify_templates()
        self.output_dir.mkdir(parents=True, exist_ok=True)
        logger.info(f"DocumentGenerator initialized. Output: {self.output_dir}. Ollama available: {self.ollama_available} with model '{self.ollama_model}'.")

    def _check_ollama_availability(self) -> bool:
        if ollama is None:
            logger.warning("Ollama library is not installed. LLM features will be disabled.")
            print("Warning: Ollama library not installed. Document customization will be basic.")
            return False
        try:
            ollama.list()
            # Check if the specific model is available
            models = ollama.list().get('models', [])
            if not any(m['name'].startswith(self.ollama_model) for m in models):
                logger.warning(f"Ollama model '{self.ollama_model}' not found locally. Please run `ollama pull {self.ollama_model}`.")
                print(f"Warning: Ollama model '{self.ollama_model}' not found. LLM features might be limited or fail.")
                # Depending on strictness, could return False here. For now, assume user might fix it.
            logger.info(f"Successfully connected to Ollama. Configured model: {self.ollama_model}")
            return True
        except Exception as e:
            logger.warning(f"Ollama is not available or error during check: {e}. LLM features will be significantly limited or disabled.")
            print(f"Warning: Ollama not responding. Ensure Ollama is running (e.g., `ollama serve`) and the model '{self.ollama_model}' is pulled. Document customization will be basic.")
            return False

    def _verify_templates(self) -> None:
        if not utils.ensure_file_exists(str(self.base_resume_docx_path), "Base resume template"):
            raise FileNotFoundError(f"Base resume template not found: {self.base_resume_docx_path}")
        if not utils.ensure_file_exists(str(self.base_cover_letter_docx_path), "Base cover letter template"):
            raise FileNotFoundError(f"Base cover letter template not found: {self.base_cover_letter_docx_path}")
        logger.info("Base resume and cover letter templates verified.")

    def _load_template(self, template_path: Path) -> Document:
        try:
            logger.debug(f"Loading template from: {template_path}")
            return Document(template_path)
        except Exception as e:
            logger.error(f"Error loading .docx template {template_path}: {e}", exc_info=True)
            raise

    def _call_ollama(self, prompt: str, purpose: str) -> Optional[str]:
        if not self.ollama_available:
            logger.warning(f"Ollama not available, skipping LLM call for {purpose}.")
            return None # Fallback handled by caller
        try:
            logger.info(f"Sending prompt to Ollama ({self.ollama_model}) for {purpose}...")
            # logger.debug(f"Ollama prompt: {prompt}") # Can be very verbose
            
            response = ollama.generate(model=self.ollama_model, prompt=prompt, stream=False)
            
            generated_text = response.get('response', '').strip()
            if not generated_text:
                logger.warning(f"Ollama returned empty response for {purpose}.")
                return None
            logger.info(f"Ollama generated text for {purpose} (snippet): {generated_text[:150]}...")
            return generated_text
        except Exception as e:
            logger.error(f"Error calling Ollama for {purpose}: {e}", exc_info=True)
            print(f"Error communicating with Ollama for {purpose}: {e}. Using fallback content.")
            return None

    def _extract_job_keywords_with_ollama(self, job_description: str, job_title: str) -> List[str]:
        if not job_description: return []
        prompt = (
            f"Analyze the following job description for the role of '{job_title}'. "
            f"Extract the top 10-15 most important technical skills, tools, software, and key responsibilities mentioned. "
            f"Focus on concrete nouns and noun phrases that represent qualifications or duties. "
            f"Return them as a comma-separated list. Job Description:\n\n{job_description[:3000]}" # Limit context
        )
        response_text = self._call_ollama(prompt, f"keyword extraction for {job_title}")
        if response_text:
            keywords = [kw.strip().lower() for kw in response_text.split(',') if kw.strip() and len(kw) > 2]
            logger.info(f"Keywords extracted by Ollama for '{job_title}': {keywords}")
            return list(set(keywords)) # Unique keywords
        return []

    def _generate_custom_section_with_ollama(self, section_type: str, base_text: str, job_description: str, job_title: str, company_name: str, profile_skills: List[str], profile_keywords: List[str]) -> str:
        if not self.ollama_available:
            logger.warning(f"Ollama not available. Returning base text for {section_type}.")
            return base_text # Fallback to base text

        job_extracted_keywords = self._extract_job_keywords_with_ollama(job_description, job_title)
        
        # Combine profile skills and keywords for context
        user_context_skills = list(set(profile_skills + profile_keywords))

        # Identify skills from job_extracted_keywords that are NOT in user_context_skills
        new_skills_to_consider = [kw for kw in job_extracted_keywords if kw.lower() not in (skill.lower() for skill in user_context_skills)]

        prompt = ""
        if section_type == "resume_summary":
            prompt = (
                f"You are an expert resume writer. Rewrite the following resume summary to be more impactful and tailored for the job of '{job_title}' at '{company_name}'.\n"
                f"Original Summary: \"{base_text}\"\n"
                f"User's Skills/Keywords: {', '.join(user_context_skills)}\n"
                f"Key aspects from the Job Description: {', '.join(job_extracted_keywords)}\n"
                f"Instructions: Create a concise (2-3 sentences) professional summary. Highlight how the user's skills and experience align with the key aspects of the job. "
                f"Naturally weave in relevant keywords from the job description. Do not just list keywords. "
                f"If the job requires skills like '{', '.join(new_skills_to_consider[:3])}' which are not explicitly in the user's list, try to phrase the summary to show adaptability or related experience if possible, but prioritize existing skills. "
                f"Output only the revised summary text."
            )
        elif section_type == "cover_letter_body":
            prompt = (
                f"You are an expert cover letter writer. Craft 1-2 compelling body paragraphs for a cover letter for the role of '{job_title}' at '{company_name}'.\n"
                f"The applicant's key skills and keywords are: {', '.join(user_context_skills)}\n"
                f"The job description emphasizes these aspects: {', '.join(job_extracted_keywords)}\n"
                f"Instructions: Write 1-2 paragraphs (approx 3-5 sentences total). Focus on demonstrating how the applicant's background and skills directly address the needs of the role as suggested by the job's key aspects. "
                f"Express enthusiasm for the company and specific role. "
                f"Naturally integrate relevant keywords from the job description. Do not just list keywords. "
                f"If the job requires skills like '{', '.join(new_potential_skills[:3])}' which are not explicitly in the user's list, you can suggest how existing skills are transferable or show eagerness to learn, but focus on current strengths. "
                f"Output only the crafted cover letter body paragraphs."
            )
        
        if not prompt:
            return base_text

        customized_text = self._call_ollama(prompt, f"{section_type} generation")

        if customized_text:
            if new_potential_skills:
                mentioned_new_skills = [skill for skill in new_potential_skills if skill.lower() in customized_text.lower()]
                if mentioned_new_skills:
                    question = (f"Ollama's suggestion for '{section_type}' includes concepts related to '{', '.join(mentioned_new_skills)}' which might be new. "
                                f"Do you want to use this AI-generated text? (Review carefully):\n---\n{customized_text}\n---\nUse this text?")
                    if not self._ask_user_confirmation(question, default_yes=True):
                        logger.info(f"User rejected Ollama's suggestion for {section_type} due to new skills. Using base text.")
                        return base_text
            return customized_text
        
        logger.warning(f"Ollama failed to generate text for {section_type}. Using base text.")
        return base_text


    def _get_replacements(self, job_data: Dict[str, Any]) -> Dict[str, str]:
        replacements = {}
        profile = self.profile_data
        job_title = job_data.get("title", "The Position")
        company_name = job_data.get("company", "The Company")
        job_description = job_data.get("description", job_data.get("summary", ""))

        # Basic Profile Info
        replacements["{{FULL_NAME}}"] = profile.get("name", "Your Name")
        replacements["{{FIRST_NAME}}"] = profile.get("first_name", profile.get("name", "Your Name").split(" ")[0])
        replacements["{{LAST_NAME}}"] = profile.get("last_name", profile.get("name", "Your Name").split(" ")[-1] if " " in profile.get("name", "") else "")
        replacements["{{EMAIL}}"] = profile.get("email", "your.email@example.com")
        replacements["{{PHONE}}"] = profile.get("phone", "555-555-5555")
        replacements["{{USER_LOCATION}}"] = profile.get("location", "Your City, ST")
        replacements["{{LINKEDIN_URL}}"] = profile.get("linkedin_url", "")
        replacements["{{PORTFOLIO_URL}}"] = profile.get("portfolio_url", profile.get("github", ""))
        replacements["{{WEBSITE_URL}}"] = profile.get("website_url", "")
        
        addr = profile.get("address", {})
        replacements["{{STREET_ADDRESS}}"] = addr.get("street", "")
        replacements["{{CITY_STATE_ZIP}}"] = f"{addr.get('city', '')}, {addr.get('state', '')} {addr.get('zip', '')}".strip(", ")
        
        # Job Info
        replacements["{{COMPANY_NAME}}"] = company_name
        replacements["{{JOB_TITLE}}"] = job_title
        replacements["{{JOB_LOCATION}}"] = job_data.get("location", "N/A")
        replacements["{{JOB_SOURCE_PLATFORM}}"] = job_data.get("source", "the company website") # Scraper should provide this
        replacements["{{DATE}}"] = datetime.now().strftime("%B %d, %Y")
        
        # Cover Letter specific (can be generic if not found)
        replacements["{{HIRING_MANAGER_NAME_OR_TITLE}}"] = job_data.get("hiring_manager", "Hiring Team")
        replacements["{{SALUTATION_RECIPIENT}}"] = job_data.get("hiring_manager_salutation", "Hiring Team") # e.g. "Mr. Smith" or "Hiring Team"
        replacements["{{COMPANY_ADDRESS_LINE1}}"] = job_data.get("company_address_line1", "") # Might need to be scraped or defaulted
        replacements["{{COMPANY_CITY_STATE_ZIP}}"] = job_data.get("company_city_state_zip", "")

        # Skills list (simple join for now, could be Ollama-enhanced for prioritization)
        profile_skills = profile.get("skills", [])
        job_extracted_keywords = job_data.get("job_extracted_keywords", []) # From scraper's Ollama call
        
        # Combine and unique skills, prioritizing job-relevant ones
        combined_skills = list(dict.fromkeys(profile_skills + [kw for kw in job_extracted_keywords if kw in profile.get("keywords", [])]))
        replacements["{{SKILLS_LIST}}"] = ", ".join(filter(None, combined_skills)) if combined_skills else "Relevant technical and professional skills."

        # Ollama-generated sections
        base_resume_summary = profile.get("summary_statement", "A dedicated professional seeking new opportunities.")
        replacements["{{RESUME_SUMMARY_OLLAMA}}"] = self._generate_custom_section_with_ollama(
            "resume_summary", base_resume_summary, job_description, job_title, company_name,
            profile_skills, profile.get("keywords", [])
        )

        # For cover letter, we might use a more generic base or specific placeholders in template
        base_cover_letter_body = (f"I am writing to express my keen interest in the {job_title} position at {company_name}. "
                                  "My background and skills make me a strong candidate.") # Generic base
        replacements["{{COVER_LETTER_BODY_OLLAMA}}"] = self._generate_custom_section_with_ollama(
            "cover_letter_body", base_cover_letter_body, job_description, job_title, company_name,
            profile_skills, profile.get("keywords", [])
        )

        logger.debug(f"Generated replacements for {job_title} at {company_name}: {list(replacements.keys())}")
        return replacements

    def _replace_text_in_document(self, doc: Document, replacements: Dict[str, str]):
        """Replaces placeholders in paragraphs, tables, headers, and footers."""
        
        # Helper to replace in elements containing paragraphs (like body, cell, header)
        def replace_in_container(container):
            for paragraph in container.paragraphs:
                # Iterate through a copy of runs because replacing text can change the run structure
                # This is a simplified replacement. For complex formatting within placeholders,
                # a more sophisticated run-level replacement is needed.
                # For now, we assume placeholders are not split across runs or heavily formatted internally.
                inline = paragraph.runs
                for i in range(len(inline)):
                    text = inline[i].text
                    for key, value in replacements.items():
                        if key in text:
                            text = text.replace(key, str(value))
                    inline[i].text = text # Update the run's text
            
            # Also handle tables within the container (e.g. tables in headers/footers)
            for table in container.tables:
                for row in table.rows:
                    for cell in row.cells:
                        replace_in_container(cell) # Recursive call for cells

        # Process main body
        replace_in_container(doc)

        # Process headers and footers
        for section in doc.sections:
            for header_footer_type in ['header', 'first_page_header', 'even_page_header', 
                                       'footer', 'first_page_footer', 'even_page_footer']:
                container = getattr(section, header_footer_type, None)
                if container:
                    replace_in_container(container)


    def _generate_and_convert_document(self, job_data: Dict[str, Any], template_path: Path, doc_type: str) -> Tuple[Optional[Path], Optional[Path]]:
        """
        Generates a customized DOCX document, then attempts to convert it to PDF.
        Returns (path_to_docx, path_to_pdf_or_docx_if_failed).
        path_to_pdf_or_docx_if_failed will be PDF if conversion is successful, else DOCX.
        """
        job_title_safe = job_data.get('title', 'N/A')
        company_safe = job_data.get('company', 'N/A')
        logger.info(f"Generating {doc_type} for '{job_title_safe}' at '{company_safe}' using template {template_path.name}")

        profile_name_sanitized = utils.sanitize_filename(self.profile_data.get("_profile_name", "Profile"))
        job_title_sanitized = utils.sanitize_filename(job_title_safe)
        company_sanitized = utils.sanitize_filename(company_safe)
        
        base_filename = f"{doc_type}_{profile_name_sanitized}_{job_title_sanitized}_{company_sanitized}"
        docx_filename = f"{base_filename}.docx"
        custom_docx_path = self.output_dir / docx_filename

        if custom_docx_path.exists():
            question = f"Customized {doc_type} '{docx_filename}' already exists. Overwrite?"
            if not self._ask_user_confirmation(question, default_yes=False):
                logger.info(f"Using existing customized {doc_type}: {custom_docx_path}")
                # Attempt PDF conversion for existing DOCX if PDF is preferred
                pdf_to_use = self._handle_pdf_conversion_for_existing_docx(custom_docx_path, base_filename, doc_type)
                return custom_docx_path, pdf_to_use

        doc = self._load_template(template_path)
        replacements = self._get_replacements(job_data)
        self._replace_text_in_document(doc, replacements)
        
        try:
            doc.save(str(custom_docx_path))
            logger.info(f"Customized {doc_type} (DOCX) saved to: {custom_docx_path}")
        except Exception as e:
            logger.error(f"Error saving generated {doc_type} DOCX {custom_docx_path}: {e}", exc_info=True)
            raise # Re-raise to be caught by main application loop for proper logging
        
        # Now convert the newly generated DOCX to PDF
        pdf_to_upload = self._handle_pdf_conversion_for_existing_docx(custom_docx_path, base_filename, doc_type)
        return custom_docx_path, pdf_to_upload

    def _handle_pdf_conversion_for_existing_docx(self, docx_path: Path, base_filename: str, doc_type: str) -> Path:
        """
        Handles PDF conversion for a given DOCX file.
        Prompts user on failure. Returns path to PDF if successful, else path to DOCX.
        """
        pdf_filename = f"{base_filename}.pdf"
        pdf_path = self.output_dir / pdf_filename

        if utils.convert_doc_to_pdf_windows(str(docx_path), str(pdf_path)):
            logger.info(f"Successfully converted {doc_type} to PDF: {pdf_path}")
            return pdf_path
        else:
            logger.warning(f"Failed to convert {doc_type} '{docx_path.name}' to PDF.")
            choice = ask_user_choice(
                f"PDF conversion for {doc_type} '{docx_path.name}' failed. How to proceed?",
                ["Retry PDF conversion (ensure MS Word is closed and working)", 
                 "Upload the DOCX file instead", 
                 "Skip this job entirely"]
            )
            if choice == 1: # Retry
                if utils.convert_doc_to_pdf_windows(str(docx_path), str(pdf_path)):
                    logger.info(f"Successfully converted {doc_type} to PDF on retry: {pdf_path}")
                    return pdf_path
                else:
                    logger.error("PDF conversion failed on retry. Will use DOCX.")
                    return docx_path
            elif choice == 2: # Upload DOCX
                logger.info(f"User chose to upload DOCX for {doc_type}: {docx_path}")
                return docx_path
            else: # Skip job (choice 3 or None)
                logger.info(f"User chose to skip job due to PDF conversion failure for {doc_type}.")
                raise Exception("UserSkippedJobDueToPDFConversionFailure") # Special exception

    def generate_resume(self, job_data: Dict[str, Any]) -> Tuple[Optional[Path], Optional[Path]]:
        """Generates resume. Returns (path_to_docx, path_to_pdf_or_docx_for_upload)."""
        return self._generate_document(job_data, self.base_resume_docx_path, "Resume")

    def generate_cover_letter(self, job_data: Dict[str, Any]) -> Tuple[Optional[Path], Optional[Path]]:
        """Generates cover letter. Returns (path_to_docx, path_to_pdf_or_docx_for_upload)."""
        return self._generate_document(job_data, self.base_cover_letter_docx_path, "CoverLetter")


if __name__ == '__main__':
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    logger.info("DocumentGenerator direct execution for testing.")

    current_script_dir = Path(__file__).resolve().parent
    dummy_profiles_base_dir = current_script_dir / "profiles" 
    dummy_output_base_dir = current_script_dir / "output" 
    profile_name_for_test = "TestUserProfileOllama" 
    
    # Specific profile folder for this test
    profile_folder_for_test = dummy_profiles_base_dir / profile_name_for_test
    dummy_profile_specific_output_dir = dummy_output_base_dir / profile_name_for_test / "documents"
    
    profile_folder_for_test.mkdir(parents=True, exist_ok=True)
    dummy_profile_specific_output_dir.mkdir(parents=True, exist_ok=True)

    test_profile = {
        "_profile_name": profile_name_for_test, # Added for consistency with main flow
        "name": "Dr. Ollama TestUser",
        "first_name": "Ollama", "last_name": "TestUser",
        "email": "ollama.test@example.com", "phone": "800-OLLAMA-AI",
        "location": "Localhost, LLM Land",
        "linkedin_url": "https://linkedin.com/in/ollamatest", "portfolio_url": "https://github.com/ollamatest",
        "summary_statement": "An innovative AI enthusiast with foundational experience in large language models and document automation. Eager to apply cutting-edge techniques to real-world problems.",
        "skills": ["Python", "Ollama Integration", "DOCX Templating", "Prompt Engineering", "Natural Language Processing"],
        "keywords": ["AI", "LLM", "Automation", "Python", "Document Generation", "Mistral"],
        "resume_docx": "OllamaTest_Resume_Template.docx",
        "cover_letter_docx": "OllamaTest_CoverLetter_Template.docx",
        "ollama_model": "mistral"
    }
    
    # Save dummy profile JSON
    with open(profile_folder_for_test / f"{profile_name_for_test}.json", "w") as f:
        json.dump(test_profile, f, indent=2)

    # Create dummy resume template file
    resume_template_file_path = profile_folder_for_test / test_profile["resume_docx"]
    doc_resume = Document()
    doc_resume.add_heading("{{FULL_NAME}}", level=1)
    doc_resume.add_paragraph("{{EMAIL}} | {{PHONE}} | {{USER_LOCATION}}")
    doc_resume.add_paragraph("LinkedIn: {{LINKEDIN_URL}} | Portfolio: {{PORTFOLIO_URL}}")
    doc_resume.add_heading("Summary", level=2)
    doc_resume.add_paragraph("{{RESUME_SUMMARY_OLLAMA}}") 
    doc_resume.add_heading("Skills", level=2)
    doc_resume.add_paragraph("{{SKILLS_LIST}}") 
    doc_resume.add_paragraph(f"\nApplied for: {{JOB_TITLE}} at {{COMPANY_NAME}} on {{DATE}}.")
    doc_resume.save(resume_template_file_path)
    logger.info(f"Created dummy resume template: {resume_template_file_path}")

    # Create dummy cover letter template file
    cover_letter_template_file_path = profile_folder_for_test / test_profile["cover_letter_docx"]
    doc_cl = Document()
    doc_cl.add_paragraph("{{DATE}}")
    doc_cl.add_paragraph("\n{{FULL_NAME}}\n{{USER_LOCATION}}")
    doc_cl.add_paragraph("\n{{HIRING_MANAGER_NAME_OR_TITLE}}\n{{COMPANY_NAME}}")
    doc_cl.add_paragraph("\nDear {{SALUTATION_RECIPIENT}},")
    doc_cl.add_paragraph("{{COVER_LETTER_BODY_OLLAMA}}")
    doc_cl.add_paragraph("Thank you for your consideration.")
    doc_cl.add_paragraph("\nSincerely,\n{{FULL_NAME}}")
    doc_cl.save(cover_letter_template_file_path)
    logger.info(f"Created dummy cover letter template: {cover_letter_template_file_path}")

    test_job = {
        "title": "AI Prompt Engineer",
        "company": "InnovateLLM Corp.",
        "location": "Remote",
        "description": "Seeking a creative AI Prompt Engineer to develop and refine prompts for our Mistral-based LLMs. Must have Python skills and experience with document automation. Familiarity with NLP concepts is a plus. You will work on generating high-quality text for various applications.",
        "summary": "AI Prompt Engineer role focusing on Mistral LLMs, Python, and NLP.",
        "keywords_found": ["AI", "Prompt Engineer", "Mistral", "LLM", "Python", "NLP"] # Scraper might provide this
    }

    print("\n--- IMPORTANT: Ensure Ollama is running with the 'mistral' model pulled. ---")
    print("--- (e.g., `ollama pull mistral` and `ollama serve` or `ollama run mistral`) ---")
    print("--- Also, ensure MS Word is installed for PDF conversion if on Windows. ---")
    input("Press Enter to continue test if Ollama and Word (if applicable) are ready...")

    try:
        generator = DocumentGenerator(
            profile_data=test_profile,
            profile_folder_path=profile_folder_for_test,
            output_dir_for_profile_docs=dummy_profile_specific_output_dir,
        )
        
        if generator.ollama_available:
            logger.info("Ollama is available, proceeding with LLM-enhanced generation.")
        else:
            logger.warning("Ollama not available. Test will use fallback mechanisms for LLM content.")

        # Generate resume
        generated_resume_docx, resume_for_upload = generator.generate_resume(test_job)
        if generated_resume_docx:
            logger.info(f"Test resume DOCX generated: {generated_resume_docx}")
            logger.info(f"Resume file for upload: {resume_for_upload} (Type: {Path(resume_for_upload).suffix})")
        else:
            logger.error("Test resume generation failed.")


        # Generate cover letter
        generated_cl_docx, cl_for_upload = generator.generate_cover_letter(test_job)
        if generated_cl_docx:
            logger.info(f"Test cover letter DOCX generated: {generated_cl_docx}")
            logger.info(f"Cover letter file for upload: {cl_for_upload} (Type: {Path(cl_for_upload).suffix})")
        else:
            logger.error("Test cover letter generation failed.")
            
        print(f"\nGenerated documents are in: {dummy_profile_specific_output_dir}")
        print("Please review them.")

    except Exception as e_main:
        logger.error(f"An error occurred during the DocumentGenerator test run: {e_main}", exc_info=True)
    finally:
        logger.info("DocumentGenerator test finished.")
        # Clean up dummy files (optional, for repeated testing)
        # import shutil
        # if dummy_profiles_base_dir.exists(): shutil.rmtree(dummy_profiles_base_dir)
        # if dummy_output_base_dir.exists(): shutil.rmtree(dummy_output_base_dir)
