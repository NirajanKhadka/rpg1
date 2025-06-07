#!/usr/bin/env python3
"""
Job Scraper Module
-----------------
Provides functionality to scrape job listings from Indeed and Eluta.ca using Playwright.
Integrates with Ollama for smart keyword extraction and comparison against profile keywords.
"""

import os
import re
import json
import logging
import asyncio
import time
from abc import ABC, abstractmethod
from typing import Dict, List, Optional, Any, Set, Tuple
from pathlib import Path
from datetime import datetime
from urllib.parse import urlparse, urljoin, quote_plus

try:
    from playwright.async_api import async_playwright, Page, Browser, BrowserContext, TimeoutError as PlaywrightTimeoutError
except ImportError:
    print("Playwright not found. Please install it with 'pip install playwright'")
    print("Then run: 'playwright install chromium' (or other browsers if needed)")
    raise

try:
    import ollama
except ImportError:
    print("ollama library not found. Please install it with 'pip install ollama'")
    ollama = None # Will be checked before use

# Assuming utils.py contains ask_user_choice and sanitize_filename
try:
    from utils import ask_user_choice, sanitize_filename # If ask_user_choice is in utils
except ImportError:
    # Fallback if utils.ask_user_choice is not available (e.g. for standalone testing)
    def ask_user_choice(question: str, choices: List[str]) -> Optional[int]:
        print(f"\n{question}")
        for i, choice in enumerate(choices):
            print(f"  {i+1}) {choice}")
        while True:
            try:
                user_input = input(f"Enter your choice (1-{len(choices)}): ").strip()
                if not user_input: continue
                choice_num = int(user_input)
                if 1 <= choice_num <= len(choices):
                    return choice_num
                else:
                    print(f"Invalid choice. Please enter a number between 1 and {len(choices)}.")
            except ValueError:
                print("Invalid input. Please enter a number.")
            except Exception: return None
    
    def sanitize_filename(filename: str, max_length: int = 200) -> str:
        if not isinstance(filename, str): filename = str(filename)
        base, ext = os.path.splitext(filename)
        sanitized_base = "".join(char for char in base if ord(char) >= 32)
        sanitized_base = re.sub(r'[<>:"/\\|?*]', '_', sanitized_base)
        sanitized_base = re.sub(r'[\s_]+', '_', sanitized_base).strip('_')
        if len(sanitized_base) > max_length: sanitized_base = sanitized_base[:max_length]
        sanitized_ext = ""
        if ext:
            sanitized_ext = "." + ext.lstrip('.').strip()
            sanitized_ext = re.sub(r'[<>:"/\\|?*]', '_', sanitized_ext)
        final_filename = sanitized_base + sanitized_ext
        if not final_filename.strip('.'):
            final_filename = f"sanitized_file_{datetime.now().strftime('%Y%m%d%H%M%S')}{sanitized_ext or '.dat'}"
        return final_filename


logger = logging.getLogger(__name__)

class JobScraper(ABC):
    """Base class for job scrapers."""

    def __init__(self, experience_level: str, ollama_model: str = "mistral"):
        self.logger = logging.getLogger(self.__class__.__name__)
        self.playwright_instance = None
        self.browser = None
        self.context = None
        self.page = None
        self.job_cache_file = Path("cache") / f"{self.__class__.__name__.lower()}_job_cache.json"
        self.seen_job_ids: Set[str] = set() # Stores job_id strings
        self.load_job_cache()
        self.experience_level = experience_level.lower() if experience_level else "entry level"
        self.ollama_model = ollama_model
        self.ollama_available = self._check_ollama_availability()

    def _check_ollama_availability(self) -> bool:
        if ollama is None:
            logger.warning("Ollama library is not installed. LLM keyword features will be basic.")
            print("Warning: Ollama library not installed. Using basic keyword matching.")
            return False
        try:
            ollama.list()
            logger.info(f"Successfully connected to Ollama. Using model: {self.ollama_model} for keyword extraction.")
            return True
        except Exception as e:
            logger.warning(f"Ollama is not available or model '{self.ollama_model}' not found: {e}. Using basic keyword matching.")
            print(f"Warning: Ollama not responding or model '{self.ollama_model}' missing. Ensure Ollama is running and the model is pulled. Falling back to basic keyword matching.")
            return False

    async def _call_ollama_for_keywords(self, text_content: str, job_title: str) -> List[str]:
        if not self.ollama_available or not text_content:
            return []
        
        prompt = (
            f"Extract the most relevant technical skills, tools, and key responsibilities from the following job description for a '{job_title}'. "
            f"Focus on concrete nouns and noun phrases. Return them as a comma-separated list. Job Description:\n\n{text_content[:2000]}" # Limit context size
        )
        try:
            logger.debug(f"Sending keyword extraction prompt to Ollama for job: {job_title}")
            response = await asyncio.to_thread(ollama.generate, model=self.ollama_model, prompt=prompt, stream=False)
            extracted_text = response.get('response', '').strip()
            if extracted_text:
                keywords = [kw.strip().lower() for kw in extracted_text.split(',') if kw.strip()]
                # Basic filtering for very short/common words - can be improved
                keywords = [kw for kw in keywords if len(kw) > 2 and kw not in ["and", "the", "for", "with"]]
                logger.info(f"Ollama extracted keywords for '{job_title}': {keywords[:10]}...") # Log first 10
                return list(set(keywords)) # Unique keywords
            return []
        except Exception as e:
            logger.error(f"Error calling Ollama for keyword extraction: {e}", exc_info=True)
            return []

    def _compare_keywords(self, job_extracted_keywords: List[str], profile_keywords: List[str]) -> Tuple[List[str], List[str]]:
        if not profile_keywords:
            return [], [] # No profile keywords to compare against
        if not job_extracted_keywords: # If Ollama failed or job desc was empty
            return [], list(profile_keywords) # All profile keywords are missing

        profile_keywords_lower = {kw.lower() for kw in profile_keywords}
        job_keywords_lower = {kw.lower() for kw in job_extracted_keywords}
        
        found_in_job = list(profile_keywords_lower.intersection(job_keywords_lower))
        missing_from_job = list(profile_keywords_lower.difference(job_keywords_lower))
        
        # Return original casing for found keywords if possible, otherwise lowercased
        # This is a simple approach; more sophisticated mapping could be used
        found_original_case = [next((pk for pk in profile_keywords if pk.lower() == fkw), fkw) for fkw in found_in_job]
        missing_original_case = [next((pk for pk in profile_keywords if pk.lower() == mkw), mkw) for mkw in missing_from_job]

        return found_original_case, missing_original_case

    async def initialize_browser(self, headless: bool = True, browser_type: str = "chromium"):
        if self.browser: # Already initialized
            return
        try:
            self.playwright_instance = await async_playwright().start()
            browser_map = {
                "chromium": self.playwright_instance.chromium,
                "firefox": self.playwright_instance.firefox,
                "webkit": self.playwright_instance.webkit,
            }
            browser_launcher = browser_map.get(browser_type.lower(), self.playwright_instance.chromium)
            self.browser = await browser_launcher.launch(headless=headless)
            self.context = await self.browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
                viewport={"width": 1920, "height": 1080},
                java_script_enabled=True,
            )
            self.page = await self.context.new_page()
            self.page.set_default_timeout(30000) # 30 seconds
            logger.info(f"Playwright browser ({browser_type}, headless={headless}) initialized.")
        except Exception as e:
            logger.error(f"Failed to initialize Playwright browser: {e}", exc_info=True)
            await self.close_browser()
            raise

    async def close_browser(self):
        if self.page: await self.page.close()
        if self.context: await self.context.close()
        if self.browser: await self.browser.close()
        if self.playwright_instance: await self.playwright_instance.stop()
        self.page, self.context, self.browser, self.playwright_instance = None, None, None, None
        logger.info("Playwright browser closed.")

    def load_job_cache(self):
        self.job_cache_file.parent.mkdir(parents=True, exist_ok=True)
        if self.job_cache_file.exists():
            try:
                with open(self.job_cache_file, 'r', encoding='utf-8') as f:
                    self.seen_job_ids = set(json.load(f).get("seen_job_ids", []))
                logger.info(f"Loaded {len(self.seen_job_ids)} job IDs from cache: {self.job_cache_file}")
            except Exception as e:
                logger.error(f"Error loading job cache {self.job_cache_file}: {e}")
                self.seen_job_ids = set()
        else:
            logger.info(f"No job cache found at {self.job_cache_file}. Starting fresh.")

    def save_job_cache(self):
        self.job_cache_file.parent.mkdir(parents=True, exist_ok=True)
        try:
            with open(self.job_cache_file, 'w', encoding='utf-8') as f:
                json.dump({"seen_job_ids": list(self.seen_job_ids)}, f)
            logger.info(f"Saved {len(self.seen_job_ids)} job IDs to cache: {self.job_cache_file}")
        except Exception as e:
            logger.error(f"Error saving job cache {self.job_cache_file}: {e}")

    def is_job_seen(self, job_id: str) -> bool:
        return job_id in self.seen_job_ids

    def mark_job_as_seen(self, job_id: str):
        self.seen_job_ids.add(job_id)

    async def _should_skip_based_on_experience(self, job_title: str) -> bool:
        """
        Checks if a job title suggests a seniority level mismatching the profile's experience_level.
        Prompts user if a potential mismatch is found.
        """
        title_lower = job_title.lower()
        
        senior_keywords = ["senior", "sr.", "lead", "principal", "director", "manager", "vp", "head of", "chief"]
        entry_keywords = ["junior", "jr.", "entry level", "intern", "trainee", "graduate"]

        is_senior_job = any(skw in title_lower for skw in senior_keywords)
        is_entry_job = any(ekw in title_lower for ekw in entry_keywords)

        if self.experience_level not in ["senior", "lead", "principal", "manager", "director"]: # User is not senior
            if is_senior_job:
                logger.info(f"Potential seniority mismatch: Job '{job_title}' (seems senior) vs Profile '{self.experience_level}'.")
                if not (ask_user_choice(f"Job title '{job_title}' sounds senior. Your level is '{self.experience_level}'. Still consider?", ["Yes", "No, skip"]) == 1):
                    logger.info(f"User chose to skip senior job: {job_title}")
                    return True
        elif self.experience_level in ["senior", "lead", "principal", "manager", "director"]: # User is senior
            if is_entry_job:
                logger.info(f"Potential seniority mismatch: Job '{job_title}' (seems entry-level) vs Profile '{self.experience_level}'.")
                if not (ask_user_choice(f"Job title '{job_title}' sounds entry-level. Your level is '{self.experience_level}'. Still consider?", ["Yes", "No, skip"]) == 1):
                    logger.info(f"User chose to skip entry-level job: {job_title}")
                    return True
        return False

    async def _get_job_details_page_content(self, job_url: str) -> Optional[str]:
        if not self.page or not job_url:
            return None
        try:
            await self.page.goto(job_url, wait_until="domcontentloaded", timeout=20000)
            # A more robust way to wait for dynamic content might be needed for some sites
            await self.page.wait_for_timeout(2000) # Small delay for JS rendering
            content = await self.page.content()
            logger.debug(f"Fetched content for job details page: {job_url[:80]}...")
            return content
        except PlaywrightTimeoutError:
            logger.warning(f"Timeout loading job details page: {job_url}")
        except Exception as e:
            logger.error(f"Error fetching job details page {job_url}: {e}")
        return None

    @abstractmethod
    async def scrape_jobs(self, search_keywords: List[str], search_location: str, profile_keywords: List[str], num_jobs_to_fetch: int, last_job_id_processed: Optional[str] = None) -> List[Dict[str, Any]]:
        pass

    def _generate_job_id(self, job_data: Dict[str, Any], source: str) -> str:
        """Generates a unique job ID, preferring platform-specific IDs if available."""
        # Try to find a platform-specific ID in the URL or attributes
        url = job_data.get('url', '')
        if source == "indeed" and 'jk=' in url:
            return f"indeed_{url.split('jk=')[1].split('&')[0]}"
        if source == "eluta" and '/r/' in url: # Eluta often has /r/job-id
             match = re.search(r'/r/([^/?]+)', url)
             if match: return f"eluta_{match.group(1)}"
        
        # Fallback to hashing title and company if no specific ID found
        title = job_data.get('title', 'N/A').strip()
        company = job_data.get('company', 'N/A').strip()
        # Add location to hash for more uniqueness if needed, but can also lead to more "new" jobs if location string varies slightly
        # location = job_data.get('location', 'N/A').strip()
        # hash_input = f"{title}_{company}_{location}"
        hash_input = f"{title}_{company}"
        return f"{source}_{hash(hash_input)}"


class IndeedScraper(JobScraper):
    BASE_URL = "https://www.indeed.com"

    async def scrape_jobs(self, search_keywords: List[str], search_location: str, profile_keywords: List[str], num_jobs_to_fetch: int, last_job_id_processed: Optional[str] = None) -> List[Dict[str, Any]]:
        if not self.page: await self.initialize_browser()
        
        query = " ".join(search_keywords)
        location_query = quote_plus(search_location)
        jobs_found = []
        start_index = 0
        processed_this_session_count = 0
        
        # Resuming logic: if last_job_id_processed is provided, we need to find it first.
        # This is complex with pagination. For now, we'll rely on the seen_job_ids cache primarily.
        # A more robust resume would involve storing the last search URL/page number.
        
        while processed_this_session_count < num_jobs_to_fetch:
            search_url = f"{self.BASE_URL}/jobs?q={quote_plus(query)}&l={location_query}&start={start_index}"
            logger.info(f"Scraping Indeed: {search_url}")
            try:
                await self.page.goto(search_url, wait_until="domcontentloaded", timeout=45000)
                await self.page.wait_for_selector('ul.jobsearch-ResultsList', timeout=20000)
            except PlaywrightTimeoutError:
                logger.error(f"Timeout loading Indeed search results page: {search_url}")
                break # Stop if search page doesn't load
            except Exception as e:
                logger.error(f"Error navigating to Indeed search page {search_url}: {e}")
                break

            job_cards = await self.page.query_selector_all('div.job_seen_beacon') # More reliable selector
            if not job_cards:
                logger.info("No more job cards found on Indeed page.")
                break # No more jobs on this page or subsequent pages

            for card in job_cards:
                if processed_this_session_count >= num_jobs_to_fetch:
                    break
                try:
                    title_element = await card.query_selector('h2.jobTitle > a')
                    company_element = await card.query_selector('span.companyName')
                    location_element = await card.query_selector('div.companyLocation')
                    summary_element = await card.query_selector('div.job-snippet')

                    if not title_element: continue

                    job_title = (await title_element.inner_text()).strip()
                    job_url = urljoin(self.BASE_URL, await title_element.get_attribute('href'))
                    company = (await company_element.inner_text()).strip() if company_element else "N/A"
                    location = (await location_element.inner_text()).strip() if location_element else "N/A"
                    summary = (await summary_element.inner_text()).strip() if summary_element else ""
                    
                    job_id = self._generate_job_id({"title": job_title, "company": company, "url": job_url}, "indeed")

                    if self.is_job_seen(job_id):
                        logger.debug(f"Skipping already seen Indeed job: {job_title} (ID: {job_id})")
                        continue
                    
                    if await self._should_skip_based_on_experience(job_title):
                        self.mark_job_as_seen(job_id) # Mark as seen even if skipped by experience to avoid re-prompting
                        continue

                    # Fetch full description for keyword extraction
                    full_description_content = await self._get_job_details_page_content(job_url)
                    job_description_text = summary # Fallback to summary
                    if full_description_content:
                        # Basic parsing, can be improved with more specific selectors if Indeed's detail page structure is consistent
                        from bs4 import BeautifulSoup
                        soup = BeautifulSoup(full_description_content, 'html.parser')
                        desc_container = soup.find('div', id='jobDescriptionText')
                        if desc_container:
                            job_description_text = desc_container.get_text(separator='\n', strip=True)
                        else:
                            logger.warning(f"Could not find #jobDescriptionText for {job_url}, using summary for keywords.")
                    
                    job_extracted_keywords = await self._call_ollama_for_keywords(job_description_text, job_title)
                    found_keywords, missing_keywords = self._compare_keywords(job_extracted_keywords, profile_keywords)

                    job_data = {
                        "title": job_title, "company": company, "location": location, "url": job_url,
                        "summary": summary, "description": job_description_text, "job_id": job_id,
                        "source": "indeed", "date_found": datetime.now().isoformat(),
                        "job_extracted_keywords": job_extracted_keywords,
                        "keywords_found": found_keywords,
                        "keywords_missing": missing_keywords
                    }
                    jobs_found.append(job_data)
                    self.mark_job_as_seen(job_id)
                    processed_this_session_count += 1
                    logger.info(f"Scraped Indeed job: {job_title} at {company}")

                except Exception as e:
                    logger.error(f"Error scraping an Indeed job card: {e}", exc_info=True)
            
            start_index += len(job_cards) # Indeed typically shows 10-15 jobs per page
            if not await self.page.query_selector('a[data-testid="pagination-page-next"]'): # Check for next page link
                 logger.info("No 'next page' button found on Indeed. Ending scrape for this keyword set.")
                 break
            await asyncio.sleep(2) # Brief pause before next page

        self.save_job_cache()
        logger.info(f"Finished Indeed scraping. Found {len(jobs_found)} new jobs in this session.")
        return jobs_found


class ElutaScraper(JobScraper):
    BASE_URL = "https://www.eluta.ca"

    async def scrape_jobs(self, search_keywords: List[str], search_location: str, profile_keywords: List[str], num_jobs_to_fetch: int, last_job_id_processed: Optional[str] = None) -> List[Dict[str, Any]]:
        if not self.page: await self.initialize_browser()

        query = " ".join(search_keywords)
        location_query = quote_plus(search_location)
        jobs_found = []
        page_num = 1
        processed_this_session_count = 0

        while processed_this_session_count < num_jobs_to_fetch:
            # Eluta uses 'pg' parameter for pagination, starting from 1
            search_url = f"{self.BASE_URL}/search?q={quote_plus(query)}&l={location_query}&pg={page_num}"
            logger.info(f"Scraping Eluta: {search_url}")
            try:
                await self.page.goto(search_url, wait_until="domcontentloaded", timeout=45000)
                await self.page.wait_for_selector('div.organic-job', timeout=20000) # Main container for organic job listings
            except PlaywrightTimeoutError:
                logger.info(f"Timeout or no 'organic-job' found on Eluta page {page_num}. Assuming end of results.")
                break
            except Exception as e:
                logger.error(f"Error navigating to Eluta search page {search_url}: {e}")
                break

            job_cards = await self.page.query_selector_all('div.organic-job')
            if not job_cards:
                logger.info(f"No more job cards found on Eluta page {page_num}.")
                break

            for card in job_cards:
                if processed_this_session_count >= num_jobs_to_fetch:
                    break
                try:
                    title_element = await card.query_selector('h2.title a')
                    company_element = await card.query_selector('span.employer')
                    location_element = await card.query_selector('span.location')
                    summary_element = await card.query_selector('span.description') # Eluta's snippet

                    if not title_element: continue

                    job_title = (await title_element.inner_text()).strip()
                    job_url_relative = await title_element.get_attribute('href')
                    job_url = urljoin(self.BASE_URL, job_url_relative) if job_url_relative else "N/A"
                    
                    company = (await company_element.inner_text()).strip() if company_element else "N/A"
                    location = (await location_element.inner_text()).strip() if location_element else "N/A"
                    summary = (await summary_element.inner_text()).strip() if summary_element else ""

                    job_id = self._generate_job_id({"title": job_title, "company": company, "url": job_url}, "eluta")

                    if self.is_job_seen(job_id):
                        logger.debug(f"Skipping already seen Eluta job: {job_title} (ID: {job_id})")
                        continue
                    
                    if await self._should_skip_based_on_experience(job_title):
                        self.mark_job_as_seen(job_id)
                        continue

                    # Eluta job detail pages are often external. We'll use the summary for keyword extraction.
                    # For a more thorough approach, one might try to follow the link if it's still on eluta.ca
                    # or handle common ATS patterns if it redirects.
                    job_description_text = summary 
                    if self.ollama_available and not job_description_text and job_url.startswith(self.BASE_URL): # If summary is empty and URL is internal
                        logger.info(f"Fetching Eluta job details page for keywords: {job_url}")
                        detail_page_content = await self._get_job_details_page_content(job_url)
                        if detail_page_content:
                            from bs4 import BeautifulSoup
                            soup = BeautifulSoup(detail_page_content, 'html.parser')
                            # Eluta's detail page structure varies, this is a guess
                            desc_container = soup.find('div', class_='job-description') or soup.find('section', id='job-description')
                            if desc_container:
                                job_description_text = desc_container.get_text(separator='\n', strip=True)
                    
                    job_extracted_keywords = await self._call_ollama_for_keywords(job_description_text or summary, job_title)
                    found_keywords, missing_keywords = self._compare_keywords(job_extracted_keywords, profile_keywords)

                    job_data = {
                        "title": job_title, "company": company, "location": location, "url": job_url,
                        "summary": summary, "description": job_description_text, "job_id": job_id,
                        "source": "eluta", "date_found": datetime.now().isoformat(),
                        "job_extracted_keywords": job_extracted_keywords,
                        "keywords_found": found_keywords,
                        "keywords_missing": missing_keywords
                    }
                    jobs_found.append(job_data)
                    self.mark_job_as_seen(job_id)
                    processed_this_session_count += 1
                    logger.info(f"Scraped Eluta job: {job_title} at {company}")

                except Exception as e:
                    logger.error(f"Error scraping an Eluta job card: {e}", exc_info=True)
            
            page_num += 1
            # Check for a "Next" button or if the number of results indicates more pages
            # Eluta's pagination might not have an obvious next button if it's the last page.
            # A more robust check would be to see if the number of jobs found is less than expected per page.
            next_page_link = await self.page.query_selector(f'a[href*="pg={page_num}"]') # Check if next page link exists
            if not next_page_link:
                 logger.info(f"No 'next page' (pg={page_num}) link found on Eluta. Ending scrape.")
                 break
            await asyncio.sleep(2)

        self.save_job_cache()
        logger.info(f"Finished Eluta scraping. Found {len(jobs_found)} new jobs in this session.")
        return jobs_found

# Placeholder for other scrapers if needed in the future
# class WorkdayScraper(JobScraper): ...
# class LinkedInScraper(JobScraper): ...


async def scrape_jobs(
    job_site: str, 
    search_keywords: List[str], 
    search_location: str, 
    profile_keywords: List[str],
    experience_level: str,
    num_jobs_to_fetch: int,
    last_job_id_processed: Optional[str] = None,
    ollama_model: str = "mistral",
    browser_type: str = "chromium",
    headless: bool = True
) -> List[Dict[str, Any]]:
    """
    Factory function to get and run the appropriate job scraper.
    """
    scraper_map = {
        "indeed": IndeedScraper,
        "eluta": ElutaScraper,
        # Add other scrapers here
    }
    scraper_class = scraper_map.get(job_site.lower())

    if not scraper_class:
        logger.error(f"Unsupported job scraping site: {job_site}")
        print(f"Error: Job scraping not supported for '{job_site}'. Supported: {', '.join(scraper_map.keys())}")
        return []

    scraper = scraper_class(experience_level=experience_level, ollama_model=ollama_model)
    
    try:
        await scraper.initialize_browser(headless=headless, browser_type=browser_type)
        jobs = await scraper.scrape_jobs(
            search_keywords=search_keywords,
            search_location=search_location,
            profile_keywords=profile_keywords,
            num_jobs_to_fetch=num_jobs_to_fetch,
            last_job_id_processed=last_job_id_processed
        )
        return jobs
    except Exception as e:
        logger.error(f"An error occurred during scraping for {job_site}: {e}", exc_info=True)
        return []
    finally:
        if scraper:
            await scraper.close_browser()

# Command-line interface for testing the scraper module directly
async def _cli_main():
    parser = argparse.ArgumentParser(description="Job Scraper CLI")
    parser.add_argument("job_site", choices=["indeed", "eluta"], help="Job site to scrape")
    parser.add_argument("-k", "--keywords", nargs="+", required=True, help="Job keywords (e.g., 'Python Developer')")
    parser.add_argument("-l", "--location", required=True, help="Location (e.g., 'Toronto, ON')")
    parser.add_argument("-pk", "--profile_keywords", nargs="*", default=[], help="Keywords from user profile for matching")
    parser.add_argument("-exp", "--experience", default="Entry Level", help="User's experience level")
    parser.add_argument("-n", "--num_jobs", type=int, default=5, help="Number of new jobs to fetch")
    parser.add_argument("--last_id", help="Last processed job ID for resuming", default=None)
    parser.add_argument("--output_file", help="JSON file to save results")
    parser.add_argument("--headless", action="store_true", help="Run browser in headless mode")
    parser.add_argument("--verbose", action="store_true", help="Enable verbose logging")
    
    args = parser.parse_args()

    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=log_level, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    
    console = utils.Console() if utils.RICH_AVAILABLE else utils.Console() # Use utils console
    console.print(f"[bold blue]Starting {args.job_site} scraper...[/bold blue]")
    console.print(f"Keywords: {args.keywords}, Location: {args.location}, Experience: {args.experience}")

    found_jobs = await scrape_jobs(
        job_site=args.job_site,
        search_keywords=args.keywords,
        search_location=args.location,
        profile_keywords=args.profile_keywords,
        experience_level=args.experience,
        num_jobs_to_fetch=args.num_jobs,
        last_job_id_processed=args.last_id,
        headless=args.headless
    )

    console.print(f"\n[bold green]Found {len(found_jobs)} new jobs.[/bold green]")
    for i, job in enumerate(found_jobs):
        console.print(f"\n[bold]Job {i+1}: {job.get('title')} at {job.get('company')}[/bold]")
        console.print(f"  URL: {job.get('url')}")
        console.print(f"  Job ID: {job.get('job_id')}")
        console.print(f"  Extracted Keywords: {job.get('job_extracted_keywords')}")
        console.print(f"  Profile Keywords Found: {job.get('keywords_found')}")
        console.print(f"  Profile Keywords Missing: {job.get('keywords_missing')}")


    if args.output_file:
        try:
            with open(args.output_file, 'w', encoding='utf-8') as f:
                json.dump(found_jobs, f, indent=2, ensure_ascii=False)
            console.print(f"\nResults saved to {args.output_file}")
        except Exception as e:
            console.print(f"[bold red]Error saving results to file: {e}[/bold red]")

if __name__ == "__main__":
    # This allows running the scraper independently for testing
    # Ensure utils.py is in the same directory or accessible via PYTHONPATH
    # Example: python job_scraper.py indeed -k "Software Engineer" -l "Remote" -pk "Python" "API" "Django" -n 3 --verbose
    asyncio.run(_cli_main())
