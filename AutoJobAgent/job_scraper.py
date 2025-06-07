import json
import random
import re
import time
from pathlib import Path
from typing import Dict, Generator, List, Optional, Union

import ollama
import requests
from bs4 import BeautifulSoup
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError
from rich.console import Console

console = Console()

class JobScraper:
    """
    JobScraper class that handles scraping jobs from various job sites.
    Supports Indeed, Eluta, and can be extended for other job sites.
    """
    
    def __init__(self, profile: Dict, start: int = 0):
        """
        Initialize the JobScraper with a profile and starting index.
        
        Args:
            profile: User profile dictionary with search parameters
            start: Starting index for pagination
        """
        self.profile = profile
        self.current_index = start
        self.keywords = profile.get("keywords", [])
        self.location = profile.get("location", "")
        self.ollama_model = profile.get("ollama_model", "mistral:7b")
        
        # Extract city from location for better search results
        self.city = self.location.split(",")[0].strip() if self.location else ""
        
        # Convert keywords list to search string
        if isinstance(self.keywords, list):
            self.search_terms = " ".join(self.keywords)
        else:
            self.search_terms = str(self.keywords)
    
    def batched(self, batch_size: int) -> Generator[List[Dict], None, None]:
        """
        Generate batches of jobs with the specified batch size.
        
        Args:
            batch_size: Number of jobs to include in each batch
            
        Yields:
            List of job dictionaries in each batch
        """
        batch = []
        
        while True:
            try:
                # Try to get jobs from Indeed first
                for job in self.scrape_indeed():
                    # Extract keywords using Ollama
                    try:
                        job = self.extract_keywords(job)
                    except Exception as e:
                        console.print(f"[yellow]Warning: Failed to extract keywords: {e}[/yellow]")
                    
                    batch.append(job)
                    
                    if len(batch) >= batch_size:
                        yield batch
                        batch = []
                        break
                
                # If batch is still not full, try Eluta
                if len(batch) < batch_size:
                    for job in self.scrape_eluta():
                        # Extract keywords using Ollama
                        try:
                            job = self.extract_keywords(job)
                        except Exception as e:
                            console.print(f"[yellow]Warning: Failed to extract keywords: {e}[/yellow]")
                        
                        batch.append(job)
                        
                        if len(batch) >= batch_size:
                            yield batch
                            batch = []
                            break
                
                # If we still have jobs in the batch, yield them
                if batch:
                    yield batch
                    batch = []
                
                # Increment the current index for pagination
                self.current_index += batch_size
                
                # Add a delay to avoid rate limiting
                time.sleep(random.uniform(1.5, 3.0))
                
            except Exception as e:
                console.print(f"[bold red]Error in job scraping: {e}[/bold red]")
                # If there was an error but we have jobs, yield them
                if batch:
                    yield batch
                    batch = []
                time.sleep(5)  # Longer delay after error
    
    def scrape_indeed(self) -> Generator[Dict, None, None]:
        """
        Scrape jobs from Indeed.
        
        Yields:
            Job dictionaries with title, company, location, url, summary, and site
        """
        # Construct the Indeed URL with search parameters
        base_url = "https://ca.indeed.com/jobs"
        params = {
            "q": self.search_terms,
            "l": self.city,
            "start": self.current_index
        }
        
        # Construct the URL
        url_params = "&".join([f"{k}={v}" for k, v in params.items()])
        url = f"{base_url}?{url_params}"
        
        try:
            # Make the request with a user agent to avoid being blocked
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
            }
            response = requests.get(url, headers=headers, timeout=15)
            response.raise_for_status()
            
            # Parse the HTML
            soup = BeautifulSoup(response.text, "html.parser")
            
            # Find all job cards
            job_cards = soup.select("div.job_seen_beacon")
            
            if not job_cards:
                console.print("[yellow]No job cards found on Indeed. Might be rate limited or page structure changed.[/yellow]")
                return
            
            # Process each job card
            for card in job_cards:
                try:
                    # Extract job title
                    title_elem = card.select_one("h2.jobTitle a span")
                    if not title_elem:
                        title_elem = card.select_one("h2.jobTitle span")
                    title = title_elem.text.strip() if title_elem else "Unknown Title"
                    
                    # Extract company name
                    company_elem = card.select_one("span.companyName")
                    company = company_elem.text.strip() if company_elem else "Unknown Company"
                    
                    # Extract location
                    location_elem = card.select_one("div.companyLocation")
                    location = location_elem.text.strip() if location_elem else "Unknown Location"
                    
                    # Extract job URL
                    url_elem = card.select_one("h2.jobTitle a")
                    job_url = "https://ca.indeed.com" + url_elem["href"] if url_elem and "href" in url_elem.attrs else ""
                    
                    # Extract job summary
                    summary_elem = card.select_one("div.job-snippet")
                    summary = summary_elem.text.strip() if summary_elem else ""
                    
                    # Create job dictionary
                    job = {
                        "title": title,
                        "company": company,
                        "location": location,
                        "url": job_url,
                        "summary": summary,
                        "site": "Indeed"
                    }
                    
                    yield job
                    
                except Exception as e:
                    console.print(f"[yellow]Error parsing Indeed job card: {e}[/yellow]")
                    continue
                
        except requests.RequestException as e:
            console.print(f"[bold red]Error scraping Indeed: {e}[/bold red]")
            return
    
    def scrape_eluta(self, browser_context=None) -> Generator[Dict, None, None]:
        """
        Scrape jobs from Eluta using Playwright.
        
        Args:
            browser_context: Optional Playwright browser context
            
        Yields:
            Job dictionaries with title, company, location, url, summary, and site
        """
        if not browser_context:
            console.print("[yellow]No browser context provided for Eluta scraping. Skipping.[/yellow]")
            return
        
        try:
            # Create a new page in the browser context
            page = browser_context.new_page()
            
            # Construct the Eluta URL with search parameters
            base_url = "https://www.eluta.ca/search"
            
            # Navigate to the search page
            search_url = f"{base_url}?q={self.search_terms}&l={self.city}&pg={self.current_index // 10 + 1}"
            page.goto(search_url, timeout=30000)
            
            # Wait for job results to load
            page.wait_for_selector("div.resultInfo", timeout=10000)
            
            # Extract job listings
            job_elements = page.query_selector_all("div.result")
            
            for job_elem in job_elements:
                try:
                    # Extract job title
                    title_elem = job_elem.query_selector("h2.resultTitle a")
                    title = title_elem.inner_text() if title_elem else "Unknown Title"
                    
                    # Extract job URL
                    job_url = title_elem.get_attribute("href") if title_elem else ""
                    if job_url and not job_url.startswith("http"):
                        job_url = f"https://www.eluta.ca{job_url}"
                    
                    # Extract company name
                    company_elem = job_elem.query_selector("div.resultCompany")
                    company = company_elem.inner_text() if company_elem else "Unknown Company"
                    
                    # Extract location
                    location_elem = job_elem.query_selector("div.resultLocation")
                    location = location_elem.inner_text() if location_elem else "Unknown Location"
                    
                    # Extract job summary
                    summary_elem = job_elem.query_selector("div.resultDesc")
                    summary = summary_elem.inner_text() if summary_elem else ""
                    
                    # Create job dictionary
                    job = {
                        "title": title,
                        "company": company,
                        "location": location,
                        "url": job_url,
                        "summary": summary,
                        "site": "Eluta"
                    }
                    
                    yield job
                    
                except Exception as e:
                    console.print(f"[yellow]Error parsing Eluta job: {e}[/yellow]")
                    continue
            
        except PlaywrightTimeoutError:
            console.print("[yellow]Timeout while loading Eluta results[/yellow]")
        except Exception as e:
            console.print(f"[bold red]Error scraping Eluta: {e}[/bold red]")
        finally:
            if 'page' in locals():
                page.close()
    
    def extract_keywords(self, job: Dict) -> Dict:
        """
        Extract keywords from job summary using Ollama.
        
        Args:
            job: Job dictionary with summary
            
        Returns:
            Job dictionary with added keywords
        """
        if not job.get("summary"):
            job["keywords"] = []
            return job
        
        # Retry up to 3 times
        for attempt in range(3):
            try:
                prompt = (
                    "List 8 technical keywords that a resume should contain "
                    "to match this job summary, return a JSON list."
                    "\n==SUMMARY==\n" + job["summary"]
                )
                
                response = ollama.generate(
                    model=self.ollama_model,
                    prompt=prompt,
                    temperature=0,
                    max_tokens=120
                )
                
                # Extract the response text
                kw_text = response["response"]
                
                # Try to parse as JSON
                try:
                    keywords = json.loads(kw_text)
                    if isinstance(keywords, list):
                        job["keywords"] = keywords
                        return job
                except json.JSONDecodeError:
                    # If not valid JSON, try to extract keywords using regex
                    matches = re.findall(r'["\'](.*?)["\']', kw_text)
                    if matches:
                        job["keywords"] = matches
                        return job
                
                # Fallback: split by commas and clean up
                keywords = [k.strip(' "\'[]') for k in kw_text.split(',')]
                job["keywords"] = [k for k in keywords if k]
                return job
                
            except Exception as e:
                console.print(f"[yellow]Attempt {attempt+1} failed to extract keywords: {e}[/yellow]")
                time.sleep(2)  # Wait before retry
        
        # If all attempts failed, return empty keywords
        job["keywords"] = []
        return job


# Helper functions for testing
def test_indeed_scraper():
    """Test function for Indeed scraper"""
    profile = {
        "keywords": ["Python", "Data Analyst"],
        "location": "Toronto, ON",
        "ollama_model": "mistral:7b"
    }
    
    scraper = JobScraper(profile)
    
    for i, job in enumerate(scraper.scrape_indeed()):
        print(f"Job {i+1}:")
        print(f"  Title: {job['title']}")
        print(f"  Company: {job['company']}")
        print(f"  Location: {job['location']}")
        print(f"  URL: {job['url']}")
        print(f"  Summary: {job['summary'][:100]}...")
        print()
        
        if i >= 5:  # Just test with 5 jobs
            break


if __name__ == "__main__":
    test_indeed_scraper()
