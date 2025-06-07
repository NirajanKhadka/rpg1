import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright
from rich.console import Console
from rich.panel import Panel
from rich.prompt import Confirm, Prompt

import document_generator
import job_scraper
import utils
from ats import detect, get_submitter

console = Console()

def parse_args():
    parser = argparse.ArgumentParser(description="AutoJobAgent - Automated job application tool")
    parser.add_argument("profile", help="Profile name to use (folder name in /profiles)")
    parser.add_argument("--batch", type=int, help="Number of jobs to process in each batch")
    parser.add_argument("--ats", choices=["workday", "icims", "greenhouse", "lever", "auto", "manual"], 
                       help="ATS system to target (or 'auto' to detect)")
    parser.add_argument("--verbose", action="store_true", help="Enable verbose logging")
    parser.add_argument("--allow-senior", action="store_true", help="Don't skip senior job postings")
    parser.add_argument("--headless", action="store_true", help="Run browser in headless mode")
    return parser.parse_args()

def select_ats():
    console.print(Panel("Select ATS System", style="bold blue"))
    options = {
        "1": "Workday",
        "2": "iCIMS",
        "3": "Greenhouse",
        "4": "Auto-detect",
        "5": "Resume last session",
        "6": "Manual mode (prepare documents only)"
    }
    
    for key, value in options.items():
        console.print(f"[bold]{key}[/bold]: {value}")
    
    choice = Prompt.ask("Select option", choices=list(options.keys()))
    
    if choice == "1":
        return "workday"
    elif choice == "2":
        return "icims"
    elif choice == "3":
        return "greenhouse"
    elif choice == "4":
        return "auto"
    elif choice == "5":
        return "resume"
    else:
        return "manual"

def prompt_continue():
    return Confirm.ask("Continue to next batch?")

def is_senior_job(job_title):
    senior_keywords = ["senior", "lead", "principal", "director", "manager", "head"]
    lower_title = job_title.lower()
    return any(keyword in lower_title for keyword in senior_keywords)

def main():
    args = parse_args()
    
    # Set up console and display banner
    console.print(Panel.fit(
        "[bold blue]AutoJobAgent[/bold blue] - Automated Job Application Tool",
        border_style="blue"
    ))
    
    # Load profile
    try:
        profile_path = Path(f"profiles/{args.profile}")
        if not profile_path.exists():
            console.print(f"[bold red]Error:[/bold red] Profile '{args.profile}' not found")
            return 1
        
        profile = utils.load_profile(args.profile)
        console.print(f"[green]Loaded profile:[/green] {profile['name']}")
    except Exception as e:
        console.print(f"[bold red]Error loading profile:[/bold red] {e}")
        return 1
    
    # Ensure profile files exist (may generate PDFs if needed)
    try:
        utils.ensure_profile_files(profile)
    except Exception as e:
        console.print(f"[bold red]Error ensuring profile files:[/bold red] {e}")
        return 1
    
    # Select ATS if not provided in args
    ats_choice = args.ats if args.ats else select_ats()
    
    # If resuming, load the previous session
    if ats_choice == "resume":
        session = utils.load_session(profile)
        ats_choice = session.get("ats", "auto")
        console.print(f"[green]Resuming session with ATS:[/green] {ats_choice}")
    else:
        session = utils.load_session(profile, ats_choice)
    
    # Set batch size from args or profile default
    batch_size = args.batch if args.batch else profile.get("batch_default", 10)
    console.print(f"[green]Batch size:[/green] {batch_size}")
    
    # Start dashboard subprocess
    console.print("[green]Starting dashboard...[/green]")
    dashboard_proc = subprocess.Popen(
        ["uvicorn", "dashboard_api:app", "--port", "8000"],
        stdout=subprocess.DEVNULL if not args.verbose else None,
        stderr=subprocess.DEVNULL if not args.verbose else None
    )
    time.sleep(1)  # Give dashboard time to start
    
    # Launch Playwright
    console.print("[green]Launching browser...[/green]")
    with sync_playwright() as p:
        # Create persistent context to maintain cookies/storage
        user_data_dir = profile_path / "playwright"
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(user_data_dir),
            headless=args.headless
        )
        
        # Initialize job scraper
        console.print("[green]Initializing job scraper...[/green]")
        scraper = job_scraper.JobScraper(profile, start=session.get("next_index", 0))
        
        try:
            # Process jobs in batches
            for batch_num, batch in enumerate(scraper.batched(batch_size)):
                console.print(Panel(f"[bold]Processing Batch #{batch_num + 1}[/bold]", style="blue"))
                
                # Filter out jobs already processed
                batch = [job for job in batch if utils.hash_job(job) not in session.get("done", [])]
                if not batch:
                    console.print("[yellow]No new jobs to process in this batch[/yellow]")
                    break
                
                console.print(f"[green]Found {len(batch)} new jobs to process[/green]")
                
                # Process each job
                for job_index, job in enumerate(batch):
                    console.print(f"\n[bold]Job {job_index + 1}/{len(batch)}:[/bold] {job['title']} at {job['company']}")
                    
                    # Skip senior jobs unless allowed
                    if not args.allow_senior and is_senior_job(job["title"]):
                        console.print(f"[yellow]Skipping senior job:[/yellow] {job['title']}")
                        utils.append_log_row(job, profile, "Skipped (Senior)", "", "", "")
                        continue
                    
                    try:
                        # Generate/customize documents
                        console.print("[green]Customizing documents...[/green]")
                        pdf_cv, pdf_cl = document_generator.customize(job, profile)
                        
                        # Skip if in manual mode
                        if ats_choice == "manual":
                            console.print(f"[green]Documents prepared:[/green] {pdf_cv}, {pdf_cl}")
                            utils.append_log_row(job, profile, "Manual", pdf_cv, pdf_cl, "")
                            session.setdefault("done", []).append(utils.hash_job(job))
                            continue
                        
                        # Detect ATS if auto mode
                        detected_ats = detect(job["url"]) if ats_choice == "auto" else ats_choice
                        console.print(f"[green]Detected ATS:[/green] {detected_ats}")
                        
                        # Get submitter for the ATS
                        submitter = get_submitter(detected_ats, ctx)
                        
                        # Submit application
                        console.print("[green]Submitting application...[/green]")
                        status = submitter.submit(job, profile, pdf_cv, pdf_cl)
                        console.print(f"[bold]Status:[/bold] {status}")
                        
                        # Log result
                        utils.append_log_row(job, profile, status, pdf_cv, pdf_cl, detected_ats)
                        
                        # Mark job as done
                        session.setdefault("done", []).append(utils.hash_job(job))
                        
                    except Exception as e:
                        console.print(f"[bold red]Error processing job:[/bold red] {e}")
                        if args.verbose:
                            import traceback
                            traceback.print_exc()
                
                # Update session state
                session["next_index"] = scraper.current_index
                utils.save_session(profile, session)
                
                # Check for pause signal
                if utils.check_pause_signal():
                    console.print("[yellow]Pause signal detected[/yellow]")
                    break
                
                # Prompt to continue
                if not prompt_continue():
                    console.print("[yellow]User requested to stop[/yellow]")
                    break
            
            console.print(Panel("[bold green]Job processing complete![/bold green]", style="green"))
            
        except KeyboardInterrupt:
            console.print("\n[yellow]Interrupted by user[/yellow]")
        
        finally:
            # Clean up
            ctx.close()
            dashboard_proc.terminate()
            console.print("[green]Dashboard stopped[/green]")
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
