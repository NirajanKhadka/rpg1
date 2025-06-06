#!/usr/bin/env python3
"""
AutoJobAgent - Main Orchestrator
---------------------------------
Handles CLI, profile management, ATS selection, job scraping,
Ollama-powered document customization, DOCX-to-PDF conversion,
Playwright-based application submission, Excel logging,
FastAPI dashboard interaction, and session management.
"""

import os
import sys
import json
import logging
import argparse
import asyncio
import multiprocessing
import time
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Any, Tuple, Set
import signal
import webbrowser

try:
    from rich.console import Console
    from rich.prompt import Prompt, Confirm, IntPrompt
    from rich.panel import Panel
    from rich.text import Text
    from rich.table import Table
    RICH_AVAILABLE = True
except ImportError:
    RICH_AVAILABLE = False
    # Fallback for input if rich is not available
    class Prompt:
        @staticmethod
        def ask(prompt_text, choices=None, default=None):
            if choices:
                prompt_text += f" ({'/'.join(choices)})"
            if default:
                prompt_text += f" [{default}]"
            return input(f"{prompt_text}: ")

    class Confirm:
        @staticmethod
        def ask(prompt_text, default=False):
            response = input(f"{prompt_text} (y/n) [{'y' if default else 'n'}]: ").strip().lower()
            if not response:
                return default
            return response == 'y'

    class IntPrompt:
        @staticmethod
        def ask(prompt_text, default=None):
            while True:
                try:
                    val_str = input(f"{prompt_text} [{default}]: ").strip()
                    if not val_str and default is not None:
                        return default
                    return int(val_str)
                except ValueError:
                    print("Please enter a valid number.")
    
    class Console:
        def print(self, *args, **kwargs):
            print(*args)

# Project-specific imports
from job_scraper import scrape_jobs as scrape_jobs_from_source
from document_generator import DocumentGenerator
from ats import get_ats_submitter # Factory from ats/__init__.py
import utils
from dashboard_api import start_dashboard

# --- Constants ---
CONFIG_FILE_NAME = "config.json"
SESSION_FILE_NAME_TEMPLATE = "session.json" # Stored in profile directory
DEFAULT_APPLICATIONS_LOG_FILE_NAME = "applications.xlsx"
CURRENT_JOB_STATUS_FILE_NAME = "current_job_live.json"
DEFAULT_BATCH_SIZE = 10
SUPPORTED_SCRAPING_SITES = ["indeed", "eluta"]
# Order matters for the menu
ATS_CHOICES_MAP = {
    1: ("Workday", "workday"),
    2: ("iCIMS", "icims"),
    3: ("Greenhouse", "greenhouse"),
    4: ("Lever", "lever"), # Added Lever as per prompt
    5: ("Manual Job Logging Only", "manual"),
    6: ("Resume Last Paused Session", "resume"),
    7: ("Exit", "exit")
}

# --- Global variable for session state (used by signal handler) ---
current_session_state: Optional[Dict[str, Any]] = None
interrupted = False # Flag for Ctrl+C
console = Console() if RICH_AVAILABLE else Console() # Fallback console

# --- Signal Handler for Ctrl+C ---
def signal_handler(sig, frame):
    global interrupted
    interrupted = True
    console.print("\n[bold yellow]Ctrl+C detected! Attempting to save session and exit gracefully...[/bold yellow]")
    # Actual saving is handled in the main loop's finally block or specific checkpoints

signal.signal(signal.SIGINT, signal_handler)

# --- Argument Parsing ---
def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="AutoJobAgent: Smart Local Job Application Assistant")
    parser.add_argument("profile_folder_name", help="Name of the profile folder in 'profiles/' (e.g., Nirajan).")
    parser.add_argument("--config", default=CONFIG_FILE_NAME, help=f"Path to config file (default: {CONFIG_FILE_NAME}).")
    parser.add_argument("--verbose", "-v", action="store_true", help="Enable verbose DEBUG logging.")
    return parser.parse_args()

# --- Profile and Configuration Management ---
def load_profile(profile_folder_path: Path, profile_folder_name: str) -> Optional[Dict[str, Any]]:
    profile_json_path = profile_folder_path / f"{profile_folder_name}.json"
    if not utils.ensure_file_exists(str(profile_json_path), f"Profile JSON for '{profile_folder_name}'"):
        return None

    try:
        with open(profile_json_path, "r", encoding='utf-8') as f:
            profile_data = json.load(f)
    except json.JSONDecodeError as e:
        logging.error(f"Invalid JSON in profile file {profile_json_path}: {e}")
        console.print(f"[bold red]Error: Profile file '{profile_json_path.name}' contains invalid JSON.[/bold red]")
        return None
    except Exception as e:
        logging.error(f"Error loading profile {profile_json_path}: {e}", exc_info=True)
        console.print(f"[bold red]Error: Could not load profile '{profile_json_path.name}'.[/bold red]")
        return None

    required_fields = ["name", "location", "keywords", "skills", "experience_level", 
                       "resume_docx", "cover_letter_docx", "email"] # PDF fields are now optional at load
    missing_fields = [field for field in required_fields if field not in profile_data]
    if missing_fields:
        logging.error(f"Profile {profile_folder_name}.json is missing required fields: {', '.join(missing_fields)}")
        console.print(f"[bold red]Error: Profile '{profile_folder_name}.json' is missing: {', '.join(missing_fields)}.[/bold red]")
        return None
    
    profile_data["_profile_folder_path"] = str(profile_folder_path) # Store path for convenience
    profile_data["_profile_name"] = profile_folder_name # Store name for convenience
    
    # Ensure base DOCX files exist
    base_resume_docx = profile_folder_path / profile_data["resume_docx"]
    base_cl_docx = profile_folder_path / profile_data["cover_letter_docx"]
    if not utils.ensure_file_exists(str(base_resume_docx), "Base resume DOCX"): return None
    if not utils.ensure_file_exists(str(base_cl_docx), "Base cover letter DOCX"): return None

    # Check and convert base PDF files if specified or missing
    for doc_type in ["resume", "cover_letter"]:
        docx_field = f"{doc_type}_docx"
        pdf_field = f"{doc_type}_pdf"
        
        docx_filename = profile_data.get(docx_field)
        pdf_filename = profile_data.get(pdf_field)

        if not docx_filename: # Should not happen due to required_fields check, but good practice
            logging.error(f"Profile missing '{docx_field}'. Cannot proceed with PDF check for {doc_type}.")
            continue

        docx_path = profile_folder_path / docx_filename
        
        if pdf_filename:
            pdf_path = profile_folder_path / pdf_filename
            if not pdf_path.exists():
                console.print(f"[yellow]Base {doc_type} PDF ('{pdf_filename}') not found. Attempting to convert from DOCX.[/yellow]")
                if utils.convert_doc_to_pdf_windows(str(docx_path), str(pdf_path)):
                    console.print(f"[green]Successfully converted and saved base {doc_type} PDF: {pdf_path}[/green]")
                else:
                    console.print(f"[bold red]Failed to convert base {doc_type} DOCX to PDF. You may need to do this manually or ensure MS Word is correctly set up.[/bold red]")
                    # Continue without base PDF, application-specific PDFs will still be attempted
        else:
            # PDF field is missing, try to generate a default PDF name and convert
            default_pdf_filename = Path(docx_filename).stem + ".pdf"
            pdf_path = profile_folder_path / default_pdf_filename
            console.print(f"[yellow]Profile JSON missing '{pdf_field}'. Attempting to convert '{docx_filename}' to '{default_pdf_filename}'.[/yellow]")
            if utils.convert_doc_to_pdf_windows(str(docx_path), str(pdf_path)):
                console.print(f"[green]Successfully converted and saved base {doc_type} PDF: {pdf_path}[/green]")
                profile_data[pdf_field] = default_pdf_filename # Update profile data in memory
                # Consider prompting user to save this change back to the JSON file
            else:
                console.print(f"[bold red]Failed to convert base {doc_type} DOCX to PDF. You may need to do this manually.[/bold red]")

    logging.info(f"Profile '{profile_folder_name}' loaded and base documents checked.")
    return profile_data

# --- Session Management ---
def get_session_file_path(profile_folder_path: Path) -> Path:
    return profile_folder_path / SESSION_FILE_NAME_TEMPLATE

def save_session_state(session_data: Dict[str, Any], profile_folder_path: Path):
    global current_session_state
    current_session_state = session_data
    session_file = get_session_file_path(profile_folder_path)
    try:
        with open(session_file, 'w', encoding='utf-8') as f:
            json.dump(session_data, f, indent=4)
        logging.info(f"Session state saved to {session_file}.")
    except Exception as e:
        logging.error(f"Failed to save session state to {session_file}: {e}")
        console.print(f"[yellow]Warning: Could not save session state to {session_file}.[/yellow]")

def load_session_state(profile_folder_path: Path) -> Optional[Dict[str, Any]]:
    session_file = get_session_file_path(profile_folder_path)
    if not session_file.exists():
        logging.info(f"Session file {session_file} not found. Cannot resume.")
        return None
    try:
        with open(session_file, 'r', encoding='utf-8') as f:
            session_data = json.load(f)
        logging.info(f"Session state loaded from {session_file}.")
        return session_data
    except Exception as e:
        logging.error(f"Failed to load session state from {session_file}: {e}")
        console.print(f"[yellow]Error: Could not load session state from {session_file}. Starting fresh if possible.[/yellow]")
        return None

# --- Application Log Management --- (Simplified, assuming utils.py has robust versions)
def get_applications_log_path(config: Dict[str, Any], logs_dir: Path) -> Path:
    log_filename = config.get("applications_log_filename", DEFAULT_APPLICATIONS_LOG_FILE_NAME)
    return logs_dir / log_filename

def load_applications_log(log_path: Path) -> pd.DataFrame:
    log_columns = ["Profile", "Job Title", "Company", "Location", "Job Link", "ATS", 
                   "Resume Used", "Cover Letter Used", "Timestamp", "Status", "Notes",
                   "Keywords Found", "Keywords Missing", "Job ID", "Source"]
    if log_path.exists():
        try:
            df = pd.read_excel(log_path)
            for col in log_columns:
                if col not in df.columns:
                    df[col] = "" # Add missing columns with empty string
            logging.info(f"Applications log loaded from {log_path}.")
            return df
        except Exception as e:
            logging.error(f"Error loading applications log {log_path}: {e}. A new log will be created.")
    logging.info(f"Applications log not found at {log_path}. Creating a new one.")
    return pd.DataFrame(columns=log_columns)

def save_applications_log(df: pd.DataFrame, log_path: Path) -> None:
    try:
        df.to_excel(log_path, index=False)
        logging.info(f"Applications log saved to {log_path}.")
    except Exception as e:
        logging.error(f"Error saving applications log to {log_path}: {e}")
        console.print(f"[yellow]Warning: Could not save applications log to {log_path}.[/yellow]")

def get_processed_job_identifiers(df_log: pd.DataFrame, profile_name_filter: Optional[str] = None) -> Set[Tuple[str, str]]:
    if df_log.empty or "Profile" not in df_log.columns or "Job ID" not in df_log.columns:
        return set()
    temp_df = df_log.copy()
    if profile_name_filter:
        temp_df = temp_df[temp_df["Profile"] == profile_name_filter]
    # Ensure 'Job ID' is string, handle NaN by filling with a placeholder then converting
    return set(tuple(x) for x in temp_df[["Profile", "Job ID"]].fillna("N/A_ID").astype(str).values)


# --- User Interaction ---
def display_profile_summary(profile_data: Dict[str, Any], profile_folder_path: Path):
    console.print(Panel(f"[bold cyan]Profile Loaded: {profile_data['name']}[/bold cyan]\n"
                        f"Location: {profile_data.get('location', 'N/A')}\n"
                        f"Keywords: {', '.join(profile_data.get('keywords', []))[:100]}...\n"
                        f"Resume: {profile_folder_path / profile_data.get('resume_docx', 'N/A')}\n"
                        f"Cover Letter: {profile_folder_path / profile_data.get('cover_letter_docx', 'N/A')}",
                        title="Profile Summary", border_style="blue"))

# --- Dashboard Management ---
def start_dashboard_process_handler(port: int) -> Optional[multiprocessing.Process]:
    try:
        dashboard_process = multiprocessing.Process(target=start_dashboard, args=("0.0.0.0", port), daemon=True)
        dashboard_process.start()
        logging.info(f"Dashboard process starting on http://localhost:{port} (PID: {dashboard_process.pid}).")
        time.sleep(3)
        if not dashboard_process.is_alive():
            logging.error("Dashboard process failed to start or terminated prematurely.")
            console.print("[bold red]Error: Could not start the dashboard. Check dashboard_api.py logs.[/bold red]")
            return None
        webbrowser.open_new_tab(f"http://localhost:{port}")
        console.print(f"[green]Dashboard should be opening at http://localhost:{port}[/green]")
        return dashboard_process
    except Exception as e:
        logging.error(f"Failed to start dashboard process: {e}", exc_info=True)
        console.print(f"[bold red]Error starting dashboard: {e}[/bold red]")
        return None

def update_live_dashboard_status(status_file_path: Path, profile_name: str, job_title: Optional[str], company: Optional[str], status: str,
                                 url: Optional[str], keywords_found: List[str], keywords_missing: List[str],
                                 ats_target: Optional[str] = None, progress_percent: Optional[int] = None) -> None:
    live_status = {
        "profile": profile_name, "job_title": job_title, "company": company, "status": status,
        "url": url, "keywords_found": keywords_found or [], "keywords_missing": keywords_missing or [],
        "ats_target": ats_target, "progress_percent": progress_percent,
        "timestamp": datetime.now().isoformat()
    }
    try:
        with open(status_file_path, 'w', encoding='utf-8') as f: json.dump(live_status, f, indent=2)
        logging.debug(f"Live dashboard status updated: {status} for {job_title or 'N/A'}")
    except Exception as e:
        logging.warning(f"Could not write live dashboard status to {status_file_path}: {e}")

def clear_live_dashboard_status(status_file_path: Path) -> None:
    idle_status = {
        "profile": None, "job_title": None, "company": None, "status": "Idle",
        "url": None, "keywords_found": [], "keywords_missing": [], "ats_target": None, "progress_percent": None,
        "timestamp": datetime.now().isoformat()
    }
    try:
        with open(status_file_path, 'w', encoding='utf-8') as f: json.dump(idle_status, f, indent=2)
        logging.info("Live dashboard status cleared to Idle.")
    except Exception as e:
        logging.warning(f"Could not clear live dashboard status at {status_file_path}: {e}")

# --- Main Application Logic ---
async def main_application_loop(session_data: Dict[str, Any], config: Dict[str, Any], args: argparse.Namespace):
    global interrupted # Allow modification of the global interrupted flag
    profile_folder_path = Path(session_data["profile_folder_path"])
    profile_data = session_data["profile_data"]
    
    output_dir = Path(config.get("output_dir", "output"))
    logs_dir = output_dir / config.get("logs_dir_name", "logs")
    user_documents_dir = output_dir / profile_data["_profile_name"] / "documents"
    utils.create_directories([str(user_documents_dir)]) # Ensure it exists

    applications_log_path = get_applications_log_path(config, logs_dir)
    applications_df = load_applications_log(applications_log_path)
    
    # Combine overall processed jobs with session-specific processed jobs
    processed_job_identifiers = get_processed_job_identifiers(applications_df, profile_data["_profile_name"])
    session_processed_tuples = set(tuple(item) for item in session_data.get("processed_job_identifiers_session", []))
    processed_job_identifiers.update(session_processed_tuples)

    doc_generator = DocumentGenerator(
        profile_data=profile_data,
        profiles_base_dir=str(profile_folder_path),
        output_dir_for_profile_docs=str(user_documents_dir), # Customized docs go here
        ollama_model=profile_data.get("ollama_model", "mistral")
    )

    live_status_file = logs_dir / CURRENT_JOB_STATUS_FILE_NAME

    continue_processing_batches = True
    while continue_processing_batches and not interrupted:
        if session_data["current_job_index_in_batch"] >= len(session_data["jobs_in_current_batch_list"]):
            if session_data["job_site_scraped"] == "manual_entry":
                console.print("[yellow]Manual logging mode. No more jobs in current 'batch'. Choose 'Exit' or start a new session for scraping.[/yellow]")
                break 
            
            console.print(f"\n--- Scraping new batch of jobs from [cyan]{session_data['job_site_scraped']}[/cyan] ---")
            update_live_dashboard_status(live_status_file, profile_data["_profile_name"], None, None, f"Scraping {session_data['job_site_scraped']}...", None, [], [], session_data["target_ats_for_application"])
            
            scraped_jobs: List[Dict[str, Any]] = await scrape_jobs_from_source(
                job_site=session_data["job_site_scraped"],
                keywords=profile_data.get("keywords", []),
                location=profile_data.get("location", ""),
                experience_level=profile_data.get("experience_level", "Entry Level"),
                batch_size=session_data["current_batch_size"],
                last_job_id=session_data.get("last_job_id_processed_on_site") 
            )
            if not scraped_jobs:
                console.print("[yellow]No new jobs found in this scrape attempt.[/yellow]")
                if not (ask_user_choice("Try scraping again?", ["Yes", "No"]) == 1):
                    continue_processing_batches = False; break
                else: continue
            session_data["jobs_in_current_batch_list"] = scraped_jobs
            session_data["current_job_index_in_batch"] = 0
            save_session_state(session_data, get_session_file_path(profile_folder_path))

        batch_jobs = session_data["jobs_in_current_batch_list"]
        job_idx = session_data["current_job_index_in_batch"]
        
        while job_idx < len(batch_jobs) and not interrupted:
            job_details = batch_jobs[job_idx]
            job_id_scraped = job_details.get("job_id", f"{job_details.get('source', 'unknown')}_{hash(job_details.get('url', job_details.get('title')))}")
            job_details["job_id"] = job_id_scraped
            
            update_live_dashboard_status(live_status_file, profile_data["_profile_name"], 
                                         job_details.get('title'), job_details.get('company'), "Evaluating",
                                         job_details.get('url'), job_details.get('keywords_found', []), job_details.get('keywords_missing', []),
                                         session_data["target_ats_for_application"], int((job_idx + 1) / len(batch_jobs) * 100))

            if (profile_data["_profile_name"], str(job_id_scraped)) in processed_job_identifiers:
                logging.info(f"Job ID {job_id_scraped} for profile {profile_data['_profile_name']} already processed. Skipping.")
                console.print(f"\n[yellow]Skipping (already processed):[/yellow] {job_details.get('title')} at {job_details.get('company')}")
                session_data["last_job_id_processed_on_site"] = job_id_scraped
                job_idx += 1
                session_data["current_job_index_in_batch"] = job_idx
                save_session_state(session_data, get_session_file_path(profile_folder_path))
                continue

            console.print(Panel(
                f"[bold blue]Processing Job {job_idx + 1} of {len(batch_jobs)} in current batch:[/bold blue]\n"
                f"  Title: [bold]{job_details.get('title', 'N/A')}[/bold]\n"
                f"  Company: {job_details.get('company', 'N/A')}\n"
                f"  Location: {job_details.get('location', 'N/A')}\n"
                f"  URL: [link={job_details.get('url','')}]{job_details.get('url', 'N/A')}[/link]\n"
                f"  Scraped from: {job_details.get('source', session_data['job_site_scraped'])}",
                title="Current Job", border_style="green"
            ))
            
            detected_ats = utils.detect_ats_from_url(job_details.get('url', ''))
            console.print(f"  Detected ATS from URL: [cyan]{detected_ats if detected_ats else 'Unknown'}[/cyan]")
            
            final_ats_for_submission = session_data["target_ats_for_application"]
            if final_ats_for_submission == "manual":
                console.print("[yellow]Manual logging mode selected. Preparing documents only.[/yellow]")
            elif detected_ats and detected_ats != session_data["target_ats_for_application"]:
                ats_mismatch_choice = ask_user_choice(
                    f"Detected ATS '{detected_ats}' differs from your target '{session_data['target_ats_for_application']}'. How to proceed?",
                    [f"Attempt application via detected ATS: {detected_ats}", 
                     f"Stick to target ATS: {session_data['target_ats_for_application']}",
                     "Log Manually for this job",
                     "Skip this job"]
                )
                if ats_mismatch_choice == 1: final_ats_for_submission = detected_ats
                elif ats_mismatch_choice == 3: final_ats_for_submission = "manual"
                elif ats_mismatch_choice == 4 or ats_mismatch_choice is None:
                    application_status, application_notes = "Skipped (ATS Mismatch)", f"User skipped due to ATS mismatch (detected {detected_ats}, target {session_data['target_ats_for_application']})"
                    # Log skip and continue (logging consolidated below)
                    job_idx += 1; session_data["current_job_index_in_batch"] = job_idx
                    session_data["last_job_id_processed_on_site"] = job_id_scraped
                    session_data.setdefault("processed_job_identifiers_session", []).append((args.profile_folder_name, str(job_id_scraped)))
                    save_session_state(session_data, get_session_file_path(profile_folder_path))
                    new_log_entry_data = {"Profile": args.profile_folder_name, "Job Title": job_details.get('title'), "Company": job_details.get('company'), "Location": job_details.get('location'), "Job Link": job_details.get('url'), "ATS": final_ats_for_submission, "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Status": application_status, "Notes": application_notes, "Job ID": str(job_id_scraped), "Source": job_details.get('source', session_data['job_site_scraped'])}
                    applications_df = pd.concat([applications_df, pd.DataFrame([new_log_entry_data])], ignore_index=True)
                    save_applications_log(applications_df, applications_log_path)
                    continue
            
            console.print(f"  Targeting ATS for application: [bold magenta]{final_ats_for_submission}[/bold magenta]")

            application_status = "Pending"
            application_notes = ""
            final_resume_path_str = "N/A"
            final_cover_letter_path_str = "N/A"

            apply_choice = ask_user_choice(f"Apply to '{job_details.get('title')}'?", ["Yes", "No, skip this job"])
            if apply_choice == 1:
                update_live_dashboard_status(live_status_file, args.profile_folder_name, job_details.get('title'), job_details.get('company'), "Preparing Documents...", job_details.get('url'), job_details.get('keywords_found', []), job_details.get('keywords_missing', []), final_ats_for_submission)
                try:
                    job_data_for_docs = job_details.copy()
                    # Keywords found by scraper are already in job_details if scraper supports it
                    
                    console.print("[cyan]Generating customized resume...[/cyan]")
                    custom_resume_docx_path = doc_generator.generate_resume(job_data_for_docs) # This now returns DOCX path
                    console.print(f"Customized resume DOCX: {custom_resume_docx_path}")
                    
                    console.print("[cyan]Generating customized cover letter...[/cyan]")
                    custom_cl_docx_path = doc_generator.generate_cover_letter(job_data_for_docs) # This now returns DOCX path
                    console.print(f"Customized cover letter DOCX: {custom_cl_docx_path}")

                    # Convert to PDF
                    resume_to_upload_path = custom_resume_docx_path
                    cl_to_upload_path = custom_cl_docx_path
                    
                    if final_ats_for_submission != "manual": # PDF conversion mainly for automated uploads
                        console.print("[cyan]Converting documents to PDF...[/cyan]")
                        # Resume PDF
                        resume_pdf_name = f"{Path(custom_resume_docx_path).stem}.pdf"
                        resume_pdf_path = user_documents_dir / resume_pdf_name
                        if utils.convert_doc_to_pdf_windows(str(custom_resume_docx_path), str(resume_pdf_path)):
                            console.print(f"[green]Resume PDF generated: {resume_pdf_path}[/green]")
                            resume_to_upload_path = resume_pdf_path
                        else:
                            pdf_fail_choice = ask_user_choice(
                                "Resume PDF conversion failed. How to proceed?",
                                ["Retry conversion", "Upload DOCX instead", "Skip this job"]
                            )
                            if pdf_fail_choice == 1: # Retry (could implement retry logic)
                                console.print("[yellow]Retrying PDF conversion for resume...[/yellow]")
                                if utils.convert_doc_to_pdf_windows(str(custom_resume_docx_path), str(resume_pdf_path)):
                                    resume_to_upload_path = resume_pdf_path
                                else:
                                    console.print("[red]Retry failed. Will use DOCX for resume.[/red]")
                            elif pdf_fail_choice == 3:
                                application_status, application_notes = "Skipped (PDF Fail)", "User skipped due to resume PDF conversion failure."
                                # Break from this job processing, go to logging
                                raise Exception("UserSkippedJob") # Special exception to break and log
                            # else (choice 2 or failed retry): resume_to_upload_path remains DOCX
                        
                        # Cover Letter PDF
                        cl_pdf_name = f"{Path(custom_cl_docx_path).stem}.pdf"
                        cl_pdf_path = user_documents_dir / cl_pdf_name
                        if utils.convert_doc_to_pdf_windows(str(custom_cl_docx_path), str(cl_pdf_path)):
                            console.print(f"[green]Cover Letter PDF generated: {cl_pdf_path}[/green]")
                            cl_to_upload_path = cl_pdf_path
                        else:
                            pdf_fail_choice_cl = ask_user_choice(
                                "Cover Letter PDF conversion failed. How to proceed?",
                                ["Retry conversion", "Upload DOCX instead", "Skip this job"]
                            )
                            if pdf_fail_choice_cl == 1:
                                console.print("[yellow]Retrying PDF conversion for cover letter...[/yellow]")
                                if utils.convert_doc_to_pdf_windows(str(custom_cl_docx_path), str(cl_pdf_path)):
                                    cl_to_upload_path = cl_pdf_path
                                else:
                                    console.print("[red]Retry failed. Will use DOCX for cover letter.[/red]")
                            elif pdf_fail_choice_cl == 3:
                                application_status, application_notes = "Skipped (PDF Fail)", "User skipped due to cover letter PDF conversion failure."
                                raise Exception("UserSkippedJob")
                            # else (choice 2 or failed retry): cl_to_upload_path remains DOCX
                    
                    final_resume_path_str = str(resume_to_upload_path.resolve())
                    final_cover_letter_path_str = str(cl_to_upload_path.resolve())

                    if final_ats_for_submission != "manual":
                        update_live_dashboard_status(live_status_file, args.profile_folder_name, job_details.get('title'), job_details.get('company'), f"Submitting via {final_ats_for_submission}...", job_details.get('url'), job_details.get('keywords_found', []), job_details.get('keywords_missing', []), final_ats_for_submission)
                        
                        submitter = ApplicationSubmitter.get_ats_submitter(
                            ats_name=final_ats_for_submission,
                            profile_data=profile_data,
                            output_dir=str(user_output_dir),
                            browser_type=profile_data.get("browser", "chromium")
                        )
                        if submitter:
                            application_status, application_notes = await submitter.submit_application(
                                job_details, resume_to_upload_path, cl_to_upload_path
                            )
                        else:
                            application_status, application_notes = "Failed", f"No submitter available for ATS: {final_ats_for_submission}. Please apply manually."
                            console.print(f"[red]No submitter for {final_ats_for_submission}. Please apply manually.[/red]")
                    else: # Manual logging mode for this job
                        application_status, application_notes = "Pending (Manual Log)", "Job logged for manual application. Documents generated."
                        console.print(f"[green]Documents generated for manual application to {job_details.get('title')}.[/green]")
                        console.print(f"  Resume: {final_resume_path_str}")
                        console.print(f"  Cover Letter: {final_cover_letter_path_str}")
                        console.print(f"  Apply at: {job_details.get('url')}")

                except Exception as e:
                    if str(e) == "UserSkippedJob": # Custom exception to handle skipping from PDF failure
                        pass # Status already set
                    else:
                        logging.error(f"Error during document generation or submission for {job_details.get('title')}: {e}", exc_info=True)
                        application_status, application_notes = "Failed", f"Error: {str(e)[:150]}"
                    console.print(f"[bold red]Error processing application: {application_notes}[/bold red]")
                else: # Skipped by user
                    application_status, application_notes = "Skipped (User)", "User chose to skip."
                
                console.print(f"Application Status: [bold { 'green' if 'Applied' in application_status else ('yellow' if 'Skipped' in application_status or 'Pending' in application_status else 'red') }]{application_status}[/bold { 'green' if 'Applied' in application_status else ('yellow' if 'Skipped' in application_status or 'Pending' in application_status else 'red') }] - {application_notes}")

                # Log to Excel
                new_log_entry_data = {
                    "Profile": args.profile_folder_name, "Job Title": job_details.get('title'), 
                    "Company": job_details.get('company'), "Location": job_details.get('location'), 
                    "Job Link": job_details.get('url'), "ATS": final_ats_for_submission, 
                    "Resume Used": final_resume_path_str, "Cover Letter Used": final_cover_letter_path_str,
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Status": application_status,
                    "Notes": application_notes,
                    "Keywords Found": ", ".join(job_details.get('keywords_found', [])),
                    "Keywords Missing": ", ".join(job_details.get('keywords_missing', [])),
                    "Job ID": str(job_id_scraped), 
                    "Source": job_details.get('source', session_data['job_site_scraped'])
                }
                log_columns = ["Profile", "Job Title", "Company", "Location", "Job Link", "ATS", 
                               "Resume Used", "Cover Letter Used", "Timestamp", "Status", "Notes",
                               "Keywords Found", "Keywords Missing", "Job ID", "Source"]
                for col in log_columns:
                    if col not in new_log_entry_data: new_log_entry_data[col] = ""
                
                applications_df = pd.concat([applications_df, pd.DataFrame([new_log_entry_data])], ignore_index=True)
                save_applications_log(applications_df, applications_log_path)
                
                session_data.setdefault("processed_job_identifiers_session", []).append((args.profile_folder_name, str(job_id_scraped)))
                processed_job_identifiers.add((args.profile_folder_name, str(job_id_scraped))) 
                
                session_data["last_job_id_processed_on_site"] = job_id_scraped
                set_overall_last_processed_job_id(config, args.profile_folder_name, session_data["job_site_scraped"], job_id_scraped, args.config)
                
                job_idx += 1
                session_data["current_job_index_in_batch"] = job_idx
                save_session_state(session_data, get_session_file_path(profile_folder_path))

                if interrupted:
                    console.print("[yellow]Processing interrupted. Will save session and exit after this job.[/yellow]")
                    break # Break from processing jobs in batch
            
            if interrupted:
                continue_processing_batches = False # Ensure outer loop also exits
                break

            # End of current batch processing
            console.print(f"\n--- Batch of {len(batch_jobs)} jobs processed (or attempted). ---")
            post_batch_choice = ask_user_choice(
                "Batch complete! What next?",
                ["Continue to next batch", "Pause and resume later", "Exit application"]
            )
            if post_batch_choice == 2: # Pause
                console.print("[yellow]Pausing session. Run again and choose 'Resume Last Session' to continue.[/yellow]")
                save_session_state(session_data, get_session_file_path(profile_folder_path))
                continue_processing_batches = False
            elif post_batch_choice == 3 or post_batch_choice is None: # Exit
                console.print("[yellow]Exiting application.[/yellow]")
                continue_processing_batches = False
            # If 1 (Continue), loop will try to fetch new batch if current one is exhausted.
            # If current batch not exhausted (e.g. due to an interruption mid-batch that wasn't a full stop), it will continue.

        logging.info("Job application session finished.")
        console.print("\n[bold green]Job application session finished.[/bold green]")

    except FileNotFoundError as e:
        logging.error(f"Critical file not found: {e}", exc_info=True)
        console.print(f"[bold red]Error: A critical file was not found: {e}[/bold red]")
    except KeyboardInterrupt: # Should be caught by signal handler, but as a fallback
        logging.info("KeyboardInterrupt caught in main. Session state should have been saved by handler.")
        console.print("\n[yellow]Exiting due to KeyboardInterrupt.[/yellow]")
    except Exception as e:
        logging.error(f"An unexpected error occurred in main: {e}", exc_info=True)
        console.print(f"[bold red]An unexpected error occurred: {e}[/bold red]")
        if session_data:
            save_session_state(session_data, get_session_file_path(profile_folder_path))
            console.print(f"[yellow]Attempted to save session state to {get_session_file_path(profile_folder_path)} due to error.[/yellow]")
    finally:
        logging.info("Cleaning up and exiting.")
        clear_live_dashboard_status(live_status_file)
        if session_data and not continue_processing_batches and post_batch_choice != 2 and not interrupted : # If exited cleanly, not paused
            session_file_to_clean = get_session_file_path(profile_folder_path)
            if session_file_to_clean.exists():
                try: 
                    session_file_to_clean.unlink()
                    logging.info(f"Cleaned up session file {session_file_to_clean} on normal exit.")
                except OSError as e_clean:
                    logging.warning(f"Could not delete session file {session_file_to_clean}: {e_clean}")
        elif interrupted and current_session_state: # Ensure save on interrupt if not already saved by loop end
             save_session_state(current_session_state, get_session_file_path(profile_folder_path))


        if dashboard_process and dashboard_process.is_alive():
            logging.info("Terminating dashboard process...")
            dashboard_process.terminate()
            dashboard_process.join(timeout=5)
            if dashboard_process.is_alive():
                logging.warning("Dashboard process did not terminate gracefully. Forcing kill.")
                dashboard_process.kill()
            logging.info("Dashboard process stopped.")
        console.print("[bold blue]Application assistant has shut down.[/bold blue]")

if __name__ == "__main__":
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    multiprocessing.freeze_support() 
    
    args = parse_arguments()
    config = utils.load_config(args.config) # Load config early for paths
    if not config:
        config = {
            "output_dir": "output", "profiles_dir": "profiles", "logs_dir_name": "logs",
            "applications_log_filename": DEFAULT_APPLICATIONS_LOG_FILE_NAME,
            "last_processed_jobs_overall": {}
        }
    
    # Initialize session_data here to pass to main_application_loop
    # This part will be refined by the menu logic inside main_application_loop
    # For now, this is a placeholder to allow the structure to work.
    # The actual session_data initialization or loading happens *inside* main_application_loop
    # based on user choices.
    
    # We need profile_folder_path for session file path, so load profile minimally first
    profiles_base_dir = Path(config.get("profiles_dir", "profiles"))
    profile_folder_path = profiles_base_dir / args.profile_folder_name
    
    # Dummy session_data for now; it will be properly initialized or loaded in the loop
    session_data_for_main_loop = {
        "profile_folder_path": str(profile_folder_path),
        "profile_data": {}, # Will be loaded inside the loop
        # Other fields will be set based on user input or loaded session
    }

    asyncio.run(main_application_loop(session_data_for_main_loop, config, args))
