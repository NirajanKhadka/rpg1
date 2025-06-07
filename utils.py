#!/usr/bin/env python3
"""
Utilities Module
---------------
Common utility functions for the Auto Job Application Assistant.
Handles logging, directory creation, configuration management,
path manipulation, and DOCX to PDF conversion on Windows.
"""

import os
import json
import logging
import re
from pathlib import Path
from typing import Dict, List, Optional, Any
from datetime import datetime

# Conditional import for comtypes, as it's Windows-specific
try:
    import comtypes.client
except ImportError:
    comtypes = None # Will be checked before use

logger = logging.getLogger(__name__)

def setup_logging(logs_dir_str: str, profile_name: str, log_level: int = logging.INFO) -> None:
    """
    Set up logging configuration. Logs to both console and a file.

    Args:
        logs_dir_str: Directory path (as string) to store log files.
        profile_name: Name of the user profile, used in the log filename.
        log_level: Logging level (e.g., logging.INFO, logging.DEBUG).
    """
    logs_path = Path(logs_dir_str)
    try:
        logs_path.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        print(f"Error creating logs directory {logs_path}: {e}. Logs may not be saved to file.")
        # Proceed with console logging

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file_name = f"{profile_name}_{timestamp}.log"
    log_file_path = logs_path / log_file_name

    # Remove any existing handlers from the root logger to avoid duplicate logs
    # if this function is called multiple times (e.g., in tests or re-runs).
    root_logger = logging.getLogger()
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)

    # Configure basic logging to console
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[logging.StreamHandler()]
    )

    # Add file handler only if directory was created or already exists
    if logs_path.exists() and logs_path.is_dir():
        try:
            file_handler = logging.FileHandler(log_file_path, encoding='utf-8')
            file_handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
            logging.getLogger().addHandler(file_handler) # Add to root logger
            logging.info(f"File logging enabled. Log file: {log_file_path}")
        except Exception as e:
            # Use print here as logging might not be fully set up for file if this fails
            print(f"Warning: Could not set up file logging to {log_file_path}: {e}")
            logging.error(f"Failed to set up file handler for logs at {log_file_path}: {e}")
            
    logging.info(f"Logging initialized for profile: {profile_name}. Log level: {logging.getLevelName(log_level)}")


def create_directories(directory_paths: List[str]) -> bool:
    """
    Create directories if they don't exist. Handles nested directories.

    Args:
        directory_paths: List of directory paths (as strings) to create.

    Returns:
        True if all directories were created or already existed, False otherwise.
    """
    all_successful = True
    for dir_path_str in directory_paths:
        path_obj = Path(dir_path_str)
        try:
            path_obj.mkdir(parents=True, exist_ok=True)
            logging.debug(f"Directory verified/created: {path_obj}")
        except OSError as e:
            logging.error(f"Failed to create directory {path_obj}: {e}")
            print(f"Error: Could not create directory {path_obj}. Please check permissions and path validity.")
            all_successful = False
        except Exception as e: 
            logging.error(f"An unexpected error occurred creating directory {path_obj}: {e}")
            print(f"Error: An unexpected error occurred creating directory {path_obj}.")
            all_successful = False
    return all_successful


def load_config(config_path_str: str) -> Optional[Dict[str, Any]]:
    """
    Load configuration from a JSON file.

    Args:
        config_path_str: Path (as string) to the configuration file.

    Returns:
        Configuration dictionary or None if file doesn't exist or is invalid.
    """
    config_file = Path(config_path_str)
    if not config_file.exists():
        logging.info(f"Configuration file not found: {config_file}")
        return None
    if not config_file.is_file():
        logging.error(f"Configuration path exists but is not a file: {config_file}")
        print(f"Error: Configuration path {config_file.name} is not a file.")
        return None

    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        logging.info(f"Configuration loaded from {config_file}")
        return config
    except json.JSONDecodeError as e:
        logging.error(f"Invalid JSON in configuration file {config_file}: {e}")
        print(f"Error: Configuration file {config_file.name} contains invalid JSON. Please check its format.")
        return None
    except OSError as e:
        logging.error(f"OS error loading configuration from {config_file}: {e}")
        print(f"Error: Could not read configuration file {config_file.name} due to an OS error.")
        return None
    except Exception as e:
        logging.error(f"Unexpected error loading configuration from {config_file}: {e}")
        print(f"Error: An unexpected error occurred while loading {config_file.name}.")
        return None


def save_config(config_data: Dict[str, Any], config_path_str: str) -> bool:
    """
    Save configuration to a JSON file.

    Args:
        config_data: Configuration dictionary to save.
        config_path_str: Path (as string) to save the configuration file.

    Returns:
        True if successful, False otherwise.
    """
    config_file = Path(config_path_str)
    try:
        config_file.parent.mkdir(parents=True, exist_ok=True)
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=4)
        logging.info(f"Configuration saved to {config_file}")
        return True
    except OSError as e:
        logging.error(f"OS error saving configuration to {config_file}: {e}")
        print(f"Error: Could not write configuration to {config_file.name} due to an OS error.")
        return False
    except TypeError as e:
        logging.error(f"Type error saving configuration to {config_file}: {e}. Check data types.")
        print(f"Error: Could not save configuration due to non-serializable data: {e}")
        return False
    except Exception as e:
        logging.error(f"Unexpected error saving configuration to {config_file}: {e}")
        print(f"Error: An unexpected error occurred while saving configuration to {config_file.name}.")
        return False


def convert_doc_to_pdf_windows(doc_path_str: str, pdf_path_str: str) -> bool:
    """
    Convert a DOCX file to PDF using Microsoft Word (Windows only).
    Returns True on success, False on failure.
    """
    if comtypes is None:
        print("❌ PDF conversion skipped: comtypes library not available (likely not on Windows or not installed).")
        logging.error("comtypes library not available for PDF conversion.")
        return False

    doc_path_abs = str(Path(doc_path_str).resolve())
    pdf_path_abs = str(Path(pdf_path_str).resolve())
    
    # Ensure output directory for PDF exists
    Path(pdf_path_abs).parent.mkdir(parents=True, exist_ok=True)

    word = None  # Initialize word to None
    try:
        logging.info(f"Attempting to convert '{doc_path_abs}' to '{pdf_path_abs}' using MS Word.")
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False  # Run Word in the background
        
        # Check if the document is already open by another process (simple check)
        # This is a very basic check and might not be foolproof.
        # A more robust solution might involve checking for lock files or more advanced COM error handling.
        try:
            # Try to open with read-only to see if it's locked
            with open(doc_path_abs, 'rb') as f_test:
                pass
        except IOError: # Typically FileNotFoundError or PermissionError
            logging.warning(f"File {doc_path_abs} might be in use or inaccessible. Conversion might fail.")

        doc = word.Documents.Open(doc_path_abs)
        # wdFormatPDF = 17
        doc.SaveAs(pdf_path_abs, FileFormat=17)
        doc.Close(False) # False means don't save changes to original .docx
        logging.info(f"Successfully converted '{doc_path_str}' to '{pdf_path_str}'.")
        return True
    except Exception as e:
        logging.error(f"PDF conversion failed for '{doc_path_str}': {e}", exc_info=True)
        print(f"❌ PDF conversion failed for {Path(doc_path_str).name}: {e}")
        return False
    finally:
        if word:
            try:
                word.Quit()
            except Exception as e_quit:
                logging.error(f"Error quitting Word application: {e_quit}")
        # Ensure comtypes is uninitialized to release resources
        comtypes.CoUninitialize()


def ensure_file_exists(file_path_str: str, purpose: str = "file") -> bool:
    """
    Check if a file exists and is actually a file. Logs and prints an error if not.
    """
    path_obj = Path(file_path_str)
    if not path_obj.exists():
        msg = f"Required {purpose} not found: {path_obj}"
        logging.error(msg)
        print(f"Error: {msg}")
        return False
    if not path_obj.is_file():
        msg = f"Path for {purpose} exists but is not a file: {path_obj}"
        logging.error(msg)
        print(f"Error: {msg}")
        return False
    return True


def sanitize_filename(filename: str, max_length: int = 200) -> str:
    """
    Sanitizes a filename to remove/replace invalid characters and control length,
    making it safe for most file systems (including Windows).
    """
    if not isinstance(filename, str):
        filename = str(filename)

    base, ext = os.path.splitext(filename)
    sanitized_base = "".join(char for char in base if ord(char) >= 32)
    sanitized_base = re.sub(r'[<>:"/\\|?*]', '_', sanitized_base)
    sanitized_base = re.sub(r'[\s_]+', '_', sanitized_base).strip('_')
    
    if len(sanitized_base) > max_length:
        sanitized_base = sanitized_base[:max_length]
    
    sanitized_ext = ""
    if ext:
        sanitized_ext = "." + ext.lstrip('.').strip()
        sanitized_ext = re.sub(r'[<>:"/\\|?*]', '_', sanitized_ext)
        
    final_filename = sanitized_base + sanitized_ext
    
    if not final_filename.strip('.'):
        final_filename = f"sanitized_file_{datetime.now().strftime('%Y%m%d%H%M%S')}{sanitized_ext or '.dat'}"
        logging.warning(f"Original filename '{filename}' sanitized to a generic name '{final_filename}' due to extensive invalid characters.")

    return final_filename

# Example usage (typically not run when imported as a module)
if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
    
    print("--- Testing Utility Functions ---")

    test_root = Path("./temp_utils_test_root")
    test_root.mkdir(exist_ok=True)

    # Test directory creation
    print("\nTesting directory creation:")
    test_dirs = [str(test_root / "level1/level2"), str(test_root / "another_dir")]
    if create_directories(test_dirs):
        print(f"Successfully created/verified test directories: {test_dirs}")
        for d_str in test_dirs:
            d_path = Path(d_str)
            assert d_path.exists() and d_path.is_dir(), f"Directory {d_str} was not created or is not a directory."
    else:
        print("Directory creation failed for one or more paths.")

    # Test config load/save
    print("\nTesting config load/save:")
    test_config_path = str(test_root / "test_config.json")
    test_data = {"setting1": "value1", "nested": {"num": 123, "bool": True}}
    
    if save_config(test_data, test_config_path):
        print(f"Config saved to {test_config_path}")
        loaded_data = load_config(test_config_path)
        if loaded_data == test_data:
            print(f"Config loaded successfully and matches saved data: {loaded_data}")
        else:
            print(f"Error: Loaded config data does not match saved data. Loaded: {loaded_data}")
    else:
        print(f"Failed to save config to {test_config_path}")

    # Test ensure_file_exists
    print("\nTesting ensure_file_exists:")
    print(f"Checking existing config file: {ensure_file_exists(test_config_path, 'test config')}")
    print(f"Checking non-existent file: {ensure_file_exists(str(test_root / 'non_existent.txt'), 'dummy file')}")
    print(f"Checking a directory (should fail as not a file): {ensure_file_exists(str(test_root), 'directory as file')}")

    # Test sanitize_filename
    print("\nTesting sanitize_filename:")
    filenames_to_test = [
        "My Document: Final Version?.docx", "file/with/slashes.txt",
        " leading_and_trailing_spaces.pdf ", "file*with|invalid<chars>.md",
        "CON.txt", "very_long_filename_" + "a"*200 + ".zip",
        "file_with_你好世界_unicode.txt", "../../../../etc/passwd", "image..png", ""
    ]
    for fn in filenames_to_test:
        sanitized = sanitize_filename(fn)
        print(f"Original: '{fn}' -> Sanitized: '{sanitized}'")

    # Test DOCX to PDF conversion (WINDOWS ONLY, REQUIRES MS WORD)
    if os.name == 'nt' and comtypes is not None:
        print("\nTesting DOCX to PDF conversion (Windows with MS Word only):")
        sample_docx_path = test_root / "sample_for_pdf.docx"
        sample_pdf_path = test_root / "sample_converted.pdf"
        
        # Create a dummy docx file
        try:
            from docx import Document as DocxDocument
            doc = DocxDocument()
            doc.add_paragraph("This is a test document for PDF conversion.")
            doc.save(str(sample_docx_path))
            print(f"Created dummy DOCX: {sample_docx_path}")

            if convert_doc_to_pdf_windows(str(sample_docx_path), str(sample_pdf_path)):
                print(f"Successfully converted to PDF: {sample_pdf_path}")
                assert sample_pdf_path.exists(), "PDF file was not created."
            else:
                print("PDF conversion failed. Check if MS Word is installed and accessible.")
        except ImportError:
            print("Skipping DOCX creation for PDF test: python-docx not found (should be in requirements).")
        except Exception as e:
            print(f"Error during PDF conversion test setup or execution: {e}")
    else:
        print("\nSkipping DOCX to PDF conversion test (not on Windows or comtypes not available).")

    # Clean up (optional)
    # import shutil
    # if test_root.exists():
    #     shutil.rmtree(test_root)
    #     print(f"\nCleaned up temporary test directory: {test_root}")
    
    print("\n--- Utility Functions Test Complete ---")
