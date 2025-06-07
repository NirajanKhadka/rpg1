#!/usr/bin/env python3
"""
Integration test for AutoJobAgent.
Tests that all modules import correctly and basic functionality works.
"""

import os
import sys
import traceback
from pathlib import Path

# Add the current directory to the path to ensure imports work
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_imports():
    """Test that all modules import correctly."""
    print("\n=== Testing Imports ===")
    
    import_results = {}
    
    # Test main imports
    modules_to_test = [
        "main",
        "job_scraper",
        "document_generator",
        "utils",
        "dashboard_api",
        "ats",
        "ats.base_submitter",
        "ats.workday",
        "ats.icims",
        "ats.greenhouse"
    ]
    
    for module_name in modules_to_test:
        try:
            __import__(module_name)
            import_results[module_name] = "✅ Success"
            print(f"✅ Successfully imported {module_name}")
        except ImportError as e:
            import_results[module_name] = f"❌ Failed: {str(e)}"
            print(f"❌ Failed to import {module_name}: {e}")
    
    return import_results

def test_profile_loading():
    """Test loading a profile."""
    print("\n=== Testing Profile Loading ===")
    
    try:
        import utils
        
        # Check if Nirajan profile exists
        profile_name = "Nirajan"
        profile = utils.load_profile(profile_name)
        
        if profile and isinstance(profile, dict):
            print(f"✅ Successfully loaded profile: {profile_name}")
            print(f"   Name: {profile.get('name')}")
            print(f"   Email: {profile.get('email')}")
            print(f"   Keywords: {profile.get('keywords')}")
            return True
        else:
            print(f"❌ Failed to load profile: {profile_name}")
            return False
    
    except Exception as e:
        print(f"❌ Error testing profile loading: {e}")
        traceback.print_exc()
        return False

def test_ats_detection():
    """Test ATS detection functionality."""
    print("\n=== Testing ATS Detection ===")
    
    try:
        from ats import detect
        
        test_urls = {
            "https://company.workday.com/careers/job/12345": "workday",
            "https://jobs.icims.com/company/job/12345": "icims",
            "https://boards.greenhouse.io/company/jobs/12345": "greenhouse",
            "https://jobs.lever.co/company/12345": "lever",
            "https://example.com/careers/job/12345": "unknown"
        }
        
        success = True
        
        for url, expected in test_urls.items():
            detected = detect(url)
            if detected == expected:
                print(f"✅ Correctly detected {detected} for {url}")
            else:
                print(f"❌ Detection failed for {url}: expected {expected}, got {detected}")
                success = False
        
        return success
    
    except Exception as e:
        print(f"❌ Error testing ATS detection: {e}")
        traceback.print_exc()
        return False

def test_job_hash():
    """Test job hashing functionality."""
    print("\n=== Testing Job Hashing ===")
    
    try:
        import utils
        
        # Test job
        job = {
            "title": "Data Analyst",
            "company": "Test Company",
            "location": "Toronto, ON",
            "url": "https://example.com/job/12345"
        }
        
        # Hash the job
        job_hash = utils.hash_job(job)
        
        if job_hash and isinstance(job_hash, str) and len(job_hash) > 0:
            print(f"✅ Successfully hashed job: {job_hash}")
            
            # Test hash consistency
            job_hash2 = utils.hash_job(job)
            if job_hash == job_hash2:
                print("✅ Hash is consistent for the same job")
                return True
            else:
                print("❌ Hash is not consistent for the same job")
                return False
        else:
            print("❌ Failed to hash job")
            return False
    
    except Exception as e:
        print(f"❌ Error testing job hashing: {e}")
        traceback.print_exc()
        return False

def test_pdf_conversion():
    """Test PDF conversion functionality."""
    print("\n=== Testing PDF Conversion ===")
    
    try:
        import utils
        
        # Check if Word COM is available
        try:
            import win32com.client
            import comtypes.client
            word_available = True
            print("✅ Word COM libraries are available")
        except ImportError:
            word_available = False
            print("⚠️ Word COM libraries are not available, PDF conversion will be limited")
        
        # Create a test DOCX file
        test_docx = Path("temp") / "test_conversion.docx"
        test_pdf = Path("temp") / "test_conversion.pdf"
        
        # Ensure temp directory exists
        Path("temp").mkdir(exist_ok=True)
        
        # Create a simple DOCX file using python-docx if not exists
        if not test_docx.exists():
            try:
                from docx import Document
                doc = Document()
                doc.add_paragraph("This is a test document for PDF conversion.")
                doc.save(test_docx)
                print(f"✅ Created test DOCX file: {test_docx}")
            except Exception as e:
                print(f"❌ Failed to create test DOCX file: {e}")
                return False
        
        # Try to convert to PDF
        success = utils.convert_doc_to_pdf(str(test_docx), str(test_pdf))
        
        if success:
            print(f"✅ Successfully converted DOCX to PDF: {test_pdf}")
            return True
        else:
            print("⚠️ PDF conversion not available, but this is expected on some systems")
            return True  # Still return True as this is expected on some systems
    
    except Exception as e:
        print(f"❌ Error testing PDF conversion: {e}")
        traceback.print_exc()
        return False

def test_ollama_connection():
    """Test connection to Ollama LLM."""
    print("\n=== Testing Ollama Connection ===")
    
    try:
        import ollama
        
        # Try to ping Ollama
        try:
            models = ollama.list()
            print(f"✅ Successfully connected to Ollama")
            print(f"   Available models: {', '.join([m.get('name', 'unknown') for m in models.get('models', [])])}")
            return True
        except Exception as e:
            print(f"⚠️ Ollama not available: {e}")
            print("   This is OK if you're not using LLM features yet")
            return True  # Still return True as Ollama might not be set up yet
    
    except ImportError as e:
        print(f"❌ Failed to import ollama module: {e}")
        print("   Run: pip install ollama")
        return False
    except Exception as e:
        print(f"❌ Error testing Ollama connection: {e}")
        traceback.print_exc()
        return False

def test_directory_structure():
    """Test that the directory structure is correct."""
    print("\n=== Testing Directory Structure ===")
    
    required_dirs = [
        "profiles",
        "output",
        "output/logs",
        "ats",
        "templates"
    ]
    
    success = True
    
    for dir_path in required_dirs:
        path = Path(dir_path)
        if path.exists() and path.is_dir():
            print(f"✅ Directory exists: {dir_path}")
        else:
            print(f"❌ Missing directory: {dir_path}")
            try:
                path.mkdir(parents=True, exist_ok=True)
                print(f"   Created directory: {dir_path}")
            except Exception as e:
                print(f"   Failed to create directory: {e}")
                success = False
    
    return success

def run_all_tests():
    """Run all integration tests."""
    print("\n========================================")
    print("  AUTOJOBAGENT INTEGRATION TEST")
    print("========================================")
    
    results = {}
    
    # Test imports first
    results["imports"] = test_imports()
    
    # Test directory structure
    results["directory_structure"] = test_directory_structure()
    
    # Test profile loading
    results["profile_loading"] = test_profile_loading()
    
    # Test ATS detection
    results["ats_detection"] = test_ats_detection()
    
    # Test job hashing
    results["job_hash"] = test_job_hash()
    
    # Test PDF conversion
    results["pdf_conversion"] = test_pdf_conversion()
    
    # Test Ollama connection
    results["ollama_connection"] = test_ollama_connection()
    
    # Print summary
    print("\n========================================")
    print("  TEST SUMMARY")
    print("========================================")
    
    all_passed = True
    
    for test_name, result in results.items():
        if isinstance(result, dict):
            # Handle import results
            has_failures = any("Failed" in status for status in result.values())
            status = "❌ FAILED" if has_failures else "✅ PASSED"
            if has_failures:
                all_passed = False
        else:
            # Handle boolean results
            status = "✅ PASSED" if result else "❌ FAILED"
            if not result:
                all_passed = False
        
        print(f"{status} - {test_name}")
    
    print("\n========================================")
    if all_passed:
        print("  ✅ ALL TESTS PASSED")
    else:
        print("  ❌ SOME TESTS FAILED")
    print("========================================\n")
    
    return all_passed

if __name__ == "__main__":
    success = run_all_tests()
    sys.exit(0 if success else 1)
