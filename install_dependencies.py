"""
install_dependencies.py
=======================
Comprehensive Dependency Installer & Environment Doctor for
Re-Inspection & Quality Intelligence Suite Visualizer.

This script ensures a completely fresh user can run the dashboard immediately by:
  1. Validating Python version (3.8+ required, 3.10-3.13 recommended)
  2. Upgrading and configuring pip
  3. Installing all required packages (pandas, openpyxl, xlrd, python-pptx, pillow, xlsxwriter, lxml)
  4. Testing GUI subsystem (Tkinter)
  5. Verifying project modules and required folders/templates
"""

import sys
import os
import subprocess
import importlib
import platform

# Ensure UTF-8 or safe console output on Windows
if hasattr(sys.stdout, 'reconfigure'):  
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass
if hasattr(sys.stderr, 'reconfigure'):
    try:
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# Comprehensive list of packages needed for all features
REQUIRED_PACKAGES = [
    ("pandas", "pandas>=2.0.0", "Data manipulation & analytics engine"),
    ("openpyxl", "openpyxl>=3.1.2", "Excel .xlsx reader & cell formatting"),
    ("xlrd", "xlrd>=2.0.1", "Legacy Excel .xls data reader"),
    ("xlsxwriter", "xlsxwriter>=3.1.0", "High-performance Excel export"),
    ("pptx", "python-pptx>=1.0.0", "PowerPoint report generation & automation"),
    ("PIL", "pillow>=10.0.0", "Image processing for dashboard charts"),
    ("lxml", "lxml>=4.9.0", "Fast XML parser for Office OpenXML"),
    ("numpy", "numpy>=1.24.0", "Numerical matrix computations"),
]

def print_header(title):
    print("\n" + "=" * 72)
    print(f"  {title}")
    print("=" * 72)

def check_python_environment():
    """Verify that Python version meets minimum requirements."""
    major, minor, micro = sys.version_info[:3]
    version_str = f"{major}.{minor}.{micro}"
    arch = platform.architecture()[0]
    
    print(f"  [*] Python Version : {version_str} ({arch})")
    print(f"  [*] Executable Path: {sys.executable}")
    print(f"  [*] Platform       : {platform.system()} {platform.release()}")
    
    if (major, minor) < (3, 8):
        print(f"\n  [!] CRITICAL: Python 3.8 or higher is required. Detected: {version_str}")
        print("      Please install a newer Python version from https://www.python.org/downloads/")
        return False
        
    print("  [OK] Python version is compatible.")
    return True

def upgrade_pip():
    """Ensure pip is available and updated."""
    print("\n  [*] Checking pip package manager...")
    try:
        # First ensure pip is installed/boostrapped
        subprocess.check_call(
            [sys.executable, "-m", "ensurepip", "--default-pip"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
    except Exception:
        pass

    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--upgrade", "pip", "--quiet"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        print("  [OK] pip is ready and up to date.")
    except Exception as e:
        print(f"  [i] pip upgrade skipped ({e}). Proceeding with current pip...")

def run_pip_install(args):
    """Execute a pip install command and return True if successful."""
    cmd = [sys.executable, "-m", "pip", "install"] + args
    try:
        res = subprocess.run(cmd, check=True)
        return res.returncode == 0
    except subprocess.CalledProcessError:
        return False

def install_packages():
    """Install all required packages with intelligent fallback strategies."""
    print("\n  [*] Installing required packages...")
    current_dir = os.path.dirname(os.path.abspath(__file__))
    req_file = os.path.join(current_dir, "requirements.txt")
    
    pkg_specs = [pkg[1] for pkg in REQUIRED_PACKAGES]
    
    # Strategy 1: Install from requirements.txt
    if os.path.exists(req_file):
        print("      Strategy 1: Installing from requirements.txt...")
        if run_pip_install(["-r", req_file]):
            print("  [OK] Successfully installed dependencies.")
            return True
            
        print("  [!] Direct install failed. Trying with --user flag...")
        # Strategy 2: User-level installation (for restricted environments)
        if run_pip_install(["--user", "-r", req_file]):
            print("  [OK] Successfully installed dependencies with --user.")
            return True

    # Strategy 3: Install individual package specs
    print("      Strategy 3: Installing packages directly...")
    if run_pip_install(pkg_specs):
        print("  [OK] Successfully installed dependencies.")
        return True

    # Strategy 4: Fallback with trusted hosts (corporate proxies / SSL inspection)
    print("      Strategy 4: Retrying with trusted host flags...")
    trusted_flags = [
        "--trusted-host", "pypi.org",
        "--trusted-host", "files.pythonhosted.org",
        "--trusted-host", "pypi.python.org",
        "--user"
    ] + pkg_specs
    
    if run_pip_install(trusted_flags):
        print("  [OK] Successfully installed dependencies using trusted hosts.")
        return True

    print("  [!] Automated installation failed. Please check network/proxy connection.")
    return False

def verify_all():
    """Verify package imports, GUI capability, and workspace assets."""
    print_header("Comprehensive Environment Verification")
    all_ok = True

    # 1. Verify Third-Party Libraries
    print("  --- 1. Python Packages ---")
    for import_name, pip_name, desc in REQUIRED_PACKAGES:
        try:
            mod = importlib.import_module(import_name)
            ver = getattr(mod, "__version__", "installed")
            print(f"  [OK]   {pip_name:<24} v{ver:<10} ({desc})")
        except ImportError:
            print(f"  [FAIL] {pip_name:<24} MISSING    ({desc})")
            all_ok = False

    # 2. Verify Desktop GUI Engine (Tkinter)
    print("\n  --- 2. Desktop GUI Subsystem ---")
    try:
        import tkinter
        # Test creating and destroying a headless root window
        root = tkinter.Tk()
        root.withdraw()
        root.update_idletasks()
        root.destroy()
        print(f"  [OK]   tkinter (GUI Engine)    Tk v{tkinter.TkVersion}   (Native GUI window system verified)")
    except Exception as e:
        print(f"  [FAIL] tkinter (GUI Engine)    ERROR: {e}")
        print("         Notice: On Windows, please ensure 'tcl/tk and IDLE' is enabled in Python installer.")
        all_ok = False

    # 3. Verify Local Project Modules
    current_dir = os.path.dirname(os.path.abspath(__file__))
    if current_dir not in sys.path:
        sys.path.insert(0, current_dir)

    print("\n  --- 3. Core Application Modules ---")
    for mod_name in ["reinspection_suite", "pptx_generator"]:
        try:
            importlib.import_module(mod_name)
            print(f"  [OK]   {mod_name + '.py':<24} Loaded successfully")
        except Exception as e:
            print(f"  [FAIL] {mod_name + '.py':<24} ERROR: {e}")
            all_ok = False

    # 4. Verify Project Folders and Templates
    print("\n  --- 4. Workspace Directories & Templates ---")
    essential_dirs = ["Re-Inspection", "FTT", "HFPA", "Output"]
    for d in essential_dirs:
        full_p = os.path.join(current_dir, d)
        if not os.path.exists(full_p):
            try:
                os.makedirs(full_p, exist_ok=True)
                print(f"  [OK]   Folder '{d}/' created (ready for data)")
            except Exception as e:
                print(f"  [WARN] Folder '{d}/' could not be created: {e}")
        else:
            files_count = len(os.listdir(full_p))
            print(f"  [OK]   Folder '{d}/' present ({files_count} items)")

    # Check template files
    templates = [
        ("Color_Template.xlsx", "Defect color mapping template"),
        ("RE-INS REPORT.APR.2026.pptx", "PowerPoint report baseline template"),
    ]
    for filename, desc in templates:
        p = os.path.join(current_dir, filename)
        if os.path.exists(p):
            print(f"  [OK]   Template '{filename}' found ({desc})")
        else:
            print(f"  [WARN] Template '{filename}' not found. (Place in folder if report export is required)")

    return all_ok

def main():
    print_header("Re-Inspection & Quality Intelligence Suite Setup")
    print("  Initializing environment and verifying all prerequisites...")
    
    if not check_python_environment():
        sys.exit(1)
        
    upgrade_pip()
    
    success = install_packages()
    if not success:
        print("\n  [!] Warning: Automated package installation reported issues.")
        print("      Checking installed components anyway...")
        
    verified = verify_all()
    
    print("\n" + "=" * 72)
    if verified:
        print("  SUCCESS: Ready to run! All requirements and components are ready.")
        print("  Launch the application using:")
        print("    -> Run_Dashboard.bat  (Double-click or run from terminal)")
        print("    -> python Run_Dashboard.py")
        print("=" * 72 + "\n")
        sys.exit(0)
    else:
        print("  WARNING: One or more components need attention.")
        print("  Please check the log messages above.")
        print("=" * 72 + "\n")
        sys.exit(1)

if __name__ == "__main__":
    main()
