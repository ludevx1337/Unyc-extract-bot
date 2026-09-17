"""
Setup script for UNYC automation
Installs required dependencies and Playwright browsers
"""

import subprocess
import sys

def run_command(command):
    """Run a command and print the output"""
    print(f"Running: {command}")
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")
        print(f"Output: {e.stdout}")
        print(f"Error: {e.stderr}")
        return False

def main():
    """Main setup function"""
    print("Setting up UNYC automation environment...")
    
    # Install Python dependencies
    print("\n1. Installing Python dependencies...")
    if not run_command("pip install -r requirements.txt"):
        print("Failed to install dependencies")
        return False
    
    # Install Playwright browsers
    print("\n2. Installing Playwright browsers...")
    if not run_command("playwright install chromium"):
        print("Failed to install Playwright browsers")
        return False
    
    print("\n✅ Setup completed successfully!")
    print("\nYou can now run the automation script with:")
    print("python unyc_automation.py")

if __name__ == "__main__":
    main()