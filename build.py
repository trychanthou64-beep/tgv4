import os
import sys
import subprocess
import shutil

def run_command(command):
    print(f"Running: {command}")
    result = subprocess.run(command, shell=True)
    if result.returncode != 0:
        print(f"Command failed with exit code {result.returncode}")
        sys.exit(result.returncode)

def main():
    print("=== Telegram Bot Builder ===")
    
    # Terminate running instances to release file locks on Windows
    print("Terminating any running instances of bot.exe, gui.exe, telegram_bot.exe, or setup.exe...")
    os.system("taskkill /f /im bot.exe >nul 2>&1")
    os.system("taskkill /f /im gui.exe >nul 2>&1")
    os.system("taskkill /f /im telegram_bot.exe >nul 2>&1")
    os.system("taskkill /f /im setup.exe >nul 2>&1")
    
    # Install pyinstaller if not present
    try:
        import PyInstaller
        print("PyInstaller is already installed.")
    except ImportError:
        print("PyInstaller not found. Installing...")
        run_command("pip install pyinstaller")
        
    # Build bot.py (Console app)
    # Add README.md as a data file so the frozen app can read it
    print("\nBuilding bot.py...")
    run_command('pyinstaller --onefile --icon "app_icon.ico" --add-data "README.md;." bot.py')
    
    # Build gui.py (Windowed/GUI app) as telegram_bot.exe embedding dist/bot.exe and README.md
    print("\nBuilding telegram_bot.exe (embedding bot.exe)...")
    run_command('pyinstaller --onefile --noconsole --name "telegram_bot" --icon "app_icon.ico" --add-data "dist/bot.exe;." --add-data "README.md;." --add-data "app_icon.png;." gui.py')
    
    # Clean up temporary bot.exe and old gui.exe
    print("\nCleaning up temporary and old executables...")
    for filename in ['bot.exe', 'gui.exe']:
        filepath = os.path.join('dist', filename)
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
                print(f"Deleted {filepath}")
            except Exception as e:
                print(f"Failed to delete {filepath}: {e}")
                
    # Copy telegram_bot.exe to root directory
    src_exe = os.path.join('dist', 'telegram_bot.exe')
    dest_exe = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'telegram_bot.exe')
    if os.path.exists(src_exe):
        print(f"\nCopying compiled executable to root folder...")
        try:
            shutil.copy2(src_exe, dest_exe)
            print(f"Successfully copied to: {dest_exe}")
        except Exception as e:
            print(f"Failed to copy executable to root folder: {e}")

    # Build setup.exe using Inno Setup
    print("\nBuilding setup.exe using Inno Setup...")
    possible_paths = [
        r"C:\Program Files\Inno Setup 7\ISCC.exe",
        r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        r"C:\Program Files\Inno Setup 6\ISCC.exe"
    ]
    iscc_path = None
    for path in possible_paths:
        if os.path.exists(path):
            iscc_path = path
            break
            
    if iscc_path:
        run_command(f'"{iscc_path}" setup.iss')
    else:
        print("Error: Inno Setup compiler (ISCC.exe) not found.")
        sys.exit(1)

    print("\n=== Build Complete ===")
    print("Files available in the root folder:")
    print("  - ./telegram_bot.exe (Application)")
    print("  - ./setup.exe        (Inno Setup Installer)")
    print("\nNote: You can run setup.exe on any PC to install the app.")
    
if __name__ == "__main__":
    main()
