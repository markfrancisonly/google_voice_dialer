# Google Voice Dialer
#
# Description:
#   Registers a tel: and callto: URI handler to initiate calls using Google Voice.
#   Supports protocol registration, unregistration, building to executable,
#   dynamic Chrome app detection, and dialing via Google Voice.
#
# Usage: Install python and pip, then run:
#
#   py google_voice_dialer.py --install
#
#       Build the script into a standalone exe and install it as a TEL link
#       default app protocol handler in your AppData folder.
#
#
#   py google_voice_dialer.py --uninstall
#
#       Remove the executable and unregister the handler.
#

import sys
import webbrowser
import re
import urllib.parse
import winreg
import os
import argparse
import datetime
import subprocess
import shutil

try:
    import win32com.client as com_client
except ImportError:
    com_client = None

try:
    import win32api
    import win32gui
    import win32ui
    import win32con
except ImportError:
    pass    

PROG_ID = "Google Voice Dialer"
PROG_NAME = f"{PROG_ID}"
PROG_DESC = "Google Voice Dialer for TEL: Links"
LOG_FILENAME = f"{PROG_ID}.log"

def get_base_dir():
    """Get the base directory of the running script or executable."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def find_google_voice_shortcut():
    """Find the Google Voice shortcut path by searching in Start Menu Programs and subdirectories."""
    if com_client is None:
        return None

    try:
        base_dir = os.path.expandvars(r'%APPDATA%\Microsoft\Windows\Start Menu\Programs')
        if not os.path.exists(base_dir):
            return None

        for root, dirs, files in os.walk(base_dir):
            for file in files:
                if file == 'Google Voice.lnk':
                    return os.path.join(root, file)
        return None
    except Exception as e:
        return None

def get_google_voice_app_id():
    """Dynamically find the app_id for Google Voice PWA from the shortcut."""
    if com_client is None:
        print("pywin32 not installed. Cannot dynamically lookup app_id. Install with 'pip install pywin32'.")
        return 

    lnk_path = find_google_voice_shortcut()
    if not lnk_path:
        print("No Google Voice chrome app shortcut found.")
        return 

    try:
        shell = com_client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(lnk_path)
        if not shortcut.Arguments:
            print(f"No arguments found in {os.path.basename(lnk_path)}.")
            return
        match = re.search(r'--app-id=([a-z]{32})', shortcut.Arguments)
        if match:
            return match.group(1)
    except Exception as e:
        print(f"Error finding Google Voice chrome app_id: {e}.")
        return 

def get_chrome_proxy_path():
    """Dynamically find the path to chrome_proxy.exe via registry."""
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe")
        chrome_exe_path = winreg.QueryValueEx(key, None)[0]
        winreg.CloseKey(key)
        chrome_proxy_path = chrome_exe_path.replace('chrome.exe', 'chrome_proxy.exe')
        if os.path.exists(chrome_proxy_path):
            return chrome_proxy_path
        else:
            raise FileNotFoundError("chrome_proxy.exe not found")
    except Exception as e:
        print(f"Error finding Chrome path: {e}")
        return None

def get_google_voice_icon_location():
    """Dynamically find the icon location for Google Voice PWA from the shortcut."""
    if com_client is None:
        return None

    lnk_path = find_google_voice_shortcut()
    if not lnk_path:
        return None

    try:
        shell = com_client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(lnk_path)
        icon_location = shortcut.IconLocation
        if icon_location:
            return icon_location
        else:
            return None
    except Exception as e:
        return None
        
def register_handler(exe_path=None):
    """Register the handler for tel: protocol with capabilities for Windows 11."""
    try:
        # Determine the running file path (script or executable)
        if exe_path:
            running_path = exe_path
            runner = ''
        elif getattr(sys, 'frozen', False):
            running_path = sys.executable
            runner = ''
        else:
            running_path = os.path.abspath(__file__)
            python_exe = sys.executable
            # Use pythonw.exe to run without a console window
            runner = python_exe.replace('python.exe', 'pythonw.exe') if 'python.exe' in python_exe else python_exe
            runner = f'"{runner}" '
        
        # Use HKCU for user-specific registration
        base_key = winreg.HKEY_CURRENT_USER
        
        # Create ProgId key under HKCR (which merges to HKCU\Software\Classes)
        # But for consistency, we'll use HKCR for protocol and ProgId
        prog_id_key = winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, PROG_ID, 0, winreg.KEY_WRITE)
        winreg.SetValueEx(prog_id_key, None, 0, winreg.REG_SZ, f"URL:{PROG_NAME}")
        winreg.SetValueEx(prog_id_key, "URL Protocol", 0, winreg.REG_SZ, "")
        
        # Optional: DefaultIcon (point to running_path for icon)
        icon_key = winreg.CreateKeyEx(prog_id_key, "DefaultIcon", 0, winreg.KEY_WRITE)
        winreg.SetValueEx(icon_key, None, 0, winreg.REG_SZ, f"{running_path},0")
        winreg.CloseKey(icon_key)
        
        # Create shell\open\command under ProgId
        command_key = winreg.CreateKeyEx(prog_id_key, r"shell\open\command", 0, winreg.KEY_WRITE)
        command_value = f'{runner}"{running_path}" "%1"'
        winreg.SetValueEx(command_key, None, 0, winreg.REG_SZ, command_value)
        winreg.CloseKey(command_key)
        winreg.CloseKey(prog_id_key)
        
        # Register capabilities under HKCU\Software
        cap_key = winreg.CreateKeyEx(base_key, rf"Software\{PROG_ID}\Capabilities", 0, winreg.KEY_WRITE)
        winreg.SetValueEx(cap_key, "ApplicationName", 0, winreg.REG_SZ, PROG_NAME)
        winreg.SetValueEx(cap_key, "ApplicationDescription", 0, winreg.REG_SZ, PROG_DESC)
        # Optional: ApplicationIcon
        winreg.SetValueEx(cap_key, "ApplicationIcon", 0, winreg.REG_SZ, f"{running_path},0")
        
        # URLAssociations
        url_assoc_key = winreg.CreateKeyEx(cap_key, "URLAssociations", 0, winreg.KEY_WRITE)
        winreg.SetValueEx(url_assoc_key, "tel", 0, winreg.REG_SZ, PROG_ID)
        winreg.SetValueEx(url_assoc_key, "callto", 0, winreg.REG_SZ, PROG_ID)  
        winreg.CloseKey(url_assoc_key)
        winreg.CloseKey(cap_key)
        
        # Register the app in RegisteredApplications
        reg_apps_key = winreg.CreateKeyEx(base_key, r"Software\RegisteredApplications", 0, winreg.KEY_WRITE)
        winreg.SetValueEx(reg_apps_key, PROG_NAME, 0, winreg.REG_SZ, rf"Software\{PROG_ID}\Capabilities")
        winreg.CloseKey(reg_apps_key)
        
        print(f"Successfully registered '{PROG_ID}' tel: protocol handler.")
        print("Go to Settings > Apps > Default apps > Set defaults for applications, and search 'Google Voice Dialer'. Set it as default for 'TEL' links.")
    except PermissionError:
        print("Permission denied: Run the script as Administrator if needed (though HKCU should not require it).")
    except Exception as e:
        print(f"Error registering: {e}")

def unregister_handler(prog_id=PROG_ID, prog_name=PROG_NAME):
    """Unregister the tel: protocol handler and capabilities."""
    try:
        def delete_key_recursive(root, path):
            try:
                key = winreg.OpenKey(root, path, 0, winreg.KEY_ALL_ACCESS)
                try:
                    while True:
                        subkey = winreg.EnumKey(key, 0)
                        delete_key_recursive(key, subkey)
                except OSError:
                    pass
                winreg.CloseKey(key)
                winreg.DeleteKey(root, path)
            except FileNotFoundError:
                pass
        
        # Delete ProgId from HKCR
        delete_key_recursive(winreg.HKEY_CLASSES_ROOT, prog_id)
        
        # Delete capabilities from HKCU\Software
        delete_key_recursive(winreg.HKEY_CURRENT_USER, rf"Software\{prog_id}")
        
        # Remove from RegisteredApplications
        reg_apps_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\RegisteredApplications", 0, winreg.KEY_ALL_ACCESS)
        try:
            winreg.DeleteValue(reg_apps_key, prog_name)
        except FileNotFoundError:
            pass
        winreg.CloseKey(reg_apps_key)
        
        print(f"Successfully unregistered '{prog_id}' tel: protocol handler.")
    except PermissionError:
        print("Permission denied: Run the script as Administrator to unregister.")
    except Exception as e:
        print(f"Error unregistering: {e}")

def dial(phone_url):
    """Handle dialing via Google Voice."""
    if not (phone_url.lower().startswith('tel:') or phone_url.lower().startswith('callto:')):
        return

    # Extract and clean phone number (remove non-digits except +)
    phone = re.sub(r'^(tel|callto):', '', phone_url, flags=re.IGNORECASE)
    phone = re.sub(r'[^0-9+]', '', phone)
    phone = phone.strip()
    
    # Log the dial attempt
    base_dir = get_base_dir()
    log_path = os.path.join(base_dir, LOG_FILENAME)
    with open(log_path, 'a') as log_file:
        log_file.write(f"{datetime.datetime.now()} - {phone}\n")

    # Encode for URL
    encoded_phone = urllib.parse.quote(phone)

    # Google Voice dialing URL
    gv_url = f"https://voice.google.com/u/0/calls?a=nc,{encoded_phone}"

    # Dynamically get Chrome proxy path and app_id
    chrome_proxy_path = get_chrome_proxy_path()
    app_id = get_google_voice_app_id()

    if chrome_proxy_path and app_id:
        subprocess.run([chrome_proxy_path, f'--app-id={app_id}', f'--app-launch-url-for-shortcuts-menu-item={gv_url}'], check=False)
    else:
        webbrowser.open(gv_url)

def extract_and_save_icon(icon_file, icon_index, output_ico):
    large, small = win32gui.ExtractIconEx(icon_file, icon_index)
    if large:
        hicon = large[0]
    elif small:
        hicon = small[0]
    else:
        raise Exception("No icon extracted")

    ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
    ico_y = win32api.GetSystemMetrics(win32con.SM_CYICON)
    hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
    hbmp = win32ui.CreateBitmap()
    hbmp.CreateCompatibleBitmap(hdc, ico_x, ico_y)
    hdc = hdc.CreateCompatibleDC()
    hdc.SelectObject(hbmp)
    hdc.DrawIcon((0, 0), hicon)
    bmpinfo = hbmp.GetInfo()
    bmpstr = hbmp.GetBitmapBits(True)
    from PIL import Image
    img = Image.frombuffer('RGBA', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), bmpstr, 'raw', 'BGRA', 0, 1)
    img.save(output_ico, 'ICO')
    win32gui.DestroyIcon(hicon)

def install_executable():
    try:
        # Common parts
        appdata_dir = os.path.expandvars(rf'%APPDATA%\{PROG_ID}')
        os.makedirs(appdata_dir, exist_ok=True)
        target_exe = os.path.join(appdata_dir, PROG_ID + '.exe')

        if getattr(sys, 'frozen', False):
            # Running as executable, copy self to appdata
            source_exe = sys.executable
            if os.path.abspath(source_exe) != os.path.abspath(target_exe):
                shutil.copy(source_exe, target_exe)
            else:
                print("Already installed in AppData.")
        else:
            # Not frozen, build it
            # Install pywin32 if not installed
            global com_client
            if com_client is None:
                print("Installing pywin32...")
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pywin32'])
                import win32com.client as com_client

                # Re-import win32 modules after pywin32 installation
                try:
                    global win32api, win32gui, win32ui, win32con
                    import win32api
                    import win32gui
                    import win32ui
                    import win32con
                except ImportError as e:
                    raise ImportError(f"Failed to import win32 modules after installing pywin32: {e}")

            # Install pyinstaller if not installed
            try:
                subprocess.check_call(['pyinstaller', '--version'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except (subprocess.CalledProcessError, FileNotFoundError):
                print("Installing pyinstaller...")
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])

            # Install pillow if not installed
            try:
                from PIL import Image
            except ImportError:
                print("Installing pillow...")
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pillow'])
                from PIL import Image

            # Get Google Voice icon location
            icon_location = get_google_voice_icon_location()
            icon_arg = []
            icon_ico = 'google_voice.ico'
            if icon_location:
                try:
                    parts = icon_location.rsplit(',', 1)
                    if len(parts) == 2:
                        icon_file = parts[0].strip()
                        icon_index_str = parts[1].strip()
                        icon_index = int(icon_index_str)
                    else:
                        icon_file = icon_location.strip()
                        icon_index = 0
                    
                    extract_and_save_icon(icon_file, icon_index, icon_ico)
                    icon_arg = ['--icon', icon_ico]
                    print("Extracted Google Voice icon for executable.")
                except Exception as e:
                    print(f"Failed to extract icon: {e}. Building without custom icon.")

            # Build the executable
            script_path = os.path.abspath(__file__)
            build_args = ['pyinstaller', '--onefile', '--noconsole', '--name', PROG_ID] + icon_arg + [script_path]
            subprocess.check_call(build_args)
            print("Executable built successfully. Check the dist folder.")

            # Copy to AppData
            dist_exe = os.path.join(os.path.dirname(script_path), 'dist', PROG_ID + '.exe')
            shutil.copy(dist_exe, target_exe)

            # Clean up icon file if created
            if os.path.exists(icon_ico):
                os.remove(icon_ico)

        # Now, register the copied exe
        if os.path.exists(target_exe):
            register_handler(exe_path=target_exe)
            # Open Windows Settings > Apps > Default apps
            subprocess.call(['start', 'ms-settings:defaultapps'], shell=True)
        else:
            print("Executable not found at expected path.")

    except Exception as e:
        print(f"Error installing executable: {e}")

def uninstall():
    unregister_handler()
    appdata_dir = os.path.expandvars(rf'%APPDATA%\{PROG_ID}')
    if os.path.exists(appdata_dir):
        shutil.rmtree(appdata_dir)
    print("Uninstalled successfully.")

def main():
    parser = argparse.ArgumentParser(description="Google Voice Dialer")
    parser.add_argument('--install', action='store_true', help="build and/or install application")
    parser.add_argument('--uninstall', action='store_true', help="uninstall application")
    parser.add_argument('--register', action='store_true', help="register handler for TEL links")
    parser.add_argument('--unregister', action='store_true', help="unregister handler for TEL links")
    parser.add_argument('url', nargs='?', help="tel: URL to dial")

    args = parser.parse_args()

    if args.install:
        install_executable()
    elif args.uninstall:
        uninstall()
    elif args.register:
        register_handler()
    elif args.unregister:
        unregister_handler()
    elif args.url:
        dial(args.url)
    else:
        help_text = parser.format_help()
        if getattr(sys, 'frozen', False):
            win32api.MessageBox(0, help_text, "Google Voice Dialer Usage", win32con.MB_OK | win32con.MB_ICONINFORMATION)
        else:
            print(help_text)
        sys.exit(0)

if __name__ == "__main__":
    main()