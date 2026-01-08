# Google Voice Dialer
#
# Description:
#   Registers a tel: and callto: URI handler to initiate calls using Google Voice.
#   Supports protocol registration, unregistration, building to executable,
#   dynamic Chrome app detection, and dialing via Google Voice.
#
# Usage: Run the script or executable with:
#
#   python google_voice_dialer.py --install  
#
#       Copies the script to your AppData folder and registers it as a TEL link
#       default app protocol handler.
#
#
#   python google_voice_dialer.py --uninstall 
#
#       Removes the installed files and unregisters the handler.
#
# Requirements: Python 3, pywin32 (for Windows-specific features)

import argparse
import datetime
import os
import re
import shutil
import subprocess
import sys
import urllib.parse
import webbrowser
import winreg

try:
    import win32com.client as com_client
except ImportError:
    com_client = None

try:
    import win32api
    import win32con
except ImportError:
    pass

PROG_ID = "Google Voice Dialer"
PROG_NAME = f"{PROG_ID}"
PROG_DESC = "Google Voice tel: protocol handler. Dial phone numbers using Google Voice."
LOG_FILENAME = f"{PROG_ID}.log"
VERSION = "1.1.1"


def find_google_voice_shortcut():
    """Find the Google Voice shortcut path by searching in Start Menu Programs and subdirectories."""
    if com_client is None:
        return None

    try:
        base_dir = os.path.expandvars(
            r"%APPDATA%\Microsoft\Windows\Start Menu\Programs"
        )
        if not os.path.exists(base_dir):
            return None

        for root, dirs, files in os.walk(base_dir):
            for file in files:
                if file == "Google Voice.lnk":
                    return os.path.join(root, file)
        return None
    except Exception as e:
        return None


def get_google_voice_app_id():
    """Dynamically find the app_id for Google Voice PWA from the shortcut."""
    if com_client is None:
        print(
            "pywin32 not installed. Cannot dynamically lookup app_id. Install with 'pip install pywin32'."
        )
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
        match = re.search(r"--app-id=([a-z]{32})", shortcut.Arguments)
        if match:
            return match.group(1)
    except Exception as e:
        print(f"Error finding Google Voice chrome app_id: {e}.")
        return


def get_chrome_paths():
    """Find paths to chrome_proxy.exe and chrome.exe via registry and PATH."""
    proxy_path = None
    chrome_path = None
    try:
        # Try HKCU and HKLM to find Chrome paths
        paths = [
            (
                winreg.HKEY_CURRENT_USER,
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
            ),
            (
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
            ),
        ]
        for root, subkey in paths:
            try:
                key = winreg.OpenKey(root, subkey)
                chrome_exe_path = winreg.QueryValueEx(key, None)[0]
                winreg.CloseKey(key)
                if os.path.exists(chrome_exe_path):
                    chrome_path = chrome_exe_path
                    proxy_candidate = chrome_exe_path.replace(
                        "chrome.exe", "chrome_proxy.exe"
                    )
                    if os.path.exists(proxy_candidate):
                        proxy_path = proxy_candidate
                    break
            except OSError:
                pass

        # Fallback to PATH
        if not proxy_path:
            proxy_on_path = shutil.which("chrome_proxy.exe")
            if proxy_on_path:
                proxy_path = proxy_on_path
        if not chrome_path:
            chrome_on_path = shutil.which("chrome.exe")
            if chrome_on_path:
                chrome_path = chrome_on_path
    except Exception as e:
        print(f"Error finding Chrome path: {e}")
    return proxy_path, chrome_path


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


def register_handler(path: str = None):
    """Register the handler for tel: protocol with capabilities for Windows 11."""
    try:
        # Determine the running file path (script or executable)
        if path.lower().endswith(".exe"):
            running_path = path
            runner = ""
        elif getattr(sys, "frozen", False):
            running_path = sys.executable
            runner = ""
        else:

            running_path = (
                path if path.lower().endswith(".py") else os.path.abspath(__file__)
            )
            python_exe = sys.executable
            # Use pythonw.exe to run without a console window
            candidate = os.path.join(os.path.dirname(python_exe), "pythonw.exe")
            runner = (
                f'"{candidate}" ' if os.path.exists(candidate) else f'"{python_exe}" '
            )

        # Use HKCU for user-specific registration
        base_key = winreg.HKEY_CURRENT_USER

        # Create ProgId under HKCU\Software\Classes\<PROG_ID>
        classes_path = rf"Software\Classes\{PROG_ID}"
        prog_id_key = winreg.CreateKeyEx(base_key, classes_path, 0, winreg.KEY_WRITE)
        winreg.SetValueEx(prog_id_key, None, 0, winreg.REG_SZ, f"URL:{PROG_NAME}")
        winreg.SetValueEx(prog_id_key, "URL Protocol", 0, winreg.REG_SZ, "")

        # Optional: DefaultIcon (point to running_path for icon)
        icon_key = winreg.CreateKeyEx(prog_id_key, "DefaultIcon", 0, winreg.KEY_WRITE)
        winreg.SetValueEx(icon_key, None, 0, winreg.REG_SZ, f"{running_path},0")
        winreg.CloseKey(icon_key)

        # Create shell\open\command under ProgId
        command_key = winreg.CreateKeyEx(
            prog_id_key, r"shell\open\command", 0, winreg.KEY_WRITE
        )
        command_value = f'{runner}"{running_path}" "%1"'
        winreg.SetValueEx(command_key, None, 0, winreg.REG_SZ, command_value)
        winreg.CloseKey(command_key)
        winreg.CloseKey(prog_id_key)

        # Register capabilities under HKCU\Software\<PROG_ID>\Capabilities
        cap_key = winreg.CreateKeyEx(
            base_key, rf"Software\{PROG_ID}\Capabilities", 0, winreg.KEY_WRITE
        )
        winreg.SetValueEx(cap_key, "ApplicationName", 0, winreg.REG_SZ, PROG_NAME)
        winreg.SetValueEx(
            cap_key, "ApplicationDescription", 0, winreg.REG_SZ, PROG_DESC
        )
        # Optional: ApplicationIcon
        winreg.SetValueEx(
            cap_key, "ApplicationIcon", 0, winreg.REG_SZ, f"{running_path},0"
        )

        # URLAssociations
        url_assoc_key = winreg.CreateKeyEx(
            cap_key, "URLAssociations", 0, winreg.KEY_WRITE
        )
        winreg.SetValueEx(url_assoc_key, "tel", 0, winreg.REG_SZ, PROG_ID)
        winreg.SetValueEx(url_assoc_key, "callto", 0, winreg.REG_SZ, PROG_ID)
        winreg.CloseKey(url_assoc_key)
        winreg.CloseKey(cap_key)

        # Register the app in RegisteredApplications
        reg_apps_key = winreg.CreateKeyEx(
            base_key, r"Software\RegisteredApplications", 0, winreg.KEY_WRITE
        )
        winreg.SetValueEx(
            reg_apps_key,
            PROG_NAME,
            0,
            winreg.REG_SZ,
            rf"Software\{PROG_ID}\Capabilities",
        )
        winreg.CloseKey(reg_apps_key)

        print(f"Successfully registered '{PROG_ID}' tel: protocol handler.")
        print(
            "Go to Settings > Apps > Default apps > Set defaults for applications, and search 'Google Voice Dialer'. Set it as default for 'TEL' links."
        )
    except PermissionError:
        print(
            "Permission denied: Run the script as Administrator if needed (though HKCU should not require it)."
        )
    except Exception as e:
        print(f"Error registering: {e}")


def unregister_handler(prog_id=PROG_ID, prog_name=PROG_NAME):
    """Unregister the tel: protocol handler and capabilities."""
    try:

        def delete_key_recursive(root, path):
            # Deletes path and all subkeys under root using stable full paths
            try:
                with winreg.OpenKey(root, path, 0, winreg.KEY_ALL_ACCESS) as key:
                    while True:
                        sub = winreg.EnumKey(key, 0)
                        delete_key_recursive(root, path + "\\" + sub)
            except OSError:
                pass
            try:
                winreg.DeleteKey(root, path)
            except FileNotFoundError:
                pass

        # Delete ProgId from HKCU\Software\Classes\<PROG_ID>
        delete_key_recursive(winreg.HKEY_CURRENT_USER, rf"Software\Classes\{prog_id}")

        # Delete capabilities from HKCU\Software\<PROG_ID>
        delete_key_recursive(winreg.HKEY_CURRENT_USER, rf"Software\{prog_id}")

        # Remove from RegisteredApplications
        try:
            reg_apps_key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\RegisteredApplications",
                0,
                winreg.KEY_ALL_ACCESS,
            )
            try:
                winreg.DeleteValue(reg_apps_key, prog_name)
            except FileNotFoundError:
                pass
            winreg.CloseKey(reg_apps_key)
        except OSError:
            pass

        print(f"Successfully unregistered '{prog_id}' tel: protocol handler.")
    except PermissionError:
        print("Permission denied: Run the script as Administrator to unregister.")
    except Exception as e:
        print(f"Error unregistering: {e}")


def install():
    try:
        appdata_dir = os.path.expandvars(rf"%APPDATA%\{PROG_ID}")
        os.makedirs(appdata_dir, exist_ok=True)
        target_exe = os.path.join(appdata_dir, PROG_ID + ".exe")
        target_py = os.path.join(appdata_dir, "google_voice_dialer.py")

        if getattr(sys, "frozen", False):
            # Running as executable, copy self to appdata
            source = sys.executable
            target = target_exe
            if os.path.abspath(source) != os.path.abspath(target):
                shutil.copy(source, target)
            else:
                print("Already installed in AppData.")
            install_path = target
        else:
            # Running as script, copy the script to appdata
            source = os.path.abspath(__file__)
            target = target_py
            if os.path.abspath(source) != os.path.abspath(target):
                shutil.copy(source, target)
            else:
                print("Already installed in AppData.")
            install_path = target

        # Register the copied file
        if os.path.exists(install_path):
            register_handler(path=install_path)
            subprocess.call(["start", "ms-settings:defaultapps"], shell=True)
        else:
            print("File not found at expected path.")

    except Exception as e:
        print(f"Error installing: {e}")


def uninstall():
    unregister_handler()
    appdata_dir = os.path.expandvars(rf"%APPDATA%\{PROG_ID}")
    if os.path.exists(appdata_dir):
        shutil.rmtree(appdata_dir)
    print("Uninstalled successfully.")


def dial(phone_url):
    """Handle dialing via Google Voice."""
    if not (
        phone_url.lower().startswith("tel:") or phone_url.lower().startswith("callto:")
    ):
        return
    # Extract and clean phone number (preserve single leading +, strip non-digits)
    phone = re.sub(r"^(tel|callto):", "", phone_url, flags=re.IGNORECASE).strip()
    phone = urllib.parse.unquote(phone)

    # Strip everything after the first , or #
    idx_comma = phone.find(",")
    idx_hash = phone.find("#")
    if idx_comma == -1:
        idx = idx_hash
    elif idx_hash == -1:
        idx = idx_comma
    else:
        idx = min(idx_comma, idx_hash)
    if idx != -1:
        phone = phone[:idx].strip()

    plus = "+" if phone.startswith("+") else ""
    digits = re.sub(r"\D", "", phone)
    phone = plus + digits

    # Log the dial attempt
    try:
        base_dir = (
            os.path.dirname(sys.executable)
            if getattr(sys, "frozen", False)
            else os.path.dirname(os.path.abspath(__file__))
        )
        log_path = os.path.join(base_dir, LOG_FILENAME)
        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write(f"{datetime.datetime.now()} - {phone}\n")
    except Exception:
        pass

    # Encode for URL
    encoded_phone = urllib.parse.quote(phone)

    # Google Voice dialing URL
    gv_url = f"https://voice.google.com/u/0/calls?a=nc,{encoded_phone}"

    # Dynamically get Chrome proxy path and app_id
    proxy_path, chrome_path = get_chrome_paths()
    app_id = get_google_voice_app_id()

    if proxy_path and app_id:
        subprocess.run(
            [
                proxy_path,
                f"--app-id={app_id}",
                f"--app-launch-url-for-shortcuts-menu-item={gv_url}",
            ],
            check=False,
        )
    elif chrome_path:
        subprocess.run([chrome_path, f"--app={gv_url}"], check=False)
    else:
        webbrowser.open(gv_url)


def main():
    parser = argparse.ArgumentParser(description=f"Google Voice Dialer v{VERSION}")
    parser.add_argument(
        "--install", action="store_true", help="build and/or install application"
    )
    parser.add_argument(
        "--uninstall", action="store_true", help="uninstall application"
    )
    parser.add_argument(
        "--register", action="store_true", help="register handler for TEL links"
    )
    parser.add_argument(
        "--unregister", action="store_true", help="unregister handler for TEL links"
    )
    parser.add_argument("url", nargs="?", help="tel: URL to dial")

    args = parser.parse_args()

    if args.install:
        install()
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
        if getattr(sys, "frozen", False):
            win32api.MessageBox(
                0,
                help_text,
                "Google Voice Dialer Usage",
                win32con.MB_OK | win32con.MB_ICONINFORMATION,
            )
        else:
            print(help_text)
        sys.exit(0)


if __name__ == "__main__":
    main()

