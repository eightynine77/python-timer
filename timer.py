import os
import sys
import time
import subprocess
import ctypes
import msvcrt
import shutil
from colorama import init as _colorama_init
from termcolor import colored

os.system('title python timer')
MENU_EXECUTABLE = os.path.join("resource", "cmdmenusel.exe")
MENU_COLOR = "cff2"
POWERSHELL_SCRIPT = os.path.join("resource", "notification.ps1")
SOUND_FILE = os.path.join("resource", "alarm.mp3")
SOUND_FILE_PLAYER = os.path.join("resource", "ffplay.exe")

_colorama_init()

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def check_dependencies():
    """Checks if all required external files exist before starting."""
    dependencies = [MENU_EXECUTABLE, POWERSHELL_SCRIPT, SOUND_FILE, SOUND_FILE_PLAYER]
    missing_files = [f for f in dependencies if not os.path.exists(f)]
    if missing_files:
        print("[ERROR] The following required files are missing in the script's directory:")
        for f in missing_files:
            try:
                print(colored(f"- {f}", "green", attrs=["bold"]))
            except Exception:
                print(f"- {f}")
        print("\nPlease make sure all required files are in the same folder as the script.")
        print("")
        print(colored("NOTE: if you are downloading this not from the releases section of the github repo\nthen be sure to download from the releases section of the github repo since github doesn't allow file upload larger than 25 megabytes and ffplay.exe is heavier than that", "yellow", attrs=["bold"]))
        print("")
        print("download link: https://github.com/eightynine77/python-timer/releases")
        print("")
        print(colored("if you really wish to continue using this script then make sure you add ffplay.exe into this script's \"resource\" folder:", "yellow", attrs=["bold"]))
        print(colored(os.getcwd() + "\\resource", "cyan", attrs=["bold"]))
        print("")
        print("ffplay download link: https://www.ffmpeg.org/download.html")
        print("")
        os.system('pause')
        sys.exit(1)
    return True

def countdown(minutes):
    """
    Runs a countdown for a given number of minutes.
    Press 'P' to pause/resume.
    """
    total_seconds = int(minutes * 60)
    paused = False

    while total_seconds > 0:
        if msvcrt.kbhit():
            key = msvcrt.getch().decode('utf-8', errors='ignore').lower()
            if key == 'p':
                paused = not paused
                if paused:
                    print("\n-- PAUSED -- (Press 'P' again to resume)", end="")
                else:
                    print("\r-- RESUMED --                                \r", end="")

        if not paused:
            mins, secs = divmod(total_seconds, 60)
            timer_display = f"Time Remaining: {mins:02d}:{secs:02d}"
            print(timer_display, end="\r")
            time.sleep(1)
            total_seconds -= 1

    print("\n Time's up!                                    ")

def trigger_alarm_and_notification():
    """
    Plays a sound on loop, shows a PowerShell notification,
    and displays a message box. Stops the sound when the user clicks 'OK'.
    """
    print("Triggering alarm and notification...")

    creation_flags = 0x08000000  
    subprocess.Popen(
        ['powershell.exe', '-ExecutionPolicy', 'Bypass', '-File', POWERSHELL_SCRIPT],
        creationflags=creation_flags
    )

    ffplay_process = subprocess.Popen(
        [SOUND_FILE_PLAYER, '-nodisp', '-autoexit', '-loop', '0', SOUND_FILE],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )

ctypes.windll.user32.MessageBoxW(
    ctypes.windll.user32.GetForegroundWindow(),
    "Your countdown has finished. Click OK to silence the alarm.",
    "Time's Up!",
    0x00000040  
)

    try:
        ffplay_process.terminate()
    except Exception:
        pass

def single_timer_mode():
    """Handles the logic for a single countdown timer."""
    clear_screen()
    print("--- Single Timer Mode ---")
    while True:
        try:
            minutes = float(input("Enter countdown duration in minutes: "))
            if minutes > 0:
                break
            else:
                print("Please enter a positive number.")
        except ValueError:
            print("Invalid input. Please enter a number.")

    print("Press 'P' at any time to pause or resume.")
    countdown(minutes)
    trigger_alarm_and_notification()
    input("\nPress Enter to return to the main menu...")

def multiple_timer_mode():
    """Handles the logic for a sequence of countdown timers."""
    clear_screen()
    print("--- Multiple Timers Mode ---")

    while True:
        try:
            num_timers = int(input("How many timers do you want to set? "))
            if num_timers > 0:
                break
            else:
                print("Please enter a number greater than 0.")
        except ValueError:
            print("Invalid input. Please enter a whole number.")

    timers = []
    for i in range(num_timers):
        while True:
            try:
                minutes = float(input(f"Enter duration for timer #{i+1} in minutes: "))
                if minutes > 0:
                    timers.append(minutes)
                    break
                else:
                    print("Please enter a positive number.")
            except ValueError:
                print("Invalid input. Please enter a number.")

    for i, minutes in enumerate(timers):
        clear_screen()
        print(f"--- Running Timer {i+1} of {num_timers} ({minutes} minutes) ---")
        print("Press 'P' at any time to pause or resume.")
        countdown(minutes)
        trigger_alarm_and_notification()

        if i < len(timers) - 1:  
            input("\nPress Enter to start the next timer...")

    input("\nAll timers finished. Press Enter to return to the main menu...")

STARTUP_FOLDER = os.path.join(
    os.environ.get('APPDATA', r''),
    'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup'
)

def _ps_single_quote_escape(s: str) -> str:
    """Escape single quotes for embedding in PowerShell single-quoted strings."""
    return s.replace("'", "''")

def run_at_startup_menu():
    """
    Submenu to enable/disable "run at startup". Creates a .lnk in Startup that
    points to the running Python interpreter with this script as an argument.
    Uses a default shortcut name (script name + .lnk) without prompting the user.
    """

    os.makedirs(STARTUP_FOLDER, exist_ok=True)
    script_path = os.path.abspath(sys.argv[0])
    script_dir = os.path.dirname(script_path)
    default_name = os.path.splitext(os.path.basename(script_path))[0] + ".lnk"
    link_path = os.path.join(STARTUP_FOLDER, default_name)

    while True:
        clear_screen()
        print("--- Run At Startup ---")
        print(f"Shortcut name (automatic): {default_name}")
        process = subprocess.run(
            [MENU_EXECUTABLE, MENU_COLOR, "Enable Run at Startup", "Disable Run at Startup", "Go back"],
            capture_output=True, text=True
        )
        choice = process.returncode

        if choice == 0 or choice == 3:
            return

        if choice == 1:  
            clear_screen()
            print("--- Enable Run at Startup ---")
            print("")
            print("loading...")
            try:
                if os.path.exists(link_path):
                    try:
                        os.remove(link_path)
                    except Exception:
                        pass

                ps_cmd = (
                    f"$WshShell = New-Object -ComObject WScript.Shell; "
                    f"$Shortcut = $WshShell.CreateShortcut('{_ps_single_quote_escape(link_path)}'); "
                    f"$Shortcut.TargetPath = '{_ps_single_quote_escape(script_path)}'; "
                    f"$Shortcut.WorkingDirectory = '{_ps_single_quote_escape(script_dir)}'; "
                    f"$Shortcut.Save();"
                )
                subprocess.run(
                    ['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', ps_cmd],
                    check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
                )
                print(f"\nsuccess! Shortcut '{default_name}' created in Startup.")
            except subprocess.CalledProcessError as e:
                print(f"\n[ERROR] Failed to create shortcut: {e}")
            except Exception as e:
                print(f"\n[ERROR] Unexpected error: {e}")

            input("\nPress Enter to continue...")

        elif choice == 2:  
            clear_screen()
            print("--- Disable Run at Startup ---")
            if os.path.exists(link_path):
                try:
                    os.remove(link_path)
                    print(f"\nSuccess! '{default_name}' has been removed from startup.")
                except Exception as e:
                    print(f"\n[ERROR] An error occurred while trying to remove the file: {e}")
            else:
                print(f"\n[INFORMATION] The file '{default_name}' was not found in the startup folder.")
            input("\nPress Enter to continue...")

        else:
            return

def manage_reminders():
    """Handles copying or removing a Sticky Note shortcut from the Startup folder."""
    while True:
        clear_screen()
        print("--- Reminder Setup ---")
        print("This feature copies a Sticky Note shortcut to your Windows Startup folder.")

        process = subprocess.run(
            [MENU_EXECUTABLE, MENU_COLOR, "Enable Startup Reminder", "Disable Startup Reminder", "go back"],
            capture_output=True,
            text=True
        )
        choice = process.returncode

        if choice == 0 or choice == 3:
            return

        startup_folder = STARTUP_FOLDER 

        if choice == 1:  
            clear_screen()
            print("--- Enable Startup Reminder ---")
            shortcut_path = input("Please enter the full path to your Sticky Note shortcut (.lnk) file:\n> ").strip('"')

            if not os.path.exists(shortcut_path) or not shortcut_path.endswith('.lnk'):
                print("\n[ERROR] The file does not exist or is not a valid .lnk shortcut file.")
                input("\nPress Enter to continue...")
                continue

            try:
                shutil.copy(shortcut_path, startup_folder)
                print(f"\nSuccess! Reminder '{os.path.basename(shortcut_path)}' has been added to startup.")
            except Exception as e:
                print(f"\n[ERROR] An error occurred: {e}")
            input("\nPress Enter to continue...")

        elif choice == 2:  
            clear_screen()
            print("--- Disable Startup Reminder ---")
            shortcut_name = input("Enter the name of the shortcut file to remove from startup (e.g., 'Sticky Notes.lnk'):\n> ")
            target_path = os.path.join(startup_folder, shortcut_name)

            if os.path.exists(target_path):
                try:
                    os.remove(target_path)
                    print(f"\nSuccess! '{shortcut_name}' has been removed from startup.")
                except Exception as e:
                    print(f"\n[ERROR] An error occurred while trying to remove the file: {e}")
            else:
                print(f"\n[INFORMATION] The file '{shortcut_name}' was not found in the startup folder.")
            input("\nPress Enter to continue...")

def main():
    """Main function to display the menu and handle user choices."""
    check_dependencies()
    while True:
        clear_screen()
        print("====== PYTHON COUNTDOWN TIMER ======")

        process = subprocess.run(
            [MENU_EXECUTABLE, MENU_COLOR,
             "Single Timer", "Multiple Timers", "Run at Startup", "Set Reminder", "Exit"],
            capture_output=True,
            text=True
        )
        choice = process.returncode

        if choice == 0 or choice == 5:
            print("Exiting. Goodbye!")
            break

        if choice == 1:
            single_timer_mode()
        elif choice == 2:
            multiple_timer_mode()
        elif choice == 3:
            run_at_startup_menu()
        elif choice == 4:
            manage_reminders()
        else:
            continue

if __name__ == "__main__":
    main()