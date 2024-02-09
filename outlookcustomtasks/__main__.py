# --------------
# --- IMPORTS ---
# --------------

# built in
from pprint import pprint
# site
from colorama import Fore, Style, init as colorama_init
# package
from outlookcustomtasks.outlook import OutlookClient
from outlookcustomtasks.settings import get_settings

# ----------------
# --- CONSTANTS ---
# ----------------
MOVE_RESPONSE_DEFAULT = "n"

# ---------------------
# --- INITIALIZATION ---
# ---------------------
colorama_init()
_settings = get_settings()

# -------------
# --- SCRIPT ---
# -------------

olc = OutlookClient()
inbox = olc.inbox()
target_folder = olc.folder(_settings["target_folder_name"])

matches = olc.find_messages(
    folder = inbox,
    filter_by = [
        lambda m: m.Subject == _settings["subject_to_match"]
    ]
)

match_count = len(matches)

for message in matches:
    print(f"Found match: \"{message.Subject}\" from [{message.SenderEmailAddress}]({message.SenderName})] at {message.ReceivedTime}")

print(f"Found {match_count} matching emails.")

if match_count > 0:
    while True:
        move_response = input(f"Move {match_count} matching emails to the folder \"{_settings['target_folder_name']}\"? (Y/n) [{MOVE_RESPONSE_DEFAULT}]:")

        # if response is empty, use default
        if move_response == "":
            move_response = MOVE_RESPONSE_DEFAULT

        if move_response == "Y":
            olc.move_messages(matches, target_folder)
            break
        elif move_response == "n":
            print("skipping move.")
            break
        else:
            print(f"{Fore.RED}Invalid response.{Style.RESET_ALL} {Fore.YELLOW}(note: input is case-sensitive){Style.RESET_ALL}")
