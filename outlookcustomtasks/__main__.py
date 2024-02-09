# --------------
# --- IMPORTS ---
# --------------

# built in
from pprint import pprint
from typing import List
# site
import win32com.client
from alive_progress import alive_bar as progressBar
from colorama import Fore, Style, init as colorama_init
# package
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
outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace("MAPI")

# ----------------
# --- FUNCTIONS ---
# ----------------
def move_messages(messages: List[win32com.client.CDispatch], target_folder: win32com.client.CDispatch):
    message_count = len(messages)
    moved_messages = []

    try:
        with progressBar(message_count, title="moving emails", bar="filling", stats=True, enrich_print=True) as bar:
            # count = 0
            for message in messages:
                message.Move(target_folder)
                moved_messages.append(message)
                # count += 1
                # if count==20:
                #     raise RuntimeError("something wrong!")
                bar()
        print("")
    except Exception as e:
        print("Error: Failed to move all messages!")
        print(e)

    # print out basic identifying information about all messages that were successfully moved
    for idx, message in enumerate(moved_messages):
        print(f"[{idx+1:>{len(str(message_count))}}/{message_count}] Moved email received at {message.ReceivedTime} with subject \"{message.Subject}\" to the folder \"{target_folder.Name}\"")
    print("")

# -------------
# --- SCRIPT ---
# -------------

inbox = outlook.GetDefaultFolder(6)

target_folder = None
for folder in inbox.Folders:
    if folder.Name == _settings["target_folder_name"]:
        target_folder = folder
        break

if not target_folder:
    raise RuntimeError("Failed to find target folder")
else:
    print(f"Using target folder: \"{target_folder.Name}\"")

matches = []
with progressBar(len(inbox.Items), title=f"searching for matches", bar="filling", stats=True) as bar:
    for message in inbox.Items:
        if message.Subject == _settings["subject_to_match"]:
            matches.append(message)
        bar()

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
            move_messages(matches, target_folder)
            break
        elif move_response == "n":
            print("skipping move.")
            break
        else:
            print(f"{Fore.RED}Invalid response.{Style.RESET_ALL} {Fore.YELLOW}(note: input is case-sensitive){Style.RESET_ALL}")
