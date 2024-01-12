# ---------------
# --- IMPORTS ---
# ---------------

# built in
from pprint import pprint
from typing import List
# site
import win32com.client
# package
from outlookcustomtasks.settings import get_settings

# -----------------
# --- CONSTANTS ---
# -----------------
MOVE_RESPONSE_DEFAULT = "n"

# ---------------------
# --- INITIALIZATION ---
# ---------------------
_settings = get_settings()
outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace("MAPI")

def move_messages(messages: List[win32com.client.CDispatch], target_folder: win32com.client.CDispatch):
    for message in messages:
        message.Move(target_folder)

# --------------
# --- SCRIPT ---
# --------------

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
for message in inbox.Items:
    if message.Subject == _settings["subject_to_match"]:
        print(f"Found match: \"{message.Subject}\" from [{message.SenderEmailAddress}]({message.SenderName})] at {message.ReceivedTime}")
        matches.append(message)

match_count = len(matches)

print(f"Found {match_count} matching emails.")

if match_count > 0:
    while True:
        move_response = input(f"Move {match_count} matching emails to the folder \"{_settings['target_folder_name']}\"? (Y/n) [{MOVE_RESPONSE_DEFAULT}]:")

        # if response is empty, use default
        if move_response == "":
            move_response = MOVE_RESPONSE_DEFAULT

        if move_response == "Y":
            print("moving...")
            move_messages(matches, target_folder)
            print("moved!")
            break
        elif move_response == "n":
            print("skipping move.")
            break
        else:
            print("Invalid response.")
