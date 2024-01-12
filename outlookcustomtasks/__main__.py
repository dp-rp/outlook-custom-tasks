# ---------------
# --- IMPORTS ---
# ---------------

# built in
from pprint import pprint
# site
import win32com.client
# package
from outlookcustomtasks.settings import get_settings

# -----------------
# --- CONSTANTS ---
# -----------------
VALID_MOVE_REPONSE_OPTIONS = ["Y","n"]

# ---------------------
# --- INITIALIZATION ---
# ---------------------
_settings = get_settings()
outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace("MAPI")

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

match_count = 0
for message in inbox.Items:
    if message.Subject == _settings["subject_to_match"]:
        print(f"Found match: \"{message.Subject}\" from [{message.SenderEmailAddress}]({message.SenderName})] at {message.ReceivedTime}")
        match_count += 1

# move_response = None
# while not (move_response in VALID_MOVE_REPONSE_OPTIONS):
#     move_response = input(f"Found {match_count} matching emails. Move them to {TARGET_FOLDER_NAME}? (Y/n) [n]:")
#     if VALID_MOVE_REPONSE_OPTIONS
