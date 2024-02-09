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
            print(f"{Fore.YELLOW}skipping move.{Style.RESET_ALL}")
            break
        else:
            print(f"{Fore.RED}Invalid response.{Style.RESET_ALL} {Fore.YELLOW}(note: input is case-sensitive){Style.RESET_ALL}")

# # FIXME: below is temporary
# matches = olc.find_messages(
#     folder = inbox,
#     filter_by = [
#         lambda m: "Blog" in m.Subject 
#     ]
# )

# match_count = len(matches)

# for message in matches:
#     print(f"Found match: \"{message.Subject}\" from [{message.SenderEmailAddress}]({message.SenderName})] at {message.ReceivedTime}")

# print(f"Found {match_count} matching emails.")

# collected_messages = [item for item in olc.inbox().Items]
# print(len(collected_messages))
# first_level_subfolder_messages = [folder.Items for folder in olc.inbox().Folders]

# for item in first_level_subfolder_messages:
#     print(item.Subject)

# collected_messages += first_level_subfolder_messages
# print(len(collected_messages))


grouped_by_sender = {}
for message in olc.inbox().Items:
    try:
        sender_email_address = message.SenderEmailAddress

        if not sender_email_address in grouped_by_sender:
            grouped_by_sender[sender_email_address] = []
        
        grouped_by_sender[message.SenderEmailAddress].append(message)

    except:
        print(f"{Fore.YELLOW}Warning: unknown sender email address for message with subject \"{message.Subject}\"{Style.RESET_ALL}")

senders = []
for sender, sender_messages in grouped_by_sender.items():
    senders.append({
        "sender": sender,
        "messages": sender_messages,
        "message_count": len(sender_messages)
    })

# filter out senders without at least 3 messages
senders_at_least_3_messages = filter(
    lambda sender: sender["message_count"] >= 3,
    senders
)

# sort sender list by number of messages received
senders_sorted_by_message_count = sorted(
    senders_at_least_3_messages,
    key=lambda sender: sender["message_count"],
)

for sender in senders_sorted_by_message_count:
    duplicate_subject_count = 0
    for message in sender['messages']:
        message.Subject
    print(f"{sender['message_count']:<3} :{sender['sender']}")

test = olc.folder('Complete')
print(test)
