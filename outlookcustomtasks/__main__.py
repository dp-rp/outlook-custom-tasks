# --------------
# --- IMPORTS ---
# --------------

# built in
from pprint import pprint
from collections import Counter
# site
from colorama import Fore, Style, init as colorama_init
from alive_progress import alive_bar as progressBar
# package
from outlookcustomtasks.outlook import OutlookClient, conditions as cond
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
olc = OutlookClient()

# ----------------
# --- FUNCTIONS ---
# ----------------

def get_predicates_from_conditions(conditions):
    print("conditions:",conditions)
    predicates = []
    with progressBar(len(conditions), title=f"generating predicates", bar="filling", stats=True) as bar:
        for condition in conditions:
            if "subject_matches" in condition:
                predicates.append(lambda m: m.Subject == condition["subject_matches"])
                bar()
            elif "subject_starts_with" in condition:
                predicates.append(lambda m: m.Subject.startswith(condition["subject_starts_with"]))
                bar()
            elif "sender_matches" in condition:
                sender_to_match = condition["sender_matches"]
                print(condition)
                predicates.append(lambda m: getattr(m,'SenderEmailAddress',None) == sender_to_match)
                bar()
            else:
                raise RuntimeError("unsupported condition!")
    return predicates

def run_rule(rule):
    print(f"running rule {Fore.GREEN}'{rule['name']}'{Style.RESET_ALL}...")
    # gen predicates based on conditions in config
    predicates = get_predicates_from_conditions(rule["conditions"])

    # find message matches using predicates
    matches = olc.find_messages(
        folder=olc.inbox(),
        filter_by=predicates
    )

    # print out basic information about each match
    for message in matches:
        print(f"{Fore.CYAN}Found match:{Style.RESET_ALL} \"{message.Subject}\" from [{message.SenderEmailAddress}]({message.SenderName})] at {message.ReceivedTime}")

    # print number of matches found
    match_count = len(matches)
    print(f"Found {match_count} matching emails.")
    
    action_count = len(rule["actions"])

    # HACK(Denver): just a quick way to do things - confirmation prompts should be based on which actions are associated with the given rule
    # if this rule doesn't have any actions
    if action_count < 1:
        print("No actions! Continuing to next rule...")
    elif action_count == 1:
        first_action = rule["actions"][0]
        
        if "move_to_folder" in first_action:
            target_folder_name = first_action["move_to_folder"]
            if match_count > 0:
                while True:
                    move_response = input(f"Move {match_count} matching emails to the folder \"{target_folder_name}\"? (Y/n) [{MOVE_RESPONSE_DEFAULT}]:")

                    # if response is empty, use default
                    if move_response == "":
                        move_response = MOVE_RESPONSE_DEFAULT

                    if move_response == "Y":
                        target_folder = olc.folder(target_folder_name)
                        olc.move_messages(matches, target_folder)
                        break
                    elif move_response == "n":
                        print(f"{Fore.YELLOW}skipping move.{Style.RESET_ALL}")
                        break
                    else:
                        print(f"{Fore.RED}Invalid response.{Style.RESET_ALL} {Fore.YELLOW}(note: input is case-sensitive){Style.RESET_ALL}")
    else:
        raise NotImplementedError("Multiple actions not yet supported.")

def group_by_sender_email_address(messages):
    grouped_by_sender = {}

    for message in messages:
        try:
            # try to get sender's email address
            sender_email_address = message.SenderEmailAddress

            # if no emails from sender yet
            if not sender_email_address in grouped_by_sender:
                # create empty array for sender messages
                grouped_by_sender[sender_email_address] = []

            # put message in group of other messages from same sender
            grouped_by_sender[message.SenderEmailAddress].append(message)

        except:
            print(f"{Fore.YELLOW}Warning: unknown sender email address for message with subject \"{message.Subject}\"{Style.RESET_ALL}")

    return grouped_by_sender

# -------------
# --- SCRIPT ---
# -------------

# for each rule defined in settings file
for rule in _settings["rules"]:
    # run the rule
    run_rule(rule)

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

# TODO(Denver): for stuff below, adapt any not already adapted logic to the "run_rule" flow so it can be
# ... defined in the config instead of in code

# TODO(Denver): make this it's own function
sender_grouped_messages = group_by_sender_email_address(olc.inbox().Items)

senders = [
    {
        "sender": sender,
        "messages": sender_messages,
        "message_count": len(sender_messages)
    }
    for sender, sender_messages
    in sender_grouped_messages.items()
]

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
    subject_counts = Counter([m.Subject for m in sender["messages"]]).most_common()

    # print total message count for sender
    print(f"({sender['message_count']:<3}): {sender['sender']}")

    # print total messages with same subject from same sender
    for subject, _count in subject_counts:
        print(f"  [{_count:<3}] {subject}")

# find all messages from given sender where the subject starts with the given string
messages_from_sender = olc.find_messages(
    folder=olc.inbox(),
    filter_by=[
        cond.is_sender_email_address(_settings["sender_for_subject_start_matching"]),
        cond.subject_starts_with(_settings["subject_start_to_match"])
    ]
)

for m in messages_from_sender:
    print(m.subject)
