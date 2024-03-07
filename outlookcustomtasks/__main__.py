# --------------
# --- IMPORTS ---
# --------------

# built in
from pprint import pprint
from collections import Counter
from itertools import groupby
# site
from colorama import Fore, Back, Style, init as colorama_init
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
                def sender_matches(sender_to_match):
                    return lambda m: get_sender_email_address(m) == sender_to_match
                
                predicates.append(sender_matches(condition["sender_matches"]))
                bar()
            else:
                raise RuntimeError("unsupported OCT condition!")
    return predicates

def get_real_folder_idx(target_folder_name,folders):
    for idx, folder in enumerate(folders):
        if folder.Name == target_folder_name:
            return idx
    else:
        return None


def get_folders_from_targets(targets):
    all_folders = olc.inbox_folders_recursive_flat() # HACK: grabs all folders recursively up front even if while iterating through targets we've already got all the ones we need
    target_folders = []

    # for each target
    for idx, target in enumerate(targets):
        target_folder_name = target['folder_name']
        recursive = target['recursive']
        # FIXME: will just grab the first one it sees if any folders share names
        # try to find real folder with target folder name
        folder_idx = get_real_folder_idx(target_folder_name, all_folders)
        # if target_folder_name didn't match any real folder's names
        if folder_idx is None:
            raise Exception(f"Failed to find any folders with the name '{target_folder_name}'")
            
        # TODO: do stuff when matched
        if recursive is True:
            # target_folders.append(olc._get_subfolders_recursively(initial_folder_match))
            # TODO: handle recursive targets
            raise NotImplementedError("TODO")
        elif recursive is False:
            target_folders.append(all_folders[folder_idx])
        else:
            raise Exception("target being recursive must be true or false")

    # for each real folder
    # for folder in all_folders:
    #     target = targets.find(folder.name)
    #     # if the real folder name matches a target folder name
    #     if folder.Name in target_folder_names:
    #         # if the target IS recursive
    #         if targets[]
    #         # if the target ISN'T recursive
    #         target_folders.append(folder)
    
    return target_folders


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
    for m in matches:
        print(f"{Fore.CYAN}Found match:{Style.RESET_ALL} \"{m.Subject}\" from [{get_sender_email_address(m)}]({m.SenderName})] at {m.ReceivedTime}")

    # print number of matches found
    match_count = len(matches)
    print(f"Found {match_count} matching emails.")
    
    if match_count > 0:

        action_count = len(rule["actions"])

        # HACK(Denver): just a quick way to do things - confirmation prompts should be based on which actions are associated with the given rule
        # if this rule doesn't have any actions
        if action_count < 1:
            print("No actions! Continuing to next rule...")
        elif action_count == 1:
            first_action = rule["actions"][0]
            
            if "move_to_folder" in first_action:
                target_folder_name = first_action["move_to_folder"]
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
            
            elif "basic_analytics" in first_action:
                # HACK: just a quick dirty implementation
                print("\n----[ Basic Analytics ]----\n")
                
                # messages_grouped_by_sender = group_by_sender_email_address(olc.inbox().Items)

                # senders = [
                #     {
                #         "email_address": sender_email_address,
                #         "messages": sender_messages,
                #         "message_count": len(sender_messages)
                #     }
                #     for sender_email_address, sender_messages
                #     in messages_grouped_by_sender.items()
                #     # filter out senders without at least 3 messages
                #     if len(sender_messages) >= 3
                # ]

                # # sort sender list by number of messages received
                # senders_sorted_by_message_count = sorted(
                #     senders,
                #     key=lambda sender: sender["message_count"],
                # )

                # # for each sender
                # for sender in senders_sorted_by_message_count:
                #     subject_counts = Counter([m.Subject for m in sender["messages"]]).most_common()

                #     # print total message count for sender
                #     print(f"({sender['message_count']:<3}): {sender['email_address']}")

                #     # print total messages with same subject from same sender
                #     for subject, _count in subject_counts:
                #         print(f"  [{_count:<3}] {subject}")
                        
                # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                # GET TOP 10 IDENTICAL SUBJECTS REGARDLESS OF SENDER #
                # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                LIMIT = 10
                
                messages_grouped_by_subject = group_by_subject(matches)

                subjects = [
                    {
                        "subject": subject,
                        "messages": subject_messages,
                        "message_count": len(subject_messages)
                    }
                    for subject, subject_messages
                    in messages_grouped_by_subject.items()
                ]

                subjects_sorted_by_message_count = sorted(
                    subjects,
                    key=lambda s: s['message_count'],
                    reverse=True
                )[:LIMIT]

                # print(f"Top {limit} offenders for senders that have sent messages with identical subject lines:")
                print(f"\n{Fore.BLACK}{Back.GREEN} Top {LIMIT} (or less) offenders for messages with identical subject lines (regardless of sender): {Style.RESET_ALL}\n")

                for subject in subjects_sorted_by_message_count:
                    print(f"[ {subject['message_count']:>3} x ] subject: [{subject['subject']}]")
                    
                print()
                
            else:
                raise Exception(f"No supported OCT action names were found in defined actions")
                
        else:
            raise NotImplementedError("Multiple OCT actions not yet supported.")

def group_by_sender_email_address(messages):
    grouped_by_sender = {}

    with progressBar(len(messages), title=f"grouping messages by sender", bar="filling") as bar:
        for message in messages:
            # get sender's email address
            sender_email_address = get_sender_email_address(message)

            # if no emails from sender yet
            if not sender_email_address in grouped_by_sender:
                # create empty array for sender messages
                grouped_by_sender[sender_email_address] = []

            # put message in group of other messages from same sender
            grouped_by_sender[sender_email_address].append(message)
            bar()

    return grouped_by_sender

# HACK: very similar usage to group_by_sender_email_address, should DRY this out
def group_by_subject(messages):
    keyfunc = lambda m: m.Subject

    # grouped_by_subject = {
    #     subject: list(subject_messages)
    #     for subject, subject_messages
    #     in groupby(
    #         sorted(messages, key=keyfunc),
    #         keyfunc
    #     )
    # }

    grouped_by_subject = {}
    with progressBar(title=f"grouping messages by subject", bar="filling") as bar:
        for subject, subject_messages in groupby(
            sorted(messages, key=keyfunc),
            keyfunc
        ):
            grouped_by_subject[subject] = list(subject_messages)
            bar()

    return grouped_by_subject

def get_sender_email_address(message):
    try:
        return message.SenderEmailAddress
    except:
        print(f"{Fore.YELLOW}Warning: unknown sender email address for message with subject \"{message.Subject}\"{Style.RESET_ALL}")
        return None

# -------------
# --- SCRIPT ---
# -------------

# for each rule defined in settings file
if len(_settings["rules"]) < 1:
    print(f"{Fore.YELLOW}Warning: No OCT rules defined!{Style.RESET_ALL}")
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
#     print(f"Found match: \"{message.Subject}\" from [{get_sender_email_address(m)}]({message.SenderName})] at {message.ReceivedTime}")

# print(f"Found {match_count} matching emails.")

# collected_messages = [item for item in olc.inbox().Items]
# print(len(collected_messages))
# first_level_subfolder_messages = [folder.Items for folder in olc.inbox().Folders]

# for item in first_level_subfolder_messages:
#     print(item.Subject)

# collected_messages += first_level_subfolder_messages
# print(len(collected_messages))
