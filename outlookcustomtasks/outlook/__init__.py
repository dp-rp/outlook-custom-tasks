# built in
from typing import List
# site
import win32com.client
from alive_progress import alive_bar as progressBar

class OutlookClient:
    def __init__(self) -> None:
        with progressBar(title=f"connecting to Outlook locally", bar="filling") as bar:
            self._outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace("MAPI")
            bar()
        self._inbox = None

    def inbox(self):
        if self._inbox is None:
            self._inbox = self._outlook.GetDefaultFolder(6)
        return self._inbox
    
    def _get_subfolders_recursively(self, folder: win32com.client.CDispatch, acc: List[win32com.client.CDispatch]):
        # for each subfolder (might be an empty list)
        for subfolder in list(folder.Folders):
            # get any subfolder subfolders
            acc = self._get_subfolders_recursively(subfolder,acc)

        # add this folder to acc (whether it has subfolders or not)
        acc.append(folder)

        return acc

    def inbox_folders_recursive_flat(self) -> List[win32com.client.CDispatch]:
        return self._get_subfolders_recursively(self.inbox(), [])

    def folder(self, target_folder_name):
        target_folder = None
        for folder in self.inbox().Folders:
            if folder.Name == target_folder_name:
                target_folder = folder
                break

        if target_folder is None:
            raise RuntimeError("Failed to find target folder")
        else:
            print(f"Using target folder: \"{target_folder.Name}\"")

        return target_folder
    
    def find_messages(self, folders, filter_by):
        matches = []
        total_message_count = sum([len(f.Items) for f in folders])
        with progressBar(total_message_count, title=f"searching for matches", bar="filling", stats=True) as bar:
            for folder in folders:
                folder_item_count = len(folder.Items)
                folder_item_count_chars = len(str(folder_item_count))
                # for each message
                for idx, message in enumerate(folder.Items):
                    # update progress bar title
                    bar.title(f"searching for matches in '{folder.Name}' ({idx:>{folder_item_count_chars}}/{folder_item_count})")

                    # assume match until proven otherwise
                    is_match = True

                    # for each predicate
                    for _filter in filter_by:
                        # if any predicates aren't satisfied
                        if not _filter(message):
                            # mark message as not a match
                            is_match = False
                            # skip any other predicates to check
                            break

                    if is_match:
                        matches.append(message)

                    bar()

        return matches

    def move_messages(self, messages: List[win32com.client.CDispatch], target_folder: win32com.client.CDispatch):
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
