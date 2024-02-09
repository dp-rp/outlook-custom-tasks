import json

SETTINGS_FILEPATH = 'oct.settings.json'

def get_settings():
    with open(SETTINGS_FILEPATH, 'r') as f:
        settings = json.load(f)
        if settings["version"] == "1.0.0":
            return {
                "target_folder_name": settings["data"]["targetFolderName"],
                "subject_to_match": settings["data"]["subjectToMatch"],
                "subject_start_to_match": settings["data"]["subjectStartToMatch"],
                "sender_for_subject_start_matching": settings["data"]["senderForSubjectStartMatching"]
            }
