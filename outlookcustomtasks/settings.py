import json

SETTINGS_FILEPATH = 'oct.settings.json'

def load_settings():
    with open(SETTINGS_FILEPATH, 'r') as f:
        settings = json.load(f)
        if settings["version"] == "1.0.0":
            return {
                "target_folder_name": settings["data"]["targetFolderName"],
                "subject_to_match": settings["data"]["subjectToMatch"]
            }
