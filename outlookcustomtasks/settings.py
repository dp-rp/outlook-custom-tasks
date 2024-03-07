import json

SETTINGS_FILEPATH = 'oct.settings.json'

def get_settings():
    with open(SETTINGS_FILEPATH, 'r') as f:
        settings = json.load(f)
        if settings["version"] == "1.0.0":
            rules = []
            DEFAULT_TARGETS = [{"folder_name": "Inbox", "recursive": False}]
            for rule in settings["data"]["rules"]:
                rules.append({
                    "name": rule["name"],
                    "targets": rule["targets"] if "targets" in rule else DEFAULT_TARGETS,
                    "conditions": rule["conditions"],
                    "actions": rule["actions"]
                })

            return {
                "rules": rules
            }
