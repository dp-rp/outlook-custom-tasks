import json

SETTINGS_FILEPATH = 'oct.settings.json'

def get_settings():
    with open(SETTINGS_FILEPATH, 'r') as f:
        settings = json.load(f)
        if settings["version"] == "1.0.0":
            rules = []
            for rule in settings["data"]["rules"]:
                rules.append({
                    "name": rule["name"],
                    "conditions": rule["conditions"],
                    "actions": rule["actions"]
                })

            return {
                "rules": rules
            }
