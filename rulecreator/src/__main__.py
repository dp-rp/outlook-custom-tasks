import tkinter as tk
import json

RULES_FILE = "rules.json"

class RuleCreatorApp:
    def __init__(self):
        self.rules = []
        self.root = tk.Tk()
        self.root.title("Rule Creator")
        self.setup_ui()

    def save_rules(self):
        rules_data = {
            "schema_version": "1.0",
            "rules": [{"name": entry.get()} for entry in self.rules]
        }
        with open(RULES_FILE, "w") as file:
            json.dump(rules_data, file, indent=4)

    def add_rule(self):
        rule_frame = tk.Frame(self.root)
        rule_frame.pack(fill='x', padx=5, pady=5)

        rule_entry = tk.Entry(rule_frame)
        rule_entry.pack(side=tk.LEFT, expand=True, fill='x')

        edit_button = tk.Button(rule_frame, text="edit")
        edit_button.pack(side=tk.LEFT)

        delete_button = tk.Button(rule_frame, text="delete", command=lambda frame=rule_frame, entry=rule_entry: self.remove_rule(frame, entry))
        delete_button.pack(side=tk.LEFT)

        self.rules.append(rule_entry)

    def remove_rule(self, rule_frame, rule_entry):
        if rule_entry in self.rules:
            self.rules.remove(rule_entry)
            rule_frame.destroy()

    def setup_ui(self):
        # Clear the current UI elements
        for widget in self.root.winfo_children():
            widget.destroy()

        # Add the "rules" label frame
        rules_frame = tk.LabelFrame(self.root, text="rules")
        rules_frame.pack(fill='x', padx=10, pady=10)

        # Add the "add rule" and "save rules" buttons
        top_frame = tk.Frame(rules_frame)
        top_frame.pack(fill='x')

        add_button = tk.Button(top_frame, text="add rule", command=self.add_rule)
        add_button.pack(side=tk.LEFT, padx=5)

        save_button = tk.Button(top_frame, text="save rules", command=self.save_rules)
        save_button.pack(side=tk.RIGHT, padx=5)

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = RuleCreatorApp()
    app.run()
