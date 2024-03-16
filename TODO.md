# TODO

- [ ] Write basic tests.
- [ ] Make the testing environment for the tool unable to connect to the internet other than whatever is required to interface with the container to ensure connection isn't necessary for the tool to work properly.
- [ ] If Outlook takes >30 seconds to give a response to a given query, prompt the user with an error message saying something along the lines of "Outlook is taking too long to respond, do you want me to manually kill it's process, reconnect, and try again?".
- [ ] **UX:** Provide a user-friendly interface or step-by-step wizard to help users create and manage their rules without requiring technical knowledge.
- [ ] **UX:** Implement an undo feature that allows users to revert any accidental changes made to their Outlook inbox.
- [ ] **Rule Creation:** Have a standard well considered way of allowing rules to have AND/OR/XOR/ETC. relationships between conditions (but in user-friendly language) and NOTs.
- [ ] **Rule Creation:** Introduce the ability to sort emails based on multiple attributes, such as read/unread status, sender name, date, etc.
- [ ] **AI:** Develop an AI component that can analyze a user's email patterns and suggest new rules for better email management.
- [ ] **AI:** Implement a rule creation feature that allows users to create their own custom rules based on the AI component.
- [ ] **AI:** The AI should be able to identify important emails based on user interactions, such as flagging or reading habits, and recommend actions accordingly.
- [ ] **Collaboration:** Create a system for users to share their custom rules and insights with team members, allowing for collaborative email management in a professional setting.
- [ ] **DX:** Work towards compliance with the SLSA (Supply-chain Levels for Software Artifacts) specification to enhance the security and integrity of the project.
- [ ] **DX:** Include a comprehensive set of tests to validate the functionality and reliability of the tool.
- [ ] **Docs/Support:** Provide detailed documentation on how to use the tool, configure rules, and troubleshoot common issues.
- [ ] **Docs/Support:** Establish a support system or community forum where users can seek assistance and share knowledge. (e.g. Discord, Slack, etc.)
- [ ] **Config Improvements:** config improvement suggestions
  1. **Validation and Error Handling**:
     - Consider adding validation checks to ensure that settings are correctly formatted.
     - Handle potential errors gracefully (e.g., invalid conditions, missing fields).
  2. **Default Values**:
     - Include default values for settings that are optional. For instance:
       ```json
       {
           "version": "1.0.0",
           "defaultFolder": "Inbox",
           // Other settings...
       }
       ```
     - This allows users to override defaults when needed.
  3. **Configuration Overrides**:
     - Allow users to override specific settings at runtime (e.g., via environment variables or command-line arguments).
  4. **Documentation**:
     - Provide comprehensive documentation explaining each setting, its purpose, and usage.
     - Consider adding a README or inline comments within the file.
