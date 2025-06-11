# TODO: make a rule to check that the yaml file doesn't have any Bash@3 tasks where under input there is a `script:`, but there isn't a `targetType: inline` set - which makes certain things like reading from the FS fail silently

# TODO: (more broadly for YAML - check that no comments are used before what looks like code when the yaml block scalar is `>` - give a warning along lines of, "hey! you might have accidentally commented out your code here!")
