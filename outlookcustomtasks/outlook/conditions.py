def is_sender_email_address(sender_email_address: str):
    return lambda m: getattr(m,'SenderEmailAddress',None) == sender_email_address

def subject_starts_with(_starts_with: str):
    return lambda m: m.Subject.startswith(_starts_with)

def subject_matches(subject_to_match: str):
    return lambda m: m.Subject == subject_to_match

def is_unread(true_if_unread: bool):
    return lambda m: m.IsUnread is true_if_unread
