def is_sender_email_address(sender_email_address):
    return lambda m: getattr(m,'SenderEmailAddress',None) == sender_email_address

def subject_starts_with(_starts_with):
    return lambda m: m.Subject.startswith(_starts_with)

def subject_matches(subject_to_match):
    return lambda m: m.Subject == subject_to_match

def is_unread(true_if_unread):
    return lambda m: m.IsUnread == true_if_unread
