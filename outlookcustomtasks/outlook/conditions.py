def is_sender_email_address(sender_email_address):
    return lambda m: getattr(m,'SenderEmailAddress',None) == sender_email_address

def subject_starts_with(_starts_with):
    return lambda m: m.Subject.startswith(_starts_with)
