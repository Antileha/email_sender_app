import re

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")


def parse_emails(raw_text: str):
    if not raw_text:
        return []

    normalized = raw_text.replace(";", ",").replace("\n", ",").replace("\t", ",")
    parts = [x.strip() for x in normalized.split(",") if x.strip()]

    result = []
    seen = set()

    for email in parts:
        key = email.lower()
        if key not in seen:
            result.append(email)
            seen.add(key)

    return result


def validate_emails(emails):
    valid = []
    invalid = []

    for email in emails:
        if EMAIL_RE.match(email):
            valid.append(email)
        else:
            invalid.append(email)

    return valid, invalid


def find_cross_duplicates(to_list, cc_list, bcc_list):
    to_set = {x.lower() for x in to_list}
    cc_set = {x.lower() for x in cc_list}
    bcc_set = {x.lower() for x in bcc_list}

    duplicates = set()
    duplicates |= to_set & cc_set
    duplicates |= to_set & bcc_set
    duplicates |= cc_set & bcc_set

    return sorted(duplicates)