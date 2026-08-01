# -------------------------------------------------------
# Audit Panel Authentication
# -------------------------------------------------------

AUDIT_PASSWORD = "audit123"


def verify_audit_password(password):
    """
    Check whether the entered audit password is correct.
    """

    return password == AUDIT_PASSWORD