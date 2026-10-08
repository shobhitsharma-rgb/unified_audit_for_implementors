"""The state filing statuses Uzio's onboarding API actually accepts.

This is a transcription of `parseStateFilingStatus(stateCode, stateFilingStatus)`
in the onboarding service:

    onboarding-service/src/main/java/com/uzio/onboarding/validator/
        StateTaxWithholdingValidator.java

That method is the only gate. `ADPStateTaxWithholdingValidator` calls it with the
employee's **works-in** state and writes the result straight back over the record:

    String parsed = parseStateFilingStatus(empRecord.getWorksInState(), status);
    empRecord.setStateFilingStatusDesc(parsed);

An unmatched value returns null and `validateStateFilingStatus` then fails the
row -- "Invalid State filing status: 'X' for state: Y". So a value we cannot find
here is a visible import failure for that employee, not a silent drop.

Matching in Java is `stateFilingStatus.toUpperCase().trim()` -- case-insensitive
and trimmed, but **punctuation-sensitive**. `is_accepted` reproduces that exactly.
`canonical_label` is deliberately looser (it also ignores punctuation) so the tool
can repair ADP spellings like "Married, but withhold at higher single rate" that
the API would otherwise reject over one comma.

Do NOT read `filing status_code.txt` as the acceptance list -- that file is the
Uzio UI's display list, and it is narrower than what the API takes. Iowa is the
clearest example: the UI does not offer "Single", but the API maps it to IA_OTHER.

The table is baked in rather than parsed at runtime because the Java lives outside
this repo and is not on every machine. That makes it a copy that can rot, which is
what `utils/check_state_filing_status.py` is for -- the same arrangement, and the
same reasoning, as `utils/check_job_titles.py`.

Transcribed 2026-09-08: 33 states, 163 labels, 125 distinct enums.
"""
import re

# States with no wage income tax at all. The API has no `case` for them and no
# filing status can be meaningful, so the tool leaves their values alone.
NO_SIT_STATES = frozenset({
    "AK", "FL", "NH", "NV", "SD", "TN", "TX", "WA", "WY",
})

# state -> ((label, enum), ...) in the Java's own order. Labels are stored in the
# Java's spelling, which is upper-case because it compares against
# `stateFilingStatus.toUpperCase()`.
ACCEPTED = {
    "AL": (
        ('SINGLE', 'AL_SINGLE'),
        ('MARRIED JOINT', 'AL_MARRIED'),
        ('MARRIED', 'AL_MARRIED'),
        ('HEAD OF FAMILY', 'AL_HEAD_OF_HOUSEHOLD'),
        ('MARRIED FILING SEPARATELY (MS)', 'AL_MARRIED_SEPARATELY'),
        ('MARRIED FILING SEPARATELY', 'AL_MARRIED_SEPARATELY'),
        ('NO PERSONAL EXEMPTION', 'AL_NO_PERSONAL_EXEMPTION'),
    ),
    "AZ": (
        ('2.0% OF GROSS TAXABLE', '2.0'),
        ('2.5% OF GROSS TAXABLE', '2.5'),
        ('.5% OF GROSS TAXABLE', '0.5'),
        ('3.5% OF GROSS TAXABLE', '3.5'),
    ),
    "CA": (
        ('HEAD OF HOUSEHOLD', 'CA_HEAD_OF_HOUSEHOLD'),
        ('SINGLE', 'CA_SINGLE'),
        ('SINGLE OR MARRIED W/2 INCOMES', 'CA_SINGLE'),
        ('SINGLE OR MARRIED (WITH TWO OR MORE INCOMES)', 'CA_SINGLE'),
        ('MARRIED (ONE INCOME)', 'CA_MARRIED'),
    ),
    "CO": (
        ('MARRIED FILING JOINTLY', 'CO_MARRIED_JOINTLY'),
        ('SINGLE', 'CO_SINGLE'),
        ('MARRIED', 'CO_MARRIED'),
        ('MARRIED BUT WITHHOLD AT HIGHER SINGLE RATE', 'CO_MARRIED_SINGLE_RATE'),
        ('SINGLE OR MARRIED FILING SEPARATELY', 'CO_SINGLE_OR_MARRIED_SEPARATELY'),
        ('HEAD OF HOUSEHOLD', 'CO_HEAD_OF_HOUSEHOLD'),
    ),
    "DC": (
        ('SINGLE', 'DC_SINGLE'),
        ('MARRIED/DOMESTIC PARTNERS FILING JOINTLY', 'DC_MARRIED_DP_JOINTLY'),
        ('MARRIED FILING SEPARATELY', 'DC_MARRIED_SEPARATELY'),
        ('HEAD OF HOUSEHOLD', 'DC_HEAD_OF_HOUSEHOLD'),
        ('MARRIED/DOMESTIC PARTNERS FILING SEPARATELY', 'DC_MARRIED_DP_SEPARATELY'),
    ),
    "DE": (
        ('MARRIED FILING JOINTLY', 'DE_MARRIED'),
        ('MARRIED', 'DE_MARRIED'),
        ('SINGLE', 'DE_SINGLE'),
        ('MARRIED BUT WITHHOLD AS SINGLE', 'DE_MARRIED_SINGLE_RATE'),
    ),
    "GA": (
        ('C. MARRIED FILING JOINT/1 EARNER', 'GA_MARRIED_JOINT_ONE_WORKING'),
        ('C. MARRIED FILING JOINT, ONE SPOUSE WORKING', 'GA_MARRIED_JOINT_ONE_WORKING'),
        ('MARRIED FILING JOINT ONE SPOUSE WORKING', 'GA_MARRIED_JOINT_ONE_WORKING'),
        ('', 'GA_SINGLE'),
        ('A. SINGLE', 'GA_SINGLE'),
        ('SINGLE', 'GA_SINGLE'),
        ('D. HEAD OF HOUSEHOLD', 'GA_HEAD_OF_HOUSEHOLD'),
        ('E. HEAD OF HOUSEHOLD', 'GA_HEAD_OF_HOUSEHOLD'),
        ('HEAD OF HOUSEHOLD', 'GA_HEAD_OF_HOUSEHOLD'),
        ('B. MARRIED FILING SEPARATE OR MARRIED FILING JOINT, BOTH SPOUSES WORKING', 'GA_SEPARATE_MARRIED_JOINT_BOTH_WORKING'),
        ('MARRIED FILING SEPARATE OR MARRIED FILING JOINT BOTH SPOUSES WORKING', 'GA_SEPARATE_MARRIED_JOINT_BOTH_WORKING'),
        ('B. MARRIED FILING JOINT/2 EARNERS', 'GA_SEPARATE_MARRIED_JOINT_BOTH_WORKING'),
    ),
    "HI": (
        ('SINGLE', 'HI_SINGLE'),
        ('MARRIED BUT WITHHOLD AT HIGHER SINGLE RATE', 'HI_MARRIED_SINGLE_RATE'),
        ('MARRIED, BUT WITHHOLD AT HIGHER SINGLE RATE', 'HI_MARRIED_SINGLE_RATE'),
        ('MARRIED', 'HI_MARRIED'),
        ('CERTIFIED DISABLED PERSON', 'HI_DISABLED'),
        ('NONRESIDENT MILITARY SPOUSE', 'HI_NMS'),
    ),
    "IA": (
        ('OTHER', 'IA_OTHER'),
        ('SINGLE', 'IA_OTHER'),
        ('OTHER (INCLUDING SINGLE)', 'IA_OTHER'),
        ('MARRIED', 'IA_MARRIED_JOINTLY'),
        ('MARRIED FILING JOINTLY', 'IA_MARRIED_JOINTLY'),
        ('HEAD OF HOUSEHOLD', 'IA_HEAD_OF_HOUSEHOLD'),
        ('QUALIFYING SURVIVING SPOUSE', 'IA_QUALIFIED_SPOUSE'),
    ),
    "ID": (
        ('SINGLE', 'ID_SINGLE'),
        ('MARRIED', 'ID_MARRIED'),
        ('MARRIED BUT WITHHOLD AT HIGHER SINGLE RATE', 'ID_MARRIED_SINGLE_RATE'),
    ),
    "KS": (
        ('SINGLE', 'KS_SINGLE'),
        ('JOINT', 'KS_JOINT'),
    ),
    "MD": (
        ('MARRIED (SURVIVING SPOUSE OR UNMARRIED HEAD OF HOUSEHOLD)', 'MD_MARRIED'),
        ('MARRIED', 'MD_MARRIED'),
        ('SINGLE', 'MD_SINGLE'),
        ('MARRIED - TWO INCOMES', 'MD_MARRIED_SINGLE'),
        ('MARRIED, BUT WITHHOLD AT SINGLE RATE', 'MD_MARRIED_SINGLE'),
        ('MARRIED BUT WITHHOLD AT SINGLE RATE', 'MD_MARRIED_SINGLE'),
    ),
    "ME": (
        ('SINGLE OR HEAD OF HOUSEHOLD', 'ME_SINGLE'),
        ('MARRIED', 'ME_MARRIED'),
        ('MARRIED BUT WITHHOLD AT HIGHER SINGLE RATE', 'ME_MARRIED_SINGLE_RATE'),
        ('NONRESIDENT ALIEN', 'ME_NON_RESIDENT_ALIEN'),
    ),
    "NC": (
        ('SINGLE OR MARRIED FILING SEPARATELY', 'NC_SINGLE'),
        ('MARRIED FILING JOINTLY OR SURVIVING SPOUSE', 'NC_MARRIED'),
        ('HEAD OF HOUSEHOLD', 'NC_HEAD_OF_HOUSEHOLD'),
    ),
    "AR": (
        ('SINGLE', 'AR_SINGLE'),
        ('MARRIED FILING JOINTLY', 'AR_MARRIED_FILING_JOINTLY'),
        ('HEAD OF HOUSEHOLD', 'AR_HOH'),
    ),
    "NJ": (
        ('RATE C', 'NJ_SINGLE'),
        ('SINGLE', 'NJ_SINGLE'),
        ('MARRIED/CIVIL UNION COUPLE JOINT', 'NJ_MARRIED_DP_JOINTLY'),
        ('MARRIED/CIVIL UNION PARTNER SEPARATE', 'NJ_MARRIED_SEPARATELY'),
        ('HEAD OF HOUSEHOLD', 'NJ_HEAD_OF_HOUSEHOLD'),
        ('QUALIFYING WIDOW(ER)/SURVIVING CIVIL UNION PARTNER', 'NJ_QUALIFIED_WIDOW'),
    ),
    "NY": (
        ('SINGLE OR HEAD OF HOUSEHOLD', 'NY_SINGLE'),
        ('SINGLE', 'NY_SINGLE'),
        ('MARRIED', 'NY_MARRIED'),
        ('MARRIED, BUT WITHHOLD AT HIGHER SINGLE RATE', 'NY_MARRIED_WITHHOLD_SINGLE'),
        ('MARRIED BUT WITHHOLD AS SINGLE', 'NY_MARRIED_WITHHOLD_SINGLE'),
        ('HEAD OF HOUSEHOLD', 'NY_HEAD_OF_HOUSEHOLD'),
    ),
    "OK": (
        ('SINGLE', 'OK_SINGLE'),
        ('MARRIED, AT SINGLE RATE', 'OK_MARRIED_SINGLE_RATE'),
        ('MARRIED BUT WITHHOLD AS SINGLE', 'OK_MARRIED_SINGLE_RATE'),
        ('MARRIED', 'OK_MARRIED'),
        ('NON-RESIDENT ALIEN', 'OK_NRA'),
    ),
    "SC": (
        ('SINGLE', 'SC_SINGLE'),
        ('MARRIED', 'SC_MARRIED'),
        ('MARRIED, BUT AT A SINGLE RATE', 'SC_MARRIED_SINGLE_RATE'),
        ('MARRIED BUT WITHHOLD AT HIGHER SINGLE RATE', 'SC_MARRIED_SINGLE_RATE'),
    ),
    "CT": (
        ('A', 'A'),
        ('B', 'B'),
        ('C', 'C'),
        ('D', 'D'),
        ('E', 'E'),
        ('F', 'F'),
        ('NO FORM', 'NO_FORM'),
    ),
    "VT": (
        ('MARRIED/CIVIL UNION FILING JOINTLY', 'VT_MARRIED'),
        ('SINGLE', 'VT_SINGLE'),
        ('MARRIED/CIVIL UNION FILING SEPARATELY', 'VT_MARRIED_FILING_SEPERATELY'),
        ('MARRIED, BUT WITHHOLD AT HIGHER SINGLE RATE', 'VT_MARRIED_SINGLE_RATE'),
    ),
    "WV": (
        ('SINGLE', 'SINGLE_ONE_EXEMPTION'),
        ('SINGLE WITH ONE EXEMPTION', 'SINGLE_ONE_EXEMPTION'),
        ('MARRIED, AT SINGLE RATE', 'MARRIED_ONE_EXEMPTION'),
        ('MARRIED AT SINGLE RATE', 'MARRIED_ONE_EXEMPTION'),
        ('MARRIED WITH ONE EXEMPTION', 'MARRIED_ONE_EXEMPTION'),
        ('MARRIED WITH TWO EXEMPTIONS', 'MARRIED_TWO_EXEMPTIONS'),
        ('MARRIED', 'MARRIED_NO_EXEMPTION'),
        ('MARRIED WITH NO EXEMPTION', 'MARRIED_NO_EXEMPTION'),
    ),
    "WI": (
        ('SINGLE', 'WI_SINGLE'),
        ('MARRIED', 'WI_MARRIED'),
        ('MARRIED BUT WITHHOLD AT HIGHER SINGLE RATE', 'WI_MARRIED_SINGLE_RATE'),
    ),
    "MN": (
        ('MARRIED, BUT WITHHOLD AT HIGHER SINGLE RATE', 'MN_MARRIED_SINGLE_RATE'),
        ('MARRIED BUT WITHHOLD AT HIGHER SINGLE RATE', 'MN_MARRIED_SINGLE_RATE'),
        ('SINGLE; MARRIED, BUT LEGALLY SEPARATED; OR SPOUSE IS A NRA', 'MN_SINGLE'),
        ('SINGLE, MARRIED BUT LEGALLY SEPARATED OR SPOUSE IS A NONRESIDENT ALIEN', 'MN_SINGLE'),
        ('MARRIED', 'MN_MARRIED'),
    ),
    "MO": (
        ('SINGLE OR MARRIED SPOUSE WORKS OR MARRIED FILING SEPARATE', 'MO_SINGLE'),
        ('HEAD OF HOUSEHOLD', 'MO_HEAD_OF_HOUSEHOLD'),
        ('MARRIED (SPOUSE DOES NOT WORK)', 'MO_MARRIED'),
    ),
    "MS": (
        ('MARRIED - SPOUSE EMPLOYED', 'MS_M2'),
        ('MARRIED (SPOUSE IS EMPLOYED)', 'MS_M2'),
        ('MARRIED (SPOUSE NOT EMPLOYED)', 'MS_M1'),
        ('SINGLE', 'MS_SINGLE'),
        ('HEAD OF FAMILY', 'MS_HEAD_OF_HOUSEHOLD'),
    ),
    "MT": (
        ('REGULAR', 'MT_SINGLE'),
        ('SINGLE OR MARRIED FILING SEPARATELY', 'MT_SINGLE'),
        ('MARRIED FILING JOINTLY OR QUALIFYING WIDOWER', 'MT_MARRIED'),
        ('MARRIED FILING JOINTLY OR QUALIFYING SURVIVING SPOUSE', 'MT_MARRIED'),
        ('HEAD OF HOUSEHOLD', 'MT_HEAD_OF_HOUSEHOLD'),
    ),
    "NE": (
        ('MARRIED FILING JOINTLY OR QUALIFYING WIDOW(ER)', 'NE_MARRIED'),
        ('SINGLE', 'NE_SINGLE'),
        ('MARRIED, AT SINGLE RATE', 'NE_SINGLE'),
    ),
    "ND": (
        ('SINGLE', 'ND_SINGLE'),
        ('MARRIED', 'ND_MARRIED'),
        ('MARRIED BUT WITHHOLD AT HIGHER SINGLE RATE', 'ND_MARRIED_SINGLE_RATE'),
        ('SINGLE OR MARRIED FILING SEPARATELY', 'ND_SINGLE_MARRIED_SEPARATELY'),
        ('HEAD OF HOUSEHOLD', 'ND_HEAD_OF_HOUSEHOLD'),
        ('MARRIED FILING JOINTLY OR QUALIFYING SURVIVING SPOUSE', 'ND_MARRIED_JOINTLY'),
    ),
    "LA": (
        ('NO DEDUCTION', 'LA_NO_DEDUCTION'),
        ('SINGLE OR MARRIED FILING SEPARATELY', 'LA_SINGLE_OR_MARRIED'),
        ('MARRIED FILING JOINTLY, QUALIFYING SURVIVING SPOUSE, OR HEAD OF HOUSEHOLD', 'LA_MARRIED_FILING_JOINTLY_HOH'),
    ),
    "OR": (
        ('SINGLE', 'OR_SINGLE'),
        ('MARRIED', 'OR_MARRIED'),
        ('MARRIED BUT WITHHOLD AT HIGHER SINGLE RATE', 'OR_MARRIED_SINGLE_RATE'),
    ),
    "UT": (
        ('SINGLE OR MARRIED FILING SEPARATELY', 'UT_SINGLE'),
        ('MARRIED FILING JOINTLY OR QUALIFYING WIDOW(ER)', 'UT_MARRIED'),
        ('HEAD OF HOUSEHOLD', 'UT_HEAD_OF_HOUSEHOLD'),
    ),
    "NM": (
        ('MARRIED', 'NM_MARRIED'),
        ('MARRIED FILING JOINTLY OR QUALIFYING SURVIVING SPOUSE', 'NM_MARRIED'),
        ('SINGLE', 'NM_SINGLE'),
        ('SINGLE OR MARRIED FILING SEPARATELY', 'NM_SINGLE'),
        ('MARRIED BUT WITHHOLD AS SINGLE', 'NM_MARRIED_SINGLE'),
        ('MARRIED BUT WITHHOLD AT HIGHER SINGLE RATE', 'NM_MARRIED_SINGLE'),
        ('HEAD OF HOUSEHOLD', 'NM_HEAD_OF_HOUSEHOLD'),
    ),
}


def _depunct(value):
    """Upper-case, strip every non-alphanumeric run to a single space."""
    return re.sub(r"[^A-Z0-9]+", " ", str(value).upper()).strip()


# Punctuation-insensitive index, built once. First spelling wins, so a repair
# always lands on the label the Java lists first. Several states do carry two
# spellings that collapse together ("MARRIED, AT SINGLE RATE" / "MARRIED AT
# SINGLE RATE"), but both sides map to the same enum, so the choice is cosmetic.
# The checker fails only when collapsing labels map to DIFFERENT enums, which
# would make a repair a coin flip.
def _build_loose():
    out = {}
    for state, pairs in ACCEPTED.items():
        index = {}
        for label, _ in pairs:
            if label:
                index.setdefault(_depunct(label), label)
        out[state] = index
    return out


_LOOSE = _build_loose()


def has_table(state):
    """True when the API has an acceptance list for this state."""
    return _norm_state(state) in ACCEPTED


def is_no_sit(state):
    return _norm_state(state) in NO_SIT_STATES


def accepted_labels(state):
    """The labels to offer in a dropdown, in the API's own order.

    The empty label is left out: Georgia maps "" to GA_SINGLE, which is a reason
    to leave GA blanks alone, not something to offer as a choice.
    """
    return [label for label, _ in ACCEPTED.get(_norm_state(state), ()) if label]


def is_accepted(state, label):
    """Exactly what the API does: upper-case, trim, exact lookup."""
    pairs = ACCEPTED.get(_norm_state(state))
    if pairs is None:
        return False
    key = str(label or "").upper().strip()
    return any(key == lbl for lbl, _ in pairs)


def canonical_label(state, label):
    """The API's spelling of `label` when only punctuation differs, else None.

    Returns None when the value already matches exactly -- callers want to know
    "does this need rewriting", and an exact match does not.
    """
    state = _norm_state(state)
    if state not in ACCEPTED or is_accepted(state, label):
        return None
    return _LOOSE[state].get(_depunct(label))


def enum_for(state, label):
    """The enum the API would store, or None."""
    key = str(label or "").upper().strip()
    for lbl, enum in ACCEPTED.get(_norm_state(state), ()):
        if key == lbl:
            return enum
    return None


def _norm_state(state):
    return str(state or "").strip().upper()
