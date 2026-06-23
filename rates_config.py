from datetime import date

RATES = {
    "pre_jul2026": {
        "label":             "Vigente hasta 30 Jun 2026",
        "valid_from":        date(2000, 1, 1),
        "valid_until":       date(2026, 6, 30),
        "bp_margin":         3.0,
        "ingram_margin":     2.1,
        "cvr_renew_bp":      2.0,
        "cvr_renew_ingram":  2.0,
        "cvr_ontime_bp":     2.0,
        "cvr_ontime_ingram": 1.0,
    },
    "jul2026": {
        "label":             "Vigente desde 1 Jul 2026",
        "valid_from":        date(2026, 7, 1),
        "valid_until":       date(9999, 12, 31),
        "bp_margin":         2.0,
        "ingram_margin":     1.0,
        "cvr_renew_bp":      2.0,
        "cvr_renew_ingram":  1.0,
        "cvr_ontime_bp":     2.0,
        "cvr_ontime_ingram": 1.0,
    },
}

def get_rates(reference_date: date = None) -> dict:
    if reference_date is None:
        reference_date = date.today()
    for key in sorted(RATES, key=lambda k: RATES[k]["valid_from"], reverse=True):
        r = RATES[key]
        if r["valid_from"] <= reference_date <= r["valid_until"]:
            return {**r, "rate_key": key}
    last = sorted(RATES, key=lambda k: RATES[k]["valid_from"])[-1]
    return {**RATES[last], "rate_key": last}
