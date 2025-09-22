from fuzzywuzzy import process, fuzz
from helpers.linkProcessing import normalize

def best_match(query, index):
    candidates = list(index.keys())
    if not candidates:
        return None, None
    key, score = process.extractOne(normalize(query), candidates, scorer=fuzz.token_set_ratio)
    if score < 70:
        return None, None
    # Return the display name and the first launch spec (lnk/uwp)
    entry = index[key][0]
    return entry.get("display", key), entry
