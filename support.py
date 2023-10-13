import json


def load_json_data(name, default=None):
    with open(name) as f:
        out = json.load(f)
        if out:
            return out
    return default


def save_result(name, results):
    with open(name, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=4)
