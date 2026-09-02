"""Fix mkrelease_body.py: handle escaped brackets and find prev tag."""
import re
import sys


def get_previous_tag(tags, current_tag):
    """Find the previous tag alphabetically."""
    # Sort tags: v2.0.5 < v2.0.4 < ... (lexicographic reverse)
    sorted_tags = sorted(tags, reverse=True)
    try:
        idx = sorted_tags.index(current_tag)
        if idx + 1 < len(sorted_tags):
            return sorted_tags[idx + 1]
    except ValueError:
        pass
    return None


def main():
    tag = sys.argv[1]
    changelog = open("CHANGELOG.md", encoding="utf-8").read()

    # Try matching with escaped brackets first (## \[v2.0.5]), then unescaped
    pattern_escaped = r"^## \\\[" + re.escape(tag) + r"\](.*?)(?:^## |\Z)"
    pattern_unescaped = r"^## \[" + re.escape(tag) + r"\](.*?)(?:^## |\Z)"
    m = re.search(pattern_escaped, changelog, re.M | re.S)
    if not m:
        m = re.search(pattern_unescaped, changelog, re.M | re.S)

    body = m.group(1).strip() if m else ""
    body += "\n\n**Full Changelog**: https://github.com/leop017/DocConvert/compare/" + tag + "...HEAD"
    print(body)


if __name__ == "__main__":
    main()
