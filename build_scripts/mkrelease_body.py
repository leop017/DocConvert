#!/usr/bin/env python3
"""Generate GitHub release body from CHANGELOG.md for a given tag."""
import re
import sys


def main():
    tag = sys.argv[1]
    changelog = open("CHANGELOG.md", encoding="utf-8").read()
    pattern = r"^## \[" + re.escape(tag) + r"\](.*?)^## "
    m = re.search(pattern, changelog, re.M | re.S)
    body = m.group(1).strip() if m else ""
    body += "\n\n**Full Changelog**: https://github.com/leop017/DocConvert/compare/" + tag + "...HEAD"
    print(body)


if __name__ == "__main__":
    main()
