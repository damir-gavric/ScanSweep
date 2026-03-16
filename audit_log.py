from collections import defaultdict
from datetime import datetime
from pathlib import Path


class AuditLog:
    def __init__(self, src, dst, profile_name, quote_language, output_format, options):
        self.src = src
        self.dst = dst
        self.profile_name = profile_name
        self.quote_language = quote_language
        self.output_format = output_format
        self.options = options
        self.created_at = datetime.now().isoformat(timespec="seconds")
        self.stats = defaultdict(int)
        self.changes = defaultdict(list)
        self.notes = []

    def bump(self, key, amount=1):
        self.stats[key] += amount

    def add_note(self, message):
        self.notes.append(message)

    def record_change(self, category, before, after, context=""):
        if before == after:
            return
        self.changes[category].append(
            {
                "before": before,
                "after": after,
                "context": context,
            }
        )
        self.bump(f"{category}_count")

    def save(self, path=None):
        target = Path(path) if path else Path(self.dst).with_suffix(".audit.md")
        target.write_text(self.to_markdown(), encoding="utf-8")
        return target

    def to_markdown(self):
        lines = [
            "# Audit Log",
            "",
            "## Metadata",
            f"- Time: `{self.created_at}`",
            f"- Input: `{self.src}`",
            f"- Output: `{self.dst}`",
            f"- Profile: `{self.profile_name}`",
            f"- Quote style: `{self.quote_language}`",
            f"- Output format: `{self.output_format}`",
            "",
            "## Enabled Options",
        ]

        for key, value in sorted(self.options.items()):
            lines.append(f"- {key}: `{value}`")

        lines.extend(["", "## Summary"])
        if self.stats:
            for key, value in sorted(self.stats.items()):
                lines.append(f"- {key}: `{value}`")
        else:
            lines.append("- No tracked changes")

        if self.notes:
            lines.extend(["", "## Notes"])
            for note in self.notes:
                lines.append(f"- {note}")

        if self.changes:
            lines.extend(["", "## Detailed Changes"])
            for category in sorted(self.changes):
                lines.extend(["", f"### {category}"])
                for index, change in enumerate(self.changes[category], start=1):
                    if change["context"]:
                        lines.append(f"{index}. Context: {change['context']}")
                    else:
                        lines.append(f"{index}.")
                    lines.append(f"   - Before: `{change['before']}`")
                    lines.append(f"   - After: `{change['after']}`")

        lines.append("")
        return "\n".join(lines)
