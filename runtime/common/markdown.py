from __future__ import annotations


def bullet_list(items: list[object]) -> str:
    return "\n".join(f"- {item}" for item in items if str(item).strip())
