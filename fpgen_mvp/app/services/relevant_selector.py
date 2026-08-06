import re

from sqlalchemy.orm import Session

from app import repositories

_STOPWORDS = {
    "and", "the", "for", "with", "our", "its", "this", "that", "from", "into",
    "project", "projects", "design", "plan", "phase", "new", "tbd",
}


def _load_projects(db: Session) -> list[dict]:
    return repositories.list_reference_projects(db)


def _tokens(*values) -> set[str]:
    """Normalize free text and keyword lists to a comparable token set."""
    out: set[str] = set()
    for value in values:
        items = value if isinstance(value, (list, tuple, set)) else [value]
        for item in items:
            for token in re.split(r"\W+", str(item or "").lower()):
                if len(token) > 2 and token not in _STOPWORDS:
                    out.add(token)
    return out


def select_relevant_projects(
    db: Session, parsed: dict, selected_ids: list[str], limit: int = 3
) -> list[dict]:
    projects = _load_projects(db)
    if not projects:
        return []

    if selected_ids:
        selected_set = set(selected_ids)
        chosen = [project for project in projects if project.get("id") in selected_set]
        return chosen[:limit]

    project = parsed.get("project", {})
    # The pipeline writes `type`; `project_type` is the legacy key.
    p_type = (project.get("type") or project.get("project_type") or "").lower().strip()
    location = (project.get("location") or "").lower().strip()
    p_type_tokens = _tokens(p_type)
    keywords = _tokens(parsed.get("keywords") or [])

    scored: list[tuple[int, dict]] = []
    for candidate in projects:
        score = 0
        candidate_type = (candidate.get("project_type") or "").lower().strip()
        candidate_location = (candidate.get("location") or "").lower().strip()
        candidate_keywords = _tokens(candidate.get("keywords", []), candidate.get("name", ""))

        if p_type and candidate_type:
            if p_type == candidate_type:
                score += 5
            elif p_type_tokens & _tokens(candidate_type):
                score += 3
        if location and location != "tbd" and location in candidate_location:
            score += 3
        score += min(len(keywords & candidate_keywords), 5)
        scored.append((score, candidate))

    scored.sort(key=lambda item: item[0], reverse=True)
    return [item[1] for item in scored[:limit]]
