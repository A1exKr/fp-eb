from sqlalchemy.orm import Session

from app import repositories


def _load_projects(db: Session) -> list[dict]:
    return repositories.list_reference_projects(db)


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

    p_type = (parsed.get("project", {}).get("project_type") or "").lower()
    location = (parsed.get("project", {}).get("location") or "").lower()
    keywords = set((parsed.get("keywords") or []))
    keywords = {word.lower() for word in keywords}

    scored: list[tuple[int, dict]] = []
    for project in projects:
        score = 0
        project_type = (project.get("project_type") or "").lower()
        project_location = (project.get("location") or "").lower()
        project_keywords = {word.lower() for word in project.get("keywords", [])}

        if p_type and p_type == project_type:
            score += 5
        if location and location in project_location:
            score += 3
        score += len(keywords.intersection(project_keywords))
        scored.append((score, project))

    scored.sort(key=lambda item: item[0], reverse=True)
    return [item[1] for item in scored[:limit]]
