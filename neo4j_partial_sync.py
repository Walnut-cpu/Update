from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
from typing import Any

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_WORKBOOK_PATH = BASE_DIR / "FundiGraph.xlsx"

DEFAULT_NEO4J_URL = "http://localhost:7474"
DEFAULT_NEO4J_USER = "neo4j"
DEFAULT_NEO4J_PASSWORD = "Yk1213903395!"

NODE_LABEL_BY_COLUMN = {
    "First": "First",
    "Second": "Second",
    "Third": "Third",
    "Forth": "Forth",
    "Anatomical_location": "Anatomical_location",
    "Examination": "Examination",
    "OCT_sign": "OCT_sign",
    "Symptom": "Symptom",
    "Physical_sign": "Physical_sign",
    "Gene": "Gene",
    "Differential_diagnosis": "Differential_diagnosis",
    "Complication": "Complication",
    "Etiology": "Etiology",
    "Treatment_general": "Treatment_general",
    "Treatment_drug": "Treatment_drug",
    "Treatment_usage": "Treatment_usage",
    "Treatment_surgery": "Treatment_surgery",
    "Treatment_indications": "Treatment_indications",
    "Treatment_contraindications": "Treatment_contraindications",
    "Related_disease": "Related_disease",
    "Staging_typing": "Staging_typing",
    "Synonym_3": "Synonym_3",
    "Synonym_4": "Synonym_4",
    "Age_of_onset": "Age_of_onset",
    "High_risk_population": "High_risk_population",
    "Medical_history": "Medical_history",
}

DIRECT_RELATION_SPECS = [
    ("Anatomical_location", "Located in", "Anatomical_location"),
    ("Examination", "Requires examination", "Examination"),
    ("OCT_sign", "Has oct sign", "OCT_sign"),
    ("Gene", "Related gene", "Gene"),
    ("Symptom", "Has symptom", "Symptom"),
    ("Physical_sign", "Has physical sign", "Physical_sign"),
    ("Differential_diagnosis", "Needs distinguished from", "Differential_diagnosis"),
    ("Complication", "May cause", "Complication"),
    ("Etiology", "Caused by", "Etiology"),
    ("Related_disease", "Related to", "Related_disease"),
    ("Age_of_onset", "Onset during", "Age_of_onset"),
    ("High_risk_population", "Affects population", "High_risk_population"),
    ("Medical_history", "Related history", "Medical_history"),
]

TREATMENT_COLUMNS = [
    "Treatment_general",
    "Treatment_drug",
    "Treatment_usage",
    "Treatment_surgery",
    "Treatment_indications",
    "Treatment_contraindications",
]

ROW_MATCH_COLUMNS = ["First", "Second", "Third", "Forth", "Synonym_3", "Synonym_4", "Staging_typing"]
EMPTY_MARKERS = {"", "entity", "entity class", "Entity", "Entity Class", "nan", "None"}


def clean_value(value: Any) -> str | None:
    if value is None:
        return None
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text or text in EMPTY_MARKERS:
        return None
    return text


def normalize_text(value: Any) -> str:
    text = clean_value(value)
    return text.casefold() if text else ""


def escape_identifier(name: str) -> str:
    return name.replace("`", "``")


def load_neo4j_settings() -> dict[str, str]:
    return {
        "url": os.getenv("NEO4J_URL", "").strip() or DEFAULT_NEO4J_URL,
        "user": os.getenv("NEO4J_USER", "").strip() or DEFAULT_NEO4J_USER,
        "password": os.getenv("NEO4J_PASSWORD", "").strip() or DEFAULT_NEO4J_PASSWORD,
    }


def load_graph():
    try:
        from py2neo import Graph
    except ImportError as exc:
        raise RuntimeError("py2neo is not installed in the current Python environment.") from exc

    settings = load_neo4j_settings()
    graph = Graph(settings["url"], auth=(settings["user"], settings["password"]))
    graph.run("RETURN 1").evaluate()
    return graph, {"url": settings["url"], "user": settings["user"]}


def read_workbook_dataframe(workbook_path: Path) -> pd.DataFrame:
    df = pd.read_excel(workbook_path)
    return df.applymap(clean_value)


def ensure_expected_columns(df: pd.DataFrame) -> None:
    missing = [column for column in NODE_LABEL_BY_COLUMN if column not in df.columns]
    missing += [column for _target, column, _label in DIRECT_RELATION_SPECS if column not in df.columns]
    for column in ["Contain", "Classified as", "Same as", "Treated with"]:
        if column not in df.columns:
            missing.append(column)
    if missing:
        raise RuntimeError(f"Workbook is missing expected columns: {sorted(set(missing))}")


def merge_node(graph, label: str, name: str) -> None:
    query = (
        f"MERGE (n:`{escape_identifier(label)}` {{name: $name}}) "
        "ON CREATE SET n.label = $label "
        "ON MATCH SET n.label = coalesce(n.label, $label)"
    )
    graph.run(query, name=name, label=label)


def merge_relationship(
    graph,
    source_label: str,
    source_name: str,
    rel_type: str,
    target_label: str,
    target_name: str,
) -> bool:
    if not all([source_label, source_name, rel_type, target_label, target_name]):
        return False

    merge_node(graph, source_label, source_name)
    merge_node(graph, target_label, target_name)
    query = (
        f"MATCH (s:`{escape_identifier(source_label)}` {{name: $source_name}}) "
        f"MATCH (t:`{escape_identifier(target_label)}` {{name: $target_name}}) "
        f"MERGE (s)-[r:`{escape_identifier(rel_type)}`]->(t)"
    )
    graph.run(query, source_name=source_name, target_name=target_name)
    return True


def sync_row_record(graph, row_number: int, record: dict[str, Any]) -> dict[str, Any]:
    cleaned = {key: clean_value(value) for key, value in record.items()}
    nodes_merged: list[dict[str, str]] = []
    relationships_merged: list[dict[str, str]] = []

    for column, label in NODE_LABEL_BY_COLUMN.items():
        value = cleaned.get(column)
        if not value:
            continue
        merge_node(graph, label, value)
        nodes_merged.append({"label": label, "name": value})

    contain_type = cleaned.get("Contain")
    if contain_type:
        if cleaned.get("First") and cleaned.get("Second"):
            if merge_relationship(graph, "First", cleaned["First"], contain_type, "Second", cleaned["Second"]):
                relationships_merged.append({"source": cleaned["First"], "relation": contain_type, "target": cleaned["Second"]})
        if cleaned.get("Second") and cleaned.get("Third"):
            if merge_relationship(graph, "Second", cleaned["Second"], contain_type, "Third", cleaned["Third"]):
                relationships_merged.append({"source": cleaned["Second"], "relation": contain_type, "target": cleaned["Third"]})
        elif cleaned.get("First") and cleaned.get("Third"):
            if merge_relationship(graph, "First", cleaned["First"], contain_type, "Third", cleaned["Third"]):
                relationships_merged.append({"source": cleaned["First"], "relation": contain_type, "target": cleaned["Third"]})
        if cleaned.get("Third") and cleaned.get("Forth"):
            if merge_relationship(graph, "Third", cleaned["Third"], contain_type, "Forth", cleaned["Forth"]):
                relationships_merged.append({"source": cleaned["Third"], "relation": contain_type, "target": cleaned["Forth"]})

    target_1_label = "Forth" if cleaned.get("Forth") else "Third"
    target_1_name = cleaned.get("Forth") or cleaned.get("Third")

    synonym_type = cleaned.get("Same as")
    if synonym_type:
        if cleaned.get("Third") and cleaned.get("Synonym_3"):
            if merge_relationship(graph, "Third", cleaned["Third"], synonym_type, "Synonym_3", cleaned["Synonym_3"]):
                relationships_merged.append({"source": cleaned["Third"], "relation": synonym_type, "target": cleaned["Synonym_3"]})
        if cleaned.get("Forth") and cleaned.get("Synonym_4"):
            if merge_relationship(graph, "Forth", cleaned["Forth"], synonym_type, "Synonym_4", cleaned["Synonym_4"]):
                relationships_merged.append({"source": cleaned["Forth"], "relation": synonym_type, "target": cleaned["Synonym_4"]})

    target_2_label = target_1_label
    target_2_name = target_1_name
    staging_type = cleaned.get("Classified as")
    if target_1_name and cleaned.get("Staging_typing") and staging_type:
        if merge_relationship(graph, target_1_label, target_1_name, staging_type, "Staging_typing", cleaned["Staging_typing"]):
            relationships_merged.append({"source": target_1_name, "relation": staging_type, "target": cleaned["Staging_typing"]})
        target_2_label = "Staging_typing"
        target_2_name = cleaned["Staging_typing"]

    if target_2_name:
        for target_column, relation_column, target_label in DIRECT_RELATION_SPECS:
            target_value = cleaned.get(target_column)
            relation_value = cleaned.get(relation_column)
            if not target_value or not relation_value:
                continue
            if merge_relationship(graph, target_2_label, target_2_name, relation_value, target_label, target_value):
                relationships_merged.append({"source": target_2_name, "relation": relation_value, "target": target_value})

        treatment_relation = cleaned.get("Treated with")
        if treatment_relation:
            for treatment_column in TREATMENT_COLUMNS:
                treatment_value = cleaned.get(treatment_column)
                if not treatment_value:
                    continue
                if merge_relationship(graph, target_2_label, target_2_name, treatment_relation, treatment_column, treatment_value):
                    relationships_merged.append({"source": target_2_name, "relation": treatment_relation, "target": treatment_value})

    return {
        "row_number": row_number,
        "status": "synced",
        "node_count": len(nodes_merged),
        "relationship_count": len(relationships_merged),
        "nodes": nodes_merged,
        "relationships": relationships_merged,
        "anchor": {
            "First": cleaned.get("First"),
            "Second": cleaned.get("Second"),
            "Third": cleaned.get("Third"),
            "Forth": cleaned.get("Forth"),
            "Staging_typing": cleaned.get("Staging_typing"),
        },
    }


def filter_dataframe(
    df: pd.DataFrame,
    row_numbers: list[int] | None = None,
    disease_names: list[str] | None = None,
) -> tuple[pd.DataFrame, list[int]]:
    if row_numbers:
        valid_pairs = []
        for row_number in row_numbers:
            dataframe_index = row_number - 2
            if 0 <= dataframe_index < len(df):
                valid_pairs.append((row_number, dataframe_index))
        if not valid_pairs:
            return df.iloc[0:0], []
        ordered_indices = [pair[1] for pair in valid_pairs]
        ordered_rows = [pair[0] for pair in valid_pairs]
        return df.iloc[ordered_indices].copy(), ordered_rows

    selected_df = df
    if disease_names:
        wanted = {normalize_text(name) for name in disease_names if clean_value(name)}
        mask = pd.Series([False] * len(df))
        for column in ROW_MATCH_COLUMNS:
            if column not in df.columns:
                continue
            mask = mask | df[column].apply(lambda value: normalize_text(value) in wanted)
        selected_df = df[mask].copy()

    row_list = [index + 2 for index in selected_df.index.tolist()]
    return selected_df, row_list


def sync_workbook_rows(
    workbook_path: Path | str = DEFAULT_WORKBOOK_PATH,
    row_numbers: list[int] | None = None,
    disease_names: list[str] | None = None,
) -> dict[str, Any]:
    workbook = Path(workbook_path)
    df = read_workbook_dataframe(workbook)
    ensure_expected_columns(df)
    selected_df, resolved_rows = filter_dataframe(df, row_numbers=row_numbers, disease_names=disease_names)

    if selected_df.empty:
        return {
            "status": "skipped",
            "reason": "no rows matched the requested partial sync scope",
            "workbook": str(workbook),
            "requested_rows": row_numbers or [],
            "requested_diseases": disease_names or [],
            "row_count": 0,
            "rows": [],
        }

    graph, connection_info = load_graph()
    row_results = []
    for row_number, (_, series) in zip(resolved_rows, selected_df.iterrows()):
        row_results.append(sync_row_record(graph, row_number, series.to_dict()))

    return {
        "status": "ok",
        "connection": connection_info,
        "workbook": str(workbook),
        "requested_rows": row_numbers or [],
        "requested_diseases": disease_names or [],
        "row_count": len(row_results),
        "synced_rows": resolved_rows,
        "node_count": sum(item["node_count"] for item in row_results),
        "relationship_count": sum(item["relationship_count"] for item in row_results),
        "rows": row_results,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Partially sync selected FundiGraph workbook rows into Neo4j.")
    parser.add_argument("--workbook", default=str(DEFAULT_WORKBOOK_PATH))
    parser.add_argument("--rows", nargs="*", type=int, help="Excel row numbers to sync, e.g. --rows 125 126")
    parser.add_argument(
        "--disease",
        nargs="*",
        help="Disease names to sync by matching First/Second/Third/Forth/Synonym/Staging columns.",
    )
    args = parser.parse_args()

    result = sync_workbook_rows(
        workbook_path=Path(args.workbook),
        row_numbers=args.rows or None,
        disease_names=args.disease or None,
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
