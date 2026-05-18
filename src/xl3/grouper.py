from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Literal

from .value_model import canonical_string, is_empty


@dataclass
class GroupNode:
    key_value: str
    children: list[GroupNode] = field(default_factory=list)
    rows: list[dict[str, Any]] = field(default_factory=list)


def partition_by_group_keys(
    rows: list[dict[str, Any]], keys: list[str]
) -> list[GroupNode]:
    if not keys:
        return []

    top: list[GroupNode] = []
    top_index: dict[str, GroupNode] = {}
    child_indexes: dict[int, dict[str, GroupNode]] = {}

    for row in rows:
        parent_list = top
        parent_index = top_index
        node: GroupNode | None = None
        for key in keys:
            value = canonical_string(row.get(key))
            next_node = parent_index.get(value)
            if next_node is None:
                next_node = GroupNode(key_value=value)
                parent_list.append(next_node)
                parent_index[value] = next_node
            node = next_node
            child_index = child_indexes.setdefault(id(next_node), {})
            parent_list = next_node.children
            parent_index = child_index
        assert node is not None
        node.rows.append(row)

    _prune_empty_leaves(top)
    return top


def _prune_empty_leaves(nodes: list[GroupNode]) -> None:
    for i in range(len(nodes) - 1, -1, -1):
        node = nodes[i]
        if node.children:
            _prune_empty_leaves(node.children)
            if not node.children:
                del nodes[i]
        elif all(all(is_empty(v) for v in row.values()) for row in node.rows):
            del nodes[i]


@dataclass
class EmitEvent:
    kind: Literal["data", "subtotal"]
    row: dict[str, Any] | None = None
    level: int = 0
    group_rows: list[dict[str, Any]] = field(default_factory=list)


def plan_emission_events(tree: list[GroupNode], key_count: int) -> list[EmitEvent]:
    events: list[EmitEvent] = []
    _walk_and_emit(tree, 0, key_count, events)
    return events


def _walk_and_emit(
    nodes: list[GroupNode],
    depth: int,
    key_count: int,
    events: list[EmitEvent],
    out_rows: list[dict[str, Any]] | None = None,
) -> None:
    for node in nodes:
        if node.children:
            group_rows: list[dict[str, Any]] = []
            _walk_and_emit(node.children, depth + 1, key_count, events, group_rows)
        else:
            group_rows = node.rows
            for row in node.rows:
                events.append(EmitEvent(kind="data", row=row))
        events.append(
            EmitEvent(kind="subtotal", level=key_count - depth, group_rows=group_rows)
        )
        if out_rows is not None:
            out_rows.extend(group_rows)
