from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, List, Tuple

from pptx.enum.shapes import MSO_SHAPE_TYPE


@dataclass
class TableColumnMap:
    shape: object
    key_to_index: Dict[str, int]
    removed_cols: List[int]
    group_to_index: Dict[str, int]
    image_slots: Dict[str, Tuple[int, int, int, int]]


COL_MARKER_RE = re.compile(r"{{COL[:_](?P<key>[A-Z_]+)}}")
IMG_MARKER_RE = re.compile(r"{{IMG[:_](?P<key>[A-Z_]+)}}")
SHAPE_COL_MARKER_RE = re.compile(r"COL_(?P<key>[A-Z_]+)")
SHAPE_IMG_MARKER_RE = re.compile(r"IMG_(?P<key>[A-Z_]+)")
SHAPE_KEY_SUFFIXES = ("_ARROW", "_FLECHA", "_ICON")
STACK_MARKER_RE = re.compile(
    r"STACK_(?P<group>\\d+)(?:_(?P<item>\\d+))?(?:_GAP(?P<gap>\\d+))?",
    re.IGNORECASE,
)
STACK_GAP = 120_000

GROUP_LABELS = {
    "LACTEOS": "LACTEOS",
    "VEGETALES": "VEGETALES",
    "FRUTAS": "FRUTAS",
    "ALMIDONES": "ALMIDONES",
    "PROTEINAS": "PROTEINAS",
    "PROTEICOS": "PROTEINAS",
    "GRASAS": "GRASAS",
}


def normalize_label(text: str) -> str:
    normalized = text.upper()
    for src, dst in (
        ("Á", "A"),
        ("É", "E"),
        ("Í", "I"),
        ("Ó", "O"),
        ("Ú", "U"),
        ("Ü", "U"),
    ):
        normalized = normalized.replace(src, dst)
    return " ".join(normalized.split())


def normalize_shape_key(key: str) -> str:
    for suffix in SHAPE_KEY_SUFFIXES:
        if key.endswith(suffix):
            return key[: -len(suffix)]
    return key


def key_to_group(key: str) -> str:
    if "_" in key:
        return key.split("_", 1)[1]
    return key


def replace_in_text(text: str, replacements: Dict[str, str]) -> str:
    updated = text
    for key, value in replacements.items():
        if key in updated:
            updated = updated.replace(key, value)
    return updated


def replace_in_text_frame(text_frame, replacements: Dict[str, str]) -> None:
    for paragraph in text_frame.paragraphs:
        if not paragraph.runs:
            continue
        original = "".join(run.text for run in paragraph.runs)
        if not original:
            continue
        updated = replace_in_text(original, replacements)
        if updated == original:
            continue
        paragraph.runs[0].text = updated
        for run in paragraph.runs[1:]:
            run.text = ""


def remove_empty_paragraphs(text_frame) -> None:
    to_remove = []
    for paragraph in text_frame.paragraphs:
        text = "".join(run.text for run in paragraph.runs).strip()
        if not text:
            to_remove.append(paragraph._element)
    for element in to_remove:
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)


def strip_col_markers(text_frame) -> None:
    for paragraph in text_frame.paragraphs:
        if not paragraph.runs:
            continue
        original = "".join(run.text for run in paragraph.runs)
        if not original:
            continue
        updated = COL_MARKER_RE.sub("", original)
        updated = " ".join(updated.split())
        if updated == original:
            continue
        paragraph.runs[0].text = updated
        for run in paragraph.runs[1:]:
            run.text = ""
    remove_empty_paragraphs(text_frame)


def strip_img_markers(text_frame) -> None:
    for paragraph in text_frame.paragraphs:
        if not paragraph.runs:
            continue
        original = "".join(run.text for run in paragraph.runs)
        if not original:
            continue
        updated = IMG_MARKER_RE.sub("", original)
        updated = " ".join(updated.split())
        if updated == original:
            continue
        paragraph.runs[0].text = updated
        for run in paragraph.runs[1:]:
            run.text = ""
    remove_empty_paragraphs(text_frame)


def should_hide_shape(shape, placeholder_values: Dict[str, int]) -> bool:
    name = getattr(shape, "name", "") or ""
    for key in SHAPE_COL_MARKER_RE.findall(name) + SHAPE_IMG_MARKER_RE.findall(name):
        normalized = normalize_shape_key(key)
        placeholder = f"{{{{{normalized}}}}}"
        if placeholder_values.get(placeholder, 0) == 0:
            return True
    return False


def remove_shape(shape) -> None:
    element = shape._element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def remove_table_columns(table, col_indices: List[int]) -> None:
    tbl = table._tbl
    grid = tbl.tblGrid
    for col_idx in col_indices:
        grid_cols = getattr(grid, "gridCol_lst", None)
        if grid_cols is None:
            grid_cols = list(grid.iterchildren())
        if col_idx < 0 or col_idx >= len(grid_cols):
            continue
        grid.remove(grid_cols[col_idx])
        for row in table.rows:
            cells = getattr(row._tr, "tc_lst", None)
            if cells is None:
                cells = list(row._tr.iterchildren())
            if col_idx < len(cells):
                row._tr.remove(cells[col_idx])


def remove_empty_table_rows(table, shape) -> None:
    tbl = table._tbl
    rows = list(table.rows)
    rows_to_remove = []
    for idx, row in enumerate(rows):
        if idx == 0:
            continue
        if all(not cell.text.strip() for cell in row.cells):
            rows_to_remove.append(idx)
    if not rows_to_remove:
        return
    tr_list = getattr(tbl, "tr_lst", None)
    if tr_list is None:
        tr_list = list(tbl.iterchildren())
    for idx in sorted(rows_to_remove, reverse=True):
        if idx < len(tr_list):
            tbl.remove(tr_list[idx])
    shape.height = sum(row.height for row in table.rows)


def find_image_slots(table, shape) -> Dict[str, Tuple[int, int, int, int]]:
    slots: Dict[str, Tuple[int, int, int, int]] = {}
    if not table.rows or not table.columns:
        return slots

    widths = [col.width for col in table.columns]
    heights = [row.height for row in table.rows]
    if any(h is None for h in heights):
        avg_height = int(shape.height / len(heights))
        heights = [h if h is not None else avg_height for h in heights]

    col_lefts = []
    acc = shape.left
    for width in widths:
        col_lefts.append(acc)
        acc += width

    row_tops = []
    acc = shape.top
    for height in heights:
        row_tops.append(acc)
        acc += height

    for row_idx, row in enumerate(table.rows):
        for col_idx, cell in enumerate(row.cells):
            text = cell.text
            markers = IMG_MARKER_RE.findall(text)
            if not markers:
                continue
            for key in markers:
                left = col_lefts[col_idx]
                top = row_tops[row_idx]
                slots[key] = (left, top, widths[col_idx], heights[row_idx])
            strip_img_markers(cell.text_frame)

    return slots


def replace_in_shape(
    shape,
    replacements: Dict[str, str],
    placeholder_values: Dict[str, int],
    slide_width: int,
    slide_shapes,
    table_maps: List[TableColumnMap],
) -> None:
    if should_hide_shape(shape, placeholder_values):
        remove_shape(shape)
        return

    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for subshape in shape.shapes:
            replace_in_shape(
                subshape,
                replacements,
                placeholder_values,
                slide_width,
                slide_shapes,
                table_maps,
            )
        return

    if shape.has_table:
        table = shape.table
        cols_to_hide = set()
        key_to_index: Dict[str, int] = {}
        group_to_index: Dict[str, int] = {}
        for col_idx in range(len(table.columns)):
            for row in table.rows:
                cell_text = row.cells[col_idx].text
                markers = COL_MARKER_RE.findall(cell_text)
                if markers:
                    for key in markers:
                        key_to_index.setdefault(key, col_idx)
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder_values.get(placeholder, 0) == 0:
                            cols_to_hide.add(col_idx)
                    strip_col_markers(row.cells[col_idx].text_frame)
                for placeholder, value in placeholder_values.items():
                    if placeholder in cell_text:
                        key_to_index.setdefault(
                            placeholder.strip("{}"), col_idx)
                        if value == 0:
                            cols_to_hide.add(col_idx)
            if table.rows:
                header_text = table.rows[0].cells[col_idx].text
                normalized_header = normalize_label(header_text)
                for label, group in GROUP_LABELS.items():
                    if label and label in normalized_header:
                        group_to_index.setdefault(group, col_idx)

        if cols_to_hide:
            remove_table_columns(table, sorted(cols_to_hide, reverse=True))
            if slide_width and shape._parent is slide_shapes:
                table_width = sum(col.width for col in table.columns)
                shape.width = table_width
                shape.left = int((slide_width - shape.width) / 2)
            group_to_index = {
                group: idx -
                sum(1 for removed in cols_to_hide if removed < idx)
                for group, idx in group_to_index.items()
            }

        remove_empty_table_rows(table, shape)
        image_slots = find_image_slots(table, shape)
        table_maps.append(
            TableColumnMap(
                shape=shape,
                key_to_index=key_to_index,
                removed_cols=sorted(cols_to_hide),
                group_to_index=group_to_index,
                image_slots=image_slots,
            )
        )

        for row in table.rows:
            for cell in row.cells:
                replace_in_text_frame(cell.text_frame, replacements)
        return

    if shape.has_text_frame:
        replace_in_text_frame(shape.text_frame, replacements)


def text_frame_contains(text_frame, tokens: List[str]) -> bool:
    for paragraph in text_frame.paragraphs:
        if any(token in paragraph.text for token in tokens):
            return True
    return False


def shape_contains_tokens(shape, tokens: List[str]) -> bool:
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return any(shape_contains_tokens(subshape, tokens) for subshape in shape.shapes)
    if shape.has_text_frame and text_frame_contains(shape.text_frame, tokens):
        return True
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                if text_frame_contains(cell.text_frame, tokens):
                    return True
    return False


def iter_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for subshape in iter_shapes(shape.shapes):
                yield subshape


def align_marked_shapes(
    slide, placeholder_values: Dict[str, int], table_maps: List[TableColumnMap]
) -> None:
    if not table_maps:
        return

    table_centers: List[Tuple[TableColumnMap, Dict[str, int]]] = []
    for table_map in table_maps:
        table = table_map.shape.table
        widths = [col.width for col in table.columns]
        removed = table_map.removed_cols
        centers: Dict[str, int] = {}
        for key, original_idx in table_map.key_to_index.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder_values.get(placeholder, 0) == 0:
                continue
            shift = sum(1 for idx in removed if idx < original_idx)
            new_idx = original_idx - shift
            if new_idx < 0 or new_idx >= len(widths):
                continue
            left = table_map.shape.left + sum(widths[:new_idx])
            center = left + int(widths[new_idx] / 2)
            centers[key] = center
        for group, idx in table_map.group_to_index.items():
            if idx < 0 or idx >= len(widths):
                continue
            left = table_map.shape.left + sum(widths[:idx])
            centers[group] = left + int(widths[idx] / 2)
        table_centers.append((table_map, centers))

    if not table_centers:
        return

    for shape in iter_shapes(slide.shapes):
        name = getattr(shape, "name", "") or ""
        img_keys = SHAPE_IMG_MARKER_RE.findall(name)
        if img_keys:
            key = normalize_shape_key(img_keys[0])
            shape_center_y = shape.top + int(shape.height / 2)
            best_slot = None
            best_distance = None
            for table_map in table_maps:
                slot = table_map.image_slots.get(key)
                if slot is None:
                    slot = table_map.image_slots.get(key_to_group(key))
                if slot is None:
                    continue
                _, slot_top, _, slot_height = slot
                slot_center_y = slot_top + int(slot_height / 2)
                distance = abs(shape_center_y - slot_center_y)
                if best_distance is None or distance < best_distance:
                    best_distance = distance
                    best_slot = slot
            if best_slot is not None:
                left, top, width, height = best_slot
                shape.left = int(left + (width - shape.width) / 2)
                shape.top = int(top + (height - shape.height) / 2)
            continue
        keys = SHAPE_COL_MARKER_RE.findall(name)
        if not keys:
            continue
        key = normalize_shape_key(keys[0])
        shape_center_y = shape.top + int(shape.height / 2)
        best_center = None
        best_distance = None
        for table_map, centers in table_centers:
            center = centers.get(key)
            if center is None:
                center = centers.get(key_to_group(key))
            if center is None:
                continue
            table_center_y = table_map.shape.top + \
                int(table_map.shape.height / 2)
            distance = abs(shape_center_y - table_center_y)
            if best_distance is None or distance < best_distance:
                best_distance = distance
                best_center = center
        if best_center is None:
            continue
        shape.left = int(best_center - shape.width / 2)


def apply_vertical_stack(slide) -> None:
    groups = {}
    group_gaps = {}
    for shape in slide.shapes:
        name = getattr(shape, "name", "") or ""
        match = STACK_MARKER_RE.search(name)
        if not match:
            continue
        group = int(match.group("group"))
        item = match.group("item")
        order = int(item) if item else 0
        gap = match.group("gap")
        if gap:
            group_gaps[group] = int(gap)
        groups.setdefault(group, []).append((order, shape))

    if len(groups) < 2:
        return

    for group, items in groups.items():
        items.sort(key=lambda item: (item[0], getattr(item[1], "name", "")))
        groups[group] = [shape for _, shape in items]

    ordered_groups = sorted(groups.items(), key=lambda item: item[0])
    first_group = ordered_groups[0][1]
    first_top, first_bottom = group_visual_bounds(first_group)
    current_top = first_bottom + \
        group_gaps.get(ordered_groups[0][0], STACK_GAP)

    for group_id, shapes in ordered_groups[1:]:
        group_top, group_bottom = group_visual_bounds(shapes)
        delta = current_top - group_top
        for shape in shapes:
            shape.top = int(shape.top + delta)
        current_top = group_bottom + delta + \
            group_gaps.get(group_id, STACK_GAP)


def shape_visual_bounds(shape) -> Tuple[int, int]:
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        tops = []
        bottoms = []
        for subshape in shape.shapes:
            top, bottom = shape_visual_bounds(subshape)
            tops.append(top)
            bottoms.append(bottom)
        if tops and bottoms:
            return min(tops), max(bottoms)
    return shape.top, shape.top + shape.height


def group_visual_bounds(shapes) -> Tuple[int, int]:
    tops = []
    bottoms = []
    for shape in shapes:
        top, bottom = shape_visual_bounds(shape)
        tops.append(top)
        bottoms.append(bottom)
    return min(tops), max(bottoms)


def slide_contains_tokens(slide, tokens: List[str]) -> bool:
    return any(shape_contains_tokens(shape, tokens) for shape in slide.shapes)
