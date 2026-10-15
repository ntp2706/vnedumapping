"""Map homeroom teachers to subject assignments in an XLS workbook."""

from __future__ import annotations

import argparse
import logging
import os
import re
import tempfile
import unicodedata
from collections import defaultdict
from pathlib import Path

import xlrd
import xlwt


LOGGER = logging.getLogger(__name__)
INPUT_FIRST_ROW = 2
OUTPUT_FIRST_ROW = 12
INPUT_CLASS_COLUMN = 3
INPUT_TEACHER_COLUMN = 7
OUTPUT_TEACHER_COLUMN = 1
OUTPUT_SUBJECT_COLUMN = 2
OUTPUT_SEMESTER_1_COLUMN = 3
OUTPUT_SEMESTER_2_COLUMN = 4
TONE_MARKS = frozenset("\u0300\u0301\u0303\u0309\u0323")
VOWELS = frozenset("aeiouy")
SUBJECTS_BY_GRADE = {
    1: ("Tiếng Việt", "Toán", "Đạo đức", "TN-XH", "Hoạt động trải nghiệm"),
    2: ("Tiếng Việt", "Toán", "Đạo đức", "TN-XH", "Hoạt động trải nghiệm"),
    3: ("Tiếng Việt", "Toán", "Đạo đức", "TN-XH", "Hoạt động trải nghiệm"),
    4: (
        "Tiếng Việt",
        "Toán",
        "Khoa học",
        "Lịch sử và Địa lí",
        "Đạo đức",
        "Kĩ thuật",
        "Hoạt động trải nghiệm",
    ),
    5: (
        "Tiếng Việt",
        "Toán",
        "Khoa học",
        "Lịch sử và Địa lí",
        "Đạo đức",
        "Kĩ thuật",
        "Hoạt động trải nghiệm",
    ),
}
SUBJECT_NAMES = frozenset(
    subject for subjects in SUBJECTS_BY_GRADE.values() for subject in subjects
)
GRADE_PREFIX = re.compile(
    r"^(?:lop\s*)?(?:(1|2|3|4|5)(?!\d)|(mot|hai|ba|bon|nam)(?![a-z]))"
)
GRADE_WORDS = {
    "mot": 1,
    "hai": 2,
    "ba": 3,
    "bon": 4,
    "nam": 5,
}


def normalize_name(value: object) -> str:
    """Normalize case, whitespace, punctuation and Vietnamese tone placement."""
    decomposed = unicodedata.normalize("NFD", str(value or "").casefold())
    normalized = []
    vowel_group = []

    def flush_vowel_group() -> None:
        if vowel_group:
            tone_marks = [character for character in vowel_group if character in TONE_MARKS]
            normalized.extend(character for character in vowel_group if character not in TONE_MARKS)
            normalized.extend(tone_marks)
            vowel_group.clear()

    for character in decomposed:
        if character in TONE_MARKS:
            if vowel_group:
                vowel_group.append(character)
        elif character.isspace():
            flush_vowel_group()
        elif character in VOWELS:
            vowel_group.append(character)
        elif unicodedata.category(character).startswith("M") and vowel_group:
            vowel_group.append(character)
        elif character.isalnum():
            flush_vowel_group()
            normalized.append(character)
    flush_vowel_group()
    return "".join(normalized)


def normalize_class_name(value: object) -> str:
    """Remove Vietnamese marks while retaining separators for grade parsing."""
    decomposed = unicodedata.normalize("NFD", str(value or "").casefold())
    without_marks = "".join(
        character
        for character in decomposed
        if not unicodedata.category(character).startswith("M")
    )
    return re.sub(r"\s+", " ", without_marks).strip()


def teacher_name(value: object) -> str:
    """Return the teacher name before the optional /username suffix."""
    return str(value or "").split("/", maxsplit=1)[0].strip()


def grade_from_class(class_name: object) -> int | None:
    """Extract grades 1-5 from numeric or Vietnamese word class labels."""
    normalized = normalize_class_name(class_name)
    match = GRADE_PREFIX.match(normalized)
    if match is None:
        return None
    numeric_grade, word_grade = match.groups()
    return int(numeric_grade) if numeric_grade else GRADE_WORDS[word_grade]


def copy_style(source_book, style_index: int) -> xlwt.XFStyle:
    """Convert an xlrd cell style to an xlwt style."""
    source_xf = source_book.xf_list[style_index]
    style = xlwt.XFStyle()

    source_font = source_book.font_list[source_xf.font_index]
    font = xlwt.Font()
    for attribute in (
        "name",
        "height",
        "bold",
        "italic",
        "underline",
        "struck_out",
        "colour_index",
        "escapement",
        "family",
        "charset",
        "outline",
        "shadow",
        "character_spacing",
        "word_spacing",
    ):
        if hasattr(source_font, attribute):
            setattr(font, attribute, getattr(source_font, attribute))
    style.font = font

    source_alignment = source_xf.alignment
    alignment = xlwt.Alignment()
    for attribute in ("horz", "vert", "wrap", "rotation", "shrink_to_fit", "indent"):
        if hasattr(source_alignment, attribute):
            setattr(alignment, attribute, getattr(source_alignment, attribute))
    style.alignment = alignment

    source_border = source_xf.border
    borders = xlwt.Borders()
    for side in ("left", "right", "top", "bottom", "diag"):
        source_side = getattr(source_border, side, None)
        if source_side is not None:
            destination_side = getattr(borders, side)
            destination_side.line_style = getattr(source_side, "line_style", 0)
            destination_side.colour_index = getattr(source_side, "colour_index", 0)
    style.borders = borders

    source_background = source_xf.background
    pattern = xlwt.Pattern()
    pattern.pattern = source_background.fill_pattern
    pattern.pattern_fore_colour = source_background.pattern_colour_index
    pattern.pattern_back_colour = source_background.background_colour_index
    style.pattern = pattern

    number_format = source_book.format_map.get(source_xf.format_key)
    style.num_format_str = number_format.format_str if number_format else "General"
    return style


def read_homeroom_teachers(input_path: Path) -> list[tuple[str, str]]:
    """Read teacher names from column H and class names from column D."""
    workbook = xlrd.open_workbook(
        str(input_path),
        formatting_info=False,
        ignore_workbook_corruption=True,
    )
    try:
        if not workbook.sheets():
            raise ValueError(f"File đầu vào không có trang tính: {input_path}")
        sheet = workbook.sheet_by_index(0)
        if sheet.ncols <= max(INPUT_CLASS_COLUMN, INPUT_TEACHER_COLUMN):
            raise ValueError(
                f"File đầu vào cần có dữ liệu đến cột H: {input_path}"
            )

        teachers = []
        for row_index in range(INPUT_FIRST_ROW - 1, sheet.nrows):
            teacher = str(sheet.cell_value(row_index, INPUT_TEACHER_COLUMN) or "").strip()
            class_name = str(sheet.cell_value(row_index, INPUT_CLASS_COLUMN) or "").strip()
            if teacher and class_name:
                teachers.append((teacher, class_name))
        return teachers
    finally:
        workbook.release_resources()


def copy_row(
    source_sheet,
    row_index: int,
    values: list[object],
    destination_sheet,
    destination_row: int,
    source_book,
    style_cache: dict[int, xlwt.XFStyle],
    merged_coordinates: set[tuple[int, int]],
) -> None:
    """Copy a row's values and cell formatting, excluding merged cells."""
    row_info = source_sheet.rowinfo_map.get(row_index)
    if row_info:
        row = destination_sheet.row(destination_row)
        row.height = row_info.height
        row.hidden = row_info.hidden

    for column_index, value in enumerate(values):
        if (row_index, column_index) in merged_coordinates:
            continue
        cell = source_sheet.cell(row_index, column_index)
        if cell.xf_index not in style_cache:
            style_cache[cell.xf_index] = copy_style(source_book, cell.xf_index)
        destination_sheet.write(
            destination_row,
            column_index,
            value,
            style_cache[cell.xf_index],
        )


def map_sheet(
    source_book,
    source_sheet,
    destination_sheet,
    teacher_classes: dict[str, list[str]],
    semester: int,
    style_cache: dict[int, xlwt.XFStyle],
) -> tuple[int, int, int, dict[int, int]]:
    """Build the output rows and assignments for the first worksheet."""
    if source_sheet.ncols <= OUTPUT_SEMESTER_2_COLUMN:
        raise ValueError("File phân công cần có dữ liệu đến cột E.")

    merged_coordinates = {
        (row_index, column_index)
        for first_row, last_row, first_column, last_column in source_sheet.merged_cells
        for row_index in range(first_row, last_row)
        for column_index in range(first_column, last_column)
    }
    output_rows = []
    original_to_output_row = {}
    mapped_count = 0
    skipped_classes = 0
    seen_teacher_keys = set()
    row_index = 0

    while row_index < source_sheet.nrows:
        values = list(source_sheet.row_values(row_index))
        values.extend([""] * (source_sheet.ncols - len(values)))
        teacher_key = ""
        if row_index >= OUTPUT_FIRST_ROW - 1:
            teacher_key = normalize_name(
                teacher_name(values[OUTPUT_TEACHER_COLUMN])
            )
        if teacher_key:
            seen_teacher_keys.add(teacher_key)
        class_names = teacher_classes.get(teacher_key, [])

        grade = grade_from_class(class_names[0]) if class_names else None
        if not class_names or grade is None:
            if class_names and grade is None:
                skipped_classes += 1
                LOGGER.warning(
                    "Không xác định được khối lớp của %s: %s",
                    teacher_name(values[OUTPUT_TEACHER_COLUMN]),
                    class_names[0],
                )
            original_to_output_row[row_index] = len(output_rows)
            output_rows.append((row_index, values))
            row_index += 1
            continue

        if len({normalize_class_name(name) for name in class_names}) > 1:
            skipped_classes += 1
            LOGGER.warning(
                "Giáo viên %s có nhiều lớp chủ nhiệm trong file đầu vào; bỏ qua để tránh ghi đè.",
                teacher_name(values[OUTPUT_TEACHER_COLUMN]),
            )
            original_to_output_row[row_index] = len(output_rows)
            output_rows.append((row_index, values))
            row_index += 1
            continue

        subjects = SUBJECTS_BY_GRADE[grade]
        previous_classes = {}
        first_subject = values[OUTPUT_SUBJECT_COLUMN]
        if first_subject in SUBJECT_NAMES:
            previous_classes[first_subject] = (
                values[OUTPUT_SEMESTER_1_COLUMN],
                values[OUTPUT_SEMESTER_2_COLUMN],
            )

        next_row = row_index + 1
        while next_row < source_sheet.nrows:
            following_values = source_sheet.row_values(next_row)
            teacher_value = (
                following_values[OUTPUT_TEACHER_COLUMN]
                if OUTPUT_TEACHER_COLUMN < len(following_values)
                else ""
            )
            subject_value = (
                following_values[OUTPUT_SUBJECT_COLUMN]
                if OUTPUT_SUBJECT_COLUMN < len(following_values)
                else ""
            )
            if teacher_value or subject_value not in SUBJECT_NAMES:
                break
            previous_classes.setdefault(
                subject_value,
                (
                    following_values[OUTPUT_SEMESTER_1_COLUMN]
                    if OUTPUT_SEMESTER_1_COLUMN < len(following_values)
                    else "",
                    following_values[OUTPUT_SEMESTER_2_COLUMN]
                    if OUTPUT_SEMESTER_2_COLUMN < len(following_values)
                    else "",
                ),
            )
            next_row += 1

        class_name = class_names[0]
        for subject_index, subject in enumerate(subjects):
            if subject_index == 0:
                subject_values = list(values)
                original_to_output_row[row_index] = len(output_rows)
            else:
                subject_values = [""] * source_sheet.ncols

            old_classes = previous_classes.get(subject, ("", ""))
            subject_values[OUTPUT_SUBJECT_COLUMN] = subject
            subject_values[OUTPUT_SEMESTER_1_COLUMN] = (
                class_name if semester == 1 else old_classes[0]
            )
            subject_values[OUTPUT_SEMESTER_2_COLUMN] = (
                class_name if semester == 2 else old_classes[1]
            )
            output_rows.append((row_index, subject_values))

        mapped_count += 1
        for consumed_row in range(row_index + 1, next_row):
            original_to_output_row[consumed_row] = len(output_rows)
        row_index = next_row

    unmatched_count = 0
    for key, class_names in teacher_classes.items():
        if key not in seen_teacher_keys:
            unmatched_count += len(class_names)
            LOGGER.warning(
                "Không tìm thấy giáo viên trong file phân công: %s",
                class_names[0],
            )

    if len(output_rows) > 65536:
        raise ValueError("Số hàng sau khi thêm môn vượt giới hạn định dạng XLS.")

    for output_index, (style_row, values) in enumerate(output_rows):
        copy_row(
            source_sheet,
            style_row,
            values,
            destination_sheet,
            output_index,
            source_book,
            style_cache,
            merged_coordinates,
        )

    for first_row, last_row, first_column, last_column in source_sheet.merged_cells:
        first_output_row = original_to_output_row[first_row]
        last_output_row = original_to_output_row[last_row - 1]
        top_left = source_sheet.cell(first_row, first_column)
        if top_left.xf_index not in style_cache:
            style_cache[top_left.xf_index] = copy_style(
                source_book, top_left.xf_index
            )
        destination_sheet.write_merge(
            first_output_row,
            last_output_row,
            first_column,
            last_column - 1,
            source_sheet.cell_value(first_row, first_column),
            style_cache[top_left.xf_index],
        )
    return mapped_count, skipped_classes, unmatched_count, original_to_output_row


def map_workbook(
    input_path: Path,
    template_path: Path,
    output_path: Path,
    semester: int,
) -> tuple[int, int, int]:
    """Write homeroom subject assignments for the selected semester."""
    if semester not in (1, 2):
        raise ValueError("Học kỳ phải là 1 hoặc 2.")
    homeroom_teachers = read_homeroom_teachers(input_path)
    if not homeroom_teachers:
        raise ValueError(f"Không tìm thấy giáo viên/lớp trong file: {input_path}")

    teacher_classes: dict[str, list[str]] = defaultdict(list)
    for teacher, class_name in homeroom_teachers:
        key = normalize_name(teacher)
        if key:
            teacher_classes[key].append(class_name)

    source_book = xlrd.open_workbook(
        str(template_path),
        formatting_info=True,
    )
    destination_book = xlwt.Workbook()
    style_cache: dict[int, xlwt.XFStyle] = {}
    mapped_count = 0
    skipped_classes = 0
    unmatched_count = 0

    try:
        if not source_book.sheets():
            raise ValueError(f"File phân công không có trang tính: {template_path}")

        for sheet_index, source_sheet in enumerate(source_book.sheets()):
            destination_sheet = destination_book.add_sheet(
                source_sheet.name,
                cell_overwrite_ok=True,
            )
            for column_index, column_info in source_sheet.colinfo_map.items():
                destination_sheet.col(column_index).width = column_info.width
                destination_sheet.col(column_index).hidden = column_info.hidden

            if sheet_index == 0:
                mapped_count, skipped_classes, unmatched_count, _ = map_sheet(
                    source_book,
                    source_sheet,
                    destination_sheet,
                    teacher_classes,
                    semester,
                    style_cache,
                )
            else:
                merged_coordinates = {
                    (row_index, column_index)
                    for first_row, last_row, first_column, last_column in source_sheet.merged_cells
                    for row_index in range(first_row, last_row)
                    for column_index in range(first_column, last_column)
                }
                for row_index in range(source_sheet.nrows):
                    copy_row(
                        source_sheet,
                        row_index,
                        list(source_sheet.row_values(row_index)),
                        destination_sheet,
                        row_index,
                        source_book,
                        style_cache,
                        merged_coordinates,
                    )
                for first_row, last_row, first_column, last_column in source_sheet.merged_cells:
                    top_left = source_sheet.cell(first_row, first_column)
                    if top_left.xf_index not in style_cache:
                        style_cache[top_left.xf_index] = copy_style(
                            source_book, top_left.xf_index
                        )
                    destination_sheet.write_merge(
                        first_row,
                        last_row - 1,
                        first_column,
                        last_column - 1,
                        source_sheet.cell_value(first_row, first_column),
                        style_cache[top_left.xf_index],
                    )
    finally:
        source_book.release_resources()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xls",
            dir=output_path.parent,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)
        destination_book.save(str(temporary_path))
        os.replace(temporary_path, output_path)
    finally:
        if temporary_path is not None and temporary_path.exists():
            os.remove(temporary_path)

    return mapped_count, skipped_classes, unmatched_count


def prompt_semester() -> int:
    """Ask for semester 1 or 2 until a valid value is provided."""
    while True:
        try:
            value = input("Nhập học kỳ (1 hoặc 2): ").strip()
        except EOFError as error:
            raise SystemExit("Cần nhập học kỳ 1 hoặc 2.") from error
        if value in ("1", "2"):
            return int(value)
        print("Giá trị không hợp lệ. Vui lòng nhập 1 hoặc 2.")


def main() -> None:
    base_directory = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        type=Path,
        default=base_directory / "lop_hoc.xls",
        help="File danh sách lớp và giáo viên chủ nhiệm.",
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=base_directory / "phancong_chuyenmon.xls",
        help="File phân công ban đầu.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="File kết quả (mặc định ghi đè file --template).",
    )
    parser.add_argument(
        "--semester",
        type=int,
        choices=(1, 2),
        help="Học kỳ cần cập nhật; nếu bỏ trống sẽ hỏi từ bàn phím.",
    )
    arguments = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    output_path = arguments.output or arguments.template
    semester = arguments.semester or prompt_semester()
    mapped_count, skipped_classes, unmatched_count = map_workbook(
        arguments.input,
        arguments.template,
        output_path,
        semester,
    )
    LOGGER.info(
        "Đã ghi phân công học kỳ %d cho %d dòng giáo viên; "
        "bỏ qua %d lớp không hợp lệ, không tìm thấy %d giáo viên.",
        semester,
        mapped_count,
        skipped_classes,
        unmatched_count,
    )


if __name__ == "__main__":
    main()