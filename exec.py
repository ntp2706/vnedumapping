"""Map student data from the input workbook into the vnEdu XLS template."""

from __future__ import annotations

import argparse
import os
import shutil
import tempfile
from copy import copy
from pathlib import Path

import openpyxl
import xlrd
import xlwt


SKIPPED_STATUSES = {
	"chuyển đi kỳ 1",
	"thôi họ kỳ 1",
	"thôi học kỳ 1",
	"chuyển đi kỳ 2",
	"thôi học kỳ 2",
	"chuyển đi trong hè",
	"thôi học trong hè",
}


# Source and destination columns are 1-based Excel column numbers.
COLUMN_MAP = {
	"C": "D",
	"D": "F",
	"E": "G",
	"F": "I",
	"G": "AD",
	"J": "H",
	"P": "J",
	"Q": "R",
	"R": "Q",
	"U": "O",
	"T": "U",
	"V": "X",
	"W": "W",
	"X": "V",
	"Z": "AP",
	"AB": "AU",
	"AC": "AF",
	"AP": "AA",
	"AQ": "AB",
	"AR": "AC",
	"AS": "AI",
	"BI": "AJ",
	"BJ": "AK",
	"BK": "AL",
	"BM": "AM",
	"BN": "AN",
	"BO": "AO",
}


def column_number(column: str) -> int:
	"""Return a 1-based number for an Excel column name."""
	result = 0
	for character in column.upper():
		result = result * 26 + ord(character) - ord("A") + 1
	return result


def clean_text(value) -> str:
	return "" if value is None else str(value).strip()


def split_current_address(value) -> tuple[str, str, str]:
	"""Split detail, commune and province from the right-hand side."""
	parts = [part.strip() for part in clean_text(value).split(",") if part.strip()]
	if len(parts) >= 3:
		return ", ".join(parts[:-2]), parts[-2], parts[-1]
	if len(parts) == 2:
		return parts[0], "", parts[1]
	if len(parts) == 1:
		return parts[0], "", ""
	return "", "", ""


def copy_style(book, source_xf_index: int):
	"""Convert an xlrd cell style to an xlwt style."""
	source_xf = book.xf_list[source_xf_index]
	style = xlwt.XFStyle()

	source_font = book.font_list[source_xf.font_index]
	font = xlwt.Font()
	for attribute in (
		"name", "height", "bold", "italic", "underline", "struck_out",
		"colour_index", "escapement", "family", "charset", "outline",
		"shadow", "character_spacing", "word_spacing",
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
	border = xlwt.Borders()
	for side in ("left", "right", "top", "bottom", "diag"):
		source_side = getattr(source_border, side, None)
		if source_side is not None:
			destination_side = getattr(border, side)
			destination_side.line_style = getattr(source_side, "line_style", 0)
			destination_side.colour_index = getattr(source_side, "colour_index", 0)
	style.borders = border

	source_pattern = source_xf.background
	pattern = xlwt.Pattern()
	pattern.pattern = source_pattern.fill_pattern
	pattern.pattern_fore_colour = source_pattern.pattern_colour_index
	pattern.pattern_back_colour = source_pattern.background_colour_index
	style.pattern = pattern

	source_format = book.format_map.get(source_xf.format_key)
	style.num_format_str = source_format.format_str if source_format else "General"
	return style


def copy_template(template_path: Path):
	"""Copy all sheets, values, dimensions and basic formatting from the XLS template."""
	source_book = xlrd.open_workbook(str(template_path), formatting_info=True)
	destination_book = xlwt.Workbook()
	style_cache = {}

	for source_sheet in source_book.sheets():
		destination_sheet = destination_book.add_sheet(source_sheet.name)
		for row_index in range(source_sheet.nrows):
			if source_sheet.name == "ds học sinh" and row_index >= 8:
				continue
			if source_sheet.rowinfo_map.get(row_index):
				row_info = source_sheet.rowinfo_map[row_index]
				destination_sheet.row(row_index).height = row_info.height
				destination_sheet.row(row_index).hidden = row_info.hidden
			for column_index in range(source_sheet.ncols):
				cell = source_sheet.cell(row_index, column_index)
				if cell.xf_index not in style_cache:
					style_cache[cell.xf_index] = copy_style(source_book, cell.xf_index)
				destination_sheet.write(row_index, column_index, cell.value, style_cache[cell.xf_index])

		for column_index, column_info in source_sheet.colinfo_map.items():
			destination_sheet.col(column_index).width = column_info.width
			destination_sheet.col(column_index).hidden = column_info.hidden
	return source_book, destination_book


def map_workbooks(input_path: Path, template_path: Path, output_path: Path) -> int:
	input_book = openpyxl.load_workbook(input_path, data_only=False, read_only=True)
	input_sheet = input_book.active
	source_book, output_book = copy_template(template_path)
	output_sheet = output_book.get_sheet("ds học sinh")

	source_columns = {column: column_number(column) for column in COLUMN_MAP}
	destination_columns = {column: column_number(column) for column in COLUMN_MAP.values()}
	output_row = 8  # zero-based row 9
	sequence = 1

	for input_row in range(7, input_sheet.max_row + 1):
		student_code = clean_text(input_sheet.cell(input_row, column_number("C")).value)
		student_name = clean_text(input_sheet.cell(input_row, column_number("D")).value)
		if not student_code and not student_name:
			continue

		status = clean_text(input_sheet.cell(input_row, column_number("H")).value).casefold()
		if status in SKIPPED_STATUSES:
			continue

		for source_column, destination_column in COLUMN_MAP.items():
			value = input_sheet.cell(input_row, source_columns[source_column]).value
			output_sheet.write(output_row, destination_columns[destination_column] - 1, value)

		detail, commune, province = split_current_address(
			input_sheet.cell(input_row, column_number("Y")).value
		)
		output_sheet.write(output_row, column_number("K") - 1, detail)
		output_sheet.write(output_row, column_number("M") - 1, commune)
		output_sheet.write(output_row, column_number("N") - 1, province)
		output_sheet.write(output_row, column_number("A") - 1, sequence)
		output_sheet.write(output_row, column_number("AE") - 1, "Không")
		output_row += 1
		sequence += 1

	output_path.parent.mkdir(parents=True, exist_ok=True)
	with tempfile.NamedTemporaryFile(delete=False, suffix=".xls", dir=output_path.parent) as temporary:
		temporary_path = Path(temporary.name)
	try:
		output_book.save(str(temporary_path))
		shutil.move(str(temporary_path), str(output_path))
	finally:
		if temporary_path.exists():
			os.remove(temporary_path)
	return sequence - 1


def process_file(input_path: Path, template_path: Path, output_path: Path) -> int:
	"""Process one input workbook and write one output workbook."""
	return map_workbooks(input_path, template_path, output_path)


def main() -> None:
	parser = argparse.ArgumentParser(description=__doc__)
	parser.add_argument("--input", type=Path, default=Path("Danh_sach_hoc_sinh.xlsx"))
	parser.add_argument("--template", type=Path, default=Path("vnedu.xls"))
	parser.add_argument("--output", type=Path, default=Path("vnedu.xls"))
	arguments = parser.parse_args()
	count = process_file(arguments.input, arguments.template, arguments.output)
	print(f"Đã map {count} học sinh vào {arguments.output}")


if __name__ == "__main__":
	main()
