"""Web server for sequential vnEdu workbook conversion."""

from __future__ import annotations

import re
import shutil
import tempfile
import threading
import uuid
from pathlib import Path

from flask import Flask, jsonify, request, send_file, send_from_directory
from werkzeug.utils import secure_filename

from exec import process_file


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "vnedu.xls"
JOBS_DIR = BASE_DIR / "web_jobs"
ALLOWED_EXTENSION = ".xlsx"
PROCESS_LOCK = threading.Lock()

app = Flask(__name__)


def output_filename(input_name: str) -> str:
	"""Use the uploaded input name as the corresponding XLS output name."""
	name = secure_filename(input_name)
	stem = Path(name).stem or "vnedu_output"
	stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._") or "vnedu_output"
	return f"{stem}.xls"


@app.get("/")
def home():
	return send_file(BASE_DIR / "index.html")


@app.post("/api/process")
def process_uploads():
	files = request.files.getlist("files")
	if not files:
		return jsonify({"error": "Vui lòng chọn ít nhất một file .xlsx."}), 400

	for uploaded_file in files:
		if not uploaded_file.filename or Path(uploaded_file.filename).suffix.lower() != ALLOWED_EXTENSION:
			return jsonify({"error": "Chỉ chấp nhận file có phần mở rộng .xlsx."}), 400

	job_id = uuid.uuid4().hex
	job_dir = JOBS_DIR / job_id
	input_dir = job_dir / "input"
	output_dir = job_dir / "output"
	input_dir.mkdir(parents=True, exist_ok=False)
	output_dir.mkdir()
	results = []

	try:
		with PROCESS_LOCK:
			used_names = set()
			for index, uploaded_file in enumerate(files, start=1):
				original_name = secure_filename(uploaded_file.filename)
				input_name = original_name or f"input_{index}.xlsx"
				input_path = input_dir / input_name
				if input_path.name in used_names:
					input_path = input_dir / f"{input_path.stem}_{index}{input_path.suffix}"
				used_names.add(input_path.name)
				uploaded_file.save(input_path)

				name = output_filename(input_path.name)
				if name in {item["filename"] for item in results}:
					name = f"{Path(name).stem}_{index}.xls"
				output_path = output_dir / name
				count = process_file(input_path, TEMPLATE_PATH, output_path)
				results.append({"filename": name, "count": count})
		return jsonify({"job_id": job_id, "files": results})
	except Exception as error:
		shutil.rmtree(job_dir, ignore_errors=True)
		return jsonify({"error": f"Không thể xử lý file: {error}"}), 500


@app.get("/download/<job_id>/<path:filename>")
def download(job_id: str, filename: str):
	return send_from_directory(JOBS_DIR / job_id / "output", filename, as_attachment=True)


if __name__ == "__main__":
	JOBS_DIR.mkdir(exist_ok=True)
	app.run(host="127.0.0.1", port=5000, debug=False)