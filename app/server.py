"""园区月度电量试算 —— Flask 网页。

流程：上传 XLSX -> 选择本楼栋工作表并填写园区总计单元格 -> 下载试算 XLSX。
"""
from __future__ import annotations

import os
import uuid

from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)
from openpyxl import load_workbook

from .analyzer import FormulaError, WorkbookAnalyzer, parse_total_target
from .builder import build_filename, build_workbook, save_workbook

DATA_DIR = os.environ.get("DATA_DIR", os.path.join(os.getcwd(), "data"))
MAX_CONTENT_LENGTH = 20 * 1024 * 1024  # 20MB


def create_app() -> Flask:
    app = Flask(__name__)
    app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-me")
    app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH
    os.makedirs(DATA_DIR, exist_ok=True)

    def _path(file_id: str) -> str:
        # 防止路径穿越
        safe = uuid.UUID(file_id).hex
        return os.path.join(DATA_DIR, f"{safe}.xlsx")

    @app.get("/")
    def index():
        return render_template("index.html")

    @app.post("/upload")
    def upload():
        file = request.files.get("file")
        if file is None or not file.filename:
            flash("请选择要上传的 XLSX 文件。", "error")
            return redirect(url_for("index"))
        if not file.filename.lower().endswith(".xlsx"):
            flash("仅支持 .xlsx 格式（旧版 .xls 请先用 Excel 另存为 xlsx）。", "error")
            return redirect(url_for("index"))
        file_id = uuid.uuid4()
        dest = _path(str(file_id))
        file.save(dest)
        try:
            wb = load_workbook(dest, data_only=False, read_only=True)
            sheets = wb.sheetnames
            wb.close()
        except Exception as exc:  # openpyxl 对坏文件抛多种异常
            os.remove(dest)
            flash(f"无法解析该 XLSX 文件：{exc}", "error")
            return redirect(url_for("index"))
        return render_template(
            "select.html",
            file_id=file_id,
            filename=file.filename,
            sheets=sheets,
        )

    @app.post("/export")
    def export():
        file_id = request.form.get("file_id", "")
        building_sheet = request.form.get("building_sheet", "")
        total_raw = request.form.get("total_cell", "")
        try:
            src_path = _path(file_id)
        except (ValueError, AttributeError):
            flash("无效的上传会话，请重新上传文件。", "error")
            return redirect(url_for("index"))
        if not os.path.exists(src_path):
            flash("上传会话已过期，请重新上传文件。", "error")
            return redirect(url_for("index"))

        try:
            total_sheet, total_coord = parse_total_target(total_raw, building_sheet)
        except FormulaError as exc:
            flash(str(exc), "error")
            return redirect(url_for("index"))

        wb = None
        try:
            wb = load_workbook(src_path, data_only=False)
            analyzer = WorkbookAnalyzer(wb, building_sheet, total_sheet, total_coord)
            result = analyzer.analyze()
            out_wb = build_workbook(wb, result)
            payload = save_workbook(out_wb)
        except FormulaError as exc:
            # 超出支持范围或成环：指出单元格并停止导出
            return render_template("error.html", message=str(exc)), 422
        except KeyError as exc:
            return render_template("error.html", message=f"引用了不存在的工作表或单元格：{exc}"), 422
        finally:
            if wb is not None:
                wb.close()

        original_name = request.form.get("filename", "电量表.xlsx")
        import io

        return send_file(
            io.BytesIO(payload),
            as_attachment=True,
            download_name=build_filename(original_name, building_sheet),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    @app.errorhandler(413)
    def too_large(_):
        flash("文件超过 20MB 限制。", "error")
        return redirect(url_for("index"))

    return app


app = create_app()

if __name__ == "__main__":  # pragma: no cover
    app.run(host="0.0.0.0", port=8000)
