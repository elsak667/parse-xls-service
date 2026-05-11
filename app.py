#!/usr/bin/env python3
"""极简 .xls 解析服务 - 筛选招商中心日程
部署到 Render.com 免费层: https://render.com/docs/free#networking-and-execution
"""
import re
import os
import tempfile
from flask import Flask, request, jsonify

try:
    import xlrd
    XLRD_OK = True
except ImportError:
    XLRD_OK = False

app = Flask(__name__)

@app.route("/health")
def health():
    return jsonify({"status": "ok", "xlrd": XLRD_OK})

@app.route("/parse_xls", methods=["POST"])
def parse_xls():
    if not XLRD_OK:
        return jsonify({"error": "xlrd not available"}), 500

    # 接收 .xls 文件
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "no file provided"}), 400

    # 写临时文件
    with tempfile.NamedTemporaryFile(suffix=".xls", delete=False) as tmp:
        tmp.write(file.read())
        tmp_path = tmp.name

    try:
        wb = xlrd.open_workbook(tmp_path)
        ws = wb.sheet_by_name("Sheet0")

        events = []
        for r in range(1, ws.nrows):
            row = ws.row_values(r)
            date_val = str(row[0]).strip() if row[0] else ""
            if not date_val:
                continue
            date_str = date_val[:10]

            for col_idx, period in [(1, "上午"), (2, "下午")]:
                cell = str(row[col_idx]) if col_idx < len(row) and row[col_idx] else ""
                # 支持 \r\n 或 \n 分隔
                items = re.split(r'[\r\n]+', cell)
                for item in items:
                    item = item.strip()
                    if not item:
                        continue
                    if "招商中心" in item:
                        time_match = re.match(r"^(\d{2}:\d{2})\s+(.+)", item)
                        if time_match:
                            time_str = time_match.group(1)
                            content = time_match.group(2)
                        else:
                            time_str = ""
                            content = item
                        events.append({
                            "date": date_str,
                            "period": period,
                            "time": time_str,
                            "content": content,
                            "raw": item
                        })

        return jsonify({"ok": True, "count": len(events), "events": events})
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        os.unlink(tmp_path)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
