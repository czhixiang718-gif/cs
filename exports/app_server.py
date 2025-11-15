import os
import json
import subprocess
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

ROOT = os.path.dirname(os.path.abspath(__file__))
MYSQL = r"C:\Program Files\MariaDB 12.0\bin\mysql.exe"
if not os.path.exists(MYSQL):
    MYSQL = "mysql"
DB = "ledger_db"
USER = os.environ.get("LEDGER_APP_USER", "ledger_app")
PWD = os.environ.get("LEDGER_APP_PWD", "App!12345")
ADMIN_FILE = os.path.join(ROOT, "admin.json")
SESSIONS = {}
RESET_TOKENS = {}

def load_admin_creds():
    default = {"username": "zhaofan", "password": "zhaofan9766"}
    if os.path.exists(ADMIN_FILE):
        try:
            with open(ADMIN_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                u = data.get("username") or default["username"]
                p = data.get("password") or default["password"]
                return {"username": u, "password": p}
        except Exception:
            return default
    else:
        try:
            with open(ADMIN_FILE, "w", encoding="utf-8") as f:
                json.dump(default, f, ensure_ascii=False)
        except Exception:
            pass
        return default

_creds = load_admin_creds()
ADMIN_USER = _creds["username"]
ADMIN_PASS = _creds["password"]
SESSION_TTL = 600

def run_sql(sql, expect_tsv=False):
    env = os.environ.copy()
    env["MYSQL_PWD"] = PWD
    args = [MYSQL, "-u", USER, "-D", DB]
    if expect_tsv:
        args += ["-B", "-e", sql]
        p = subprocess.run(args, env=env, capture_output=True)
        if p.returncode != 0:
            return None, p.stderr.decode("utf-8", errors="ignore")
        return p.stdout.decode("utf-8", errors="ignore"), None
    else:
        args += ["-e", sql]
        p = subprocess.run(args, env=env, capture_output=True)
        if p.returncode != 0:
            return False, p.stderr.decode("utf-8", errors="ignore")
        return True, None

def tsv_to_json(tsv):
    lines = [l for l in tsv.splitlines() if l.strip()]
    if not lines:
        return []
    header = lines[0].split("\t")
    rows = []
    for line in lines[1:]:
        parts = line.split("\t")
        obj = {}
        for i, k in enumerate(header):
            obj[k] = parts[i] if i < len(parts) else ""
        rows.append(obj)
    return rows

class Handler(SimpleHTTPRequestHandler):
    def translate_path(self, path):
        return os.path.join(ROOT, path.lstrip("/"))

    def do_GET(self):
        if self.path.startswith("/api/"):
            if self.path.startswith("/api/export/"):
                self.handle_export()
            else:
                self.handle_api_get()
        else:
            if self.path == "/manage.html" and not self.get_session_user():
                self.send_response(302)
                self.send_header("Location", "/login.html")
                self.end_headers()
            else:
                super().do_GET()

    def do_POST(self):
        if self.path.startswith("/api/"):
            self.handle_api_write("POST")
        else:
            self.send_error(404)

    def do_PUT(self):
        if self.path.startswith("/api/"):
            self.handle_api_write("PUT")
        else:
            self.send_error(404)

    def do_DELETE(self):
        if self.path.startswith("/api/"):
            self.handle_api_write("DELETE")
        else:
            self.send_error(404)

    def json_response(self, obj, status=200):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def read_json(self):
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length) if length > 0 else b""
        if not raw:
            return {}
        return json.loads(raw.decode("utf-8"))

    def get_session_user(self):
        cookie = self.headers.get("Cookie")
        if not cookie:
            return None
        parts = [p.strip() for p in cookie.split(";")]
        val = None
        for p in parts:
            if p.startswith("SESSIONID="):
                val = p.split("=",1)[1]
                break
        if not val:
            return None
        info = SESSIONS.get(val)
        if not info:
            return None
        import time
        last = info.get("last", 0)
        if last < time.time() - SESSION_TTL:
            try:
                SESSIONS.pop(val, None)
            except Exception:
                pass
            return None
        info["last"] = time.time()
        return info.get("user")

    def handle_api_get(self):
        parsed = urlparse(self.path)
        path = parsed.path
        qs = parse_qs(parsed.query)
        if path == "/api/health":
            self.json_response({"ok": True, "auth": True if self.get_session_user() else False})
            return
        if path == "/api/admin_profile":
            if not self.get_session_user():
                self.json_response({"error": "unauthorized"}, 401); return
            self.json_response({"username": ADMIN_USER})
            return
        if path == "/api/projects":
            tsv, err = run_sql("SELECT id,name,contract_date,total_price,created_at,updated_at FROM projects ORDER BY id", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500)
                return
            self.json_response(tsv_to_json(tsv))
            return
        if path == "/api/payments":
            pid = qs.get("project_id", [None])[0]
            if not pid:
                self.json_response({"error": "project_id required"}, 400)
                return
            tsv, err = run_sql(f"SELECT payments.id, payments.project_id, projects.name AS project_name, payments.payment_date, payments.amount, payments.created_at FROM payments LEFT JOIN projects ON projects.id=payments.project_id WHERE payments.project_id={int(pid)} ORDER BY payments.id", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500)
                return
            self.json_response(tsv_to_json(tsv))
            return
        if path == "/api/payments_all":
            tsv, err = run_sql("SELECT payments.id, projects.name AS project_name, payments.payment_date, payments.amount FROM payments LEFT JOIN projects ON projects.id=payments.project_id ORDER BY payments.payment_date, payments.id", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500)
                return
            self.json_response(tsv_to_json(tsv))
            return
        if path == "/api/report/project_finance_summary":
            tsv, err = run_sql("SELECT project_id,name,contract_date,total_price,payment_count,payment_total_amount,outstanding_amount FROM project_finance_summary ORDER BY project_id", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500); return
            self.json_response(tsv_to_json(tsv)); return
        if path == "/api/report/global_finance_totals":
            tsv, err = run_sql("SELECT total_projects,total_price_all_projects,total_outstanding_all_projects,total_payments_amount,total_payments_count FROM global_finance_totals", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500); return
            self.json_response(tsv_to_json(tsv)); return
        if path == "/api/report/monthly_contracts":
            tsv, err = run_sql("SELECT ym,project_count,contract_amount FROM monthly_contracts ORDER BY ym", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500); return
            self.json_response(tsv_to_json(tsv)); return
        if path == "/api/report/monthly_payments":
            tsv, err = run_sql("SELECT ym,payment_count,payment_amount FROM monthly_payments ORDER BY ym", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500); return
            self.json_response(tsv_to_json(tsv)); return
        if path == "/api/detail/monthly_contracts":
            tsv, err = run_sql("SELECT DATE_FORMAT(contract_date,'%Y-%m') AS ym,id,name,contract_date,total_price FROM projects ORDER BY contract_date,id", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500); return
            self.json_response(tsv_to_json(tsv)); return
        if path == "/api/detail/monthly_payments":
            tsv, err = run_sql("SELECT DATE_FORMAT(payments.payment_date,'%Y-%m') AS ym,payments.id,projects.name AS project_name,payments.payment_date,payments.amount FROM payments LEFT JOIN projects ON projects.id=payments.project_id ORDER BY payments.payment_date,payments.id", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500); return
            self.json_response(tsv_to_json(tsv)); return
        if path == "/api/report/quarterly_contracts":
            tsv, err = run_sql("SELECT yq,project_count,contract_amount FROM quarterly_contracts ORDER BY yq", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500); return
            self.json_response(tsv_to_json(tsv)); return
        if path == "/api/report/quarterly_payments":
            tsv, err = run_sql("SELECT yq,payment_count,payment_amount FROM quarterly_payments ORDER BY yq", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500); return
            self.json_response(tsv_to_json(tsv)); return
        if path == "/api/detail/quarterly_contracts":
            tsv, err = run_sql("SELECT CONCAT(YEAR(contract_date), '-Q', QUARTER(contract_date)) AS yq,id,name,contract_date,total_price FROM projects ORDER BY contract_date,id", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500); return
            self.json_response(tsv_to_json(tsv)); return
        if path == "/api/detail/quarterly_payments":
            tsv, err = run_sql("SELECT CONCAT(YEAR(payments.payment_date), '-Q', QUARTER(payments.payment_date)) AS yq,payments.id,projects.name AS project_name,payments.payment_date,payments.amount FROM payments LEFT JOIN projects ON projects.id=payments.project_id ORDER BY payments.payment_date,payments.id", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500); return
            self.json_response(tsv_to_json(tsv)); return
        if path == "/api/report/yearly_contracts":
            tsv, err = run_sql("SELECT yr,project_count,contract_amount FROM yearly_contracts ORDER BY yr", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500); return
            self.json_response(tsv_to_json(tsv)); return
        if path == "/api/report/yearly_payments":
            tsv, err = run_sql("SELECT yr,payment_count,payment_amount FROM yearly_payments ORDER BY yr", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500); return
            self.json_response(tsv_to_json(tsv)); return
        if path == "/api/detail/yearly_contracts":
            tsv, err = run_sql("SELECT YEAR(contract_date) AS yr,id,name,contract_date,total_price FROM projects ORDER BY contract_date,id", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500); return
            self.json_response(tsv_to_json(tsv)); return
        if path == "/api/detail/yearly_payments":
            tsv, err = run_sql("SELECT YEAR(payments.payment_date) AS yr,payments.id,projects.name AS project_name,payments.payment_date,payments.amount FROM payments LEFT JOIN projects ON projects.id=payments.project_id ORDER BY payments.payment_date,payments.id", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500); return
            self.json_response(tsv_to_json(tsv)); return
        if path == "/api/report/overdue_projects_90d":
            tsv, err = run_sql("SELECT project_id,name,contract_date,total_price,paid_amount,outstanding_amount,days_since_contract FROM overdue_projects_90d ORDER BY days_since_contract DESC", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500); return
            self.json_response(tsv_to_json(tsv)); return
        self.send_error(404)

    def handle_export(self):
        parsed = urlparse(self.path)
        path = parsed.path
        if path == "/api/export/all.xlsx":
            sheets = {}
            gf_tsv, e1 = run_sql("SELECT total_projects,total_price_all_projects,total_outstanding_all_projects,total_payments_amount,total_payments_count FROM global_finance_totals", expect_tsv=True)
            pfs_tsv, e2 = run_sql("SELECT project_id,name,contract_date,total_price,payment_count,payment_total_amount,outstanding_amount FROM project_finance_summary ORDER BY project_id", expect_tsv=True)
            od_tsv, e3 = run_sql("SELECT project_id,name,contract_date,total_price,paid_amount,outstanding_amount,days_since_contract FROM overdue_projects_90d ORDER BY days_since_contract DESC", expect_tsv=True)
            mc_tsv, e4 = run_sql("SELECT ym,project_count,contract_amount FROM monthly_contracts ORDER BY ym", expect_tsv=True)
            mcd_tsv, e5 = run_sql("SELECT DATE_FORMAT(contract_date,'%Y-%m') AS ym,id,name,contract_date,total_price FROM projects ORDER BY contract_date,id", expect_tsv=True)
            mp_tsv, e6 = run_sql("SELECT ym,payment_count,payment_amount FROM monthly_payments ORDER BY ym", expect_tsv=True)
            mpd_tsv, e7 = run_sql("SELECT DATE_FORMAT(payments.payment_date,'%Y-%m') AS ym,payments.id,projects.name AS project_name,payments.payment_date,payments.amount FROM payments LEFT JOIN projects ON projects.id=payments.project_id ORDER BY payments.payment_date,payments.id", expect_tsv=True)
            qc_tsv, e8 = run_sql("SELECT yq,project_count,contract_amount FROM quarterly_contracts ORDER BY yq", expect_tsv=True)
            qcd_tsv, e9 = run_sql("SELECT CONCAT(YEAR(contract_date), '-Q', QUARTER(contract_date)) AS yq,id,name,contract_date,total_price FROM projects ORDER BY contract_date,id", expect_tsv=True)
            qp_tsv, e10 = run_sql("SELECT yq,payment_count,payment_amount FROM quarterly_payments ORDER BY yq", expect_tsv=True)
            qpd_tsv, e11 = run_sql("SELECT CONCAT(YEAR(payments.payment_date), '-Q', QUARTER(payments.payment_date)) AS yq,payments.id,projects.name AS project_name,payments.payment_date,payments.amount FROM payments LEFT JOIN projects ON projects.id=payments.project_id ORDER BY payments.payment_date,payments.id", expect_tsv=True)
            yc_tsv, e12 = run_sql("SELECT yr,project_count,contract_amount FROM yearly_contracts ORDER BY yr", expect_tsv=True)
            ycd_tsv, e13 = run_sql("SELECT YEAR(contract_date) AS yr,id,name,contract_date,total_price FROM projects ORDER BY contract_date,id", expect_tsv=True)
            yp_tsv, e14 = run_sql("SELECT yr,payment_count,payment_amount FROM yearly_payments ORDER BY yr", expect_tsv=True)
            ypd_tsv, e15 = run_sql("SELECT YEAR(payments.payment_date) AS yr,payments.id,projects.name AS project_name,payments.payment_date,payments.amount FROM payments LEFT JOIN projects ON projects.id=payments.project_id ORDER BY payments.payment_date,payments.id", expect_tsv=True)
            errs = [x for x in [e1,e2,e3,e4,e5,e6,e7,e8,e9,e10,e11,e12,e13,e14,e15] if x]
            if errs:
                self.json_response({"error": errs[0]}, 500); return
            sheets["全量合计"] = tsv_to_json(gf_tsv)
            sheets["项目汇总明细"] = tsv_to_json(pfs_tsv)
            sheets["逾期未回款>90天"] = tsv_to_json(od_tsv)
            sheets["月签约汇总"] = tsv_to_json(mc_tsv)
            sheets["月签约明细"] = tsv_to_json(mcd_tsv)
            sheets["月回款汇总"] = tsv_to_json(mp_tsv)
            sheets["月回款明细"] = tsv_to_json(mpd_tsv)
            sheets["季度签约汇总"] = tsv_to_json(qc_tsv)
            sheets["季度签约明细"] = tsv_to_json(qcd_tsv)
            sheets["季度回款汇总"] = tsv_to_json(qp_tsv)
            sheets["季度回款明细"] = tsv_to_json(qpd_tsv)
            sheets["年度签约汇总"] = tsv_to_json(yc_tsv)
            sheets["年度签约明细"] = tsv_to_json(ycd_tsv)
            sheets["年度回款汇总"] = tsv_to_json(yp_tsv)
            sheets["年度回款明细"] = tsv_to_json(ypd_tsv)
            data = self.build_xlsx_multi(sheets)
            fname = "ledger_all.xlsx"
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f"attachment; filename={fname}")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        self.send_error(404)

    def build_xlsx(self, projects_rows, payments_rows):
        import io, zipfile, datetime
        def ct_xml():
            return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>""").encode("utf-8")
        def rels_root_xml():
            return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>""").encode("utf-8")
        def wb_xml():
            return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr/>
  <sheets>
    <sheet name="项目" sheetId="1" r:id="rId1"/>
    <sheet name="回款" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>""").encode("utf-8")
        def wb_rels_xml():
            return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>""").encode("utf-8")
        def styles_xml():
            return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></styleSheet>""").encode("utf-8")
        def core_xml():
            return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Ledger Export</dc:title>
  <dc:creator>ledger_app</dc:creator>
  <cp:lastModifiedBy>ledger_app</cp:lastModifiedBy>
</cp:coreProperties>""").encode("utf-8")
        def app_xml():
            return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Ledger</Application>
</Properties>""").encode("utf-8")
        def col_letter(idx):
            s = ""
            idx0 = idx
            while True:
                s = chr(ord('A') + (idx0 % 26)) + s
                idx0 = idx0 // 26 - 1
                if idx0 < 0:
                    break
            return s
        def sheet_xml(headers, rows):
            def esc(s):
                return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
            parts = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
                     "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
                     "<sheetData>"]
            # header row
            r = 1
            cells = []
            for c_idx, h in enumerate(headers):
                addr = f"{col_letter(c_idx)}{r}"
                cells.append(f"<c r=\"{addr}\" t=\"inlineStr\"><is><t>{esc(h)}</t></is></c>")
            parts.append(f"<row r=\"{r}\">{''.join(cells)}</row>")
            # data rows
            for i, row in enumerate(rows, start=2):
                cells = []
                for c_idx, h in enumerate(headers):
                    v = row.get(h, "")
                    addr = f"{col_letter(c_idx)}{i}"
                    try:
                        if v is None:
                            cells.append(f"<c r=\"{addr}\"/>")
                        else:
                            # numeric
                            float_val = float(v)
                            cells.append(f"<c r=\"{addr}\"><v>{float_val}</v></c>")
                    except Exception:
                        cells.append(f"<c r=\"{addr}\" t=\"inlineStr\"><is><t>{esc(v)}</t></is></c>")
                parts.append(f"<row r=\"{i}\">{''.join(cells)}</row>")
            parts.append("</sheetData></worksheet>")
            return "".join(parts).encode("utf-8")
        # prepare data
        proj_headers = ["id","name","contract_date","total_price"]
        pay_headers = ["id","project_name","payment_date","amount"]
        proj_rows = projects_rows
        pay_rows = payments_rows
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct_xml())
            z.writestr("_rels/.rels", rels_root_xml())
            z.writestr("xl/workbook.xml", wb_xml())
            z.writestr("xl/_rels/workbook.xml.rels", wb_rels_xml())
            z.writestr("xl/styles.xml", styles_xml())
            z.writestr("docProps/core.xml", core_xml())
            z.writestr("docProps/app.xml", app_xml())
            z.writestr("xl/worksheets/sheet1.xml", sheet_xml(proj_headers, proj_rows))
            z.writestr("xl/worksheets/sheet2.xml", sheet_xml(pay_headers, pay_rows))
        return buf.getvalue()

    def build_xlsx_multi(self, sheets_dict):
        import io, zipfile
        def ct_xml(sheet_count):
            parts = [
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">",
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>",
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>",
                "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>",
                "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>",
                "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>",
                "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>",
            ]
            for i in range(1, sheet_count+1):
                parts.append(f"<Override PartName=\"/xl/worksheets/sheet{i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>")
            parts.append("</Types>")
            return "".join(parts).encode("utf-8")
        def rels_root_xml():
            return ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>""").encode("utf-8")
        def wb_xml(sheet_names):
            sheets_xml = "".join([f"<sheet name=\"{s}\" sheetId=\"{i}\" r:id=\"rId{i}\"/>" for i,s in enumerate(sheet_names, start=1)])
            return (f"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    f"<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><workbookPr/><sheets>{sheets_xml}</sheets></workbook>").encode("utf-8")
        def wb_rels_xml(sheet_count):
            rels = [f"<Relationship Id=\"rId{i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{i}.xml\"/>" for i in range(1, sheet_count+1)]
            rels.append("<Relationship Id=\"rId999\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>")
            return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                    + "".join(rels) + "</Relationships>").encode("utf-8")
        def styles_xml():
            return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></styleSheet>").encode("utf-8")
        def core_xml():
            return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:title>Ledger Export</dc:title><dc:creator>ledger_app</dc:creator><cp:lastModifiedBy>ledger_app</cp:lastModifiedBy></cp:coreProperties>").encode("utf-8")
        def app_xml():
            return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Application>Ledger</Application></Properties>").encode("utf-8")
        def col_letter(idx):
            s = ""
            x = idx
            while True:
                s = chr(ord('A') + (x % 26)) + s
                x = x // 26 - 1
                if x < 0:
                    break
            return s
        def sheet_xml(headers_disp, headers_keys, rows):
            def esc(s):
                return (str(s).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"))
            parts = ["<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>","<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"]
            r = 1
            cells = []
            for c,h in enumerate(headers_disp):
                addr = f"{col_letter(c)}{r}"
                cells.append(f"<c r=\"{addr}\" t=\"inlineStr\"><is><t>{esc(h)}</t></is></c>")
            parts.append(f"<row r=\"{r}\">{''.join(cells)}</row>")
            for i,row in enumerate(rows, start=2):
                cells = []
                for c,key in enumerate(headers_keys):
                    v = row.get(key, "")
                    addr = f"{col_letter(c)}{i}"
                    try:
                        if v is None:
                            cells.append(f"<c r=\"{addr}\"/>")
                        else:
                            fv = float(v)
                            cells.append(f"<c r=\"{addr}\"><v>{fv}</v></c>")
                    except Exception:
                        cells.append(f"<c r=\"{addr}\" t=\"inlineStr\"><is><t>{esc(v)}</t></is></c>")
                parts.append(f"<row r=\"{i}\">{''.join(cells)}</row>")
            parts.append("</sheetData></worksheet>")
            return "".join(parts).encode("utf-8")
        label_map = {
            'global_finance_totals.csv': {
                'total_projects':'项目总数','total_price_all_projects':'所有项目合计总价','total_outstanding_all_projects':'所有合计未回款总价','total_payments_amount':'回款总额','total_payments_count':'回款笔数'
            },
            'project_finance_summary.csv': {
                'project_id':'项目ID','name':'项目名称','contract_date':'签约时间','total_price':'总价','payment_count':'回款笔数','payment_total_amount':'回款总额','outstanding_amount':'未回款总额'
            },
            'monthly_contracts.csv': {'ym':'年月','project_count':'项目数','contract_amount':'签约金额'},
            'monthly_contracts_detail.csv': {'ym':'年月','id':'项目ID','name':'项目名称','contract_date':'签约时间','total_price':'总价'},
            'monthly_payments.csv': {'ym':'年月','payment_count':'回款笔数','payment_amount':'回款金额'},
            'monthly_payments_detail.csv': {'ym':'年月','id':'回款ID','project_name':'项目名称','payment_date':'回款日期','amount':'金额'},
            'quarterly_contracts.csv': {'yq':'季度','project_count':'项目数','contract_amount':'签约金额'},
            'quarterly_contracts_detail.csv': {'yq':'季度','id':'项目ID','name':'项目名称','contract_date':'签约时间','total_price':'总价'},
            'quarterly_payments.csv': {'yq':'季度','payment_count':'回款笔数','payment_amount':'回款金额'},
            'quarterly_payments_detail.csv': {'yq':'季度','id':'回款ID','project_name':'项目名称','payment_date':'回款日期','amount':'金额'},
            'yearly_contracts.csv': {'yr':'年份','project_count':'项目数','contract_amount':'签约金额'},
            'yearly_contracts_detail.csv': {'yr':'年份','id':'项目ID','name':'项目名称','contract_date':'签约时间','total_price':'总价'},
            'yearly_payments.csv': {'yr':'年份','payment_count':'回款笔数','payment_amount':'回款金额'},
            'yearly_payments_detail.csv': {'yr':'年份','id':'回款ID','project_name':'项目名称','payment_date':'回款日期','amount':'金额'},
            'overdue_projects_90d.csv': {'project_id':'项目ID','name':'项目名称','contract_date':'签约时间','total_price':'总价','paid_amount':'已回款','outstanding_amount':'未回款','days_since_contract':'签约至今(天)'}
        }
        sheet_names = list(sheets_dict.keys())
        sheet_files = {
            "全量合计":"global_finance_totals.csv",
            "项目汇总明细":"project_finance_summary.csv",
            "逾期未回款>90天":"overdue_projects_90d.csv",
            "月签约汇总":"monthly_contracts.csv",
            "月签约明细":"monthly_contracts_detail.csv",
            "月回款汇总":"monthly_payments.csv",
            "月回款明细":"monthly_payments_detail.csv",
            "季度签约汇总":"quarterly_contracts.csv",
            "季度签约明细":"quarterly_contracts_detail.csv",
            "季度回款汇总":"quarterly_payments.csv",
            "季度回款明细":"quarterly_payments_detail.csv",
            "年度签约汇总":"yearly_contracts.csv",
            "年度签约明细":"yearly_contracts_detail.csv",
            "年度回款汇总":"yearly_payments.csv",
            "年度回款明细":"yearly_payments_detail.csv"
        }
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct_xml(len(sheet_names)))
            z.writestr("_rels/.rels", rels_root_xml())
            z.writestr("xl/workbook.xml", wb_xml(sheet_names))
            z.writestr("xl/_rels/workbook.xml.rels", wb_rels_xml(len(sheet_names)))
            z.writestr("xl/styles.xml", styles_xml())
            z.writestr("docProps/core.xml", core_xml())
            z.writestr("docProps/app.xml", app_xml())
            for i, name in enumerate(sheet_names, start=1):
                rows = sheets_dict.get(name, [])
                keys = list(rows[0].keys()) if rows else []
                labels = label_map.get(sheet_files[name], {})
                headers_disp = [labels.get(k, k) for k in keys]
                z.writestr(f"xl/worksheets/sheet{i}.xml", sheet_xml(headers_disp, keys, rows))
        return buf.getvalue()

    def handle_api_write(self, method):
        parsed = urlparse(self.path)
        path = parsed.path
        body = self.read_json()
        global ADMIN_PASS
        if path == "/api/login" and method == "POST":
            u = (body.get("username") or "").strip()
            p = (body.get("password") or "").strip()
            if u == ADMIN_USER and p == ADMIN_PASS:
                import secrets
                token = secrets.token_hex(16)
                import time
                SESSIONS[token] = {"user": u, "last": time.time()}
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Set-Cookie", f"SESSIONID={token}; Path=/; HttpOnly")
                data = json.dumps({"success": True}).encode("utf-8")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return
            else:
                self.json_response({"error": "invalid"}, 401)
                return
        if path == "/api/logout" and method == "POST":
            cookie = self.headers.get("Cookie")
            if cookie and "SESSIONID=" in cookie:
                try:
                    token = [p.strip() for p in cookie.split(";") if p.strip().startswith("SESSIONID=")][0].split("=",1)[1]
                    SESSIONS.pop(token, None)
                except Exception:
                    pass
            self.send_response(200)
            self.send_header("Set-Cookie", "SESSIONID=; Path=/; Max-Age=0")
            self.end_headers()
            return
        if path == "/api/forgot_password" and method == "POST":
            email = (body.get("email") or "").strip()
            import secrets, time
            token = secrets.token_hex(16)
            RESET_TOKENS[token] = {"email": email, "exp": time.time()+3600}
            link = f"/reset.html?token={token}"
            self.json_response({"success": True, "link": link})
            return
        if path == "/api/reset_password" and method == "POST":
            token = (body.get("token") or "").strip()
            newp = (body.get("new_password") or "").strip()
            info = RESET_TOKENS.get(token)
            import time
            if not info or info.get("exp",0) < time.time():
                self.json_response({"error": "invalid_token"}, 400)
                return
            RESET_TOKENS.pop(token, None)
            ADMIN_PASS = newp or ADMIN_PASS
            try:
                with open(ADMIN_FILE, "w", encoding="utf-8") as f:
                    json.dump({"username": ADMIN_USER, "password": ADMIN_PASS}, f, ensure_ascii=False)
            except Exception:
                pass
            self.json_response({"success": True})
            return
        if path == "/api/update_admin" and method == "POST":
            if not self.get_session_user():
                self.json_response({"error": "unauthorized"}, 401); return
            newu = (body.get("username") or "").strip() or ADMIN_USER
            newp = (body.get("password") or "").strip() or ADMIN_PASS
            globals()["ADMIN_USER"] = newu
            globals()["ADMIN_PASS"] = newp
            try:
                with open(ADMIN_FILE, "w", encoding="utf-8") as f:
                    json.dump({"username": ADMIN_USER, "password": ADMIN_PASS}, f, ensure_ascii=False)
            except Exception:
                pass
            self.json_response({"success": True})
            return
        if path == "/api/projects" and method == "POST":
            if not self.get_session_user():
                self.json_response({"error": "unauthorized"}, 401); return
            name = (body.get("name", "") or "").replace("'", "''")
            date = body.get("contract_date", "") or ""
            price = body.get("total_price", 0) or 0
            if not name or not date:
                self.json_response({"error": "name and contract_date required"}, 400); return
            ok, err = run_sql(f"INSERT INTO projects(name,contract_date,total_price) VALUES('{name}','{date}',{float(price)})")
            if not ok:
                self.json_response({"error": err}, 500)
                return
            self.json_response({"success": True})
            return
        if path.startswith("/api/projects/") and method == "PUT":
            if not self.get_session_user():
                self.json_response({"error": "unauthorized"}, 401); return
            pid = int(path.split("/")[-1])
            fields = []
            if "name" in body:
                fields.append(f"name='{body['name'].replace("'","''")}'")
            if "contract_date" in body:
                fields.append(f"contract_date='{body['contract_date']}'")
            if "total_price" in body:
                fields.append(f"total_price={float(body['total_price'])}")
            if not fields:
                self.json_response({"error": "no fields"}, 400)
                return
            ok, err = run_sql(f"UPDATE projects SET {', '.join(fields)} WHERE id={pid}")
            if not ok:
                self.json_response({"error": err}, 500)
                return
            self.json_response({"success": True})
            return
        if path.startswith("/api/projects/") and method == "DELETE":
            if not self.get_session_user():
                self.json_response({"error": "unauthorized"}, 401); return
            pid = int(path.split("/")[-1])
            ok, err = run_sql(f"DELETE FROM projects WHERE id={pid}")
            if not ok:
                self.json_response({"error": err}, 500)
                return
            self.json_response({"success": True})
            return
        if path == "/api/payments" and method == "POST":
            if not self.get_session_user():
                self.json_response({"error": "unauthorized"}, 401); return
            project_id = int(body.get("project_id", 0))
            payment_date = body.get("payment_date", "")
            amount = float(body.get("amount", 0))
            ok, err = run_sql(f"INSERT INTO payments(project_id,payment_date,amount) VALUES({project_id},'{payment_date}',{amount})")
            if not ok:
                self.json_response({"error": err}, 500)
                return
            self.json_response({"success": True})
            return
        if path.startswith("/api/payments/") and method == "PUT":
            if not self.get_session_user():
                self.json_response({"error": "unauthorized"}, 401); return
            pid = int(path.split("/")[-1])
            fields = []
            if "payment_date" in body:
                fields.append(f"payment_date='{body['payment_date']}'")
            if "amount" in body:
                fields.append(f"amount={float(body['amount'])}")
            if not fields:
                self.json_response({"error": "no fields"}, 400)
                return
            ok, err = run_sql(f"UPDATE payments SET {', '.join(fields)} WHERE id={pid}")
            if not ok:
                self.json_response({"error": err}, 500)
                return
            self.json_response({"success": True})
            return
        if path.startswith("/api/payments/") and method == "DELETE":
            if not self.get_session_user():
                self.json_response({"error": "unauthorized"}, 401); return
            pid = int(path.split("/")[-1])
            ok, err = run_sql(f"DELETE FROM payments WHERE id={pid}")
            if not ok:
                self.json_response({"error": err}, 500)
                return
            self.json_response({"success": True})
            return
        self.send_error(404)

def run(port=8000):
    os.chdir(ROOT)
    with ThreadingHTTPServer(("", port), Handler) as httpd:
        httpd.serve_forever()

if __name__ == "__main__":
    port_env = os.environ.get("PORT") or os.environ.get("LEDGER_PORT")
    try:
        port = int(port_env) if port_env else 8000
    except ValueError:
        port = 8000
    run(port)