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
            self.handle_api_get()
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

    def handle_api_get(self):
        parsed = urlparse(self.path)
        path = parsed.path
        qs = parse_qs(parsed.query)
        if path == "/api/health":
            self.json_response({"ok": True})
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
            tsv, err = run_sql(f"SELECT id,project_id,payment_date,amount,created_at FROM payments WHERE project_id={int(pid)} ORDER BY id", expect_tsv=True)
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
        if path == "/api/report/overdue_projects_90d":
            tsv, err = run_sql("SELECT project_id,name,contract_date,total_price,paid_amount,outstanding_amount,days_since_contract FROM overdue_projects_90d ORDER BY days_since_contract DESC", expect_tsv=True)
            if err:
                self.json_response({"error": err}, 500); return
            self.json_response(tsv_to_json(tsv)); return
        self.send_error(404)

    def handle_api_write(self, method):
        parsed = urlparse(self.path)
        path = parsed.path
        body = self.read_json()
        if path == "/api/projects" and method == "POST":
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
            pid = int(path.split("/")[-1])
            ok, err = run_sql(f"DELETE FROM projects WHERE id={pid}")
            if not ok:
                self.json_response({"error": err}, 500)
                return
            self.json_response({"success": True})
            return
        if path == "/api/payments" and method == "POST":
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