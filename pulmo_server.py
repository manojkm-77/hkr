"""
PulmoDoc Suite — Local Standalone HTTP Application Server
Zero-dependency Python 3 standard library server for Medical PG Clinician Suite.
"""
import os, sys, json, mimetypes, webbrowser, socket
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, unquote

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
APP_DIR = os.path.join(ROOT_DIR, "pulmo_app")
STATIC_DIR = os.path.join(APP_DIR, "static")
DATA_DIR = os.path.join(APP_DIR, "data")

# Image Directories
IMAGE_DIRS = {
    "/assets/thoracic_atlas_assets/": os.path.join(ROOT_DIR, "thoracic_atlas_assets"),
    "/assets/pneumonia_figures/": os.path.join(ROOT_DIR, "pneumonia_figures"),
    "/assets/bronchiectasis_figures/": os.path.join(ROOT_DIR, "bronchiectasis_figures"),
    "/assets/asthma_assets/": os.path.join(ROOT_DIR, "asthma_assets"),
}

class PulmoHandler(BaseHTTPRequestHandler):
    def end_headers(self):
        # Enable CORS and caching headers
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Cache-Control', 'no-cache, must-revalidate')
        super().end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        path = unquote(parsed.path)

        # 1. API Endpoints
        if path.startswith("/api/"):
            self.handle_api(path)
            return

        # 2. Asset Images
        for prefix, dir_path in IMAGE_DIRS.items():
            if path.startswith(prefix):
                filename = path[len(prefix):]
                file_full_path = os.path.join(dir_path, filename)
                if os.path.exists(file_full_path) and os.path.isfile(file_full_path):
                    self.serve_file(file_full_path)
                    return
                else:
                    self.send_error(404, f"Image {filename} not found")
                    return

        # 3. Static Web Files
        if path == "/" or path == "":
            file_to_serve = os.path.join(STATIC_DIR, "index.html")
        else:
            file_to_serve = os.path.join(STATIC_DIR, path.lstrip("/"))

        if os.path.exists(file_to_serve) and os.path.isfile(file_to_serve):
            self.serve_file(file_to_serve)
        elif os.path.exists(os.path.join(ROOT_DIR, path.lstrip("/"))) and os.path.isfile(os.path.join(ROOT_DIR, path.lstrip("/"))):
            self.serve_file(os.path.join(ROOT_DIR, path.lstrip("/")))
        else:
            # Fallback to index.html for SPA routing
            index_path = os.path.join(STATIC_DIR, "index.html")
            if os.path.exists(index_path):
                self.serve_file(index_path)
            else:
                self.send_error(404, "File not found")

    def handle_api(self, path):
        api_route = path.replace("/api/", "").strip("/")
        
        file_map = {
            "radiology": "radiology_cases.json",
            "mcqs": "mcq_bank.json",
            "chapters": "chapters_summary.json",
            "crofton": "crofton_encyclopedia.json",
            "murray": "murray_compendium.json",
            "felson": "felson_roentgen_signs.json",
            "felson-full": "felson_full_book.json",
            "harrison": "harrison_clinical_algorithms.json",
            "harrison-full": "harrison_22e_full.json",
            "fishman": "fishman_2026_full.json",
            "clinical-exam": "clinical_examination_guide.json",
            "orders": "order_sheets.json",
            "pharmacology": "pharmacology.json",
            "viva": "viva_stations.json"
        }

        if api_route in file_map:
            target_file = os.path.join(DATA_DIR, file_map[api_route])
            if os.path.exists(target_file):
                with open(target_file, "r", encoding="utf-8") as f:
                    content = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(content.encode("utf-8"))
                return
            else:
                self.send_error(404, f"Dataset {file_map[api_route]} not built yet.")
                return

        elif api_route == "stats":
            stats = {
                "radiology_cases": self.count_json_records("radiology_cases.json"),
                "mcq_count": self.count_json_records("mcq_bank.json"),
                "core_chapters": self.count_json_records("chapters_summary.json"),
                "crofton_chapters": self.count_json_records("crofton_encyclopedia.json"),
                "murray_chapters": self.count_json_records("murray_compendium.json"),
                "felson_signs": self.count_json_records("felson_roentgen_signs.json"),
                "orders_count": self.count_json_records("order_sheets.json"),
                "status": "online",
                "version": "2.0.0"
            }
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps(stats).encode("utf-8"))
            return

        self.send_error(404, f"Unknown API endpoint {path}")

    def count_json_records(self, fname):
        fpath = os.path.join(DATA_DIR, fname)
        if os.path.exists(fpath):
            try:
                with open(fpath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    return len(data) if isinstance(data, list) else len(data.keys())
            except Exception:
                return 0
        return 0

    def serve_file(self, filepath):
        ctype, _ = mimetypes.guess_type(filepath)
        if not ctype:
            if filepath.endswith(".js"):
                ctype = "application/javascript"
            elif filepath.endswith(".css"):
                ctype = "text/css"
            elif filepath.endswith(".json"):
                ctype = "application/json"
            else:
                ctype = "application/octet-stream"

        try:
            with open(filepath, "rb") as f:
                content = f.read()
            self.send_response(200)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
        except Exception as e:
            self.send_error(500, f"Error reading file: {str(e)}")

    def log_message(self, format, *args):
        # Clean logging
        sys.stdout.write(f"[{self.log_date_time_string()}] {format % args}\n")
        sys.stdout.flush()

def find_available_port(start_port=8000, max_attempts=10):
    for p in range(start_port, start_port + max_attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('127.0.0.1', p)) != 0:
                return p
    return start_port

def start_server():
    port = find_available_port(8000)
    server_address = ('0.0.0.0', port)
    httpd = HTTPServer(server_address, PulmoHandler)
    local_url = f"http://127.0.0.1:{port}"
    network_url = f"http://172.28.100.184:{port}"

    print("=" * 65)
    print("  🩺 PULMODOC SUITE 2.1 — COMPREHENSIVE MEDICAL PG SUITE")
    print(f"  • Laptop / Desktop: {local_url}")
    print(f"  • Android Phone / Tablet (Same Wi-Fi): {network_url}")
    print("  Complete Multi-Textbook Library Active:")
    print("    • Harrison's 22e Full Respiratory Textbook (808 Sections)")
    print("    • Fishman's Pulmonary Diseases 2026 Updates (759 Sections)")
    print("    • Felson's Principles of Roentgenology (Complete 14 Chapters)")
    print("    • Crofton & Douglas's Respiratory Diseases (61 Chapters)")
    print("    • Murray & Nadel's Textbook 7th Ed (36 Chapters Compendium)")
    print("    • 12 Core PG Modules with 150+ Clinical Tables")
    print("    • Bedside Clinical Examination (I-P-P-A) & 5 MD Long Cases")
    print("    • 105+ High-Yield Medical PG MCQs & Spotters")
    print("    • 48 Thoracic Radiology Cases with Interactive Canvas")
    print("    • Real-time Physiological Calculators (ABG, A-a, CURB, Wells, GOLD)")
    print("    • Inpatient & RICU Emergency Protocols (10 Bundles)")
    print("    • Pharmacology Monographs & Device Techniques")
    print("    • MD Table Viva Stations")
    print("=" * 65)
    print("Opening browser automatically...")
    
    try:
        webbrowser.open(local_url)
    except Exception as e:
        print(f"Could not open browser automatically: {e}")

    print("Press Ctrl+C to stop server.\n")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping PulmoDoc Server...")
        httpd.server_close()
        sys.exit(0)

if __name__ == "__main__":
    start_server()
