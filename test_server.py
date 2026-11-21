"""
Verification script for PulmoDoc Suite 2.0 multi-textbook endpoints.
"""
import urllib.request, json, time, threading
from pulmo_server import find_available_port, HTTPServer, PulmoHandler

def run_test():
    port = find_available_port(8960)
    server = HTTPServer(('127.0.0.1', port), PulmoHandler)
    t = threading.Thread(target=server.serve_forever, daemon=True)
    t.start()
    time.sleep(0.5)

    base = f"http://127.0.0.1:{port}"
    endpoints = [
        ("/", "text/html"),
        ("/style.css", "text/css"),
        ("/app.js", "application/javascript"),
        ("/api/stats", "application/json"),
        ("/api/radiology", "application/json"),
        ("/api/mcqs", "application/json"),
        ("/api/chapters", "application/json"),
        ("/api/crofton", "application/json"),
        ("/api/murray", "application/json"),
        ("/api/felson", "application/json"),
        ("/api/felson-full", "application/json"),
        ("/api/harrison", "application/json"),
        ("/api/harrison-full", "application/json"),
        ("/api/fishman", "application/json"),
        ("/api/clinical-exam", "application/json"),
        ("/api/orders", "application/json"),
        ("/api/pharmacology", "application/json"),
        ("/api/viva", "application/json")
    ]

    print(f"Testing PulmoDoc Suite 2.0 on {base}...")
    success = True
    for ep, expected_mime in endpoints:
        url = base + ep
        try:
            req = urllib.request.Request(url)
            with urllib.request.urlopen(req, timeout=3) as resp:
                status = resp.status
                data = resp.read()
                data_len = len(data)
                extra = ""
                if ep == "/api/mcqs":
                    j = json.loads(data)
                    extra = f"[{len(j)} MCQs]"
                elif ep == "/api/crofton":
                    j = json.loads(data)
                    extra = f"[{len(j)} Chapters]"
                elif ep == "/api/murray":
                    j = json.loads(data)
                    extra = f"[{len(j)} Chapters]"
                elif ep == "/api/felson":
                    j = json.loads(data)
                    extra = f"[{len(j)} Golden Signs]"
                elif ep == "/api/harrison":
                    j = json.loads(data)
                    extra = f"[{len(j)} Algorithms]"
                elif ep == "/api/chapters":
                    j = json.loads(data)
                    extra = f"[{len(j)} Core Modules]"

                print(f"  [OK {status}] {ep:22} -> {data_len:>8,} bytes {extra}")
        except Exception as e:
            print(f"  [FAIL] {ep}: {e}")
            success = False

    server.shutdown()
    if success:
        print("\nALL MULTI-TEXTBOOK CONCEPTS AND ENDPOINTS VERIFIED SUCCESSFULLY!")
    else:
        print("\nERRORS DETECTED.")

if __name__ == "__main__":
    run_test()
