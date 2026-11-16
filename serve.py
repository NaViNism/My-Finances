#!/usr/bin/env python3
"""
FinanceFlow Local Server (Dual Port 8080 & 8081 with No-Cache Headers)
"""

import http.server
import socket
import socketserver
import threading
import os
import sys

DIRECTORY = os.path.dirname(os.path.abspath(__file__))

class NoCacheHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def end_headers(self):
        # Strict anti-cache headers for iPhone Safari
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0, post-check=0, pre-check=0')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()

    def log_message(self, format, *args):
        client_ip = self.client_address[0] if hasattr(self, 'client_address') else 'client'
        sys.stderr.write(f"[{self.log_date_time_string()}] [{client_ip}] {format % args}\n")

def get_local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('10.255.255.255', 1))
        IP = s.getsockname()[0]
    except Exception:
        IP = '127.0.0.1'
    finally:
        s.close()
    return IP

def start_server_on_port(port):
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", port), NoCacheHandler) as httpd:
        httpd.serve_forever()

def main():
    os.chdir(DIRECTORY)
    local_ip = get_local_ip()

    print("=" * 66)
    print(" 🚀 FinanceFlow Dual Server Running (No-Cache Mode)")
    print("=" * 66)
    print(f" 👉 FRESH PORT (Bypasses all old iOS cache):")
    print(f"    📱 http://{local_ip}:8081")
    print(f"    💻 http://localhost:8081")
    print("-" * 66)
    print(f" 👉 Standard Port:")
    print(f"    📱 http://{local_ip}:8080")
    print("=" * 66)

    # Start 8080 in background thread
    t8080 = threading.Thread(target=start_server_on_port, args=(8080,), daemon=True)
    t8080.start()

    # Run 8081 in main thread
    start_server_on_port(8081)

if __name__ == '__main__':
    main()
