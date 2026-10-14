import os
import sys
import secrets
import subprocess
from pathlib import Path
from typing import List, Dict, Any

from fastapi import FastAPI, HTTPException, Depends, status
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from pydantic import BaseModel

# Inisialisasi FastAPI App
app = FastAPI(
    title="Generator ST & Laporan API",
    description="Backend API dan Web Interface untuk Sinkronisasi Spreadsheet dan Generator Dokumen Word",
    version="1.0.0"
)

# Aktifkan CORS agar Google Apps Script atau frontend lain dapat mengakses API dengan leluasa
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).resolve().parent

# File Kredensial Pengguna
USER_FILE = BASE_DIR / "username.txt"
security = HTTPBasic()

def get_valid_users() -> Dict[str, str]:
    """Membaca daftar username dan password dari file username.txt."""
    users = {}
    if not USER_FILE.exists():
        USER_FILE.write_text("admin:bapperida2026\n", encoding="utf-8")
    
    with open(USER_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if ":" in line:
                u, p = line.split(":", 1)
                users[u.strip()] = p.strip()
    return users

def authenticate(credentials: HTTPBasicCredentials = Depends(security)):
    """Memvalidasi kredensial HTTP Basic Auth terhadap username.txt."""
    users = get_valid_users()
    current_username = credentials.username
    current_password = credentials.password
    
    if current_username in users:
        correct_password = users[current_username]
        if secrets.compare_digest(current_password.encode("utf-8"), correct_password.encode("utf-8")):
            return current_username
            
    raise HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Username atau Password salah.",
        headers={"WWW-Authenticate": 'Basic realm="Restricted Area - Sistem ST & Laporan"'},
    )

# Path Skrip Python
SYNC_SCRIPT = BASE_DIR / "aset" / "sh" / "ambildarispreadsheet.py"
GENLAP_SCRIPT = BASE_DIR / "aset" / "Lap" / "genlap.py"
GENST_SCRIPT = BASE_DIR / "aset" / "ST" / "aset" / "generator.py"

# Folder Output
ST_OUTPUT_DIR = BASE_DIR / "ST"
LAP_OUTPUT_DIR = BASE_DIR / "Lap"

def run_python_script(script_path: Path) -> Dict[str, Any]:
    """Menjalankan file script python menggunakan executable saat ini dan mengembalikan outputnya."""
    if not script_path.exists():
        return {
            "success": False,
            "error": f"File skrip '{script_path}' tidak ditemukan."
        }
    
    try:
        result = subprocess.run(
            [sys.executable, str(script_path)],
            cwd=str(BASE_DIR),
            capture_output=True,
            text=True,
            timeout=180
        )
        output = (result.stdout or "") + ("\n" + result.stderr if result.stderr else "")
        return {
            "success": result.returncode == 0,
            "returncode": result.returncode,
            "output": output.strip(),
            "error": result.stderr.strip() if result.returncode != 0 else None
        }
    except subprocess.TimeoutExpired:
        return {
            "success": False,
            "error": "Proses melebihi batas waktu (timeout 180 detik)."
        }
    except Exception as e:
        return {
            "success": False,
            "error": str(e)
        }

def list_generated_files(folder_path: Path) -> List[Dict[str, Any]]:
    """Mendaftar file .docx yang ada di folder output beserta ukurannya."""
    if not folder_path.exists():
        return []
    files = []
    for f in folder_path.glob("*.docx"):
        files.append({
            "name": f.name,
            "size_kb": round(f.stat().st_size / 1024, 2),
            "modified": f.stat().st_mtime
        })
    return sorted(files, key=lambda x: x["modified"], reverse=True)

# ==================== ENDPOINT API ====================

@app.get("/health")
def health_check():
    """Health check endpoint untuk Render.com dan monitoring (Publik)."""
    return {"status": "ok", "message": "Service is running healthy."}

@app.post("/api/sync")
def trigger_sync(user: str = Depends(authenticate)):
    """Menjalankan sinkronisasi data dari Google Spreadsheet."""
    res = run_python_script(SYNC_SCRIPT)
    if not res["success"]:
        raise HTTPException(status_code=500, detail=res)
    return {
        "status": "success",
        "action": "sync",
        "message": "Sinkronisasi data dari Google Spreadsheet berhasil.",
        "log": res["output"]
    }

@app.post("/api/generate-st")
def trigger_generate_st(user: str = Depends(authenticate)):
    """Menjalankan generator Surat Tugas (ST)."""
    res = run_python_script(GENST_SCRIPT)
    if not res["success"]:
        raise HTTPException(status_code=500, detail=res)
    files = list_generated_files(ST_OUTPUT_DIR)
    return {
        "status": "success",
        "action": "generate_st",
        "message": "Dokumen Surat Tugas berhasil digenerate.",
        "log": res["output"],
        "files_count": len(files),
        "files": files
    }

@app.post("/api/generate-laporan")
def trigger_generate_laporan(user: str = Depends(authenticate)):
    """Menjalankan generator Laporan (Lap)."""
    res = run_python_script(GENLAP_SCRIPT)
    if not res["success"]:
        raise HTTPException(status_code=500, detail=res)
    files = list_generated_files(LAP_OUTPUT_DIR)
    return {
        "status": "success",
        "action": "generate_laporan",
        "message": "Dokumen Laporan berhasil digenerate.",
        "log": res["output"],
        "files_count": len(files),
        "files": files
    }

@app.post("/api/run-all")
def trigger_run_all(user: str = Depends(authenticate)):
    """Menjalankan Sinkronisasi, lalu Generate ST, lalu Generate Laporan sekaligus."""
    logs = []
    
    # 1. Sync
    sync_res = run_python_script(SYNC_SCRIPT)
    logs.append(f"=== [1] SINKRONISASI SPREADSHEET ===\n{sync_res.get('output') or sync_res.get('error')}")
    if not sync_res["success"]:
        raise HTTPException(status_code=500, detail={"step": "sync", "error": sync_res.get("error"), "logs": "\n\n".join(logs)})
    
    # 2. ST
    st_res = run_python_script(GENST_SCRIPT)
    logs.append(f"=== [2] GENERATOR SURAT TUGAS ===\n{st_res.get('output') or st_res.get('error')}")
    if not st_res["success"]:
        raise HTTPException(status_code=500, detail={"step": "generate_st", "error": st_res.get("error"), "logs": "\n\n".join(logs)})
        
    # 3. Laporan
    lap_res = run_python_script(GENLAP_SCRIPT)
    logs.append(f"=== [3] GENERATOR LAPORAN ===\n{lap_res.get('output') or lap_res.get('error')}")
    if not lap_res["success"]:
        raise HTTPException(status_code=500, detail={"step": "generate_laporan", "error": lap_res.get("error"), "logs": "\n\n".join(logs)})

    return {
        "status": "success",
        "action": "run_all",
        "message": "Semua proses (Sync, ST, dan Laporan) berhasil diselesaikan!",
        "log": "\n\n".join(logs),
        "st_files": list_generated_files(ST_OUTPUT_DIR),
        "lap_files": list_generated_files(LAP_OUTPUT_DIR)
    }

@app.get("/api/files")
def get_files(user: str = Depends(authenticate)):
    """Mengambil daftar file dokumen ST dan Laporan yang tersedia."""
    return {
        "st_files": list_generated_files(ST_OUTPUT_DIR),
        "lap_files": list_generated_files(LAP_OUTPUT_DIR)
    }

@app.get("/api/download/{category}/{filename}")
def download_file(category: str, filename: str, user: str = Depends(authenticate)):
    """Mengunduh file docx yang telah dibuat."""
    if category.lower() == "st":
        target_path = ST_OUTPUT_DIR / filename
    elif category.lower() in ["lap", "laporan"]:
        target_path = LAP_OUTPUT_DIR / filename
    else:
        raise HTTPException(status_code=400, detail="Kategori tidak valid. Gunakan 'st' atau 'lap'.")
        
    if not target_path.exists() or not target_path.is_file():
        raise HTTPException(status_code=404, detail=f"File '{filename}' tidak ditemukan.")
        
    return FileResponse(
        path=str(target_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# ==================== TAMPILAN WEB (DASHBOARD) ====================

@app.get("/", response_class=HTMLResponse)
def index_page(user: str = Depends(authenticate)):
    """Halaman Antarmuka Web Dashboard Pengganti perintah.sh."""
    template_path = BASE_DIR / "templates" / "index.html"
    if not template_path.exists():
        return HTMLResponse("<h1>File templates/index.html tidak ditemukan.</h1>", status_code=500)
    
    html = template_path.read_text(encoding="utf-8")
    html = html.replace("{{username}}", user)
    return HTMLResponse(content=html)

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("app:app", host="0.0.0.0", port=port)
