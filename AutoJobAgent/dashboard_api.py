import json
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import openpyxl
from fastapi import FastAPI, HTTPException, Request, WebSocket, WebSocketDisconnect
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from rich.console import Console

# Initialize console for rich output
console = Console()

# Initialize FastAPI app
app = FastAPI(title="AutoJobAgent Dashboard")

# Set up paths
BASE_DIR = Path(__file__).parent
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"
LOGS_DIR = OUTPUT_DIR / "logs"
IPC_FILE = OUTPUT_DIR / "ipc.json"

# Ensure directories exist
OUTPUT_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)

# Set up templates
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))

# Mount static files (if they exist)
STATIC_DIR = BASE_DIR / "static"
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


# WebSocket connection manager
class ConnectionManager:
    def __init__(self):
        self.active_connections: List[WebSocket] = []

    async def connect(self, websocket: WebSocket):
        await websocket.accept()
        self.active_connections.append(websocket)

    def disconnect(self, websocket: WebSocket):
        self.active_connections.remove(websocket)

    async def broadcast(self, message: Dict):
        for connection in self.active_connections:
            await connection.send_json(message)


# Initialize connection manager
manager = ConnectionManager()


# Models
class PauseRequest(BaseModel):
    pause: bool = True


# Helper functions
def get_application_stats() -> Dict:
    """
    Get application statistics from the log file.
    
    Returns:
        Dictionary with application statistics
    """
    stats = {
        "Applied": 0,
        "Skipped": 0,
        "Failed": 0,
        "Manual": 0,
        "Total": 0
    }
    
    log_file = LOGS_DIR / "applications.xlsx"
    
    if not log_file.exists():
        return stats
    
    try:
        wb = openpyxl.load_workbook(log_file, read_only=True)
        ws = wb.active
        
        # Skip header row
        rows = list(ws.rows)[1:]
        
        # Count statuses
        for row in rows:
            status = row[6].value  # Status is in column G (index 6)
            if status:
                status = str(status)
                stats["Total"] += 1
                
                if "Applied" in status:
                    stats["Applied"] += 1
                elif "Skipped" in status:
                    stats["Skipped"] += 1
                elif "Failed" in status:
                    stats["Failed"] += 1
                elif "Manual" in status:
                    stats["Manual"] += 1
        
        return stats
    
    except Exception as e:
        console.print(f"[bold red]Error getting application stats: {e}[/bold red]")
        return stats


def get_recent_logs(limit: int = 100) -> List[Dict]:
    """
    Get recent application logs.
    
    Args:
        limit: Maximum number of logs to return
        
    Returns:
        List of log entries
    """
    logs = []
    log_file = LOGS_DIR / "applications.xlsx"
    
    if not log_file.exists():
        return logs
    
    try:
        wb = openpyxl.load_workbook(log_file, read_only=True)
        ws = wb.active
        
        # Get all rows and reverse to get most recent first
        rows = list(ws.rows)
        headers = [cell.value for cell in rows[0]]
        
        # Get data rows (skip header), most recent first
        data_rows = rows[1:][::-1][:limit]
        
        # Convert rows to dictionaries
        for row in data_rows:
            log_entry = {}
            for i, cell in enumerate(row):
                if i < len(headers):
                    # Format timestamp if it's a datetime
                    if i == 0 and isinstance(cell.value, datetime):
                        log_entry[headers[i]] = cell.value.strftime("%Y-%m-%d %H:%M:%S")
                    else:
                        log_entry[headers[i]] = cell.value
            logs.append(log_entry)
        
        return logs
    
    except Exception as e:
        console.print(f"[bold red]Error getting recent logs: {e}[/bold red]")
        return logs


def get_pause_status() -> bool:
    """
    Get the current pause status.
    
    Returns:
        True if paused, False otherwise
    """
    if not IPC_FILE.exists():
        return False
    
    try:
        with open(IPC_FILE, "r") as f:
            ipc_data = json.load(f)
        
        return ipc_data.get("pause", False)
    
    except (json.JSONDecodeError, FileNotFoundError):
        return False


def set_pause_status(pause: bool) -> bool:
    """
    Set the pause status.
    
    Args:
        pause: True to pause, False to resume
        
    Returns:
        True if successful, False otherwise
    """
    try:
        with open(IPC_FILE, "w") as f:
            json.dump({"pause": pause}, f)
        
        return True
    
    except Exception as e:
        console.print(f"[bold red]Error setting pause status: {e}[/bold red]")
        return False


# Routes
@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """
    Render the dashboard index page.
    """
    return templates.TemplateResponse(
        "index.html",
        {"request": request, "title": "AutoJobAgent Dashboard"}
    )


@app.get("/stats")
async def stats():
    """
    Get application statistics.
    """
    return get_application_stats()


@app.get("/log")
async def log(limit: int = 100):
    """
    Get recent application logs.
    
    Args:
        limit: Maximum number of logs to return
    """
    return get_recent_logs(limit)


@app.post("/action/pause")
async def pause():
    """
    Pause the job application process.
    """
    success = set_pause_status(True)
    
    if success:
        # Broadcast pause status to WebSocket clients
        await manager.broadcast({"action": "pause", "status": True})
        return {"status": "paused"}
    else:
        raise HTTPException(status_code=500, detail="Failed to set pause status")


@app.post("/action/resume")
async def resume():
    """
    Resume the job application process.
    """
    success = set_pause_status(False)
    
    if success:
        # Broadcast resume status to WebSocket clients
        await manager.broadcast({"action": "pause", "status": False})
        return {"status": "resumed"}
    else:
        raise HTTPException(status_code=500, detail="Failed to set resume status")


@app.get("/status")
async def status():
    """
    Get the current pause status.
    """
    return {"paused": get_pause_status()}


# WebSocket endpoint
@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    """
    WebSocket endpoint for real-time updates.
    """
    await manager.connect(websocket)
    
    try:
        # Send initial stats and status
        await websocket.send_json({
            "type": "stats",
            "data": get_application_stats()
        })
        
        await websocket.send_json({
            "type": "status",
            "paused": get_pause_status()
        })
        
        # Listen for messages
        while True:
            data = await websocket.receive_text()
            
            # Process commands from client
            try:
                message = json.loads(data)
                
                if message.get("command") == "get_stats":
                    await websocket.send_json({
                        "type": "stats",
                        "data": get_application_stats()
                    })
                
                elif message.get("command") == "get_logs":
                    limit = message.get("limit", 100)
                    await websocket.send_json({
                        "type": "logs",
                        "data": get_recent_logs(limit)
                    })
                
                elif message.get("command") == "get_status":
                    await websocket.send_json({
                        "type": "status",
                        "paused": get_pause_status()
                    })
                
            except json.JSONDecodeError:
                pass
            
    except WebSocketDisconnect:
        manager.disconnect(websocket)


# Event handlers
@app.on_event("startup")
async def startup_event():
    """
    Startup event handler.
    """
    console.print("[bold green]Dashboard API started[/bold green]")
    
    # Ensure IPC file exists
    if not IPC_FILE.exists():
        set_pause_status(False)


@app.on_event("shutdown")
async def shutdown_event():
    """
    Shutdown event handler.
    """
    console.print("[bold yellow]Dashboard API shutting down[/bold yellow]")


# If this file is run directly, start the server
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
