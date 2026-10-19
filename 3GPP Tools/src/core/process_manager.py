import subprocess
import json
import logging
import platform
from typing import Dict, List, Optional


class ProcessManager:
    @staticmethod
    def get_process_stats():
        """Uses PowerShell to fetch processes, including their Window Titles."""
        ps_cmd = """
        $procs = Get-Process visio, powerpnt, winword -ErrorAction SilentlyContinue
        $out = @()
        foreach ($p in $procs) {
            $isGhost = ($p.MainWindowHandle -eq 0)
            $out += @{ 
                Name = $p.Name; 
                Id = $p.Id; 
                IsGhost = $isGhost; 
                Title = $p.MainWindowTitle 
            }
        }
        $out | ConvertTo-Json -Compress
        """

        try:
            result = subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd],
                                    capture_output=True, text=True, creationflags=0x08000000, timeout=5)

            output = result.stdout.strip()
            if not output:
                return []

            data = json.loads(output)
            return data if isinstance(data, list) else [data]
        except Exception as e:
            logging.error(f"Process check failed: {e}")
            return []


    @staticmethod
    def get_open_office_documents(app_name: str) -> List[Dict]:
        """Return open user documents for Word or PowerPoint via COM."""
        if platform.system() != "Windows":
            return []

        app_key = app_name.lower()
        if app_key == "winword":
            try:
                from modules.word_tools.core.document_opener import list_open_word_documents
                return list_open_word_documents()
            except Exception as exc:
                logging.warning("Could not enumerate open Word documents: %s", exc)
                return []

        if app_key != "powerpnt":
            return []

        import gc
        import pythoncom
        import win32com.client
        app = None
        presentation = None
        pythoncom.CoInitialize()
        try:
            try:
                app = win32com.client.GetActiveObject("PowerPoint.Application")
            except Exception:
                return []
            result = []
            for presentation in app.Presentations:
                try:
                    full_name = str(presentation.FullName or "").strip()
                    name = str(presentation.Name or "").strip() or "Untitled Presentation"
                    result.append({"name": name, "full_name": full_name, "modified": None})
                except Exception:
                    continue
            return result
        finally:
            presentation = app = None
            gc.collect()
            pythoncom.CoUninitialize()

    @staticmethod
    def close_office_documents(app_name: str, excluded_paths=None, discard_changes: bool = False) -> Dict:
        """Gracefully close Word/PowerPoint documents, optionally discarding changes."""
        excluded = {str(x).casefold() for x in (excluded_paths or []) if x}
        app_key = app_name.lower()

        if app_key == "winword":
            from modules.word_tools.core.document_opener import close_open_word_documents
            return close_open_word_documents(excluded_paths=excluded_paths, discard_changes=discard_changes)

        summary = {"closed": 0, "excluded": 0, "failed": []}
        if app_key != "powerpnt" or platform.system() != "Windows":
            return summary

        import gc
        import pythoncom
        import win32com.client
        app = None
        presentations = None
        presentation = None
        pythoncom.CoInitialize()
        try:
            try:
                app = win32com.client.GetActiveObject("PowerPoint.Application")
            except Exception:
                return summary
            presentations = [app.Presentations.Item(i) for i in range(app.Presentations.Count, 0, -1)]
            for presentation in presentations:
                try:
                    full_name = str(presentation.FullName or "").strip()
                    key = full_name.casefold() if full_name else str(presentation.Name).casefold()
                    if key in excluded:
                        summary["excluded"] += 1
                        continue
                    if discard_changes:
                        # PowerPoint has no Close(saveChanges) argument. Mark it saved so
                        # Close() does not prompt and therefore deliberately discards edits.
                        presentation.Saved = True
                    presentation.Close()
                    summary["closed"] += 1
                except Exception as exc:
                    summary["failed"].append(str(exc))
            return summary
        finally:
            presentation = presentations = app = None
            gc.collect()
            pythoncom.CoUninitialize()

    @staticmethod
    def kill_process(pid: int):
        """Kills a single specific process by its PID."""
        try:
            subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                           creationflags=0x08000000, check=True)
            return True
        except subprocess.CalledProcessError:
            return False

    @staticmethod
    def kill_app_ghosts(app_name: str):
        """Kills only the headless ghost processes for a specific application."""
        data = ProcessManager.get_process_stats()
        killed = 0

        for p in data:
            if p["Name"].lower() == app_name.lower() and p["IsGhost"]:
                ProcessManager.kill_process(p["Id"])
                killed += 1
        return killed

    @staticmethod
    def kill_app_all(app_name: str):
        """Kills ALL processes (both active and ghost) for a specific application."""
        data = ProcessManager.get_process_stats()
        killed = 0

        for p in data:
            if p["Name"].lower() == app_name.lower():
                ProcessManager.kill_process(p["Id"])
                killed += 1
        return killed