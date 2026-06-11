import subprocess
import json
import logging


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
                                    capture_output=True, text=True, creationflags=0x08000000)

            output = result.stdout.strip()
            if not output:
                return []

            data = json.loads(output)
            return data if isinstance(data, list) else [data]
        except Exception as e:
            logging.error(f"Process check failed: {e}")
            return []

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