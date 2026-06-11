import subprocess
import json
import logging


class ProcessManager:
    @staticmethod
    def get_process_stats():
        """
        Uses PowerShell to fetch Visio, PowerPoint, and Word processes.
        If MainWindowHandle is 0, the process is a headless "Ghost".
        """
        ps_cmd = """
        $procs = Get-Process visio, powerpnt, winword -ErrorAction SilentlyContinue
        $out = @()
        foreach ($p in $procs) {
            $out += @{ Name = $p.Name; Id = $p.Id; IsGhost = ($p.MainWindowHandle -eq 0) }
        }
        $out | ConvertTo-Json -Compress
        """

        try:
            # 0x08000000 prevents the black command prompt window from flashing on screen
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
    def kill_processes(app_name: str, ghosts_only: bool = True):
        """Kills processes matching the executable name. Can target only ghosts."""
        data = ProcessManager.get_process_stats()
        killed = 0

        for p in data:
            if p["Name"].lower() == app_name.lower():
                if not ghosts_only or p["IsGhost"]:
                    try:
                        subprocess.run(["taskkill", "/F", "/PID", str(p["Id"])],
                                       creationflags=0x08000000, check=True)
                        killed += 1
                    except subprocess.CalledProcessError:
                        pass
        return killed