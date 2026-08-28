"""
Keeps the dashboard updated continuously. Two modes depending on
config.USE_GRAPH_API:

- Graph API mode (True): there's no local file to "watch" for changes
  since data lives in the cloud, so this polls on a short fixed interval
  instead (config.GRAPH_POLL_SECONDS). Each poll is a live read from the
  cloud, so it reflects other people's open browser sessions.
- Local file mode (False): watches the local folder and rebuilds instantly
  when a tracked file changes, falling back to a periodic safety-net
  refresh regardless.

Leave this running in the background. Runs continuously until closed.

Usage:
    python watch_and_update.py
"""
import os
import time
import logging

import config
import run_dashboard_update as updater

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("dashboard_watch.log"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)

# Belt-and-suspenders / polling interval depending on mode.
FORCE_REFRESH_SECONDS = 45 * 60
GRAPH_POLL_SECONDS = getattr(config, "GRAPH_POLL_SECONDS", 120)


def main():
    source = getattr(config, "DATA_SOURCE", "local")
    if source in ("csv", "graph"):
        run_polling_mode(source)
    else:
        run_local_file_watch_mode()


def run_polling_mode(source):
    label = "Power Automate CSV exports" if source == "csv" else "Microsoft cloud (Graph API)"
    log.info(f"{label} mode - polling every {GRAPH_POLL_SECONDS}s")
    log.info("Leave this window open. Press Ctrl+C to stop.")
    while True:
        try:
            updater.main()
        except Exception:
            log.exception("Rebuild failed - will retry next poll")
        time.sleep(GRAPH_POLL_SECONDS)


def run_local_file_watch_mode():
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler

    DEBOUNCE_SECONDS = 5
    TRACKED_FILES = set(config.FILES.values())

    class DefectLogHandler(FileSystemEventHandler):
        def __init__(self):
            self.last_change = 0
            self.pending = False

        def _touched(self, path):
            filename = os.path.basename(path)
            return (not filename.startswith("~$")) and filename in TRACKED_FILES

        def on_modified(self, event):
            if not event.is_directory and self._touched(event.src_path):
                log.info(f"Change detected: {os.path.basename(event.src_path)}")
                self.last_change = time.time()
                self.pending = True

        on_created = on_modified

        def on_moved(self, event):
            if not event.is_directory and self._touched(event.dest_path):
                log.info(f"Change detected: {os.path.basename(event.dest_path)}")
                self.last_change = time.time()
                self.pending = True

    handler = DefectLogHandler()
    observer = Observer()
    observer.schedule(handler, config.ONEDRIVE_FOLDER, recursive=False)
    observer.start()
    log.info(f"Watching {config.ONEDRIVE_FOLDER} for changes to: {', '.join(TRACKED_FILES)}")
    log.info("Leave this window open. Press Ctrl+C to stop.")

    last_build_time = time.time()
    try:
        updater.main()
    except Exception:
        log.exception("Initial build failed - check config.py, then save the file "
                       "again to trigger a retry")

    try:
        while True:
            time.sleep(1)
            should_rebuild = False
            reason = ""

            if handler.pending and (time.time() - handler.last_change) >= DEBOUNCE_SECONDS:
                handler.pending = False
                should_rebuild = True
                reason = "file change detected"
            elif (time.time() - last_build_time) >= FORCE_REFRESH_SECONDS:
                should_rebuild = True
                reason = f"periodic refresh ({FORCE_REFRESH_SECONDS//60} min safety net)"

            if should_rebuild:
                log.info(f"Rebuilding dashboard ({reason})...")
                try:
                    updater.main()
                    last_build_time = time.time()
                except Exception:
                    log.exception("Rebuild failed - dashboard not updated this time, "
                                  "will retry on next change or safety-net interval")
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


if __name__ == "__main__":
    main()
