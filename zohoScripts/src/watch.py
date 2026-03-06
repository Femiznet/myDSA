import time
import queue
import threading
import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# 1. Setup the Queue and Worker
file_queue = queue.Queue()

def worker():
    while True:
        file_path = file_queue.get()
        if file_path is None: break
        
        try:
            print(f"Processing: {file_path}")
            subprocess.run(["python", "send_mail.py", file_path], check=True)
        except Exception as e:
            print(f"Error sending mail for {file_path}: {e}")
        finally:
            # THIS MUST RUN EVEN IF THERE WAS AN ERROR
            file_queue.task_done() 

# 2. Define the Watcher logic
class NewFileHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory:
            print(f"Queued: {event.src_path}")
            file_queue.put(event.src_path)

def On_My_Watch():
    """Main watcher loop."""
    watch_path = r"C:\Users\DELL\devFiles\tickets"
    event_handler = NewFileHandler()
    
    observer = Observer()
    observer.schedule(event_handler, watch_path, recursive=True)
    
    print(f"Watching for new files in: {watch_path}")
    observer.start()
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nStopping watcher...")
        observer.stop()
    observer.join()

# 3. Execution Entry Points
def run_watcher():
    """Call this from your Zoho script to start in background."""
    # Start the worker thread to process the queue
    threading.Thread(target=worker, daemon=True).start()
    
    # Start the observer thread
    watcher_thread = threading.Thread(target=On_My_Watch, daemon=True)
    watcher_thread.start()
    print("Watcher and Worker threads started.")


if __name__ == "__main__":
    # If running this file directly
    threading.Thread(target=worker, daemon=True).start()
    On_My_Watch()
