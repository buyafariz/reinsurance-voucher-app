import os
import time

LOCK_TIMEOUT = 30  # detik

def acquire_lock(lock_path):
    start = time.time()

    while os.path.exists(lock_path):
        if time.time() - start > LOCK_TIMEOUT:
            raise TimeoutError("Lock timeout – kemungkinan deadlock")
        time.sleep(0.3)

    with open(lock_path, "w") as f:
        f.write(str(os.getpid()))


def release_lock(lock_path):
    if os.path.exists(lock_path):
        os.remove(lock_path)
