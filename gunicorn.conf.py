# One worker + threads: the weekly-reminder scheduler runs inside the worker,
# and SQLite is happiest with a single writer process.
bind = "0.0.0.0:8000"
workers = 1
threads = 4
timeout = 60


def post_worker_init(worker):
    from app import start_scheduler
    start_scheduler()
