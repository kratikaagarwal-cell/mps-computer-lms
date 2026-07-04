import os

# Render's free/starter tiers give very limited RAM & shared CPU.
# Using multiprocessing.cpu_count() here is dangerous: Render's container
# often reports more cores than your plan actually gets, so the old
# "(cpu_count * 2) + 1" formula can spawn far more worker PROCESSES than
# your RAM allows -> Render OOM-kills the app -> 502 Bad Gateway.
#
# Fix: a small, fixed number of workers, using threads (gthread) for
# concurrency instead of extra processes. Threads share one process's
# memory AND one database connection pool, so this also keeps total DB
# connections low and predictable (see app.py pool settings).
workers = 2
worker_class = "gthread"
threads = 6
timeout = 90
keepalive = 5

# Render assigns the port dynamically via the PORT env var - don't hardcode it.
bind = "0.0.0.0:" + os.environ.get("PORT", "5000")

accesslog = "-"
errorlog = "-"
loglevel = "info"
graceful_timeout = 30
proc_name = "mps-lms"
