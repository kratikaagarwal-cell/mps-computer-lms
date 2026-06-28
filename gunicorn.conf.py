import multiprocessing

workers = (multiprocessing.cpu_count() * 2) + 1
threads = 4
worker_class = "sync"
timeout = 60
keepalive = 5
bind = "0.0.0.0:5000"
accesslog = "-"
errorlog = "-"
loglevel = "info"
graceful_timeout = 30
proc_name = "mps-lms"
