import psutil, os, time
p = psutil.Process(os.getpid())
for _ in range(5):
    print(p.cpu_percent(interval=1), p.memory_info().rss)
    time.sleep(1)
