import threading, time, signal

from datetime import timedelta

from openpyxl import Workbook
workbook = Workbook()
sheet = workbook.active
WAIT_TIME_SECONDS = 4
t = 0
y = 1
filepath="C:/Users/jl/OneDrive - Interel/Desktop/Test/demo.xlsx"

class ProgramKilled(Exception):
    pass

def function():
    print (time.ctime())
    sheet.cell(row=y,column=1).value = t

def signal_handler(signum, frame):
    raise ProgramKilled

class Job(threading.Thread):
    def __init__(self, interval, execute, *args, **kwargs):
        threading.Thread.__init__(self)
        self.daemon = False
        self.stopped = threading.Event()
        self.interval = interval
        self.execute = execute
        self.args = args
        self.kwargs = kwargs

    def stop(self):
                self.stopped.set()
                self.join()
    def run(self):
            while not self.stopped.wait(self.interval.total_seconds()):
                self.execute(*self.args, **self.kwargs)

sheet['A1'] = 'Time Elapsed(In Seconds)'
sheet['B1'] = 'Internal Temperature(In Celsius)'
sheet['C1'] = 'External Temperature(In Celsius)'

if __name__ == "__main__":
    signal.signal(signal.SIGTERM, signal_handler)
    signal.signal(signal.SIGINT, signal_handler)
    job = Job(interval=timedelta(seconds=WAIT_TIME_SECONDS), execute=function)
    job.start()

    while True:
          try:
              time.sleep(WAIT_TIME_SECONDS)
              y+=1
              t+=2
          except ProgramKilled:
              print ("Program killed: running cleanup code")
              job.stop()
              break

workbook.save(filepath)
