from datetime import date
import sys
import pandas as pd
from openpyxl import load_workbook
import xlsxwriter
import time
import os
import shutil
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import PatternMatchingEventHandler
from datetime import datetime
pd.__version__

if __name__ == "__main__":
    patterns = ["*"]
    ignore_patterns = None
    ignore_directories = True
    case_sensitive = True
    my_event_handler = PatternMatchingEventHandler(patterns, ignore_patterns, ignore_directories, case_sensitive)

def on_created(event):
    print(f"File: {event.src_path} has been found")
    head, tail = os.path.split(event.src_path)
    path_to_file = r'C:\Users\andyv\Documents\Test\Processed\{}'.format(tail)
    path_to_file2 = r'C:\Users\andyv\Documents\Test\Not applicable\{}'.format(tail)
    path = Path(path_to_file)
    path2 = Path(path_to_file2)
    if path.is_file() or path2.is_file():
        print(f'The file {path_to_file} exists')
    else:  
        if event.src_path.endswith('.xlsx'):
            
            
            event.dest_path = "C:\\Users\\andyv\\Documents\\Test\\Processed"
            data = pd.read_excel(event.src_path)
            Master = pd.read_excel(r"C:\Users\andyv\Documents\Test\master.xlsx", engine= 'openpyxl')
            writer = pd.ExcelWriter(r"C:\Users\andyv\Documents\Test\master.xlsx", engine = 'xlsxwriter')
            data.to_excel(writer, sheet_name = 'Sheet1')
            Master.to_excel(writer, sheet_name = 'original')
            writer.save()
            writer.close()
            # assert os.path.isfile(data)
            # with open(data, "r") as f: 
            #     pass


            shutil.move(event.src_path, event.dest_path)
        else:
            event.dest_path = "C:\\Users\\andyv\\Documents\\Test\\Not applicable"
            shutil.move(event.src_path, event.dest_path)
    # print(f"File moved from: {event.src_path} to {event.dest_path} according to conditions")
 

def on_deleted(event):
    print(f"{event.src_path} has been removed")
 
def on_modified(event):
    print(f"{event.src_path} has been modified")
 
def on_moved(event):
    print(f"File moved from: {event.src_path} to {event.dest_path} according to conditions")

my_event_handler.on_created = on_created
my_event_handler.on_deleted = on_deleted
my_event_handler.on_modified = on_modified
my_event_handler.on_moved = on_moved

path = "C:\\Users\\andyv\\Documents\\Test"
go_recursively = True
my_observer = Observer()
my_observer.schedule(my_event_handler, path, recursive=go_recursively)

my_observer.start()
try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    my_observer.stop()
    my_observer.join()


