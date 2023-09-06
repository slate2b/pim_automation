@if (@CodeSection == @Batch) @then

@echo off

cd c:\program files\google\chrome\application

start chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\Users\Public"

goto: EOF

@end