# Performance tracker

Tracks time spent on opened windows and make a result to Excel file. It's using Windows API (Win32) to get titles of windows. 

## Example Usage

After starting the program is running as it's showed here:
```
Program is running...(Type 'quit' to exit)
> 
```
If you finish your work, type quit and then save it with command:
```
python tracker.py save
```
It will make an excel file in path which is defined in the source code and it'll look similar to this table:

| 16.06.21 - 17:38:36  |  |
| ------------- | ------------- |
| File Explorer  | 0:00:11  |
| Brave  | 0:01:19  |
| Google Chrome  | 0:00:06  |
| Untitled - Notepad  | 0:00:02  |

Cell A1 is always the date and time of saving.


