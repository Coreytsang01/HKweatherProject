# HKweatherProject

The objective of this project is to automatically download and process the specific data from Hong Kong Observatory. 

# Getting started
Follow these instructions to setup and run the application in your local development environment.
1. Create a virtual environment: https://docs.python.org/3/library/venv.html
2. Activate the virtual environment with `./venv/Scripts/activate`
3. In the terminal of your IDE or device, navigate to the folder containing the file `setup.py` and `mainRun.py`
4. Install the required dependencies with 
```bash
pip install -r requirements.txt
```
5. Run `setup.py` and it helps you to create directory `temp file` and `Master Data.xlsx` for program use
6. Run `mainRun.py` to see whether it can download and clean the data, importing into `Master Data.xlsx`

# Task Scheduler Setup
if you would like to run the program from time to time automatically, it is recommanded to create a `.bat` to be run in Task Scheduler because the python file cannot be run directly in Task Scheduler.

Please follow these instructions to setup.
1. Open the Task Scheduler
2. Click `Create Task` 

![image](https://user-images.githubusercontent.com/115489179/235808805-548ef0c0-0178-46dc-926d-123ebeb1f1d4.png)

3. Input the corresponding information, Run only when user is logged in and run with highest privileges

![image](https://user-images.githubusercontent.com/115489179/235808325-4425646b-337f-47c3-a56f-b4f8d23e71f1.png)

4. Click into `Trigger` page, click `New` and follow the below setup
Under this setup, the program will be run every five minutes indefinitely
(Please note that the Start time is the moment, you create the task)

![image](https://user-images.githubusercontent.com/115489179/235810135-3ac60d15-d005-4ed0-87d2-8da173111012.png)

5. Click into `Actions` page, click `New` and follow the below setup
- Program/Script: C:\Windows\System32\cmd.exe
- Add arguments(Optional): /C start "" /MIN F:\python\001_ObservatoryProject\run_script.bat ^& exit
- /C start "" /MIN is to minimize the `cmd.exe` window which will pop up from time to time 
- F:\python\001_ObservatoryProject\run_script.bat is the directory of the `.bat` (please input the absolute path)
- ^& exit is to exit the `cmd.exe` 
![image](https://user-images.githubusercontent.com/115489179/235811778-db2e1519-1242-4948-8ed2-2fe8afc2cb07.png)

6. Click OK in properties

![image](https://user-images.githubusercontent.com/115489179/235812230-df9a691a-a128-4537-8a74-101cb37a0e89.png)

After setup the schedule, you can enable the task and run it once to test it whether it is working

![image](https://user-images.githubusercontent.com/115489179/235812478-2266c2d7-36a1-488b-8506-879500c59185.png)



