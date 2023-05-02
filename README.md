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

