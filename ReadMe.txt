Built by Daniel Karkhut

Code automation for EES023 Weather Notebook Assignment. We collect hourly data for each day of an entire month. The assignment however only asks for 4 hourly recordings per day, preferably one at 1:51 am, 7:51 am, 1:51 pm, and 7:51 pm. If one of these times isn't available, which is likely because the data is not very clean, then the next nearest recording is taken. Can see the problem by opening the Excel document and comparing sheets 'sheet1' and 'CleanData'

Instructions to run:

1. Ensure script and excel file are in the same folder

2. Name excel file 'WeatherNotebook.xlsx'

3. Ensure sheet with data is called 'Sheet1'

4. In terminal, go to folder and type 'python3 ConvertHourly.py'

5. Done! A new sheet should have appeared with 4 hourly recordings per day labeled with the date 
