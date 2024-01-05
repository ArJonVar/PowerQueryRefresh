#region imports
import smartsheet
import traceback
import sys
import os
import concurrent.futures
from smartsheet.exceptions import ApiError
from smartsheet_grid import grid
import time
from globals import smartsheet_token
from logger import ghetto_logger
from datetime import datetime
import win32com.client as win32
import pythoncom
from datetime import datetime
#endregion

class PQRefresher():
    '''Explain Class'''
    def __init__(self, config):
        self.config = config
        self.smartsheet_token=config.get('smartsheet_token')
        self.sheet_id=config.get('sheet_id')
        grid.token=smartsheet_token
        self.sheet_instance = grid(self.sheet_id)
        self.smart = smartsheet.Smartsheet(access_token=self.smartsheet_token)
        self.smart.errors_as_exceptions(True)
        self.start_time = time.time()
        self.log=ghetto_logger("pqrefresh.py")
    #region helpers
    def refresh_power_query(self, excel_file_path):
        '''opens file from path, refreshes, saves, closes'''
        pythoncom.CoInitialize()
        error = ""
        try:
            excel_app = win32.gencache.EnsureDispatch('Excel.Application')
            excel_app.Interactive = False
            excel_app.Visible = True #eventually set to false for deployment
            workbook = excel_app.Workbooks.Open(excel_file_path)
            # Disable background refresh and refresh all connections
            try:
                for connection in workbook.Connections:
                    connection.OLEDBConnection.BackgroundQuery = False
            except:
                self.log.log('could not turn off background refresh')
            try:
                workbook.RefreshAll()
                time.sleep(2)
            except AttributeError:
                error= "FILE OPEN ERROR: could not refresh because the file is open somewhere"
                return error
            workbook.Save()
            workbook.Close()
            excel_app.Quit()
        except Exception as e:
            error='LOCAL ERROR: failed because of bug on deployment computer, check logs for further details'
            self.log.log(e)
            return error
        finally:
            pythoncom.CoUninitialize() 

        return error
    def handle_pqrefresh_wtimeout(self, excel_file_path):
        '''wrapper for pq refresh function to handle conditional timeouts'''
        file_size_kb = os.path.getsize(excel_file_path) / 1024  # Convert bytes to kilobytes
        if file_size_kb < 1000:
            timeout = 20 * 60  # 20 minutes in seconds
        else:
            timeout = 45 * 60  # 45 minutes in seconds

        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
            future = executor.submit(self.refresh_power_query, excel_file_path)
            try:
                error = future.result(timeout=timeout)
                if error:
                    self.log.log(error)
                    return error
            except concurrent.futures.TimeoutError:
                error = f"TIMEOUT ERROR: File was not done refreshing in {timeout // 60} minutes"
                self.log.log(error)
                return error
        return None
    def now(self):
        '''generates a posting message that says the time/date'''
        now = datetime.now()
        dt_string = now.strftime("%m/%d %H:%M")
        return dt_string
    #endregion 
    def grab_ss_data(self):
        '''grabs data from https://app.smartsheet.com/sheets/RcqQV67RPv5CMQf7jcH48prH769QCQrChJMwr5h1?view=grid, makes a list of dicts'''
        # loop through column_df and get object with title:column_id for each column:
        self.sheet_instance.fetch_content()
        df = self.sheet_instance.df
        enabled_df = df[df['Enabled'] == True]
        configurednenabled_df = enabled_df[enabled_df['Configured'] == True]
        if datetime.today().weekday() != 5:
            configurednenabled_df = configurednenabled_df[configurednenabled_df['Reresh Frequency'] != 'Weekly']
        list_of_datadicts = configurednenabled_df.to_dict(orient='records')
        return list_of_datadicts
    def refresh_each_excel(self, list_of_datadicts):
        '''loops through paths in smartsheet and refreshes each path'''
        self.update = []
        for item in list_of_datadicts:
            path = item["Z Drive Path-to-file"]

            # Replace 'Z:\\' with 'C:\\Egnyte\\'
            path = path.replace('Z:\\', 'C:\\Egnyte\\')

            # Replace double backslashes with single backslashes
            path = path.replace('\\\\', '\\')

            # Remove surrounding double quotes, if present
            path = path.strip('"')

            # Update the path in the dictionary
            item["Z Drive Path-to-file"] = path

            error = 'Error Posting'

            if not os.path.isfile(path):
                error = 'FILE PATH ERROR: file was not found'
                self.update.append({'Name of Excel File':item['Name of Excel File'], 'Python Message': f'{self.now()} {error}'})
                continue

            error = self.handle_pqrefresh_wtimeout(path)
            
            if error:
                self.update.append({'Name of Excel File': item['Name of Excel File'], 
                          'Python Message': f'{self.now()} {error}'})
            else:
                self.update.append({'Name of Excel File': item['Name of Excel File'], 
                          'Python Message': f'{self.now()} Successful Refresh'})
        return self.update
            
    def run(self):
        '''runs main script as intended'''
        self.data = self.grab_ss_data()
        # self.update = self.refresh_each_excel(data)
        self.log.log('posting updates...')
        # self.sheet_instance.update_rows(self.update, 'Name of Excel File')
        # self.sheet_instance.handle_update_stamps()
        self.log.log('~fin~')


if __name__ == "__main__":
    config = {
        'smartsheet_token':smartsheet_token,
        'sheet_id': 6463522670071684
    }
    pqr=PQRefresher(config)
    pqr.run()