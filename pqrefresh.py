#region imports
import smartsheet
from smartsheet.exceptions import ApiError
from smartsheet_grid import grid
import time
from globals import smartsheet_token
from logger import ghetto_logger
from datetime import datetime
import win32com.client as win32
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
        self.log=ghetto_logger("func.py")
    #region helpers
    def refresh_power_query(self, excel_file_path):
        '''opens file from path, refreshes, saves, closes'''
        excel_app = win32.gencache.EnsureDispatch('Excel.Application')
        workbook = excel_app.Workbooks.Open(excel_file_path)
        
        # eventaully change to false
        excel_app.Interactive = False
        excel_app.Visible = True

        # Disable background refresh and refresh all connections
        for connection in workbook.Connections:
            connection.OLEDBConnection.BackgroundQuery = False
        workbook.RefreshAll()

        # Wait for some time to ensure refresh is complete
        time.sleep(2)  # Adjust the time as necessary

        workbook.Save()
        workbook.Close()
        excel_app.Quit()
    def refresh_power_query2(self, excel_file_path):
        '''opens file from path, refreshes, saves, closes'''
        try:
            excel_app = win32.gencache.EnsureDispatch('Excel.Application')
            workbook = excel_app.Workbooks.Open(excel_file_path)
            excel_app.Visible = True

            # Disable background refresh and refresh all connections
            for connection in workbook.Connections:
                connection.OLEDBConnection.BackgroundQuery = False
            workbook.RefreshAll()

            # Wait for some time to ensure refresh is complete
            time.sleep(2)  # Adjust the time as necessary

            workbook.Save()
        except Exception as e:
            error = f"COM Error: {e}"
            raise Exception(error)
        finally:
            try:
                workbook.Close()
            except Exception as e:
                error = f"COM Error (Workbook Close): {e}"
                raise Exception(error)
            try:
                excel_app.Quit()
            except Exception as e:
                error = f"COM Error (Excel Quit): {e}"
                raise Exception(error)
            try:
                # Explicitly release the COM objects
                excel_app.Quit()
                excel_app = None
                workbook = None
            except Exception as e:
                error = f"COM Error (Object Cleanup): {e}"
                raise Exception(error)
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
        list_of_datadicts = configurednenabled_df.to_dict(orient='records')
        return list_of_datadicts
    def refresh_each_excel(self, list_of_datadicts):
        '''loops through paths in smartsheet and refreshes each path'''
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

            try: 
                self.refresh_power_query(path)
                update = {'Name of Excel File':item['Name of Excel File'], 'Python Message': f'Refreshed {self.now()}'}
                print(update)
                self.sheet_instance.update_rows([update], 'Name of Excel File')
            except Exception as e:
                print(e)
                self.update = {'Name of Excel File':item['Name of Excel File'], 'Python Message': f'Error Posting {self.now()}'}
                print(self.update)
                self.sheet_instance.update_rows([self.update], 'Name of Excel File')

    def run(self):
        '''runs main script as intended'''
        data = self.grab_ss_data()
        self.refresh_each_excel(data)
        self.sheet_instance.handle_update_stamps()


if __name__ == "__main__":
    config = {
        'smartsheet_token':smartsheet_token,
        'sheet_id': 6463522670071684
    }
    pqr=PQRefresher(config)
    pqr.run()