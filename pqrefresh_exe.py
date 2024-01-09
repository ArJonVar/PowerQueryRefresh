#region imports
from smartsheet import Smartsheet, sheets, smartsheet
import os
import traceback
import sys
import concurrent.futures
import pandas as pd
from logger import ghetto_logger
from datetime import datetime
import win32com.client as win32
import pythoncom
from datetime import datetime
import time
import inspect
from cryptography.fernet import Fernet
#endregion

class PQRefresher():
    '''Explain Class'''
    def __init__(self, config):
        self.config = config
        self.smartsheet_token=config.get('smartsheet_token')
        self.sheet_id=config.get('sheet_id')
        grid.token=self.smartsheet_token
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
            excel_app.Visible = False #set to True for debugging
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
            self.log.log("full traceback & error:")
            self.log.log(traceback.format_exc())
            self.log.log(sys.exc_info())
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
            configurednenabled_df = configurednenabled_df[configurednenabled_df['Refresh Frequency'] != 'Weekly']
        list_of_datadicts = configurednenabled_df.to_dict(orient='records')
        return list_of_datadicts
    def refresh_each_excel(self, list_of_datadicts):
        '''loops through paths in smartsheet and refreshes each path'''
        self.update = []
        self.log.log(f"{len(list_of_datadicts)} items to update")
        for i, item in enumerate(list_of_datadicts):
            self.log.log(f"updating item {i+1}: {item['Name of Excel File']}")
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
        self.update = self.refresh_each_excel(self.data)
        self.log.log('posting updates...')
        self.sheet_instance.update_rows(self.update, 'Name of Excel File')
        self.sheet_instance.handle_update_stamps()
        self.log.log('~fin~')

class grid:
    """
    A class that interacts with Smartsheet using its API.

    This class provides functionalities such as fetching sheet content, 
    and posting new rows to a given Smartsheet sheet.

    Important:
    ----------
    Before using this class, the 'token' class attribute should be set 
    to the SMARTSHEET_ACCESS_TOKEN.

    Attributes:
    -----------
    token : str, optional
        The access token for Smartsheet API.
    grid_id : int
        ID of an existing Smartsheet sheet.
    grid_content : dict, optional
        Content of the sheet fetched from Smartsheet as a dictionary.

    Methods:
    --------
    get_column_df() -> DataFrame:
        Returns a DataFrame with details about the columns, such as title, type, options, etc.

    fetch_content() -> None:
        Fetches the sheet content from Smartsheet and sets various attributes like columns, rows, row IDs, etc.

    fetch_summary_content() -> None:
        Fetches and constructs a summary DataFrame for summary columns.

    reduce_columns(exclusion_string: str) -> None:
        Removes columns from the 'column_df' attribute based on characters/symbols provided in the exclusion_string.

    grab_posting_column_ids(filtered_column_title_list: Union[str, List[str]]="all_columns") -> None:
        Prepares a dictionary for column IDs based on their titles. Used internally for posting new rows.

    delete_all_rows() -> None:
        Deletes all rows in the current sheet.

    post_new_rows(posting_data: List[Dict[str, Any]], post_fresh: bool=False, post_to_top: bool=False) -> None:
        Posts new rows to the Smartsheet. Can optionally delete the whole sheet before posting or set the position of the new rows.

    update_rows(posting_data: List[Dict[str, Any]], primary_key: str):
        Updates rows that can be updated, posts rows that do not map to the sheet.

    grab_posting_row_ids(posting_data: List[Dict[str, Any]], primary_key: str):
        returns a new posting_data called update_data that is a dictionary whose key is the row id, and whose value is the dictionary for the row <column name>:<field value>

    
    Dependencies:
    -------------
    - smartsheet (from smartsheet-python-sdk)
    - pandas as pd
    """

    token = None

    def __init__(self, grid_id):
        # self.log=ghetto_logger("smartsheet_grid.py")
        self.grid_id = grid_id
        self.grid_content = None
        # self.log.log(self.token)
        if self.token == None:
            return "MUST SET TOKEN"
        else:
            self.smart = smartsheet.Smartsheet(access_token=self.token)
            self.smart.errors_as_exceptions(True)
#region core get requests   
    def get_column_df(self):
        '''returns a df with data on the columns: title, type, options, etc...'''
        if self.token == None:
            return "MUST SET TOKEN"
        else:
            return pd.DataFrame.from_dict(
                (self.smart.Sheets.get_columns(
                    self.grid_id, 
                    level=2, 
                    include='objectValue', 
                    include_all=True)
                ).to_dict().get("data"))
    def fetch_content(self):
        '''this fetches data, ask coby why this is seperated
        when this is done, there are now new objects created for various scenarios-- column_ids, row_ids, and the main sheet df'''
        if self.token == None:
            return "MUST SET TOKEN"
        else:
            self.grid_content = (self.smart.Sheets.get_sheet(self.grid_id)).to_dict()
            self.grid_name = (self.grid_content).get("name")
            self.grid_url = (self.grid_content).get("permalink")
            # this attributes pulls the column headers
            self.grid_columns = [i.get("title") for i in (self.grid_content).get("columns")]
            # note that the grid_rows is equivelant to the cell's 'Display Value'
            self.grid_rows = []
            if (self.grid_content).get("rows") == None:
                self.grid_rows = []
            else:
                for i in (self.grid_content).get("rows"):
                    b = i.get("cells")
                    c = []
                    for i in b:
                        l = i.get("displayValue")
                        m = i.get("value")
                        if l == None:
                            c.append(m)
                        else:
                            c.append(l)
                    (self.grid_rows).append(c)
            
            # resulting fetched content
            self.grid_rows = self.grid_rows
            if (self.grid_content).get("rows") == None:
                self.grid_row_ids = []
            else:
                self.grid_row_ids = [i.get("id") for i in (self.grid_content).get("rows")]
            self.grid_column_ids = [i.get("id") for i in (self.grid_content).get("columns")]
            self.df = pd.DataFrame(self.grid_rows, columns=self.grid_columns)
            # Should be row_id intead of id as that is less likely to be taken name space!!!
            self.df["id"]=self.grid_row_ids
            self.column_df = self.get_column_df()
    def fetch_summary_content(self):
        '''builds the summary df for summary columns'''
        if self.token == None:
            return "MUST SET TOKEN"
        else:
            self.grid_content = (self.smart.Sheets.get_sheet_summary_fields(self.grid_id)).to_dict()
            # this attributes pulls the column headers
            self.summary_params=['title','createdAt', 'createdBy', 'displayValue', 'formula', 'id', 'index', 'locked', 'lockedForUser', 'modifiedAt', 'modifiedBy', 'objectValue', 'type']
            self.grid_rows = []
            if (self.grid_content).get("data") == None:
                self.grid_rows = []
            else:
                for summary_field in (self.grid_content).get("data"):
                    row = []
                    for param in self.summary_params:
                        row_value = summary_field.get(param)
                        row.append(row_value)
                    self.grid_rows.append(row)
            if (self.grid_content).get("rows") == None:
                self.grid_row_ids = []
            else:
                self.grid_row_ids = [i.get("id") for i in (self.grid_content).get("data")]
            self.df = pd.DataFrame(self.grid_rows, columns=self.summary_params)
#endregion 
#region helpers     
    def reduce_columns(self,exclusion_string):
        """a method on a grid{sheet_id}) object
        take in symbols/characters, reduces the columns in df that contain those symbols"""
        if self.token == None:
            return "MUST SET TOKEN"
        else:
            regex_string = f'[{exclusion_string}]'
            self.column_reduction =  self.column_df[self.column_df['title'].str.contains(regex_string,regex=True)==False]
            self.reduced_column_ids = list(self.column_reduction.id)
            self.reduced_column_names = list(self.column_reduction.title)
#endregion
#region ss post
    #region new row(s)
    def grab_posting_column_ids(self, filtered_column_title_list="all_columns"):
        '''preps for ss post 
        creating a dictionary per column:
        { <title of column> : <column id> }
        filtered column title list is a list of column title str to prep for posting (if you are not posting to all columns)
        [NOT USED INDEPENDENTLY, BUT USED INSIDE OF POST_NEW_ROWS]'''

        column_df = self.get_column_df()

        if filtered_column_title_list == "all_columns":
            filtered_column_title_list = column_df['title'].tolist()
    
        self.column_id_dict = {title: column_df.loc[column_df['title'] == title]['id'].tolist()[0] for title in filtered_column_title_list}
    def delete_all_rows(self):
        '''deletes up to 400 rows in 200 row chunks by grabbing row ids and deleting them one at a time in a for loop
        [NOT USED INDEPENDENTLY, BUT USED INSIDE OF POST_NEW_ROWS]'''
        self.fetch_content()

        row_list_del = []
        for rowid in self.df['id'].to_list():
            row_list_del.append(rowid)
            # Delete rows to sheet by chunks of 200
            if len(row_list_del) > 199:
                self.smart.Sheets.delete_rows(self.grid_id, row_list_del)
                row_list_del = []
        # Delete remaining rows
        if len(row_list_del) > 0:
            self.smart.Sheets.delete_rows(self.grid_id, row_list_del) 
    def post_new_rows(self, posting_data, post_fresh = False, post_to_top=False):
        '''posts new row to sheet, does not account for various column types at the moment
        posting data is a list of dictionaries, one per row, where the key is the name of the column, and the value is the value you want to post
        then this function creates a second dictionary holding each column's id, and then posts the data one dictionary at a time (each is a row)
        post_to_top = the new row will appear on top, else it will appear on bottom
        post_fresh = first delete the whole sheet, then post (else it will just update existing sheet)
        TODO: if using post_to_top==False, I should really delete the empty rows in the sheet so it will properly post to bottom'''
        
        posting_sheet_id = self.grid_id
        column_title_list = list(posting_data[0].keys())
        try:
            self.grab_posting_column_ids(column_title_list)
        except IndexError:
            raise ValueError("Index Error reveals that your posting_data dictionary has key(s) that don't match the column names on the Smartsheet")
        if post_fresh:
            self.delete_all_rows()
        
        rows = []

        for item in posting_data:
            row = smartsheet.models.Row()
            row.to_top = post_to_top
            row.to_bottom= not(post_to_top)
            for key in self.column_id_dict:
                if item.get(key) != None:     
                    row.cells.append({
                    'column_id': self.column_id_dict[key],
                    'value': item[key]
                    })
            rows.append(row)

        self.post_response = self.smart.Sheets.add_rows(posting_sheet_id, rows)
    #endregion
    #region post timestamp
    def handle_update_stamps(self):
        '''PUBLIC grabs summary id, and then runs the function that posts the date'''
        current_date = datetime.today()
        formatted_date = current_date.strftime('%m/%d/%y')

        sum_id = self.grabrcreate_sum_id("Last API Automation", "DATE")
        self.post_to_summary_field(sum_id, formatted_date)
    def grabrcreate_sum_id(self, field_name_str, sum_type):
        '''checks if there is a DATE summary field called "Last API Automation", if Y, pulls id, if N, creates the field.
        then posts today's date to that field
        [ONLY TESTED FOR DATE FIELDS FOR NOW]'''
        # First, let's fetch the current summary fields of the sheet
        self.fetch_summary_content()

        # Check if "Last API Automation" summary field exists
        automation_field = self.df[self.df['title'] == field_name_str]

        # If it doesn't exist, create it
        if automation_field.empty:
            new_field = smartsheet.models.SummaryField({
                "title": field_name_str,
                "type": sum_type
            })
            response = self.smart.Sheets.add_sheet_summary_fields(self.grid_id, [new_field])
            # Assuming the response has the created field's data, extract its ID
            self.sum_id = response.data[0].id
        else:
            # Extract the ID from the existing field
            self.sum_id = automation_field['id'].values[0]

        return self.sum_id
    def post_to_summary_field(self, sum_id, post):
        '''posts to sum field, 
        designed to: posts date to summary column to tell ppl when the last time this script succeeded was
        [ONLY TESTED FOR DATE FIELDS FOR NOW]'''

        sum = smartsheet.models.SummaryField({
            "id": int(sum_id),
            "ObjectValue": post
        })
        resp = self.smart.Sheets.update_sheet_summary_fields(
            self.grid_id,    # sheet_id
            [sum],
            False    # rename_if_conflict
        )
    #endregion
    #region post row update
    def grab_posting_row_ids(self, posting_data, primary_key, skip_nonmatch=False):
        '''Prepares for an update by reorganizing the posting data with the row_id as the key and the value as the data.    

        Parameters:
        - posting_data: Dictionary where each key is a column name and each value is the corresponding row value for that column.
        - primary_key: A key from `posting_data` that serves as the reference to map row IDs to the posting data (must be case-sensitive match).
        - skip_nonmatch (optional, default=True): Determines the handling of non-matching primary keys. When set to `True`, rows with non-matching primary keys are ignored. When `False`, these rows are collected into a "new_rows" key in the resulting dictionary.  

        Process:
        1. Identify the value associated with the `primary_key` in `posting_data`.
        2. Search for this value in the Smartsheet to find its row_id.
        3. Return a dictionary: keys are row_ids (or "new_rows" for unmatched rows), values are the corresponding `posting_data` for each row.
        '''

        self.fetch_content()

        if not self.df.empty:
            # Mapping of the primary key values to their corresponding row IDs from the current Smartsheet data
            primary_to_row_id = dict(zip(self.df[primary_key], self.df['id']))  

            # Dictionary to hold the mapping of row IDs to their posting data
            update_data = {}
            new_rows = []   

            for data in posting_data:
                primary_value = data.get(primary_key)
                if primary_value in primary_to_row_id:
                    row_id = primary_to_row_id[primary_value]
                    update_data[row_id] = data
                elif not skip_nonmatch:
                    new_rows.append(data)   

            if new_rows:
                update_data['new_rows'] = new_rows  

            # Check if there were no matches at all
            if not update_data:
                raise ValueError(f"The primary_key '{primary_key}' had no matches in the current Smartsheet data.") 

            return update_data
        else:
            raise ValueError("Grid Instance is not appropriate for this task. Try create a new grid instance")
    def update_rows(self, posting_data, primary_key):
        '''
        Updates rows (and adds misc rows) in the Smartsheet based on the provided posting data.  

        Parameters:
        - posting_data (list of dicts)
        - primary_key (column name with ONLY unique values (to use to find column id))

        Returns:
        None. Updates and possibly adds rows in the Smartsheet.
        '''
        posting_sheet_id = self.grid_id
        column_title_list = list(posting_data[0].keys())
        try:
            self.grab_posting_column_ids(column_title_list)
        except IndexError:
            raise ValueError("Index Error reveals that your posting_data dictionary has key(s) that don't match the column names on the Smartsheet")
        self.update_data = self.grab_posting_row_ids(posting_data, primary_key)

        rows = []
        # Handle existing rows' updates
        for row_id in self.update_data.keys():
            if row_id != "new_rows":
                # Build the row to update
                new_row = smartsheet.models.Row()
                new_row.id = row_id
                for column_name in self.column_id_dict.keys():
                    # Build new cell value
                    new_cell = smartsheet.models.Cell()
                    new_cell.column_id = int(self.column_id_dict[column_name])
                    # stops error where post doesnt go through because value is "None"
                    if self.update_data[row_id].get(column_name) != None:
                        new_cell.value = self.update_data[row_id].get(column_name)
                    else:
                        new_cell.value = ""
                    new_cell.strict = False
                    new_row.cells.append(new_cell)
                rows.append(new_row)

        # Update rows
        self.update_response = self.smart.Sheets.update_rows(
          posting_sheet_id ,      # sheet_id
          rows)

        try:
            # Handle addition of new rows if the "new_rows" key is present
            self.post_new_rows(self.update_data.get('new_rows'))
        except TypeError:
            pass
    #endregion
#endregion

class ghetto_logger:
    '''to deploy in class, put self.log=ghetto_logger("<module name>.py"), then ctr f and replace print( w/ self.log.log('''
    def __init__(self, title, print = True):
        raw_now = datetime.now()
        self.print= print
        self.now = raw_now.strftime("%m/%d/%Y %H:%M:%S")
        self.first_use=True
        self.first_line_stamp  = f"{self.now}  {title}--"
        self.start_time = time.time()
        if os.name == 'nt':
            directory = os.path.dirname(r'C:\Egnyte\Shared\IT\Python\Ariel\power_query_refresh\excel\log.txt')
            logger_name = 'log.txt'
            self.path = os.path.join(directory, logger_name)
        else:
            self.path ="log.txt"

    def timestamp(self): 
        '''creates a string of minute/second from start_time until now for logging'''
        end_time = time.time()  # get the end time of the program
        elapsed_time = end_time - self.start_time  # calculate the elapsed time in seconds       

        minutes, seconds = divmod(elapsed_time, 60)  # convert to minutes and seconds       
        timestamp = "{:02d}:{:02d}".format(int(minutes), int(seconds))
        
        return timestamp
    
    def log(self, text, type = "new_line", mode="a"):
        # so lists/dictionaries/etc can be logged without issue
        text = str(text)

        function_name = inspect.currentframe().f_back.f_code.co_name
        
        try:
            module_name = inspect.getmodule(inspect.stack()[1][0]).__name__
        except:
            module_name = "__main__"

        func_stamp = f"{self.timestamp()}  {module_name}.{function_name}(): "

        if self.print == True:
            print(f"{func_stamp} {text}")

        with open(self.path, mode=mode) as file:
            if self.first_use == True:
                file.write("\n" + "\n"+ self.first_line_stamp)
                self.first_use = False
            if self.first_use == False and type == "paragraph":
                file.write(text)
            elif self.first_use == False:
                file.write("\n  " + func_stamp + text) 

if __name__ == "__main__":
    token = bytes(os.environ.get('smartsheet_token'), 'utf-8')
    key = os.environ.get('smartsheet_key')

    config = {
        'smartsheet_token':Fernet(key).decrypt(token).decode("utf-8"),
        'sheet_id': 6463522670071684
    }
    pqr=PQRefresher(config)
    pqr.run()

