import win32com.client
import os
import sys
import tkinter
import json
import http.client
import mimetypes
from urllib.parse import quote
from enum import IntFlag
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import askyesno

class ms_access_automation():
    '''Object that uses COM to communicate with MS Access to get all the code from its modules and tabulate it in a python list.'''

    @property
    def currentdb(self):
        '''gets a reference to the current database (if it is not already obtained) otherwise returns the existing reference.'''
        self._currentdb = self.ac.CurrentDb() if self._currentdb is None else self._currentdb
        return self._currentdb

    @property
    def module_names(self):
        '''Returns all the module names in the current project, and stores them for the next time they are needed.'''
        self._module_names = [module.Name for module in self.ac.CurrentProject.AllModules] if self._module_names is None else self._module_names
        return self._module_names

    @property
    def modules(self):
        '''Alias for the modules object.'''
        return self.ac.Modules

    @property
    def form_names(self):
        '''Returns all the form names in the current project, and stores them for the next time they are needed.'''
        self._form_names = [form.Name for form in self.ac.CurrentProject.AllForms] if self._form_names is None else self._form_names
        return self._form_names

    @property
    def forms(self):
        '''Alias for the forms object.'''
        return self.ac.Forms

    @property
    def form_modules(self):
        '''This wrapper alows us to use the form modules with the same syntax as regular modules.'''
        
        def p_form_modules(index):
            '''Returns the Module object for any form object'''
            return self.forms(index).Module
        
        return p_form_modules

    @property
    def query_names(self):
        '''Returns all the QueryDef names in the current project, and stores them for the next time they are needed.'''
        self._query_names = [query.Name for query in self.ac.CurrentData.AllQueries] if self._query_names is None else self._query_names
        return self._query_names

    @property
    def query_defs(self):
        '''Alias for the QueryDefs object.'''
        return self.currentdb.QueryDefs

    @property
    def table_names(self):
        '''Returns all the TableDef names in the current project, and stores them for the next time they are needed.'''
        self._table_names = [table.Name for table in self.ac.CurrentData.AllTables] if self._table_names is None else self._table_names
        return self._table_names

    @property
    def table_defs(self):
        '''Alias for the TableDefs object.'''
        return self.currentdb.TableDefs

    def run(self,displaying_prompts = True):
        '''Runs the automation. Displays console prompts by default but can be silent.'''
        
        def _open_access_file():
            '''Opens the Access application object to the database of interest, and makes it visible'''
            self.ac=win32com.client.Dispatch('Access.Application')
            self.ac.OpenCurrentDatabase(self.db_path)
            self.ac.UserControl=False
            for form in self.forms:
                self.ac.DoCmd.Close(2,form.Name)
            self.ac.Visible=True

        
        def _get_all_table_obj_data():
            '''Place all the data necessary to create a table into a list.'''

            class table_attributes(IntFlag):
                dbAttachedODBC = 536870912
                dbAttachedTable = 1073741824
                dbAttachExclusive = 65536
                dbAttachSavePWD = 131072
                dbHiddenObject = 1
                dbSystemObject = -2147483646

            def is_system_table():
                '''Determines if a table is a system table and if so returns true. Otherwise false.''' 
                return (self.table_defs[table_name].Attributes & table_attributes.dbSystemObject != 0)

            def _next_table_def():
                '''Gets the next record of tabledef related data that will be appended to the list'''
                
                def _next_field():
                    '''Gets the data related to the next field in the tabledef'''
                    field_obj_data = {
                        'name' : field.Name,
                        'type' : field.Type,
                        'required': field.Required,
                        'size' : field.Size,
                        'allow_zero_length' : field.AllowZeroLength
                    }

                    return field_obj_data

                tabledef_obj_data = {
                    'name' : table_name,
                    'fields' : []
                }

                for field in self.table_defs[table_name].Fields:
                    tabledef_obj_data['fields'] += [_next_field()]
                
                return tabledef_obj_data

            for table_name in self.table_names:
                if not is_system_table():
                    if displaying_prompts: print('Mining "' + table_name + '" for data...', end=" ")
                    self._table_data += [_next_table_def()]
                    if displaying_prompts: print('Done!!!')
        
        def _get_all_module_obj_data(names_list,obj_list,is_form):
            '''Place the module names, types and VBA code inside the list of tuples for any given list of module names and module objects.'''
                    
            def _open_obj(obj_name):
                '''Selects and runs the "open" method on the object based on its type.'''

                #get the appropriate opening routine then pass the right parameters based on object type
                open_method = self.ac.DoCmd.OpenForm if is_form else self.ac.DoCmd.OpenModule
                open_method(obj_name,1) if is_form else open_method(obj_name)
            
            def _close_obj(obj_name):
                '''Closes the object based on its type.'''

                obj_type = 2 if is_form else 5
                if is_form:
                    self.ac.DoCmd.RunCommand(58)
                else:
                    self.ac.DoCmd.Close(obj_type,obj_name,2)            

            def _mine_the_object_data(obj_name):
                '''Gets the necessary data from the object.'''

                def _has_or_is_module(obj_name):
                    '''Returns the value of HasModule for form objects or true for module objects.'''
                    return self.forms(obj_name).HasModule if is_form else True

                def _get_module_code(module):
                    '''Obtains the code contained inside a module and returns it as a string.'''
                    return module.Lines(1,module.CountOfLines)

                def _get_module_type(module):
                    '''Obtains the type of module and returns the int that represents it.'''
                    return 2 if (is_form and _has_or_is_module(obj_name)) else module.Type

                def _corrected_object_name(name):
                    '''Corrects the name of the module based on type. (Form modules are always prefaced by "Form_")'''
                    return 'Form_' + name if is_form else name
                
                # If the object is a module or is a for with a module then get its code and type.
                if _has_or_is_module(obj_name):
                    code = _get_module_code(obj_list(obj_name))
                    module_type = _get_module_type(obj_list(obj_name))  
                else:
                    code = None
                    module_type = 2

                # Correct the name if it is a form module and append the data to the list.
                name = _corrected_object_name(obj_name)
                self._module_data += [(name,module_type,code)]
            
            # For each module in the list, fetch the VBA code, and module type, and add the data to the list of tuples alongside the name.
            for name in names_list:
                if displaying_prompts: print('Mining "' + name + '" for data...', end=" ")
                _open_obj(name)
                _mine_the_object_data(name)
                _close_obj(name)
                if displaying_prompts: print('Done!!!')

            #if displaying prompts then add one new line between the prompts of this portion and the next.
            if displaying_prompts: print('\n',end='')

        def _get_all_query_obj_data():
            '''Place all query names, and SQL inside a list of tuples.'''

            for query_name in self.query_names:
                sql = self.query_defs[query_name].SQL
                try:
                    if(not self.pretty_print_sql):
                        ''' skip to end '''
                        raise Exception('Prettify disabled')
                    print('Prettifying: ' + query_name)
                    sql_urlencoded = quote(sql)
                    ''' If pretty_print_sql '''
                    conn = http.client.HTTPSConnection("sql-format.com")
                    payload = f'text={sql_urlencoded}&options=%7B%7D&caretPosition%5Bx%5D=0&caretPosition%5By%5D=1&saveHistory=false'
                    headers = {
                        'Content-Type': 'application/x-www-form-urlencoded'
                    }
                    conn.request("POST", "", payload, headers)
                    res = conn.getresponse()
                    data = res.read()
                    response_data = data.decode("utf-8")
                    response_json = json.loads(response_data)
                    if(response_json['Text']):
                        sql = response_json['Text']
                except Exception:
                    print('Could not pretty print SQL')
                except:
                    print('Could not pretty print SQL')
                self._query_data += [(query_name,sql)]

        def _display_prompts():
            '''Prints console prompts to show the developer what was mined.'''

            for module_num,(name,module_type,code) in enumerate(self._module_data):
                print('Module #' + str(module_num + 1) + ': ' + name)
                print('Type: ' + str(module_type))
                code = 'Code: No code.' if code is None else 'Code: Obtained!'
                print(code,end = '\n\n')

            for query_index, query_name in enumerate(self.query_names): 
                print('Query #' + str(query_index + 1) + ': ' + query_name)
                sql = 'SQL: Empty QueryDef.' if self.query_defs[query_name].SQL is None else 'SQL: Obtained!'
                print(sql,end = '\n\n')

        _open_access_file()
        _get_all_table_obj_data()
        _get_all_module_obj_data(self.module_names,self.modules,is_form=False)
        _get_all_module_obj_data(self.form_names,self.form_modules,is_form=True)
        _get_all_query_obj_data()
        if displaying_prompts: _display_prompts()

    def __init__(self):

        self.ac = None
        self._currentdb = None
        self._table_names = None
        self._module_names = None
        self._form_names = None
        self._query_names = None
        self._form_modules = None
        self._module_data = []
        self._query_data = []
        self._table_data = []

    def __del__(self):
        '''Ensure objects are closed when done using them.'''
        if self.ac is not None:
            self.currentdb.Close()
            self.ac.CloseCurrentDatabase()
            self.ac.Quit()

class file_export_automation():
    '''Object that can take the python list of module/query data from an ms_access automation and export each document as a file.'''

    def __init__(self):

        # The following variable translates the module type into the file extension:
        # 0 = standard module (*.bas)
        # 1 = class module (*.cls)
        # 2 = form module (*.cls)
        # Note: Microsoft's ModuleType for form modules is the same as class modules, but we need to have them differentiated,
        # for naming reasons so we defined it as "2".
        self._file_ext_definitions = ['.bas','.cls','.cls']

    def run(self):
        '''Takes all the modules in the python list of an ms_access_automation object and exports them as files.'''
        
        def _ensure_directories_exist():
            '''Creates all the necessary export directories and subdirectories if they do not exist.'''

            def _ensure_exports_directory_exists():
                '''If the exports directory doesn't exist, this method creates it.'''

                def _export_directory_exists():
                    '''Checks to see if the git_exports directory exists and returns true or false.'''
                    self._export_directory_path = os.path.join(project_directory_path,'git_exports')
                    return os.path.exists(path = self._export_directory_path)

                # If the directory doesn't exist then create it.
                if not _export_directory_exists(): os.mkdir(path = self._export_directory_path)

            def _ensure_tables_directory_exists():
                '''If the tables subdirectory doesn't exist, this method creates it.'''

                def _tables_directory_exists():
                    '''Checks to see if the git_exports directory exists and returns true or false.'''
                    self._tables_directory_path = os.path.join(self._export_directory_path,'tables')
                    return os.path.exists(path = self._tables_directory_path)

                if not _tables_directory_exists(): os.mkdir(path = self._tables_directory_path)

            def _ensure_modules_directory_exists():
                '''If the modules subdirectory doesn't exist, this method creates it.'''

                def _modules_directory_exists():
                    '''Checks to see if the git_exports directory exists and returns true or false.'''
                    self._modules_directory_path = os.path.join(self._export_directory_path,'modules')
                    return os.path.exists(path = self._modules_directory_path)

                # If the directory doesn't exist then create it.
                if not _modules_directory_exists(): os.mkdir(path = self._modules_directory_path)

            def _ensure_queries_directory_exists():
                '''If the queries subdirectory doesn't exist, this method creates it.'''

                def _queries_directory_exists():
                    '''Checks to see if the queries subdirectory exists and returns true or false.'''
                    self._queries_directory_path = os.path.join(self._export_directory_path,'queries')
                    return os.path.exists(path = self._queries_directory_path)

                # If the directory doesn't exist then create it.
                if not _queries_directory_exists(): os.mkdir(path = self._queries_directory_path)

            _ensure_exports_directory_exists()
            _ensure_tables_directory_exists()
            _ensure_modules_directory_exists()
            _ensure_queries_directory_exists()

        def _save_all_tables():
            '''Saves all the queries' SQL in text format in the queries sub directory of the exports directory.'''

            # For each QueryDef in the Accdb.
            for table in self._table_data:

                # Build the fully qualified file name.
                full_name = os.path.join(self._tables_directory_path,table["name"] + '.txt')
                
                 # Export the SQL code.
                with open(file = full_name,mode = 'w') as file:
                    file.write(json.dumps(table, indent=2))

        def _save_all_modules():
            '''Saves all the modules with the correct extension in the modules sub directory of the exports directory.'''

            # For each class module, standard module, and [non-empty] form module.
            for (file_name,module_type,code) in self._module_data:
                if code is not None:
                    
                    # Translate the module type into a file extension.
                    file_extension = self._file_ext_definitions[module_type]

                    # Build the fully qualified file name.
                    full_name = os.path.join(self._modules_directory_path, file_name + file_extension)

                    # Export the code.
                    with open(file = full_name, mode = 'w') as file:
                        file.write(code)

        def _save_all_queries():
            '''Saves all the queries' SQL in text format in the queries sub directory of the exports directory.'''

            # For each QueryDef in the Accdb.
            for (file_name,sql) in self._query_data:

                # Build the fully qualified file name.
                full_name = os.path.join(self._queries_directory_path,file_name + '.txt')

                # Export the SQL code.
                with open(file = full_name,mode = 'w') as file:
                    file.write(sql)

        project_directory_path = os.path.abspath(os.path.dirname(self.db_path))
        _ensure_directories_exist()
        
        print('Writing files to: ' + self._export_directory_path)

        _save_all_modules()
        _save_all_queries()
        _save_all_tables()

class gui():
    '''Validating inputs and (if necessary) opening a file dialog to request a valid MS Access file.'''

    def _file_is_valid(self):
        '''Checks to see if the file exists and is of the right extension. Returns true for "valid" and false otherwise'''

        # If the file exists and is of appropriate extension then return true otherwise false.
        if os.path.exists(self.db_path):
            response = (os.path.splitext(self.db_path)[1] in ['.accdb'])
        else:
            response = False

        return response 

    def ask_for_db_path(self):
        '''Asks the user for a path to the MS Access file, checks if it's valid, then--if it isn't--prompts the user for retry.'''

        def _show_file_dialog_to_get_db_path():
            '''Displays the file dialog that gets the MS Access file's path.'''
            self.db_path = askopenfilename(
                parent = self.window,
                title = 'Pick MS Access database to export.',
                filetypes = [('MS Office Databases', '*.accdb',)]
                )

        def _confirm_if_user_wants_to_retry():
            '''Gets user confirmation to retry the file selection.'''
            
            nonlocal retry

            # Build the user console prompt based on scenario:
            #   Scenario #1) An invalid file name provided. Prompt should say the file was invalid.
            #   Scenario #2) The file name was empty (cancel or close button on the dialog). Prompt should say no file selected.
            user_prompt = 'The file "' + self.db_path + '" is not valid. ' if self.db_path != '' else 'No file selected. '
            user_prompt += 'Would you like to try again?'

            # Select the dialog's icon based on the scenario:
            #   Scenario #1) Invalid file name. Display an error icon.
            #   Scenario #2) Empty file name. Display a waring icon.
            msgbox_icon = 'warning' if self.db_path == '' else 'error'

            # Prompt the user for retry confirmation. and set the non-local variable.
            retry = askyesno(
                parent = self.window,
                title = 'Invalid file.',
                message = user_prompt,
                icon = msgbox_icon
            )

        def _create_main_window():
            '''Creates the main window obj and hides it.'''
            self.window = tkinter.Tk()
            self.window.withdraw()
    
        # Create the main GUI window.
        _create_main_window()

        # No need to retry if we get it right the first time.
        retry = False

        # Show a file dialog to the user so they can pick an Accdb.
        _show_file_dialog_to_get_db_path()
        
        # If the file is not valid confirm if the user would like to try their choice again.
        if not self._file_is_valid(): _confirm_if_user_wants_to_retry()

        # If the file was valid or the user wishes to try again, then return to the begining of this routine.
        if retry: self.ask_for_db_path()    

    def __init__(self):
        self.window = None

    def __del__(self):
        '''Closes the window when the program ends.'''
        if self.window is not None: self.window.destroy()

class automation(ms_access_automation, file_export_automation, gui):
    '''Object that performs all the automations necessary to export the modules in an access database.'''

    def __init__(self):
        '''Run all the base object initiators in the appropriate order.'''
        ms_access_automation.__init__(self)
        file_export_automation.__init__(self)
        gui.__init__(self)
                
    def run(self, db_path = '', pretty_print_sql = False):
        '''Runs the automation on a given path.'''
        self.pretty_print_sql = pretty_print_sql
        def _perform_first_check():
            '''Performs a first round of checks to see if the path is valid and otherwise requests a path using the file dialog.'''
            
            # If the file is not valid then give the user a console prompt and a chance to change their choice.
            if not self._file_is_valid():
                if db_path != '': print('Inputted file path of "' + db_path + '" is invalid! Please choose a valid .accdb or cancel.')
                self.ask_for_db_path()

        def _run():
            '''Uses the inputted parameters to run the automation'''

            ms_access_automation.run(self)
            file_export_automation.run(self)

        self.db_path = db_path

        # Checks the file path selected. If it's invalid or empty then requests new file using a file dialog.
        _perform_first_check()

        # If file is valid then run the automation, otherwise (if the user canceled the file dialog without 
        # a valid file) then respond with a console prompt and exit out of the routine.
        if not self._file_is_valid():
            print('File was invalid! Export aborted.')
        else:
            _run()

    def __del__(self):
        '''Performs all the neccesary cleanup processes and closes the files.'''
        ms_access_automation.__del__(self)
        gui.__del__(self)

a = automation()
 
# Get the MS Access file's fully qualified path from command line argument (if it was provided). Otherwise pass in an
# empty string.
file_path = sys.argv[1] if len(sys.argv) > 1 else ''
pretty_print_sql = sys.argv[2] == 'True' if len(sys.argv) > 2 else False
a.run(file_path, pretty_print_sql)

