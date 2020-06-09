import os
import sys
import json
import http.client
import mimetypes
import traceback
from urllib.parse import quote
from mixins.gui import GuiMixin
from mixins.com_ops import ComOpsMixin

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
        '''Takes all the modules in the python list of an ComOpsMixin object and exports them as files.'''
        
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

class DiffWorthyExporter(ComOpsMixin, file_export_automation, GuiMixin):
    '''Object that performs all the automations necessary to export the modules in an access database.'''

    def __init__(self):
        '''Run all the base object initiators in the appropriate order.'''
        ComOpsMixin.__init__(self)
        file_export_automation.__init__(self)
        GuiMixin.__init__(self)
                
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

            ComOpsMixin.run(self)
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
        ComOpsMixin.__del__(self)
        GuiMixin.__del__(self)

try:
    DWE = DiffWorthyExporter()
    
    # Get the MS Access file's fully qualified path from command line argument (if it was provided). Otherwise pass in an
    # empty string.
    from_cmd_line= True if len(sys.argv) > 1 else False
    file_path = sys.argv[1] if len(sys.argv) > 1 else ''
    pretty_print_sql = sys.argv[2] == 'True' if len(sys.argv) > 2 else False
    DWE.run(file_path, pretty_print_sql)
    del(DWE)
except Exception:
    exc_type, exc_value, exc_traceback = sys.exc_info()
    traceback.print_exception(exc_type, exc_value, exc_traceback)
finally:
    wait_for_user = sys.argv[3] == 'True' if len(sys.argv) > 3 else False
    if wait_for_user: input("Press enter to continue...")

