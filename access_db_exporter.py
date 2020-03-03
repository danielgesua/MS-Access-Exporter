import win32com.client
import os

class ms_access_automation():
    '''object that uses COM to communicate with MS Access to get all the code from its modules and tabulate it in a python list.'''

    @property
    def module_names(self):
        self._module_names = [module.Name for module in self.ac.CurrentProject.AllModules] if self._module_names is None else self._module_names
        return self._module_names

    @property
    def modules(self):
        return self.ac.Modules

    @property
    def form_names(self):
        self._form_names = [form.Name for form in self.ac.CurrentProject.AllForms] if self._form_names is None else self._form_names
        return self._form_names

    @property
    def forms(self):
        return self.ac.Forms

    @property
    def form_modules(self):

        def p_form_modules(index):
            return self.forms(index).Module
        
        return p_form_modules

    @property
    def query_names(self):
        self._query_names = [query.Name for query in self.ac.CurrentData.AllQueries] if self._query_names is None else self._query_names
        return self._query_names

    @property
    def query_defs(self):
        return self.ac.CurrentDb().QueryDefs

    def run(self,displaying_prompts = True):
        '''runs the automation'''
        
        def _open_access_file():
            '''opens the application object and makes it visible'''
            self.ac=win32com.client.Dispatch('Access.Application')
            self.ac.OpenCurrentDatabase(self.db_path)
            self.ac.Visible=True
    
        def _get_all_module_obj_data(names_list,obj_list,is_form):
            '''place the module names, types and code inside the list of tuples for any given list of module names and objects'''
                    
            def _open_obj(obj_name):
                '''uses the corresponding automation routine to open the object based on the type of objects in the list'''

                #get the appropriate opening routine then pass the right parameters based on object type
                open_method = self.ac.DoCmd.OpenForm if is_form else self.ac.DoCmd.OpenModule
                open_method(obj_name,1) if is_form else open_method(obj_name)
            
            def _close_obj(obj_name):
                '''closes the object based on the type of objects in the list'''

                obj_type = 2 if is_form else 5
                if is_form:
                    self.ac.DoCmd.RunCommand(58)
                else:
                    self.ac.DoCmd.Close(obj_type,obj_name,2)            

            def _mine_the_object_data(obj_name):
                '''gets the necessary data from the object'''

                def _has_or_is_module(obj_name):
                    '''returns the value of HasModule for form objects or true for module objects'''
                    return self.forms(obj_name).HasModule if is_form else True

                def _get_module_code(module):
                    '''obtains the code contained inside a module and returns it as a string'''
                    return module.Lines(1,module.CountOfLines)

                def _get_module_type(module):
                    '''obtains the type of module and returns the int that represents it'''
                    return 2 if (is_form and _has_or_is_module(obj_name)) else module.Type

                def _corrected_object_name(name):
                    '''corrects the name based on type'''
                    return 'Form_' + name if is_form else name
                
                if _has_or_is_module(obj_name):
                    code = _get_module_code(obj_list(obj_name))
                    module_type = _get_module_type(obj_list(obj_name))  
                else:
                    code = None
                    module_type = 2
                name = _corrected_object_name(obj_name)
                self._module_data += [(name,module_type,code)]
            
            #for each module in the list, fetch the lines of code and add it to the list of tuples alongside the name
            for name in names_list:
                if displaying_prompts: print('Mining "' + name + '" for data...', end=" ")
                _open_obj(name)
                _mine_the_object_data(name)
                _close_obj(name)
                if displaying_prompts: print('Done!!!')

            #if displaying prompts then add one new line between the prompts of this portion and the next.
            if displaying_prompts: print('\n',end='')

        def _get_all_query_obj_data():
            '''place all query names, and sql inside a list of tuples.'''

            for query_name in self.query_names:
                sql = self.query_defs[query_name].SQL
                self._query_data += [(query_name,sql)]

        def _display_prompts():
            '''prints console prompts to show the developer what was mined'''

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
        _get_all_module_obj_data(self.module_names,self.modules,is_form=False)
        _get_all_module_obj_data(self.form_names,self.form_modules,is_form=True)
        _get_all_query_obj_data()
        if displaying_prompts: _display_prompts()

    def __init__(self):

        #Instance variables
        self.ac = None
        self._module_names = None
        self._form_names = None
        self._query_names = None
        self._form_modules = None
        self._module_data = []
        self._query_data = []

    def __del__(self):
        self.ac.CloseCurrentDatabase()
        self.ac.Quit()

class file_export_automation():
    '''object that can take the python list of module data from an ms_access automation and export each module as a file'''

    def __init__(self):
        self._file_ext_definitions = ['.bas','.cls','.cls']

    def run(self):
        '''takes all the modules in the python list of an ms_access_automation object and exports them as files'''
        
        def _ensure_directories_exist():
            '''creates all the necessary export directories and subdirectories if they do not exist'''

            def _ensure_exports_directory_exists():
                '''if the exports directory doesn't exist, it creates it'''

                def _export_directory_exists():
                    '''checks to see if the git_exports directory exists and returns true or false.'''
                    self._export_directory_path = os.path.join(project_directory_path,'git_exports')
                    return os.path.exists(path = self._export_directory_path)

                if not _export_directory_exists(): os.mkdir(path = self._export_directory_path)

            def _ensure_modules_directory_exists():
                '''if the modules subdirectory doesn't exist, it creates it'''

                def _modules_directory_exists():
                    '''checks to see if the git_exports directory exists and returns true or false.'''
                    self._modules_directory_path = os.path.join(self._export_directory_path,'modules')
                    return os.path.exists(path = self._modules_directory_path)

                if not _modules_directory_exists(): os.mkdir(path = self._modules_directory_path)

            def _ensure_queries_directory_exists():
                '''if the queries subdirectory doesn't exist, it creates it'''

                def _queries_directory_exists():
                    '''checks to see if the queries subdirectory exists and returns true or false.'''
                    self._queries_directory_path = os.path.join(self._export_directory_path,'queries')
                    return os.path.exists(path = self._queries_directory_path)

                if not _queries_directory_exists(): os.mkdir(path = self._queries_directory_path)

            _ensure_exports_directory_exists()
            _ensure_modules_directory_exists()
            _ensure_queries_directory_exists()

        def _save_all_modules():
            '''saves all the modules with the correct extension in the modules sub directory of the exports directory'''

            for (file_name,module_type,code) in self._module_data:
                if code is not None:
                    file_extension = self._file_ext_definitions[module_type]
                    full_name = os.path.join(self._modules_directory_path, file_name + file_extension)
                    with open(file = full_name, mode = 'w') as file:
                        file.write(code)

        def _save_all_queries():
            '''saves all the queries' sql in text format in the queries sub directory of the exports directory'''

            for (file_name,sql) in self._query_data:
                full_name = os.path.join(self._queries_directory_path,file_name + '.txt')
                with open(file = full_name,mode = 'w') as file:
                    file.write(sql)

        project_directory_path = os.path.abspath(os.path.dirname(self.db_path))
        _ensure_directories_exist()
        _save_all_modules()
        _save_all_queries()

class automation(ms_access_automation, file_export_automation):
    '''object that performs all the automations necessary to export the modules in an access database'''

    def __init__(self,db_path):
        self.db_path = db_path
        ms_access_automation.__init__(self)
        file_export_automation.__init__(self)

    def run(self):
        ms_access_automation.run(self)
        file_export_automation.run(self)

a = automation(r'C:\Test.accdb')
a.run()
