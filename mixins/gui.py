import tkinter
import os
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import askyesno
from abc import ABC

class GuiMixin(ABC):
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
