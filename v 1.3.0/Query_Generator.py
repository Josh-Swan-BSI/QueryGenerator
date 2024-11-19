#TODO
# create classes for individual items e.g. dropdown, text boxes and buttons. Make a new class for a text box that is very different from another text box for example
# move the functions into methods of the main application class

from tkinter import *
from tkinter import ttk
from idlelib.tooltip import Hovertip
import pyperclip
import re
import os
import zipfile
import sys
import shutil
import subprocess
from tkinter import messagebox
import requests
from bs4 import BeautifulSoup
import webbrowser
import warnings
import random

warnings.filterwarnings(action="ignore", module="urllib3")

'''
Date: 18/11/2024
Author: Josh Swan

Useful Notes
If you want to make variables in methods accessible to other method, it needs to be initialised properly i.e. do
self.variable = value instead of variable = value.

'''
'''Global Variables----------------------------------------------------------------------------------------------------------'''
RELEASE = ""
VERSION = "1.3.0"
VERSION_URL = "https://josh-swan-bsi.github.io/QG_Help_Guide/version.html"
REPO = "https://github.com/Josh-Swan-BSI/QueryGenerator/raw/refs/heads/main/Query_Generator.zip"
CONFIG_FILE = "config.txt"

# variable colours
headerColour = "#131c36"
bgColour = "#595959"
textColour = "#ffffff"
boxColour = "#DAE3F3"
lineColour = "#A6A6A6"

'''GLOBAL Functions'''
def convertVersion(version: str) -> int:
    '''
    This converts a version number e.g. 1.2.0 into an int to be able to let the program work out the latest version
    '''

    output = ""
    for c in version:
        try:
            c = int(c)
            if len(output) > 0 or c != 0:
                output = output + str(c)

        except:
            pass
    return int(output)

'''CLASS DECLARATIONS-------------------------------------------------------------------------------------------------'''
class QueryGenerator:
    '''
    This is the main class that establishes the query generator app
    '''

    def __init__(self, root):
        # used for keeping track of number of rows of input
        self.colCount = 1

        # used to extract inputs from users
        self.columnTitles = []
        self.textBoxes = []
        self.logic = []

        self.root = root
        self.root.state("zoomed") # full window
        self.root.title(f"Query Generator v{VERSION} {RELEASE}")
        self.root.iconbitmap('icon.ico')  # established the app icon
        self.root.config(background=bgColour)

        self.setup_ui() # called the setup_ui method to create the layout of the app.

    def setup_ui(self): # this sets up your app layout
        '''
        Putting this into a separate method helps organise the code into sections
        '''

        # Configuring the first frame that contains the logo---------------------------------------------------------------------------
        self.logo = PhotoImage(file="logosmall_xmas_final.png") # if you don't use self garbage collection will remove it from memory and it won't display
        colFrame = Frame(self.root, bg=headerColour)
        my_label = Label(colFrame, image=self.logo, bd=0).pack(anchor=W, padx=5, pady=3)  # bd=0 is border
        colFrame.pack(fill=X)

        # Setting up the first grid containing SQL Table text box---------------------------------------------------------------------------
        gridFrame2 = Frame(self.root, bg=bgColour)
        gridFrame2.columnconfigure(0, weight=1)
        gridFrame2.rowconfigure(0, weight=1)

        #SQL Table Label
        label1 = Label(gridFrame2, text="SQL Table", font=('Raleway', 12), bg=bgColour, fg=textColour, anchor=W, padx=5, pady=5)
        label1.grid(row=0, column=0, sticky="NSEW", padx=10, pady=10)
        Hovertip(label1, "The table you would like to query.\nUsually this is the latest table available.")

        # SQL Table TextBox
        self.textbox1 = Text(gridFrame2, padx=10, pady=10, font=('Raleway', 12), width=15, height=1, bg=boxColour)
        self.textbox1.grid(row=0, column=1, sticky="NSEW", padx=10, pady=10)
        Hovertip(self.textbox1, "The table you would like to query.\nUsually this is the latest table available.")

        # Column view selector, Standard columns or All columns
        self.columnSelect = dropdown(gridFrame2, 0, 2, 20, 10, ["Standard Columns", "All Columns"], "enabled",
                                '''All Columns:\n All columns from the SQL Server
                                \nStandard Columns:
                                -ACCode
                                -DocumentIdentifier
                                -PublicationDate
                                -Title_English
                                -CommitteeReference
                                -Descriptors_English
                                -IssuingBody
                                -Update_Monthly
                                -CountryCode
                                -Abstract_English
                                -Status
                                -InternationalRelationship
                                -Classification
                                -CrossReferences''')

        # Help Button
        documentation = Button(gridFrame2, text="Open Help Doc", font=('Raleway', 12), bg=headerColour, fg=textColour, anchor=E, padx=5,
                         pady=5, command=self.openDocumentation)
        documentation.grid(row=0, column=3, sticky="NSEW", padx=150, pady=10)
        Hovertip(documentation, "Click here to open documentation in MS Word")

        gridFrame2.pack(side=TOP, anchor=W)

        self.helperLabel = Label(gridFrame2, text='TIP: Use capitalised AND, OR, AND NOT     ', font=('Raleway', 11), bg=bgColour, fg=textColour, anchor=W, padx=5, pady=5,
                       wraplength=300, justify="left")
        self.helperLabel.grid(row=0, column=4, sticky="NSEW", padx=5, pady=5)


        # A separation line visible for the user to separate the sections---------------------------------------------------------------------------
        Frame1 = Frame(self.root, bg=lineColour)
        Frame1.pack(fill=X, pady=5)

        #add buttons here
        # Adding in the check button
        # active background - stops it flashing when you check it. Select colour is the colour of the tickbox
        checkGrid = Frame(root, bg=bgColour)
        checkGrid.columnconfigure(0, weight=1)
        checkGrid.rowconfigure(0, weight=1)
        self.checkVariable = IntVar()
        check = Checkbutton(checkGrid, text="Only valid records", variable=self.checkVariable, font=('Raleway', 11), anchor=W,
                            bg=bgColour, fg=textColour,
                            selectcolor="#2C2C2C", activebackground=bgColour,
                            activeforeground=textColour)  # https://stackoverflow.com/questions/66843817/tkinter-checkbox-click-animation-color-changes
        Hovertip(check, "Valid records are records that are not historical,\n withdrawn or intended for withdrawal")
        check.select()
        check.grid(row=0, column=0, sticky="E", pady=10)
        button1 = Button(checkGrid, text="Add Column", font=('Raleway', 12), bg=headerColour, fg=textColour, anchor=E, padx=5,
                         pady=5, command=lambda: self.addColumn())
        button1.grid(row=0, column=1, sticky="NSEW", padx=5, pady=10)
        button2 = Button(checkGrid, text="Remove Last Column", font=('Raleway', 12), bg=headerColour, fg=textColour, anchor=E,
                         padx=5, pady=5, command=lambda: self.removeColumn())
        button2.grid(row=0, column=2, sticky="NSEW", padx=5, pady=10)
        button3 = Button(checkGrid, text="Convert to SQL", font=('Raleway', 12), bg="green", fg=textColour,
                         activebackground="green", anchor=E, padx=5, pady=5, command=lambda: self.execute())
        button3.grid(row=0, column=3, sticky="NSEW", padx=200, pady=10)
        Hovertip(button3,
                 "Converts input from fields into a SQL query \nto paste in SQL Server Management Studio.\nResults are automatically copied to the clipboard.")
        self.label2 = Label(checkGrid, text="", font=('Raleway', 12, 'bold'), bg=bgColour, fg=textColour, anchor=W, padx=5, pady=5,
                       wraplength=300, justify="center")
        self.label2.grid(row=0, column=4, sticky="NSEW", padx=5, pady=10)

        checkGrid.pack(side=TOP, anchor=W, padx=10)

        # Creating the main grid that contains the Column fields---------------------------------------------------------------------------

        # creating the style for scroll bar as windows default overrides it normally

        # makes scroll bar a darker blue
        # style = ttk.Style()
        # style.element_create("My.Vertical.TScrollbar", "from", "default")
        # style.element_create("My.Vertical.Scrollbar.thumb", "from", "default")
        # style.layout("My.Vertical.TScrollbar",
        #     [('Vertical.Scrollbar.trough', {'children':
        #         [('Vertical.Scrollbar.uparrow', {'side': 'top', 'sticky': ''}),
        #          ('Vertical.Scrollbar.downarrow', {'side': 'bottom', 'sticky': ''}),
        #          ('My.Vertical.Scrollbar.thumb', {'expand': '1', 'sticky': 'ns'})],
        #     'sticky': 'ns'})])
        # style.configure("My.Vertical.TScrollbar", background="#131c36")

        #create canvas that will contain scrollbar and the main frame
        canvas = Canvas(root, bg=bgColour, highlightthickness=0)
        scrollbar = ttk.Scrollbar(orient="vertical", command=canvas.yview)#, style="My.Vertical.TScrollbar") # for the darker blue

        # When the content inside the scrollable frame changes size, adjust the canvas's scroll region
        scroll_frame = Frame(canvas)

        self.gridFrame = Frame(scroll_frame, bg=bgColour)
        self.gridFrame.pack(fill="both", expand=True, anchor="n")  # Ensure gridFrame fills the scroll_frame

        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

        self.gridFrame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        # Configure the canvas to scroll with the scrollbar
        canvas.configure(yscrollcommand=scrollbar.set)

        # Bind the mouse wheel event to scroll the canvas, regardless of where the cursor is
        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        # Pack the canvas and scrollbar into the frame
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.gridFrame.columnconfigure(0, weight=1)
        self.gridFrame.rowconfigure(0, weight=1)
        self.gridFrame.columnconfigure(1, weight=1)
        self.gridFrame.rowconfigure(1, weight=1)
        self.gridFrame.columnconfigure(2, weight=1)
        self.gridFrame.rowconfigure(2, weight=1)
        self.gridFrame.columnconfigure(3, weight=1)
        self.gridFrame.rowconfigure(3, weight=1)

        self.columnTitles.append(dropdown(self.gridFrame, 0, 1, 10, 10,
                                     ['Country Code', 'Title English', 'Descriptors English', 'Abstract English',
                                      'Title/Keywords', 'Free Text', 'Notes English', 'Document Identifier', 'AC Code', 'Publication Date'], "enabled", "").variable)
        text1 = Text(self.gridFrame, padx=10, pady=7, font=('Raleway', 10), height=1, bg=boxColour, wrap=NONE)
        self.textBoxes.append(text1)
        text1.grid(row=0, column=2, sticky="NSEW", padx=10, pady=10)
        text1.bind('<KeyRelease>', self.on_key_release)
        self.logic.append(dropdown(self.gridFrame, 0, 0, 10, 10, ["N/A"], "disabled", "").variable)

        # This command commits all the changes above in this section
        self.gridFrame.pack(side=TOP, anchor=N)

        #Snow!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        # List to store snowflake objects
        try:
            with open(CONFIG_FILE, "r") as file:
                data = file.read()

            if data == "1" or data == 1:
                snow = True
            else:
                snow = False
        except Exception as e:
            print(e)
            snow = True

        if snow:
            snowflakes = []

            # Function to create a snowflake
            def create_snowflake():
                # Get the current width of the window
                current_width = root.winfo_width()

                x = random.randint(0, current_width)  # Random horizontal position based on window width
                y = 0  # Start at the top
                size = random.randint(2, 6)  # Random snowflake size
                snowflake = canvas.create_oval(x, y, x + size, y + size, fill="white", outline="white")
                snowflakes.append(snowflake)
                root.after(100, create_snowflake)  # Create another snowflake after a short delay

            # Function to move snowflakes
            def move_snowflakes():
                for snowflake in snowflakes:
                    canvas.move(snowflake, 0, random.randint(1, 3))  # Move down by a random speed
                    x, y, x1, y1 = canvas.coords(snowflake)  # Get snowflake coordinates
                    if y > 1000:  # If the snowflake is out of the screen
                        canvas.delete(snowflake)  # Remove it from the canvas
                        snowflakes.remove(snowflake)  # Remove it from the list

                root.after(50, move_snowflakes)  # Update snowflakes every 50ms


            create_snowflake()
            move_snowflakes()

    '''FUNCTIONS---------------------------------------------------------------------------------------------------------'''
    def addColumn(self):

        # if self.colCount < 9: # this is the current row limit
        self.columnTitles.append(dropdown(self.gridFrame, self.colCount, 1, 10, 10,
                                     ['Country Code', 'Title English', 'Descriptors English', 'Abstract English',
                                  'Title/Keywords', 'Free Text', 'Notes English', 'Document Identifier', 'AC Code', 'Publication Date'], "enabled", '').variable)
        text1 = Text(self.gridFrame, padx=10, pady=7, font=('Raleway', 10), height=1, bg=boxColour, wrap=NONE)
        self.textBoxes.append(text1)
        text1.grid(row=self.colCount, column=2, sticky='NSEW', padx=10, pady=10)
        text1.bind('<KeyRelease>', self.on_key_release)
        self.logic.append(dropdown(self.gridFrame, self.colCount, 0, 10, 10, ['AND', 'OR', 'NOT'], "enabled", '').variable)

        self.colCount += 1
        return self.colCount
        # else:
        #     self.label2["fg"] = "#FF3F3F"
        #     self.label2["text"] = "You cannot add anymore rows"


    def removeColumn(self):

        if self.colCount > 1:
            self.colCount -= 1
            self.label2["text"] = ""

            for widget in self.gridFrame.grid_slaves():
                if int(widget.grid_info()[
                           "row"]) == self.colCount:  # https://stackoverflow.com/questions/23189610/how-to-remove-widgets-from-grid-in-tkinter
                    widget.grid_forget()

            del self.columnTitles[self.colCount]
            del self.textBoxes[self.colCount]
            del self.logic[self.colCount]

        return self.colCount

    def optionExtract(self, selection): # returns value of a widget
        return selection

    def removeQuotes(self, t, c):
        '''
        This removes any " or ' inputted by the user in t and then places it back in again to simplify inputting of text.
        For columns that use LIKE in SQL, no quotes are added. If the column uses CONTAINS, it will add double quotes

        :param t: the text from the text box
        :param c: the column get specified in the execute function
        :return c:
        '''

        result = ""
        for character in t:
            if character != '"' and character != "'" and character != "“" and character != "”":
                result += character

        if c in ["Title English", "Abstract English", "Descriptors English", "Notes English", "Title/Keywords", "Free Text"]:
            quote = True
        else:
            quote = False

        # (boat OR marine OR harbour) AND marine
        # (boat and harbour OR marine OR harbour) AND marine

        # within a text box locate all the ORs ANDs and AND NOTs

        if quote is True:

            #chat gpt's solution to locate " AND ", " OR ", " AND NOT " starting with AND NOT to avoid crossover --------------
            substrings = [" AND ", " OR ", " AND NOT "]
            substrings_sorted = sorted(substrings, key=len, reverse=True)

            all_locations = []

            for substring in substrings_sorted:
                start_index = result.find(substring)
                while start_index != -1:
                    end_index = start_index + len(substring) - 1
                    # Check if the start_index is already present in the locations of the longer substrings
                    overlapping = any(start_index in range(s, e + 1) for s, e in all_locations)
                    if not overlapping:
                        all_locations.append((start_index, end_index + 1))
                    start_index = result.find(substring, start_index + 1)

            # add quotes at start/end of connector words (AND, OR, AND NOT)
            locations = sorted([b for a in all_locations for b in a])
            locations = [num + i for i, num in enumerate(locations)]

            #------------------------------------------------------------------------------------------------------
            def insert_char_at_index(s, index, char='"'):
                if index > 0:

                    backup = 0
                    new_index = index
                    while new_index >= 0:
                        backup += 1
                        if s[index - backup] == ")":
                            new_index -= 1
                        else:
                            index = new_index
                            break

                if index < len(s):

                    forward = -1
                    new_index = index
                    while new_index <= len(s):
                        forward += 1
                        if s[index + forward] == "(":
                            new_index += 1
                        else:
                            index = new_index
                            break
                return s[:index] + char + s[index:]

            for i in locations:
                result = insert_char_at_index(result, i)

            print("result", result)
            for index, character in enumerate(result):
                if character != " " and character != "(" and character != ")":
                    result = result[:index] + '"' + result[index:]
                    break

            for i in range(len(result) - 1, -1, -1):
                if result[i] != " " and result[i] != "(" and result[i] != ")":
                    result = result[:i + 1] + '"' + result[i + 1:]
                    break

        return result

    def execute(self): # the main engine for converting input to SQL code

        self.helperLabel["text"] = f"TIP: Use capitalised AND, OR, AND NOT     "
        self.helperLabel["fg"] = textColour

        def extra_brackets(columnList: list, logicList: list) -> dict:
            '''
            This function loops through the columTitles list to check if there are the same columns next to one-another where
            extra brackets is needed. It returns a dictionary to state what brackets should be used.

            :param columnTitles: list, this is the input list of titles that the app gets from the input text boxes
            :return output: Dict
            '''

            output = {}

            bracket = ""
            open = False
            for index, (c, l) in enumerate(zip(columnList, logicList)):
                if index + 1 < len(columnList):  # to prevent calling an index outside of the list
                    if columnList[index + 1] == c and bracket == "":
                        if (logicList[index] in ["N/A", "AND", "OR"] and logicList[index + 1] in ["N/A", "AND",
                                                                                                  "OR"]) or \
                                (logicList[index] == "NOT" and logicList[index + 1] == "NOT"):
                            output[index] = "("
                            open = True
                        else:
                            output[index] = ""
                    elif columnList[index + 1] != c and open is True:
                        output[index] = ")"
                        open = False
                    else:
                        output[index] = ""
                elif open is True:
                    output[index] = ")"
                    open = False
                else:
                    output[index] = ""

            return output

        table = self.textbox1.get("1.0", END)
        bracketCheck = extra_brackets([c.get() for c in self.columnTitles], [l.get() for l in self.logic])

        if table != "\n":
            if self.columnSelect.variable.get() == "Standard Columns":
                rColumns = f'''

SELECT

ACCode
,DocumentIdentifier
,PublicationDate
,Title_English
,CommitteeReference
,Descriptors_English
,IssuingBody
,Update_Monthly
,CountryCode
,Abstract_English
,Status
,InternationalRelationship
,Classification
,CrossReferences

FROM {table.strip()}

WHERE \n'''
            else:
                rColumns = f'''
SELECT 

*

FROM {table.strip()}

WHERE \n'''

            result = rColumns

            if self.checkVariable.get() == 1:
                result = result + "(Update_Monthly NOT LIKE '%H%' AND Update_Monthly NOT LIKE '%I%' AND Update_Monthly NOT LIKE '%W%') AND \n"

            errors = []  # if the user has left any columns blank
            count = 0

            start = ""  # this is used to add an opening bracket for concatenating queries e.g. ((contains(Ti ... instead of (contains(Ti ...
            end = ""  # same as above but for end

            warning = {}
            for index, (c, t, l) in enumerate(zip(self.columnTitles, self.textBoxes, self.logic)):  # chatGPT told me about index, (c, t, l) with zip

                count += 1

                c = c.get()
                t = self.removeQuotes(t=t.get("1.0", END).strip().replace("\n", ""), c=c)
                l = l.get()


                for word in [" and ", " And ", " aNd ", " anD ", " AnD ", " ANd ", " aND ", " or ", " Or ", " oR ",
                             " not ", " Not ", " nOt ", " noT ",  " NoT ", " NOt ", " nOT "]:
                    if word in t:
                        if c in warning:
                            warning[c].append(word)
                        else:
                            warning[c] = [word]

                if l == "NOT":
                    l = "AND NOT"

                if t == "":
                    errors.append(str(c) + " ")

                connector = "OR"

                if l != "N/A":
                    result = result + f"\n\n{l}\n"

                if bracketCheck[index] == "(":
                    start = "("
                    end = ""
                elif bracketCheck[index] == ")":
                    start = ""
                    end = ")"
                else:
                    start = ""
                    end = ""

                if c == "Title/Keywords":
                    result = f'''{result}\n{start}(CONTAINS(Title_English, '{t}')\n{connector} CONTAINS(Descriptors_English, '{t}')){end}'''
                elif c == "Free Text":
                    result = f'''{result}\n{start}(CONTAINS(Title_English, '{t}')\n{connector} CONTAINS(Descriptors_English, '{t}')\n{connector} CONTAINS(Abstract_English, '{t}')\n{connector} CONTAINS(Notes_English, '{t}')){end}'''
                elif c in ["Title English", "Abstract English", "Descriptors English", "Notes English"]:
                    c = c.split(" ")
                    c = str(c[0]) + "_" + str(c[1])
                    result = f"{result}{start}CONTAINS({c}, '{t}'){end}\n"
                elif c == "Country Code" or c == "Document Identifier" or c == "AC Code" or c == "Publication Date":
                    c = c.split(" ")
                    c = str(c[0]) + str(c[1])
                    t = re.split(" OR | AND", t.strip().upper())
                    for count, code in enumerate(t):
                        if "(" in code or ")" in code:
                            self.label2["fg"] = "#FF3F3F"
                            self.label2["text"] = f"Brackets not supported in this column: {c}"
                            return
                        elif "*" in code:
                            self.label2["fg"] = "#FF3F3F"
                            self.label2["text"] = f"Wildcards * not supported in this column: {c}"
                            return
                        if count == 0:
                            if len(t) == 1:
                                result = result + f"({c} like '%{code.upper()}%')"
                            else:
                                result = result + f"({c} like '%{code.upper()}%' OR "
                        elif count != len(t) - 1:
                            result = result + f"{c} like '%{code.upper()}%' OR "
                        else:
                            result = result + f"{c} like '%{code.upper()}%')"
                else:
                    result = result + c + " in " + "(" + t + ") "


            result = result.strip()
            if result[len(result) - 3:] == "AND":
                result = result[:len(result) - 3]
            elif result[len(result) - 2:] == "OR":
                result = result[:len(result) - 2]
            elif result[len(result) - 7:] == "AND NOT":
                result = result[:len(result) - 7]

            if len(errors) > 0:
                self.label2["fg"] = "#FF3F3F"
                self.label2["text"] = f"Blanks in columns: {str(errors)[1:len(str(errors)) - 1]}"
                return
            for c, t in zip(self.columnTitles, self.textBoxes):
                c = c.get()
                t = t.get("1.0", END).strip()

                if '”' in str(t):

                    self.label2["fg"] = "#FF3F3F"
                    self.label2["text"] = f'Incorrect double quote ” in column {c}. Make sure you are using ".'
                    return

            for key, value in warning.items():
                warning[key] = list(set(value))

            self.label2["text"] = "Query copied to clipboard"
            self.label2["fg"] = "#ffffff"
            if len(warning) > 0:
                self.helperLabel["text"] = f"CAUTION - Non-capitalised logic may be in query: {warning}\n\nThis might make your query unstable. See help doc."
                self.helperLabel["fg"] = "#FFFF00"

            print(result)

            pyperclip.copy(result)
        else:
            self.label2["text"] = "'SQL Table' has been left blank"
            self.label2["fg"] = "#FF3F3F"

    def openDocumentation(self):
        os.system('start Documentation.docx')

    def on_key_release(self, event):
        '''
        This is the function responsible for autosizing textboxes when the amount of text exceeds the length of the box
        :return:
        '''
        # Get the widget that triggered the event
        widget = event.widget

        # Get the longest line in the Text widget
        lines = widget.get("1.0", "end-1c").split("\n")
        longest_line = max(lines, key=len)

        # If the length of the longest line exceeds the width of the Text widget
        if len(longest_line) > widget["width"]:
            # Find the nearest space character before the width limit
            break_position = longest_line.rfind(' ', 0, widget["width"])

            # If a space character is found, insert a newline there, otherwise just break at the width limit
            if break_position != -1:
                widget.insert(f"{lines.index(longest_line) + 1}.{break_position}", "\n")
            else:
                widget.insert(f"{lines.index(longest_line) + 1}.{widget['width']}", "\n")

        # Adjust the height based on the number of lines
        num_lines = int(widget.index('end-1c').split('.')[0])
        widget['height'] = num_lines

        # Reset the modified flag for the widget
        widget.edit_modified(False)


'''Tkinter Object Class DECLARATIONS------------------------------------------------------------------------------------------------'''
class dropdown:
    def __init__(self, frame, row, column, padx, pady, options, state, tooltip):
        self.frame = frame
        self.row = row
        self.column = column
        self.padx = padx
        self.pady = pady
        self.options = options
        self.tooltip = tooltip
        self.state = state

        self.variable = StringVar(self.frame)
        self.variable.set(self.options[0])
        dropdown = OptionMenu(self.frame, self.variable, *self.options)
        if self.state == "disabled":
            dropdown.configure(state="disabled")
        dropdown.config(bg=headerColour, fg=textColour, activebackground="#2A9AF2", font=('Raleway', 12),
                        highlightthickness=0)  # highlight thickness removes white line
        dropdown.grid(row=self.row, column=self.column, sticky="NSEW", padx=self.padx, pady=self.pady)

        if self.tooltip != "":
            Hovertip(dropdown, self.tooltip)


'''Setting up app and app attributes---------------------------------------------------------------------------------'''
try:
    r = requests.get(VERSION_URL, verify=False)
    soup = BeautifulSoup(r.content, "lxml")
    external_version = soup.find("h5").text

    checkVersion = convertVersion(external_version)
except Exception as e:
    message = messagebox.showerror(title="Query Generator",
                                  message=f"Unable to retrieve latest version, Query Generator will now exit. {e}")
    sys.exit()
if checkVersion > convertVersion(VERSION):

    message = messagebox.showinfo(title="Query Generator", message="A new version is available. The app will now try to update automatically.")

    try:
        def download_update(download_url):
            response = requests.get(download_url, stream=True, verify=False)

            # Folder to save the download
            user_profile = os.getenv('LOCALAPPDATA')
            save_folder = os.path.join(user_profile, "Query_Generator")

            # Ensure folder exists
            if not os.path.exists(save_folder):
                os.makedirs(save_folder)

            # Path for the downloaded ZIP file
            zip_file_path = os.path.join(save_folder, "Query_Generator.zip")

            # Save the ZIP file
            with open(zip_file_path, 'wb') as file:
                for chunk in response.iter_content(chunk_size=8192):
                    file.write(chunk)

            # Extract the ZIP file
            with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                zip_ref.extractall(save_folder)

            os.remove(zip_file_path)

            print(f"Update downloaded and extracted to: {save_folder}")


        def update_and_restart(current_exe_path, exe_dir):
            # Restart the application
            subprocess.Popen([current_exe_path], cwd=exe_dir)
            sys.exit()

        download_update(REPO)

        user_profile = os.getenv('LOCALAPPDATA')
        exe_path = os.path.join(user_profile, "Query_Generator/Query_Generator/Query_Generator.exe")
        exe_dir = os.path.join(user_profile, "Query_Generator/Query_Generator/")
        update_and_restart(current_exe_path=exe_path, exe_dir=exe_dir)

    except Exception as e:
        message = messagebox.showerror(title="Query Generator",
                                       message=f"Unable to retrieve latest version, Query Generator will now exit. Error: {e}")

else:
    if __name__ == '__main__':
        root = Tk()
        app = QueryGenerator(root)
        root.mainloop()