from tkinter.messagebox import showerror
from tkinter import *
from Mail_Distro_Settings import local_config, icon_path2


class MDG(Tk):
    __mail_list = None
    __mod_button = None
    __del_button = None
    __distro_configs = dict()

    def __init__(self):
        Tk.__init__(self)
        self.iconbitmap(icon_path2)
        self.__email_distro = StringVar()
        self.__sql_tbl = StringVar()
        self.__build()
        self.__load_gui()

    def __build(self):
        header_txt = ['Welcome to Mail Distro Settings!', 'Feel free to add, modify, delete a setting']
        # Set GUI Geometry and GUI Title
        self.geometry('420x285+630+290')
        self.title('Mail Distro Setup')
        self.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self)
        main_frame = LabelFrame(self, text='Mail List', width=444, height=140)
        list_frame = Frame(main_frame)
        input_frame = Frame(main_frame)
        control_frame = Frame(main_frame)
        button_frame = Frame(self)

        # Apply Frames into GUI
        header_frame.pack(fill="both")
        main_frame.pack(fill="both")
        list_frame.grid(row=0, column=0, rowspan=2)
        input_frame.grid(row=0, column=1, padx=5)
        control_frame.grid(row=1, column=1, padx=5)
        button_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self, text='\n'.join(header_txt), width=375, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply List Widget to list_frame
        #     Mail List
        xbar = Scrollbar(list_frame, orient='horizontal')
        ybar = Scrollbar(list_frame, orient='vertical')
        self.__mail_list = Listbox(list_frame, selectmode=SINGLE, width=30, yscrollcommand=ybar,
                                   xscrollcommand=xbar)
        xbar.config(command=self.__mail_list.xview)
        ybar.config(command=self.__mail_list.yview)
        self.__mail_list.grid(row=0, column=0, padx=8, pady=5)
        xbar.grid(row=1, column=0, sticky=W + E)
        ybar.grid(row=0, column=1, sticky=N + S)
        self.__mail_list.bind("<Down>", self.__list_action)
        self.__mail_list.bind("<Up>", self.__list_action)
        self.__mail_list.bind('<<ListboxSelect>>', self.__list_action)

        # Apply Entry Widgets to input_frame
        # Email Distro Entry Box
        label = Label(input_frame, text='Email Distro:')
        entry = Entry(input_frame, textvariable=self.__email_distro, width=15)
        label.grid(row=0, column=0, padx=2, pady=5, sticky='e')
        entry.grid(row=0, column=1, padx=7, pady=5, sticky='w')

        # Email SQL TBL Box
        label = Label(input_frame, text='SQL TBL:')
        entry = Entry(input_frame, textvariable=self.__sql_tbl, width=15)
        label.grid(row=1, column=0, padx=2, pady=5, sticky='e')
        entry.grid(row=1, column=1, padx=7, pady=5, sticky='w')

        # Apply Buttons to control_frame
        # Add, Delete, Modify buttons
        add_button = Button(control_frame, text='Add', width=10)
        self.__del_button = Button(control_frame, text='Del', width=10)
        self.__mod_button = Button(control_frame, text='Modify', width=23)
        add_button.grid(row=0, column=0, sticky='w', padx=3)
        self.__del_button.grid(row=0, column=1, sticky='w', padx=6)
        self.__mod_button.grid(row=1, column=0, columnspan=2, padx=3, pady=5, sticky='w')
        add_button.bind("<ButtonRelease-1>", self.__button_action)
        self.__mod_button.bind("<1>", self.__button_action)
        self.__del_button.bind("<ButtonRelease-1>", self.__button_action)

        # Apply button Widgets to button_frame
        #     Save Settings Button
        button = Button(button_frame, text='Save Settings', width=20, command=self.__save_settings)
        button.pack(side=LEFT, padx=7, pady=5)

        #     Exit GUI Button
        button = Button(button_frame, text='Exit GUI', width=20, command=self.destroy)
        button.pack(side=RIGHT, padx=7, pady=5)

    def __load_gui(self):
        if local_config['Distro_Configs']:
            self.__distro_configs = local_config['Distro_Configs']

            for distro in self.__distro_configs.keys():
                self.__mail_list.insert('end', distro)

            self.after_idle(self.__mail_list.select_set, 0)

    def __list_action(self, event):
        widget = event.widget

        if widget.size() > 0:
            selections = widget.curselection()

            if selections and hasattr(event, 'keysym') and (event.keysym == 'Up' and selections[0] > 0):
                self.after_idle(widget.select_clear, selections[0])
                self.after_idle(widget.select_set, selections[0] - 1)
            elif selections and hasattr(event, 'keysym') and (event.keysym == 'Down'
                                                              and selections[0] < widget.size() - 1):
                self.after_idle(widget.select_clear, selections[0])
                self.after_idle(widget.select_set, selections[0] + 1)
            elif not selections and hasattr(event, 'keysym') and event.keysym in ('Up', 'Down'):
                self.after_idle(widget.select_set, 0)

            self.after_idle(self.__button_state)

    def __button_state(self):
        button_state = DISABLED

        if self.__mail_list.size() > 0 and self.__mail_list.curselection():
            button_state = NORMAL

        self.__mod_button.configure(state=button_state)
        self.__del_button.configure(state=button_state)

    def __button_action(self, event):
        selections = self.__mail_list.curselection()
        widget = event.widget

        if widget.cget('text') == 'Add':
            self.__add_distro()
        elif widget.cget('text') == 'Modify' and selections:
            self.__modify_distro(selections[0])
        elif widget.cget('text') == 'Del' and selections:
            self.__delete_distro(selections[0])

    def __add_distro(self):
        if not self.__email_distro.get():
            showerror('Field Empty Error!', 'No value has been inputed for Email Distro', parent=self)
        elif not self.__sql_tbl.get():
            showerror('Field Empty Error!', 'No value has been inputed for SQL TBL', parent=self)
        elif self.__distro_configs and self.__email_distro.get() in self.__distro_configs.keys():
            showerror('Distro Exists!', 'Distro already exists!', parent=self)
        else:
            self.__distro_configs[self.__email_distro.get()] = self.__sql_tbl.get()
            self.__mail_list.insert('end', self.__email_distro.get())
            self.after_idle(self.__mail_list.select_set, 'end')
            self.__email_distro.set('')
            self.__sql_tbl.set('')

    def __modify_distro(self, selection):
        task_name = self.__mail_list.get(selection)

        if self.__distro_configs and task_name in self.__distro_configs.keys():
            self.__email_distro.set(task_name)
            self.__sql_tbl.set(self.__distro_configs[task_name])
            del self.__distro_configs[task_name]

        self.__mail_list.delete(selection)

    def __delete_distro(self, selection):
        task_name = self.__mail_list.get(selection)

        if self.__distro_configs and task_name in self.__distro_configs.keys():
            del self.__distro_configs[task_name]

        self.__mail_list.delete(selection)

    def __save_settings(self):
        if self.__distro_configs:
            local_config['Distro_Configs'] = self.__distro_configs
            local_config.sync()
        elif local_config['Distro_Configs']:
            del local_config['Distro_Configs']
            local_config.sync()

        self.destroy()


class MailDistroGUI(MDG):
    def __init__(self):
        MDG.__init__(self)
        self.mainloop()


if __name__ == '__main__':
    MailDistroGUI()
