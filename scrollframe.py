import customtkinter

class ScrollableCheckBoxFrame(customtkinter.CTkScrollableFrame):
    def __init__(self, master, item_list, command=None, **kwargs):
        super().__init__(master, **kwargs)

        self.command = command
        self.checkbox_list = []
        self.select_allc = None
        
        for i, item in enumerate(item_list):
            self.add_item(item,all=self.select_allc)
        

    def add_item(self, item, all, **kwargs):
        if all:
            checkbox = customtkinter.CTkCheckBox(
                self, text=item)
            checkbox.select()
            if self.command is not None:
                checkbox.configure(command=self.command)
            checkbox.grid(row=len(self.checkbox_list), column=0, pady=(0, 10), sticky='nsew')
            self.checkbox_list.append(checkbox)
        else:
            checkbox = customtkinter.CTkCheckBox(
                self, text=item)
            if self.command is not None:
                checkbox.configure(command=self.command)
            checkbox.grid(row=len(self.checkbox_list),
                          column=0, pady=(0, 10), sticky='nsew')
            self.checkbox_list.append(checkbox)

    def remove_item(self, item):
        for checkbox in self.checkbox_list:
            if item == checkbox.cget("text"):
                checkbox.destroy()
                self.checkbox_list.remove(checkbox)
                return

    def get_checked_items(self):
        return [checkbox.cget("text") for checkbox in self.checkbox_list if checkbox.get() == 1]
    
    def get_unchecked_items(self,item):
        checkbox = customtkinter.CTkCheckBox(self, text=item)
        checkbox.select()
        if self.command is not None:
            checkbox.configure(command=self.command)
        checkbox.grid(row=len(self.checkbox_list),
                      column=0, pady=(0, 10), sticky='nsew')
        self.checkbox_list.append(checkbox)
    

    

class ScrollableCheckBoxFrame1(customtkinter.CTkScrollableFrame):
    def __init__(self, master,  command=None, **kwargs):
        super().__init__(master, **kwargs)

        self.command = command
        self.list_items = []
        self.checkbox_list = []

    def add_item(self, item):
        checkbox = customtkinter.CTkCheckBox(self, text=item, onvalue=1)
        checkbox.select()
        if self.command is not None:
            checkbox.configure(command=self.command)
        checkbox.grid(row=len(self.checkbox_list),
                      column=0, pady=(0, 10), sticky='nsew')
        self.checkbox_list.append(checkbox)

    def remove_item(self, item):
        for checkbox in self.checkbox_list:
            if item == checkbox.cget("text"):
                checkbox.destroy()
                self.checkbox_list.remove(checkbox)
                return

    def get_checked_items(self):
        return [checkbox.cget("text") for checkbox in self.checkbox_list if checkbox.get() == 1]
