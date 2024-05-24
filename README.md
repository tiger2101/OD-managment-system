This project is a desktop application built using Python and Tkinter for managing on-duty leaves. The primary goal is to facilitate the import and management of student data from an Excel sheet. The application reads student names and registration numbers from the sheet, displays them in a GUI, and allows for selecting specific entries with checkboxes. This project was specifically built for colleges. With the help of this tool we can import student data from the excel sheet and select the students and write the reason why they are taking OD and press the button submit. Once submitted, it will generate a new excel sheet with the input given in the tool.

### Project Overview

1. **Setup and Initialization**: 
   - **Python and Tkinter**: The application is developed using Python with Tkinter for the GUI. The `openpyxl` library is used for handling Excel files.
   - **Project Structure**: The main script initializes the Tkinter window, configures the layout, and sets up necessary widgets like buttons, labels, and a tree view.

2. **User Interface**:
   - **File Upload**: The interface starts with a file upload button, allowing the user to select an Excel file containing student data. Upon the first upload, the file path is saved to a text file to bypass the need for re-uploading on subsequent application launches.
   - **Tree View Display**: Once the Excel file is loaded, the data (registration numbers and student names) is displayed in a tree view. Each row in the tree view has an associated checkbox for selection.
   - **Scrollbars**: Scrollbars are implemented to handle large datasets, ensuring the tree view remains navigable regardless of the amount of data.

3. **Functionality**:
   - **Data Import**: The application reads data from the selected Excel file using `openpyxl`, ignoring the header row, and populates the tree view with the data.
   - **Selection and Logging**: Users can select multiple rows using checkboxes. The selected data is logged to the console as a dictionary, with registration numbers as keys and student names as values.
   - **Persistent File Path**: The application saves the path of the last uploaded Excel file in a text file. On subsequent launches, it automatically loads data from this file, streamlining the user experience.
