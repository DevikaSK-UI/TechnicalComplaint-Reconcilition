# TechnicalComplint-Reconciliation
Overview

This repository provides a Python script for comparing CCC and EDC Excel sheets. The tool includes a graphical user interface (GUI) built with Tkinter, which facilitates easy file selection and result display. The comparison results are saved in an Excel file with color-coded discrepancies to highlight mismatches and missing records.It includes a .desktop file to create a launcher for the script and an icon image for the launcher. This setup allows you to easily run the Python script from your Linux desktop environment.

Contents :-

* Technical_Complaint_Reconciliation.py: The Python script that performs the comparison of Excel sheets.
* Technical_Complaint_Reconciliation.desktop: A desktop entry file to create a launcher for the Python script.
* TCC_Icon.jpeg: The icon image used for the launcher.

Setup Instructions

    Clone the Repository

Start by cloning this repository to your local machine:

git clone https://github.com/Anjali-Kumari-Mina/Technical_Complaint_Reconciliation.git

Navigate to the cloned directory:

cd Technical_Complaint_Reconciliation

    Prepare the Files

Download the Python Script and Icon:

Place Technical_Complaint_Reconciliation.py in a directory of your choice.
Save the TCC_Icon.jpeg icon file in the same directory or update the path in the .desktop file accordingly.

    Modify the .desktop File

    Download the Technical_Complaint_Reconciliation.desktop file from this repository.

    Open the .desktop file in a text editor.

    Update the Exec and Icon paths to match the locations of your Python script and icon. For example:

    [Desktop Entry] Name=Technical Complaint Reconciliation Comment=Compare CCC and EDC Excel sheets Exec=python3 "/path/to/your/Technical_Complaint_Reconciliation.py" Icon=/path/to/your/TCC_Icon.jpeg Terminal=false Type=Application Categories=Utility;

    Set Permissions

Make sure the .desktop file is executable:

chmod +x /path/to/your/Technical_Complaint_Reconciliation.desktop

    Move the .desktop File to Your Desktop

To create a shortcut directly on your desktop:

cp /path/to/your/Technical_Complaint_Reconciliation.desktop ~/Desktop/

If the .desktop file does not appear immediately or you encounter any issues, ensure that your desktop environment is configured to display executable .desktop files.
Usage

Running the Python Script

a)You can run the Python script directly from the terminal:

python3 /path/to/your/Technical_Complaint_Reconciliation.py

b) Using the Desktop Shortcut

You should now see a launcher on your desktop. Double-click the launcher to run the Python script with the specified icon.
