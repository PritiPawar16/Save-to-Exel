Excel Data Saver with QXlsx

A simple and efficient Qt-based application for saving data to Excel files using the QXlsx library. This project allows users to manage data and export it directly into .xlsx format, offering a 
seamless way to work with Excel files in a Qt environment.


📚 Table of Contents

Features

Prerequisites

Installation

Configuration

Usage

Project Structure

Contributing

License

Contact

🚀 Features

Save Data to Excel: Easily save data to .xlsx files using the QXlsx library.

Table Integration: Automatically export data from Qt tables or custom data models into Excel format.

Cell Customization: Support for customizing Excel cells (font, color, borders, etc.).

User-Friendly Interface: Provides an intuitive UI for managing and saving data.

Cross-Platform Support: Compatible with Windows, macOS, and Linux environments.

🛠 Prerequisites

Before you begin, ensure you have the following installed:


Qt Framework: Version 5.12 or higher

QXlsx Library: Installed and configured with the project

Qt Creator: Integrated Development Environment for Qt

C++ Compiler: Compatible with your Qt version (e.g., GCC, MSVC)

CMake (Optional): For building the project, if not using Qt Creator's internal build system

📦 Installation

1. Clone the Repository
2. 
Clone the repository from GitHub:


bash

Copy code

git clone https://github.com/PritiPawar16/Save-to-Exel.git

2. Install QXlsx Library
3. 
Option 1: Install QXlsx Manually

Download QXlsx from its official repository: QXlsx GitHub.

Extract the files into your project’s directory under a folder named QXlsx.

Option 2: Use Qt’s Package Manager (qpm)

Install the qpm package manager (if not installed):

bash

Copy code

pip install qpm

Install the QXlsx package via qpm:

bash

Copy code

qpm install com.github.qtxlsx

5. Add QXlsx to Project
6. 
Update your .pro file to include QXlsx:


pro

Copy code

# Add QXlsx headers and sources to the project

include(QXlsx/QXlsx.pri)

4. Open the Project in Qt Creator
5. 
Launch Qt Creator.

Click on File > Open File or Project....

Select the ExcelSaver.pro file from the cloned repository.

7. Build the Project
8. 
Select the correct Qt Kit.

Build the project by clicking the Build button or pressing Ctrl + B.

⚙️ Configuration

Configuring Excel Export Options

The application allows customization of Excel sheets, such as setting specific font sizes, adding borders, and merging cells.

You can modify these settings in the exportToExcel() function located in the MainWindow.cpp file, or create your own configurations as needed.
