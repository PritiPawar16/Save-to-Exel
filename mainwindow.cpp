#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "xlsxdocument.h"
#include "xlsxformat.h"
#include <QFileDialog>
#include <QMessageBox>
#include <QImage>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_addDataButton_clicked()
{
    // Open a file dialog to choose where to save the Excel file
    QString saveFilePath = QFileDialog::getSaveFileName(this, "Save Excel File", QDir::homePath(), "Excel Files (*.xlsx)");

    if (saveFilePath.isEmpty()) {
        QMessageBox::warning(this, "No File Selected", "Please select a file to save the data.");
        return;
    }

    // Create an Excel document
    QXlsx::Document xlsx;

    // Add some sample data
    xlsx.write("A1", "Header 1");
    xlsx.write("B1", "Header 2");
    xlsx.write("A2", "Data 1");
    xlsx.write("B2", "Data 2");

    // Save the Excel file at the user-specified location
    if (xlsx.saveAs(saveFilePath)) {
        QMessageBox::information(this, "Success", "Data added and Excel file saved successfully.");
    } else {
        QMessageBox::critical(this, "Error", "Failed to save the Excel file.");
    }
}

void MainWindow::on_addImageButton_clicked()
{
    // Open a file dialog to select an image
    QString imagePath = QFileDialog::getOpenFileName(this, "Select Image", "", "Images (*.png *.xpm *.jpg)");

    if (imagePath.isEmpty()) {
        QMessageBox::warning(this, "No Image", "No image was selected.");
        return;
    }

    // Open a file dialog to choose where to save the Excel file
    QString saveFilePath = QFileDialog::getSaveFileName(this, "Save Excel File", QDir::homePath(), "Excel Files (*.xlsx)");

    if (saveFilePath.isEmpty()) {
        QMessageBox::warning(this, "No File Selected", "Please select a file to save the data.");
        return;
    }

    // Load the Excel document if it exists; otherwise, create a new one
    QXlsx::Document xlsx;

    // Add some sample data (Optional, based on your requirement)
    xlsx.write("A1", "Header 1");
    xlsx.write("B1", "Header 2");

    // Load the image and insert it into the Excel document
    QImage image(imagePath);
    if (!image.isNull()) {
        // Insert the image at cell C2
        xlsx.insertImage(1, 2, image); // Coordinates are zero-based; (1, 2) is C2 in Excel
    } else {
        QMessageBox::critical(this, "Error", "Failed to load the image.");
        return;
    }

    // Save the updated Excel file at the specified location
    if (xlsx.saveAs(saveFilePath)) {
        QMessageBox::information(this, "Success", "Image added and Excel file saved successfully.");
    } else {
        QMessageBox::critical(this, "Error", "Failed to save the Excel file with the image.");
    }
}
