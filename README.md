# Cover Letter Lazy Tool

The Cover Letter Lazy Tool is a Windows Forms application written in C# that assists in generating cover letters by automating the process of creating a Word document, adding paragraphs of content, substituting placeholders, and exporting the final document.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Getting Started](#getting-started)
- [Usage](#usage)
- [Functionality](#functionality)
- [Notes](#notes)
- [License](#license)

## Prerequisites

Before using the Cover Letter Lazy Tool, ensure you have the following:

- Visual Studio (2017 or later) installed.
- Knowledge of C# programming.
- Microsoft Office Word installed on your machine.
- Microsoft Office Interop libraries installed. You can install them using NuGet Package Manager by searching for "Microsoft.Office.Interop.Word."

## Getting Started

1. Clone or download this repository to your local machine.
2. Open the Visual Studio solution file (`CoverLetterLazyTool.sln`).
3. Build the solution to restore NuGet packages and compile the application.
4. Run the application by clicking the "Start" button in Visual Studio.

## Usage

1. Launch the application to see the Windows Forms interface.
2. Click the "Create Doc" button to initialize a new Word document.
3. Input paragraphs of content in the provided textbox.
4. Click the "Insert Letter" button to insert the content into the document.
5. Enter the company name and occupation in the respective text boxes.
6. Click the "Substitute Text" button to replace placeholders with the entered information.
7. Use the "Save File" button to save the final cover letter as a Word file in a location of your choice.

## Functionality

- **Create Doc:** Click this button to create a new Word document.
- **Insert Letter:** Input paragraph content in the textbox and click this button to add it to the document.
- **Substitute Text:** Enter the company name and occupation, then click this button to replace placeholders with the provided information.
- **Save File:** Click this button to save the modified Word document to a user-defined location.

## Notes

- This application uses the Microsoft Office Interop libraries for Word document manipulation. Compatibility and behavior may vary depending on your environment and Microsoft Office version.
- Ensure you have Microsoft Office Word installed on your machine for the Interop libraries to work properly.

## License

This project is licensed under the [MIT License](LICENSE).
