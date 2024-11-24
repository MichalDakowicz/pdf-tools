# PDF Tools

A Python application built with Tkinter that provides a suite of tools for manipulating PDF files. This includes merging, splitting, rotating, extracting text, converting to and from images, and encrypting/decrypting PDFs.

## Features

-   **Merge PDFs:** Combine multiple PDF files into a single document.
-   **Split PDF:** Split a PDF file into multiple parts by specifying page ranges (e.g., 1-3,5,7-9).
-   **Rotate PDF:** Rotate pages of a PDF by 90, 180, or 270 degrees. Handles page dimension adjustments for 90 and 270-degree rotations.
-   **Extract Text:** Extract text content from a PDF and save it as a .txt or .csv file.
-   **Convert PDF to Images:** Convert each page of a PDF to individual image files (PNG format).
-   **Convert Images to PDF:** Combine multiple image files into a single PDF document.
-   **Decrypt PDF:** Decrypt a password-protected PDF file.
-   **Encrypt PDF:** Encrypt a PDF file with a password.
-   **Drag and Drop Reordering:** Reorder files in the merge and image-to-PDF lists using drag and drop.

## Requirements

-   Python 3.6 or higher
-   Tkinter (usually included with Python)
-   PyPDF2
-   Pillow (PIL)
-   pdf2image

## Installation

1. **Install required libraries:**

    ```bash
    pip install PyPDF2 Pillow pdf2image
    ```

2. **Download the script:** Download the provided Python code or clone the repository to your local machine.

## Usage

1. **Run the script:**
    ```bash
    python main.py
    ```
2. **Navigate the tabs:** Use the tabs to access different PDF manipulation functions.
3. **Follow the on-screen instructions:** Each tab provides clear instructions for using the specific tool. Browse for files, enter required information (like page ranges or passwords), and click the appropriate buttons.

## File Organization

The application creates a "pdf" folder in the same directory as the script if it doesn't already exist. This folder is used as the default location for opening and saving files. Images converted from PDFs are saved to a directory of your choice (selected using a file dialog).

## Examples

-   **Merging PDFs:** Select the "Merge PDFs" tab. Click "Add Files" and choose multiple PDF files. You can reorder files using drag and drop. Click "Merge", and choose a location and name for the merged output PDF.

-   **Splitting PDF:** Select the "Split PDF" tab. Browse and select `report.pdf`. Enter page ranges like `1-5,10,12-15` in the "Page Ranges" field. Click "Split". This will create files like `report_part1.pdf` (containing pages 1-5), `report_part2.pdf` (containing page 10), and `report_part3.pdf` (containing pages 12-15).

-   **Rotating PDF:** Select the "Rotate PDF" tab. Browse and select `document.pdf`. Click the "90" button to set the rotation angle. Click "Rotate" and choose a save location.

-   **Extracting Text:** Select the "Extract Text" tab. Browse and select `article.pdf`. Choose the desired output format (TXT or CSV). Click "Extract" and choose a save location.

-   **Converting PDF to Images:** Select the "Convert PDF to Images" tab. Browse and select `presentation.pdf`. Click "Convert". Choose a directory where the images (PNG format) will be saved.

-   **Converting Images to PDF:** Select the "Convert Images to PDF" tab. Click "Add Images" and select multiple image files. You can reorder them using drag and drop. Click "Convert", and choose a location and name for the output PDF.

-   **Decrypting PDF:** Select the "Decrypt PDF" tab. Browse and select `protected.pdf`. Enter the password in the "Enter Password" field. Click "Decrypt" and choose a save location for the decrypted PDF.

-   **Encrypting PDF:** Select the "Encrypt PDF" tab. Browse and select `confidential.pdf`. Enter the desired password in the "Enter Password" field. Click "Encrypt" and choose a save location for the encrypted PDF.

## File Organization

The application creates a "pdf" folder in the same directory as the script if it doesn't already exist. This folder is used as the default location for opening and saving files. Images converted from PDFs are saved to a directory of your choice (selected using a file dialog).

## Splitting PDF Example (More Detailed)

To split a PDF file named `mydocument.pdf`:

1. Go to the "Split PDF" tab.
2. Click the "Browse" button and select `mydocument.pdf`.
3. In the "Page Ranges" entry, specify the desired ranges. For example:
    - `1-3`: Pages 1, 2, and 3 will be extracted to a new PDF.
    - `5`: Page 5 will be extracted.
    - `7-9`: Pages 7, 8, and 9 will be extracted.
    - `1-3,5,7-9`: Multiple ranges can be specified, separated by commas.
4. Click "Split". The new PDF files will be created in the same directory as the original PDF, named like `mydocument_part1.pdf`, `mydocument_part2.pdf`, etc.
5. _(Optional)_ Merge the split files back together using the "Merge PDFs" tab.

## Error Handling

The application includes error handling to catch common issues like:

-   Incorrect passwords for decryption/encryption.
-   Invalid page ranges when splitting.
-   File not found errors.
-   Issues during PDF processing.

Error messages will be displayed in pop-up dialog boxes.

## Contributing

Contributions are welcome! Feel free to open issues or submit pull requests to improve this project.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.
