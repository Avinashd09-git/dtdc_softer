DTDC Task Scheduler

Project Overview
This project is a Flask web application designed to facilitate the upload, merging, processing, and storage of Excel files into a PostgreSQL database. The application provides multiple operations for handling different types of Excel data, including TS (Transhipment), Software TS, International TS, and Booking data. Additionally, users can download processed data from the database.

Features
- Upload Excel files** (multiple file selection supported).
- Merge Excel files** from a folder.
- Process and insert data** into a PostgreSQL database.
- Download processed data** from the database.
- User-friendly web interface** built with Bootstrap.

Technologies Used
- Backend: Flask (Python)
- Frontend: HTML, CSS, Bootstrap, JavaScript
- Database: PostgreSQL (via SQLAlchemy)
- File Processing: Pandas

---
Project Structure

1. Backend (Flask Application)
app.py:
- The main Flask application.
- Implements routes for different operations.
- Handles file uploads, merging, and database insertion.
- Uses Pandas for data manipulation.
- Uses SQLAlchemy to interact with PostgreSQL.
- Implements file download functionality.

Key Routes:
- `/` → Homepage with operation selection.
- `/operation1` → Handles **TS operations** (Upload & process TS Excel files).
- `/operation2` → Handles **Software TS** (Upload, merge, and insert into PostgreSQL).
- `/operation3` → Handles **International TS** (Processes 3rd sheet of Excel files and inserts into DB).
- `/operation4` → Handles **Booking data** (Merges and processes booking-related Excel files).
- `/download_table/<table_name>` → Exports database table data to an Excel file.

--

2. Frontend (HTML Files for UI)
`index.html`
- Homepage for selecting an operation.
- Buttons to navigate to TS, Software TS, International TS, and Booking operations.

`Ts.html`
- UI for TS operation file uploads.
- Allows users to upload Excel files and insert them into the database.

`Software Ts.html`
- UI for Software TS operation.
- Users can upload Excel folders, merge them, and store data in PostgreSQL.

`International TS.html`
- UI for International TS operation.
- Merges Excel files, extracts the 3rd sheet, processes data, and inserts it into PostgreSQL.

`Booking.html`
- UI for Booking data processing.
- Users upload an Excel folder, merge files, and process data for downloading.

`operation.html`
- Generic operation page for unimplemented features.

---

3. Static Files
- `logo.png` → Logo displayed in the web interface.
- Bootstrap CSS & JS → Enhances UI styling and responsiveness.
- FontAwesome Icons → Adds icons for a visually appealing experience.

---

How It Works
1. User uploads a folder containing Excel files.
2. Application processes and merges the files.
3. Data is inserted into the PostgreSQL database.
4. User can download the processed data as an Excel file.

---

Future Enhancements
- Implement authentication for secure access.
- Improve error handling for unsupported file formats.
- Add a progress indicator for file processing.
- Enable visualization of processed data via charts (Tableau or Power BI integration).

---

Conclusion
This project streamlines Excel data processing and database management, making it easy for users to upload, merge, store, and retrieve structured data. The combination of Flask, PostgreSQL, and Pandas ensures efficient and scalable data handling. 🚀


