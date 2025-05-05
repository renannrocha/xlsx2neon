<img width="100%" src="banner.png" alt="banner">

# xlsx2neon
This project is a Python application designed to read Excel files, extract table data, and send it to a NeonDB database. It provides a graphical user interface (GUI) using `tkinter` to select Excel files, define table configurations, and process the data for insertion into the database.

---

## Features

* **Select Excel File:** Choose an Excel file to load.
* **Table Configuration:** Dynamically specify the number of tables, their corresponding sheet names, header rows, and table names in the database.
* **Data Processing:** Read the Excel data, convert it to a DataFrame, and insert it into NeonDB.
* **Scroll Support:** If there are too many tables to configure, a scrollbar is available to make navigating the fields easier.
* **Error Handling:** Provides error messages for missing or incorrect data.
* **GUI (Graphical User Interface):** Built using `tkinter`, making it easy for non-technical users to interact with.

---

## Prerequisites

* Python 3.x
* NeonDB account and database URL (you need to set this up in your `.env` file).

---

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/yourusername/xlsx2neon.git
   cd excel-to-neondb
   ```

2. Create a virtual environment:

   ```bash
   python -m venv venv
   ```

3. Activate the virtual environment:

   * On Windows:

     ```bash
     .\venv\Scripts\activate
     ```
   * On macOS/Linux:

     ```bash
     source venv/bin/activate
     ```

4. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

5. Create a `.env` file in the root directory and add your database URL:

   ```txt
   DATABASE_URL=your_neondb_database_url
   ```

---

## Usage

1. Run the application:

   ```bash
   python app.py
   ```

2. The GUI will open, where you can:

   * **Select an Excel file** by clicking the "Selecionar" button.
   * **Specify the number of tables** you want to process and click "Confirmar".
   * For each table, enter the **sheet name**, **header row**, and **database table name**.
   * **Click "Processar e Enviar para o Banco"** to process and send the data to your NeonDB database.

---

## File Structure

```
excel-to-neondb/
├── app.py             # Main application code
├── .env               # Store your NeonDB URL (create this file)
├── requirements.txt   # Python dependencies
└── README.md          # Project documentation
```

---

## Technologies Used

* **Tkinter:** Used to build the graphical user interface.
* **Pandas:** Used for reading Excel files and processing the data into DataFrames.
* **SQLAlchemy:** Used to interact with the NeonDB database.
* **Python-dotenv:** Used for loading environment variables from the `.env` file.

---

## Troubleshooting

* **Error: `ModuleNotFoundError: No module named 'tkinter'`**

  * Tkinter should be installed by default with Python, but if you're on Linux, you may need to install it separately:

    * On Ubuntu/Debian:

      ```bash
      sudo apt-get install python3-tk
      ```
    * On macOS:

      ```bash
      brew install python-tk
      ```

* **Error: `Database connection failed`**

  * Ensure your `DATABASE_URL` in the `.env` file is correct. Double-check the format and credentials.

---

## Contributing

If you'd like to contribute to the project, feel free to fork the repository, create a new branch, and submit a pull request with your changes. Be sure to run the code and tests before submitting.

---

## License

This project is open-source and available under the [MIT License](LICENSE).

---

### Notes

* If you need to add more advanced features like validation, error handling, or different export formats, feel free to extend this project.
* For a production environment, consider using more advanced libraries for handling large Excel files or database interactions.
