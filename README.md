# Project Title
Fetch Data from PostgreSQL and Create PowerPoint Presentation

## Description
This project is a python script that fetches data from a PostgreSQL database and generates a PowerPoint presentation containing the data in a table. The script uses the psycopg2 library to connect to the database and execute a SELECT statement to fetch data, and the python-pptx library to create the PowerPoint presentation.

## Requirements
- psycopg2
- python-pptx
- PostgreSQL

## Environment
- Python 3.6
- PostgresSQL 15

## DB Create
 CREATE DATABASE mofdb
    WITH
    OWNER = postgres
    ENCODING = 'UTF8'
    LC_COLLATE = 'Malay_Malaysia.1252'
    LC_CTYPE = 'Malay_Malaysia.1252'
    TABLESPACE = pg_default
    CONNECTION LIMIT = -1
    IS_TEMPLATE = False; 
    
     INSERT INTO inisiatif_perbelanjaan (inisiatif, agensi, perbelanjaan)
VALUES ('Inisiatif A', 'Agensi A', 123.21), 
       ('Inisiatif B', 'Agensi B', 21.31),
       ('Inisiatif C', 'Agensi C', 33.00),
       ('Inisiatif D', 'Agensi D', 12.33),
       ('Inisiatif E', 'Agensi E', 0.23),
       ('Inisiatif F', 'Agensi F', 32.12),
       ('Inisiatif G', 'Agensi G', 4.21),
       ('Inisiatif H', 'Agensi H', 9.32),
       ('Inisiatif I', 'Agensi I', 123.42),
       ('Inisiatif J', 'Agensi J', 0.11); 

## Usage
1. Make sure that you have a PostgreSQL database set up and running on your localhost.
2. Install the required libraries by running the following command in your terminal: `pip install psycopg2 python-pptx`
3. Edit the script to specify the correct connection details for your database (host, port, database name, username, and password).
4. Run the script using python.
5. The script will generate a PowerPoint presentation in the same directory as the script. The file name will be in the format "Laporan_Bulanan_Generated YYYY-MM-DD HH-MM-SS.pptx"

## Configuration
- You can change the SQL query to fetch data from other table and columns.
- You can adjust the look of the table by editing the width, height, and other properties of the table and its cells.
- You can also customize the presentation by adding or editing slides, or even add a new style.

## Note
This script hardcoded and it may need some changes to work on different environments, make sure to use it as a reference.

## License
This project is open-source written by author.
