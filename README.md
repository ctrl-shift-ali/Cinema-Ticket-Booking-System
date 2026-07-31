# Cinema Ticket Booking System

This project is a small Flask-based web application for booking cinema tickets online. It lets users browse available movies, choose a showtime, select seats, and generate a PDF ticket for their booking.

## Project Overview

The application is built with:

- Python and Flask for the backend
- HTML templates for the web pages
- CSS for the styling and seat colors
- Excel for storing booking records
- ReportLab for generating PDF tickets

The app simulates a cinema booking flow with a simple user interface and a working ticket confirmation process.

## How the Project Works

### 1. Home page
When the app starts, the home page displays a list of movies with:

- movie title
- description
- room number
- available seats
- booked seats
- available showtimes and prices

Users can choose a movie and a showtime from this page.

### 2. Booking details page
After choosing a movie and time, the user enters:

- their name
- number of tickets

These details are used to move to the seat selection page.

### 3. Seat selection page
The seat selection page shows a seat map for the cinema room.

The seat map includes:

- available seats shown in green
- booked seats shown in red
- selected seats shown in purple

Users can click on available seats to select them. The chosen seats are stored temporarily in the browser and later submitted with the form.

### 4. Booking confirmation
When the user submits the booking form:

- the selected seats are validated
- the app checks whether the seats are still available
- if the seats are free, the booking is saved
- a PDF ticket is created and downloaded later

### 5. Ticket generation
Once the booking is confirmed, the app generates a PDF ticket containing:

- booking ID
- movie title
- room number
- seat label
- show time
- ticket price
- customer name

## Main Files in the Project

### CinemaManagementSystem.py
This is the main Flask application file. It contains:

- Flask routes for the pages
- logic for loading and saving bookings
- seat map generation
- ticket PDF creation
- movie and showtime data

#### templates/
This folder contains the HTML pages:

- index.html: the main movie selection page
- seats.html: the seat selection page

#### static/
This folder contains the styles used by the app:

- styles.css: controls the look of the pages, seat colors, buttons, and layout

#### data/
This folder is used for storing booking data. The project creates and manages an Excel file here for persistence.

#### tickets/
This folder stores generated PDF tickets after a booking is completed.

### Features

- browse movie listings
- view room availability
- select seats interactively
- prevent double booking of the same seat
- save bookings to Excel
- generate PDF tickets
- show booking confirmation and download link

### How to Run the Project Locally

#### 1. Install Python
Make sure Python 3.10+ is installed on your system.

#### 2. Create a virtual environment

```bash
python -m venv .venv
.venv\Scripts\activate
```

#### 3. Install dependencies

```bash
pip install -r requirements.txt
```

#### 4. Start the application

```bash
python CinemaManagementSystem.py
```

#### 5. Open the app in your browser

Go to:

```text
http://127.0.0.1:5000
```
Or use the [Cinema Management System](https://cinema-ticket-managing-system.vercel.app ) website made by me...

### Dependencies

The project uses the following Python packages:

- Flask
- openpyxl
- reportlab

These are listed in requirements.txt.

## GitHub and Deployment Notes

This project can be pushed to GitHub and deployed on a Python hosting platform such as Render, Railway, or PythonAnywhere.

GitHub Pages is not suitable for this app because it is a dynamic Flask application rather than a static website.

## Project Structure

```text
cinema_app/
├── CinemaManagementSystem.py
├── requirements.txt
├── README.md
├── .gitignore
├── templates/
│   ├── index.html
│   └── seats.html
├── static/
│   └── styles.css
├── data/
└── tickets/
```

## Notes

- The app creates its own folder structure when it runs if the folders do not already exist.
- The seat selection system is a simple demo booking flow and is designed for learning and small project use.It refreshes everytime you confirm your booking. 

## THANKYOU!
