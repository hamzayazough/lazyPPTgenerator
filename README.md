# Lazy PPT Generator

This project automatically captures slides from a Zoom presentation, performs OCR to detect slide numbers, and compiles the captured images into a PowerPoint file. It allows you to select two screen regions:

1. The region containing the slide number.
2. The region to capture for each slide.

Once the script detects a new slide number, it takes a screenshot of the specified region and adds it to a new PowerPoint slide. Finally, it saves the presentation to your **Downloads** folder.

---

## Features

- **Interactive Region Selection**: Select the on-screen area where the slide number appears and the area to capture for each slide.
- **OCR Integration**: Uses Tesseract OCR via pytesseract to detect the slide number and avoid duplicates
- **Automated PowerPoint Creation**: Inserts each captured slide into a new PowerPoint file and saves it in your Downloads folder.
- **Manual Exit**: Press **Ctrl+C** to stop the capture loop whenever you want, or exit automatically when Zoom is no longer running.

---

## Prerequisites

1. **Tesseract OCR**

   - Download and install Tesseract OCR from the [Tesseract at UB Mannheim page](https://github.com/UB-Mannheim/tesseract/wiki)

2. **Pipenv**

   - Install Pipenv if you haven’t already:
     ```bash
     pip install pipenv
     ```

---

## Installation

1. **Clone or Download** this repository.
2. **Navigate** into the project folder.
3. **Install Dependencies** with Pipenv:
   ```bash
   pipenv install
   ```

## How to start?

After completing all prerequisites and installing the dependencies, you can start the application by running the following command from your project's root directory:

```bash
pipenv run python main.py

```
