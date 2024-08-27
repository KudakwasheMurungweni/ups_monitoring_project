# UPS Monitoring Project

## Overview
This project captures runtime data from UPS devices and sends a daily email report.

## Installation

1. Clone the repository:
    ```bash
    git clone https://github.com/KudakwasheMurungweni/ups_monitoring_project.git
    ```

2. Navigate to the project directory:
    ```bash
    cd ups_monitoring_project
    ```

3. Create and activate a virtual environment:
    ```bash
    python -m venv env
    source env/bin/activate  # On Windows use `env\Scripts\activate`
    ```

4. Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```

## Configuration

1. Update `send_email.py` with your email credentials and Excel file path.
2. Adjust the UPS coordinates and Excel sheet details as needed.

## Usage

Run the script manually to test it:
```bash
python send_email.py
