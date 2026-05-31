# DocumentSystem

## Installation & Running the System

Follow these instructions to set up and run the project locally.

### Prerequisites

- [Python 3.x](https://www.python.org/downloads/)
- [pip](https://pip.pypa.io/en/stable/installation/)
- (Optional but recommended) [virtualenv](https://virtualenv.pypa.io/en/latest/)

### Installation

1. **Clone the repository:**
    ```bash
    git clone https://github.com/ajmayran/DocumentSystem.git
    cd DocumentSystem
    ```

2. **Create and activate a virtual environment:**  
   *(Optional but recommended)*
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows use: venv\Scripts\activate
    ```

3. **Install requirements:**
    ```bash
    pip install -r requirements.txt
    ```

### Database Setup

4. **Make migrations and migrate the database:**
    ```bash
    python manage.py makemigrations
    python manage.py migrate
    ```
### Create Superuser

5. **Create Admin Account:**
    ```bash
    python manage.py createsuperuser
    input username, email, password
    ```
    then run 

    ```bash
    python set_admin.py
    ```

### Running the Server

6. **Start the development server:**
    ```bash
    python manage.py runserver
    ```

6. **Visit** `http://127.0.0.1:8000/` **in your browser to access the application.**

---

Feel free to customize this section for your own deployment or extra setup steps.
