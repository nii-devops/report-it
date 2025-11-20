import os
import subprocess


def run_migrations():
    """Automate Flask-Migrate workflow."""
    
    # Receive message as input..
    message = input("Enter your flask migration message: \n")

    # 1. Run flask db init (only if not done yet)
    if not os.path.exists("migrations"):
        print("Initializing migrations directory...")
        subprocess.run(["flask", "db", "init"], check=True)

    # 2. Run flask db migrate -m "<message>"
    print("Generating migration...")
    subprocess.run(["flask", "db", "migrate", "-m", message], check=True)

    # 3. Run flask db upgrade
    print("Applying migration...")
    subprocess.run(["flask", "db", "upgrade"], check=True)

    print("Migration completed successfully!")


run_migrations()