import os
from time import sleep
from app import create_app, db
from flask_migrate import migrate, upgrade, init
from datetime import datetime, UTC

upgrade_time = datetime.now(UTC).strftime("%d-%b-%Y %H:%M:%S")

app = create_app()

with app.app_context():
    if not os.path.exists("migrations"):
        print("Initializing migrations folder...")
        sleep(2)
        init()

    print("Generating migration...")
    sleep(3)
    migrate(message=f"auto migration timestamp: {upgrade_time}")

    print("Applying migration...")
    sleep(3)
    upgrade()
    finish_time = datetime.now(UTC).strftime("%d-%b-%Y %H:%M:%S")
    print(f"Done! \nDatabase migration upgrade completed on {finish_time}")
