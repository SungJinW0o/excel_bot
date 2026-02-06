import os
import sys

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from excel_bot.auth import *  # noqa: F401,F403


if __name__ == "__main__":
    from excel_bot.auth import authorize, get_user

    user = get_user("analyst1@example.com")
    authorize(user, "run_pipeline")
    print(f"{user} authorized to run_pipeline")
