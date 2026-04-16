import os
import sys

PROXY_PORT = int(os.environ.get("PROXY_PORT", "12803"))
PROXY_API_KEY = os.environ.get("PROXY_API_KEY", "")
CHROME_PROFILE_DIR = os.environ.get(
    "CHROME_PROFILE_DIR",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "chrome-profile"),
)
# "auto" = pick by platform, "windows" = force Windows driver, "linux" = force Playwright driver
PROXY_DRIVER = os.environ.get("PROXY_DRIVER", "auto")


def is_linux_driver():
    if PROXY_DRIVER == "linux":
        return True
    if PROXY_DRIVER == "windows":
        return False
    return sys.platform.startswith("linux")
