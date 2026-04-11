# Copyright (c) 2026 Tran Xuan An
# This file is part of CAD Column Inspector Pro.
# Licensed under the MIT License.

import logging

# Configure logging for the application
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("cad_tool.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
