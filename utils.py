# utils.py

import logging

# Simple logging function
def log(message):
    logging.basicConfig(level=logging.INFO)
    logging.info(message)

# Simple version retrieval function
def get_version():
    return "1.0.0"
