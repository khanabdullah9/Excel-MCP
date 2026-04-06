import logging
import os

logging.basicConfig(
    filename= os.path.join(os.path.dirname(__file__), "app.log"),
    encoding="utf-8",
    filemode="a",
    format="{asctime} - {levelname} - {message}",
    style="{",
    datefmt="%Y-%m-%d %H:%M",
)

def log(err_msg: str):
    logging.error(err_msg)