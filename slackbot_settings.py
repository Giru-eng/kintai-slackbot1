import os
from os.path import join, dirname
from dotenv import load_dotenv

dotenv_path = join(dirname(__file__), '.env')
load_dotenv(dotenv_path)

API_TOKEN = os.environ.get("SLACKBOT_API_TOKEN")
GOOGLE_APPLICATION_CREDENTIALS = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")

DEFAULT_REPLY = "こんにちは、こちらは勤怠管理Botです！"

PLUGINS = ['plugins']
