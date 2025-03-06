import os
from utils.mellanni_modules import user_folder
if not os.path.exists(user_folder):
    os.makedirs(user_folder)