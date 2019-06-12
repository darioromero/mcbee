# Imports and initialization of variables
from contextlib import closing # this will correctly close the request
import io
import os
import dropbox
import pandas as pd
import xlrd
from utils import shared_lnk, ACCESS_TOKEN

# This approach is not limited to excel files
# Folder with files to retrieve from it
app_path = "/_McBee-PetroVisor/Rosewood Prod" 

# get token on https://www.dropbox.com/developers/apps/
dbx = dropbox.Dropbox(ACCESS_TOKEN)

# directory of local data to save files from Dropbox site
data_dir = os.path.abspath(os.path.join(os.pardir, 'data'))

# request list of files at application path
req = dbx.files_list_folder(app_path)

# downloading files from Dropbox
for i, f in enumerate(req.entries):
    # get Dropbox file name
    file_name = f.path_display.split('/')[3]

    # compose destination file name
    local_fname = os.path.join(data_dir, file_name)
    
    print(f"... {i} downloading to: {local_fname}")
    # md = dbx.files_download_to_file(local_fname, f.path_display)



