'''
Created on Jun 25, 2018

@author: haibo.li1
'''

from __future__ import unicode_literals
from __future__ import print_function

import os
import time
import sys
from time import sleep
import pyodbc
from os import getenv
import pymssql

try:
    from pywinauto import application
except ImportError:
    import os.path
    pywinauto_path = os.path.abspath(__file__)
    pywinauto_path = os.path.split(os.path.split(pywinauto_path)[0])[0]
    sys.path.append(pywinauto_path)
    from pywinauto import application

app = application.Application()    


# Launch OM Console
def launch_console():

    try:
        app.start(r"C:\Program Files\Nuance\Output Manager\Console 5.0.0.473\NSiOutputManagerConsole.exe")
        app['ConnectToOutputManagerServerDlg'].wait('ready')
        print("Starting OM console")  
    except application.ProcessNotFoundError:
        print("You must first start OM Console "\
            "Player before running this script")
        sys.exit()
        
# On the default Connect to Output Manager Server Dialogue, Click "connect" button to connect server
    
    app['ConnectToOutputManagerServerDlg']['Connect'].click()
    
    print("Clicked the Connect button, waiting for the launch of the console")
    sleep(20)

    
# if the File store has not been configured, configure File store
def config_filestore():
    dlg_store = app.window(title='localhost - Invalid File Store')
    dlg_store.Wait('ready')

    print(dlg_store)
    print(dlg_store.wrapper_object())
        
# Config_file_store:
    dlg_store['Yes'].click()

# create a shared folder under c:\omFileStore
    dir_filestore = 'C:\omFileStore'
    if not os.path.exists(dir_filestore):
        #         os.rmdir(dir_filestore)
        os.makedirs(dir_filestore)
    
# Assign the File Store path
    win = app.window(title='Output Manager Console - [localhost: 1]')
    win.Wait('ready')
  
    win['Add'].click()
    child = app.window(title='localhost - Add File Store')
    child.Wait('ready')
#    win.print_control_identifiers()
    child['Edit'].set_edit_text("C:\omFileStore")
    child['OK'].click()
    
    win_dis = app.window(title='localhost - Distributed System Warning')
    win_dis.Wait('ready')
    win_dis['Yes'].click()
    
    win['Apply'].click()
    print("File store is configured")

    
"""  Connect SQLEXPRESS using ODBC driver and pyodbc library
def connect_DB():
    print("DB test")

    conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=.\SQLEXPRESS; DATABASE=OMDB ;UID=haibo.li1; PWD=Tiger%tgb; Trusted_Connection=yes;')
    cursor = conn.cursor()
    
    cursor.execute("SELECT alertTypeId FROM dbo.AlertType")
    while 1:
        row = cursor.fetchone()
        if not row:
            break
        print(row.alertTypeId)
    
    conn.close()
"""  


""" Connect SQL Express using pumssql """


class OM:
    """
    Output Manager Server

    Besides methods, following member variables is accessible by QA:
    version: The version of product
    """
    dbUser = r'haibo.li1'
    dbPassword = 'Tiger%tgb'
    database = 'OMDB'
    dbAddr = None
    dbcur = None
    
    def __init__(self):
        
        self.dbAddr = "WAT-HLI10"
        print(self.dbAddr)
    def SQL(self, query):
        """
        Connect to OM database and execute query.
        Input: query - The query
            For example,
                om.SQL(SELECT * FROM table1)
        Return: The result from database.
        """
        if not self.dbcur:
            self.dbcur = pymssql.connect(self.dbAddr, self.dbUser, self.dbPassword, self.database).cursor()
            print(self.dbaddr)
        self.dbcur.execute(query)
        return self.dbcur
  
    
def main():
    """ launch_console()
    #   connect_DB()
    
    if app.window(title='localhost - Invalid File Store').exists():
        print("File store is not configured")
        config_filestore()
    """
    
    """om = OM()
    my_query = "SELECT fileStoreId FROM dbo.FileStore" 
    
    result_dbcur = om.SQL(my_query)
    print(result_dbcur)
    while 1:
        row = result_dbcur.fetchone()
        if not row:
            break
        print(row.alertTypeId)
    
     """
    server = 'WAT-HLI10\SQLEXPRESS'
    print(server)
    user = 'nuance\haibo.li1'
    print(user)
    password = 'Tiger%tgb'

    conn = pymssql.connect(server, user, password, 'OMDB')
    print(conn)
    cursor = conn.cursor() 
    print(cursor)  
if __name__ == "__main__":
    main()
