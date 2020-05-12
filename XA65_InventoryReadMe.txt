If you use Configuration Logging, you will need to use a UDL file in order for the History section of the script to work.  For an explanation, see http://tinyurl.com/CreateUDLFile.

The UDL file will need to be placed in the same folder as the XA65_Inventory.ps1 script.  The UDL file will need to be named XA65ConfigLog.udl.  You will need to edit the UDL file and add  ;Password=ConfigLogDatabasePassword to the end of the last line in the file.  For example, here is mine (line is one line):

Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=administrator;Initial Catalog=XA65ConfigLog;Data Source=SQL;Password=abcd1234
