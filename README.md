S7excel communication with S7 PLC  
e-mail: piotrbzdrega@yandex.com  
Used library: libnodave : http://libnodave.sourceforge.net/  
Don't forget to put libnodave.dll in your system folder like : C:\Windows\System32\  
Tested with PLCSIM Advanced v3.0 update1 & TIA Portal V16  
example: https://youtu.be/7bycvxdYJ7M  
            FUNCTIONS:  
#Max PDU data 80B ( 20 entries daveAddVarToReadRequest())  
#Cycle 1s read refreshing  
#Import/Export tags (same format like TIA)  

TODO:
-import .asc files from S7 v5 
-read/write each datatype char, signed/unsigned...
-ask if dbdw or dbd (dword/dint or real)