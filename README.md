# Creating an Excel report of SNMP community

  - Use openpyxl moduel to create an Excel report
  - Read a tab-delimited file which contains the list of devices

## Report preview
|Hostname|IP Address|SNMP Community|
|:------:|:-----------:|:---------:|
|Switch1|10.1.1.1|snmp-server community myComm RO 1|
|||snmp-server community yourComm RW 1|
|Switch2|10.1.2.1|snmp-server community myComm RO 1|
|Switch3|10.1.3.1|snmp-server community yourComm RW 1|

## Tab-delimited file
Switch1	IOS	10.1.1.1	SSH   
Switch2	IOS	10.1.22.1	SSH   
Switch3	IOS	10.3.1.1	SSH
