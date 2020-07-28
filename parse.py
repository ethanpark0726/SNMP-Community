import re

class Parse:
    def __init__(self, data):

        self.data = data

    def getSNMPCommunity(self):

        snmpCommunity = list()
        pattern = re.compile('snmp-server community')
        
        print('--- Gathering SNMP-Server Community')

        for elem in self.data:
            match = pattern.search(elem)
            if match:
                snmpCommunity.append(elem)
        
        return snmpCommunity
                
                