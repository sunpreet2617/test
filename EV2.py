import re
import dns
from dns import resolver
import socket
import smtplib

from openpyxl import load_workbook

wb =load_workbook('test.xlsx', data_only= True)
sh = wb["Sheet1"]


for row in sh['A{}:A{}'.format(sh.min_row+1,sh.max_row)]:
	for cell in row:
                
		wb.save('test.xlsx')
		addressToVerify = cell.value
		match = re.match('^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,4})$', addressToVerify)
		if match == None:
			print('Bad Syntax for '+addressToVerify)
			#raise ValueError('Bad Syntax')

		resolver = dns.resolver.Resolver()
        while line:
        ip = re.search(r"\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b", line)
        if ip:
            resolver.nameservers = [ip.group(0)]
            try:
                result = resolver.query('opnfv.org')[0]
                if result != "":
                    nameservers.append(ip.group())
            except dns.exception.Timeout:
                continue
                records = dns.resolver.query('lambdadirect.com','MX')
		mxRecord = records[0].exchange
		mxRecord = str(mxRecord)
     	
# Get local server hostname
		host = socket.gethostname()

# SMTP lib setup (use debug level for full output)
		server = smtplib.SMTP()
		server.set_debuglevel(0)

# SMTP Conversation
		server.connect(mxRecord)
		server.helo(host)
		server.mail('me@domain.com')
		print(addressToVerify)
		code, message = server.rcpt(str(addressToVerify))
		server.quit()
		print(code)
# Assume 250 as Success
# Assume 550 as Failure
		if code == 550:
			sh.cell(row=cell.row,column=2).value="Fail"

    			
		if code == 250:
			sh.cell(row=cell.row,column=2).value="Success"
				
		else:
			sh.cell(row=cell.row,column=2).value="Fail"
                        

				
wb.save('test.xlsx')

print("Done")