# -*- coding: utf-8 -*-
#!/usr/bin/python

import os
import sys
import email
import getopt
import re
import os.path
import logging  
import uuid
import xlsxwriter
import pandas as pd
from curses import wrapper
#CRITICAL ERROR WARNING	 INFO DEBUG	NOTSET
#logging.debug('debug')   
#logging.debug('info')   
#logging.warning('warn')  
#logging.error('error')

logging.basicConfig(filename = os.path.join(os.getcwd(), 'logdetails'), \
level = logging.INFO, filemode = 'w', format = '%(asctime)s - %(levelname)s: %(message)s')   


def usage():
	print '''
[*] Help Information:

[!] -d: path where eml files are Stored. Ex. (/root/Desktop/emails/)

[!] -h: help

[*] Outputs are stored in current folder!

[*] It Will Parse 
		Date
		From Name 
		From Email
		To Name
		To Email
		CC Name
		CC Email
		X-Originating-IP
		Subject
		Message
		Phone
	Read Path Recursively!


'''
def force_to_unicode(text):
    "If text is unicode, it is returned as is. If it's str, convert it to Unicode using UTF-8 encoding"
    return text if isinstance(text, unicode) else text.decode('utf8')

def decode_email(msgfile,r,c,filesheet1,atch):	
	#print atch
	if r%100==0:		
		print " [*]   "+str(r)+" Files Processed!"
	msgfile= force_to_unicode(msgfile)
	#print msgfile
	
	fp = open(msgfile)

	msg = email.message_from_file(fp)

	fp.close()
	
	logging.debug("="*60)
	
	#header
	
	#parse and decode subject
	subject = msg.get("subject")
	try:
		h_subject = email.Header.Header(subject)
		dh_subject = email.Header.decode_header(h_subject)
		subject = dh_subject[0][0]
		subjectcode = dh_subject[0][1]
		ter=subject.encode('utf-8')
		#print ter
		filesheet1.write(r+1,9,force_to_unicode(str(ter)))
		#print subject	
		if(subjectcode != None):
			#subject = unicode(subject,subjectcode)
			subject = subject.decode(subjectcode,'ignore')
			logging.debug("subject:"+ subject.encode('GBK','ignore'))
			#filesheet1.write(r+1,9,str(subject))
			#print subject
			
	except:
		logging.debug("subject:"+ subject)
		

		
	#messageid
	message_id = ""
	try:
		message_id = msg.get("Message-ID")
		logging.debug( "message_id: "+message_id)
		
	except:
		message_id = str(uuid.uuid4())
		logging.warning("Fail To Get Message_id. Use Random UUID:%s",message_id)
		
	
	
	#parse and decode from
	from_username = ""
	from_domain = ""
	try:
		hmail_from = email.utils.parseaddr(msg.get("from"))

		fromname =hmail_from[0]
		ptr = hmail_from[1] .find('@')
		from_username = hmail_from[1][:ptr]
		from_domain = hmail_from[1] [ptr:]
		logging.debug( "from: "+from_username+from_domain)		
		filesheet1.write(r+1,1,force_to_unicode(from_username+from_domain))
		filesheet1.write(r+1,0,force_to_unicode(fromname))
	except:
		from_username = ""
		from_domain = ""
		logging.debug( "from: "+from_username+from_domain)	
		filesheet1.write(r+1,1,force_to_unicode(from_username+from_domain))
		
		
	

	
	#parse and decode to
	to_username = ""
	to_domain = ""
	try:
		if(msg.get("to") !=None) :
			hmail_tos = msg.get("to").split(',')
			for hmail_to in hmail_tos:
				hmail_to = email.utils.parseaddr(hmail_to)				
				toname= hmail_to[0]
				ptr = hmail_to[1] .find('@')
				to_username = hmail_to[1][:ptr]
				to_domain = hmail_to[1] [ptr:]
				logging.debug( "to: "+to_username+to_domain)
				filesheet1.write(r+1,2,force_to_unicode(toname))
				filesheet1.write(r+1,3,force_to_unicode(to_username+to_domain))
								
	except:
		to_username = ""
		to_domain = ""
		logging.debug( "to: "+to_username+to_domain)
		filesheet1.write(r+1,3,force_to_unicode(to_username+to_domain))
		
	
		

	
	#parse and decode Cc
	cc_username = ""
	cc_domain = ""
	try:
		cclistn=[]
		ccliste=[]
		tr=[]
		if(msg.get("Cc") !=None) :
			hmail_ccs = msg.get("Cc").split(',')
			for hmail_cc in hmail_ccs:
				hmail_cc = email.utils.parseaddr(hmail_cc)	
				cclistn+=(hmail_cc[0]).split(',')					

				ptr = hmail_cc[1] .find('@')
				cc_username = hmail_cc[1][:ptr]
				cc_domain = hmail_cc[1][ptr:]
				ccliste+=(cc_username+cc_domain).split(',')
				logging.debug( "cc: "+cc_username+cc_domain)
			filesheet1.write(r+1,5,force_to_unicode(str(ccliste)))
			filesheet1.write(r+1,4,force_to_unicode(str(cclistn)))
				 
	except:
		cc_username = ""
		cc_domain = ""
		logging.debug( "cc: "+cc_username+cc_domain)
		filesheet1.write(r+1,5,cc_username+cc_domain)
		filesheet1.write(r+1,4,force_to_unicode(str(ccname)))
		
		
	
	#parse and decode Date
	hmail_date = ""
	try:
		hmail_date = email.utils.parsedate(msg.get("Date"))		
		date =str(hmail_date[2])+"/"+str(hmail_date[1])+"/"+str(hmail_date[0])+" "+str(hmail_date[3])+":"+str(hmail_date[4])		
		logging.debug( "Date: "+ str(hmail_date))
		filesheet1.write(r+1,7,force_to_unicode(str(date)))

		 
	except:
		hmail_date = ""
		logging.debug( "Date: "+ str(hmail_date))
		filesheet1.write(r+1,7,force_to_unicode(str(date)))
		
	

	#parse and decode sender ip
	IP = ""
	try:
		hmail_recv = msg.get("X-Originating-IP")
		if(hmail_recv != None):
			hmail_recv = hmail_recv.strip("[]")
			IP = hmail_recv
			 
		else :
			hmail_receiveds =  msg.get_all("Received")
			if(hmail_receiveds !=None):
				for hmail_received in hmail_receiveds:
					m = re.findall(r'from[^\n]*\[(?<![\.\d])(?:\d{1,3}\.){3}\d{1,3}(?![\.\d])',hmail_received)
					if(len(m)>=1):
						hmail_recv = re.findall(r'(?<![\.\d])(?:\d{1,3}\.){3}\d{1,3}(?![\.\d])',m[0])
						if(len(hmail_recv)>=1):
							IP = hmail_recv[0]
							#print IP
							 

		logging.debug( "IP:"+IP)
		filesheet1.write(r+1,6,force_to_unicode(IP))
	except:
		IP = ""
		logging.debug( "IP:"+IP)
		filesheet1.write(r+1,6,IP)
		 
		
				
	#header end
	logging.debug( "+"*60)
	
	#body start
	counter = 1
	for part in msg.walk():
		if part.get_content_maintype() == 'multipart':
			continue

		#content plain
		content_plain = ""
		########################################## PARSING Mobile no #################################################
		try: 
			sfind=re.compile("\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4}$")			
			if (part.get_content_type() == "text/plain"):
				####################################wrting msg#############
				#filesheet1.write(r+1,10,content_plain)

				m2=re.findall(sfind,part.get_payload())	

				if m2:
					filesheet1.write(r+1,8,m2[0])
				code = part.get_content_charset()
				if( code == None):
					content_plain = part.get_payload()
					logging.debug( content_plain)
					
					#filesheet1.write(r+1,10,content_plain)
				else:
					content_plain = part.get_payload(decode=True).decode(code,'utf8')
					#emailtxt= content_plainstring.encode('utf-8')					
					#print content_plain
					#r"http\S+", ""
					fil_text = re.sub(r"http\S+", '', content_plain, flags=re.MULTILINE)
					#fil_text=re.sub(r"_x000D_","",content_plain, flags=re.MULTILINE)
					#filet=fil_text.replace('_x000D','').split()	
					fil_text=fil_text.replace('\r\r\r\r', ' ')				
					fil_text=fil_text.replace('\r\r\r', ' ')
					fil_text=fil_text.replace('\r\r', ' ')	
					fil_text=fil_text.replace('\r', ' ')
					#fil_text=fil_text.replace('\n', ' ')							
					
					filesheet1.write(r+1,10,force_to_unicode(fil_text))
					logging.debug( content_plain.encode('GBK','ignore'))
				logging.debug( "+"*60)				

				continue
		except:
			print  "content_plain parse error"	
			
		#content html
		content_html = ""
		try:
			if (part.get_content_type() == "text/html"):
				code = part.get_content_charset()
				if( code == None):
					content_html = part.get_payload()
					logging.debug( content_html)
				else:
					try:
						content_html = part.get_payload(decode=True).decode(code,'ignore')
					except:
						content_html = part.get_payload(decode=True).decode('GBK','ignore')
					#print unicode(part.get_payload(decode=True),code)
					logging.debug( content_html.encode('GBK','ignore'))
				logging.debug( "+"*60)
				continue
		except:
			print  "content_html parse error"	
	#attachment
		try:
			if (part.get('Content-Disposition') != None):
				# this part is an attachment
				name = part.get_filename()
				
				counter += 1
				try:
					h = email.Header.Header(name)
					dh = email.Header.decode_header(h)
					filename = dh[0][0]
					filecode = dh[0][1]
					if(filecode != None):
						filename = filename.decode(filecode,'ignore')
					if not filename:
						filename = 'part-%03d' % (counter)
				except:
					filename = name
				logging.debug( 'attachment:'+ filename)
				logging.debug("+"*60)

				if(os.path.exists('./'+atch) == False):
					os.mkdir("./"+atch)
				fp = open(os.path.join('./'+atch+'/', filename), 'wb')
				fp.write(part.get_payload(decode=True))
				fp.close()

				continue
		except:
			logging.warning("attachment parse error,but continue")	

	
	logging.debug( "="*60)

	 

def main():	
	if sys.argv[1:]==[]:
		sys.argv[1:]=['-h']
		
	opts,args=getopt.getopt(sys.argv[1:], 'i:d:h')
	startdir = None
	count = 0
	
	for o, k in opts:
		if o=='-h':
			usage()
			sys.exit()		
		if o=='-d':
			startdir = k		
		emlist=[]
	
	try:

		if(startdir!=None):				
				
			for dirpath, dirnames, filenames in os.walk(startdir):	
				for dirs in dirnames:
					match = re.findall(r'[\w\.-]+@[\w\.-]+', str(dirs))
					if match:
						for emailsfolder in match:
							emlist+=match		
			#print emlist
			try:
				for emailsfolder in emlist:
					#print emailsfolder
					r=0
					c=0
					print "Scanning Folder : "+startdir+emailsfolder
					print "wait...."				
										
					workbook = xlsxwriter.Workbook(emailsfolder+'.xlsx')
					bold = workbook.add_format({'bold': True})
					filesheet1 = workbook.add_worksheet("details")
					format5 = workbook.add_format({'num_format': 'dd/mm/yy hh:mm','bold': True})
					bold = workbook.add_format({'bold': True})
					filesheet1.write(0,0,"From Name",bold)
					filesheet1.write(0,1,"From Email",bold)
					filesheet1.write(0,2,"To Name",bold)
					filesheet1.write(0,3,"TO Email",bold)
					filesheet1.write(0,4,"CC Name",bold)
					filesheet1.write(0,5,"CC Email",bold)
					filesheet1.write(0,6,"X-Originating-IP",bold)
					filesheet1.write(0,7,"Date",format5)	
					filesheet1.write(0,8,"Phone Numbers",bold) 
					filesheet1.write(0,9,"subject",bold) 
					filesheet1.write(0,10,"Message",bold) 

					for dirpath, dirnames, filenames in os.walk(startdir+emailsfolder):
						for filename in filenames:
							if filename.endswith(".eml"):
								filepath = os.path.join(dirpath, filename)
								#filepath=filepath.encoding('utf-8')
								#print "parsing : "+filepath
								logging.info("Parsing:"+filepath)
								decode_email(filepath,r,c,filesheet1,emailsfolder)
								count = count + 1
								r=r+1
								c=c+1				
					workbook.close()
					
			except:				
				pass			

		else:
			#print "Parsing:"+msgfile
			logging.info("Parsing:"+msgfile)
			#decode_email(msgfile,r,c,filesheet1)
			#count = count + 1
			#r=r+1
			#c=c+1
			print "some error in file"

	except IOError as e:
		print "IO error.QUIT",e
		logging.error("IO error:"+str(e))
	except:
		print "Unkown Error in Gettng File"
		logging.error("UnkownError in Gettng File")

	
	logging.info("Parse Email Done: "+str(count) +" Eml Files Parsed")
	print "\n"
	print "!!!!!!!!!!!!!!Extracting  Done !!!!!!!!!!!!!!!!!"
	print "\n"
	
	################################################Removing Duplicates#########################
	print "[*]!!!!!!!!!!!!!Removing Duplicates!!!!!!!!!!!!!!!"
	
	try:
		for emailsfolder in emlist:		 
			sheet="details"
			columns=["Message"]
			df= pd.read_excel(emailsfolder+".xlsx",sheet,index= False)			 
			df = df.drop_duplicates(['From Email'])			
			#df=df['Message'].replace('_x000D_',' ',inplace=True)
			df.to_excel(emailsfolder+"-Without-newline.xlsx",sheet_name=sheet, index= False)
			print "[*]"+emailsfolder+"     Done!"
		print "\n"
		print "All Done!"
		print "[*][*][*][*][*][*][*][*][*][*][*][*][*][*][*][*][*][*][*]"
	except:
		print "Some Error in finding Duplicates."

	
if __name__ == '__main__':
	main()