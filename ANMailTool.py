#!/usr/bin/env python

import sys, os, time, datetime
from getpass import getpass

from selenium import webdriver # selenium approach
from selenium.webdriver.support.ui import WebDriverWait

import smtplib # smtp approach
from email.MIMEMultipart import MIMEMultipart # smtp approach
from email.MIMEImage import MIMEImage # smtp approach
from email.MIMEText import MIMEText # smtp approach
from email.MIMEBase import MIMEBase # smtp approach

class ANMailTool:
	"""
	Tool for Anime North Artist Alley registration, specifically to assist
	timed email sending.
	NOTE: Setting default values into __init__ will be more convenient.
	
	Dependencies:
	1) "message.txt": Text file for email content in root folder.
	2) Must use gmail account to send for time accuracy!!
	3) Computer clock set to "time.nist.gov"
	
	
	Optional:
	1) "artwork/xxx.jpg": all attached images. Optional if links provided.
	2) "prereg/xxx.pdf": AN preregistration pdf.
	"""
	
	def __init__(self):
		""" Default values. """
		
		############## SET VALUES HERE BEFORE USE FOR CONVENIENCE ##############
		self.officialTime = "23:00:00" # official time email must be sent
		self.officialMail = "ancmarket@gmail.com" # table reg email
		self.testMail = "danny.yiu@gmail.com" # address for test emails
		self.fromMail = "danny.yiu@gmail.com" # your email address
		self.fromMailPass = "" # password for your email
		########################################################################
		
		# don't touch values here
		self.sendTime = None
		self.sendMail = None
		self.sendMsg = None
		self.artCheck = None
		self.mode = None # "test" or "official"
		
		# selenium vars and elements
		self.webdriver = None
		self.loginField = None
		self.passField = None
		self.loginButton = None
		self.composeButton = None 
		self.toField = None
		self.fromField = None
		self.subjectField = None
		self.msgField = None
		self.sendButton = None
		
		# debug prefix
		self.preD = "[DEBUG]"
		self.preO = "[OUTPUT]"
		self.preI = "[INPUT]"
		
	def usage(self):
		""" Print usage. """
		
		print "Usage:\n" + \
			  "--time=hh:mm:ss     Send test email at specified time.\n" + \
			  "--auto              Send official registration at official time."
	
	def pass_params(self):
		""" Params check. """
		
		if not self.fromMail or "gmail" not in self.fromMail:
			while "gmail" not in self.fromMail:
				self.fromMail = input("%sEnter your email: " % self.preI)
				if "gmail" not in self.fromMail:
					print "%sMust be a gmail account." % self.preD
		elif not self.fromMailPass:
			while not self.fromMailPass:
				self.fromMailPass = getpass(
					"%sEnter email password: " % self.preI)
		elif not self.officialTime:
			sys.exit("%sError: Registration time is not set" % self.preD)
		elif not self.officialMail:
			sys.exit("%sError: Registration email not set" % self.preD)
			
		if len(sys.argv) < 2:
			self.usage()
		elif "--time=" in sys.argv[1]:
			# get time string
			stime = sys.argv[1].split("=")[-1]
			# time values check
			for i, digit in enumerate(stime.split(":")):
				if i == 0 and int(digit) > 23:
					sys.exit("%sInvalid hour value. Correct format: HH:MM:SS" %\
							 self.preD)
				elif i > 0 and int(digit) > 59:
					sys.exit("%sInvalid time value. Correct format: HH:MM:SS" %\
							 self.preD)
			#save time value
			self.sendTime = sys.argv[1].split("=")[-1]
			if not self.testMail:
				sys.exit("%sError: Test email address can't be empty" %\
						 self.preD)
			self.sendMail = self.testMail
			self.mode = "test"
		elif "--auto" in sys.argv[1]:
			# set official values
			self.sendTime = self.officialTime
			self.sendMail = self.officialMail
			self.mode = "official"
		else:
			self.usage()
			sys.exit(1)
	
	###################### SMTP APPROACH FUNCTION STARTS HERE ##################	
		
	def smtp_mail_setup(self):
		""" Set up mail to be sent. """
		
		#read message
		try:
			raw = open('message.txt', 'r')
		except:
			sys.exit("%sError opening file \"message.txt\"." % self.preD)
		msg = raw.read()
		raw.close()
		# basic check if art link is included, and if empty message
		if msg:
			self.artCheck = [False, True]["http://" in msg or ".jpg" in msg]
		else:
			sys.exit("%sError: Message file \"message.txt\" cannot be empty." %\
					 self.preD)
		
		# compile email

		self.sendMsg = MIMEMultipart() # multipart container
		self.sendMsg['Subject'] = "Table Registration"
		self.sendMsg['To'] = self.sendMail
		self.sendMsg['From'] = self.fromMail
		self.sendMsg.preamble = "This is a multi-part message in MIME format."
		self.sendMsg.attach(MIMEText(msg))
		
		
		# attach artwork if no links given in message
		if not self.artCheck:
			# check artwork folder
			if not os.listdir("artwork"):
				exit("%sNo artwork links in message " % self.preD + \
					 "and no images in artwork folder.")
			else:
				# attach art to multipart
				fileList = os.listdir("artwork")
				print fileList
				for file in fileList:
					if file[0] != ".":
						with open('artwork/%s' % file, 'rb') as rawImg:
							img = MIMEImage(rawImg.read())
						self.sendMsg.attach(img)
				# mark artwork check as true
				self.artCheck = True
				
		else:
			print "%sGood links given." % self.preO
	
	def smtp_send_mail(self):
		""" Connect and send mail. """
		
		# dependencies check
		if self.sendTime and self.sendMail and self.sendMsg and \
			self.artCheck and self.mode:
			
			try:
				# test connection to mail server
				server = smtplib.SMTP("smtp.gmail.com", 587)
				print "%sLogging in to mail server..." % self.preO
				server.ehlo() # esmtp handling
				server.starttls() # esmtp handling
				server.login(self.fromMail, self.fromMailPass)
				print "%sTest connection successful!" % self.preD
			except:
				sys.exit("%sConnection failure at login." % self.preD)
			server.quit()
			print "%sTest connection closed." % self.preD
			
			# establish new connection at most 60 seconds before send to
			# prevent session expiry
			now = datetime.datetime.now()
			timeString = self.sendTime.split(":")
			diff = datetime.datetime(now.year, now.month, now.day,
				int(timeString[0]), int(timeString[1]), 
				int(timeString[2])) - now
			print "%sWaiting till closer time to connect..." % self.preD
			while diff.seconds > 70:
				time.sleep(5)
				diff = datetime.datetime(now.year, now.month, now.day,
					int(timeString[0]), int(timeString[1]), 
					int(timeString[2])) - datetime.datetime.now()
			# real connection
			server = smtplib.SMTP("smtp.gmail.com", 587)
			server.ehlo() # esmtp handling
			server.starttls() # esmtp handling
			server.login(self.fromMail, self.fromMailPass)
			
			# delay till actual send time
			self.delay(self.sendTime)
			print "here1"
			# send precompiled message, assuming mail_setup has run
			server.sendmail(self.fromMail, self.sendMail, 
				self.sendMsg.as_string())
			print "here2"
			# end connection
			server.quit()
			
	######################## SMTP FUNCTIONS END HERE ###########################
	
	#################### SELENIUM FUNCTIONS START HERE #########################
	def selenium_setup(self):
		"""
		Setup a selenium webdriver instance.
		"""
		try:
			self.webdriver = webdriver.Firefox() # assume default firefox path
		except:
			print "%sWebDriver initiation failure." % self.preD + \
				  " Check FireFox installation path."
	
	def selenium_mail_setup(self):
		"""
		Login and setup mail message ready to be sent.
		"""
		self.webdriver.get("http://mail.google.com")
		
		# login check
		login = None
		try:
			# check for login page elements
			self.loginField = WebDriverWait(
				self.webdriver, timeout=10).until(
				lambda webdriver: self.webdriver.find_element_by_xpath(
				'//input[@id="Email"]'))
			self.passField = self.webdriver.find_element_by_xpath(
				'//input[@id="Passwd"]')	
			self.loginButton = self.webdriver.find_element_by_xpath(
				'//input[@id="signIn"]')
			login = True
		except:
			login = False
			
		if login:
			# login block
			# enter login info
			self.loginField.send_keys(self.fromMail)
			self.passField.send_keys(self.fromMailPass)
			self.loginButton.click()
		else:
			pass
		
		# compose button
		self.composeButton = WebDriverWait(
			self.webdriver, timeout=10).until(
			lambda webdriver: self.webdriver.find_element_by_xpath(
			'//div[text()="COMPOSE"]'))
		self.composeButton.click()
		# to field
		self.toField = WebDriverWait(
			self.webdriver, timeout=10).until(
			lambda webdriver: self.webdriver.find_element_by_xpath(
			'//textarea[@class="vO" and @aria-label="Address"]'))
		self.toField.send_keys(self.sendMail)
		# subject field
		self.subjectField = WebDriverWait(
			self.webdriver, timeout=10).until(
			lambda webdriver: self.webdriver.find_element_by_xpath(
			'//input[@name="subjectbox"]'))
		self.subjectField.send_keys("Table Registration")
		# message field
		if 'message.txt' in os.listdir("."):
			with open('message.txt', 'r') as raw:
				msg = raw.read()
				if len(msg) > 1:
					self.sendMsg = msg
				else:
					self.sendMsg = "Enter message here."
		else:
			print "%sMust provide 'message.txt' for" % self.preD + \
				  "message automation (optional)."
			self.sendMsg = "Enter message here."
		# enter into gmail's message iframe	
		self.webdriver.switch_to_frame(
			self.webdriver.find_element_by_xpath(
			'//div[@aria-label="Message Body"]/iframe'))
		self.msgField = WebDriverWait(
			self.webdriver, timeout=10).until(
			lambda webdriver: self.webdriver.find_element_by_xpath(
			'//body[@class="editable LW-avf"]'))
		self.msgField.send_keys(self.sendMsg)
		self.webdriver.switch_to_default_content()
		
		# (optional) attachments
		print "%sMANUAL ACTION REQUIRED: " % self.preD + \
			  "1) Attach any files if necessary. 2) Be sure they're uploaded."
			
	def selenium_send_mail(self):
		"""
		Use selenium to send mail. Assumes mail as been setup using webdriver.
		"""
		
		# check send button exist
		try:
			self.sendButton = self.webdriver.find_element_by_xpath(
				'//div[contains(@aria-label, "Send ")]')
		except:
			try:
				self.sendButton = self.webdriver.find_element_by_xpath(
					'//div[text()="Send"]')
			except:
				sys.exit("%sCANNOT FIND SEND BUTTON! :O" % self.preD)
		# prompt ready
		raw_input("%sPRESS ENTER TO START TIMED SEND! [Enter]" % self.preD)
		self.delay(self.sendTime)
		self.sendButton.click()
			
		# close webdriver
		#self.webdriver.quit()
		

	####################### SELENIUM FUNCTIONS END HERE ########################
	def delay(self, timeString):
		""" 
		Delay until exact specified time HH:MM:SS. 
		Delay intervals will gradually tighten closer to destination time.
		"""
		
		timeString = timeString.split(":")
		now = datetime.datetime.now()
		future = datetime.datetime(
			now.year,now.month,now.day,
			int(timeString[0]),int(timeString[1]),int(timeString[2]))
		
		print "%sTarget time: %d:%d:%d:%d" % (
				self.preD, future.hour, future.minute, future.second, 0)	
		
		# wait at 1 second intervals when >15s away
		diff = future - datetime.datetime.now()	
		if diff.seconds > 15:
			print "%s%d seconds left, refreshing every 1s..." %\
					(self.preD, diff.seconds)
			while diff.seconds > 15:
				time.sleep(1)
				diff = future - datetime.datetime.now()	
		
		# wait at 0.001s intervals within 15s away
		diff = future - datetime.datetime.now()
		print "%s%d seconds left, refreshing every 0.001s..." %\
				(self.preD, diff.seconds)
		while datetime.datetime.now().second != future.second:
			time.sleep(0.001)
		
		# may delete below output block for even faster speeds
		print "%sWaited till: %d:%d:%d:%d" % (
			self.preD,
			datetime.datetime.now().hour, 
			datetime.datetime.now().minute,
			datetime.datetime.now().second, 
			datetime.datetime.now().microsecond)
			
		
	
	
if __name__ == '__main__':
	# init
	mtool = ANMailTool()
	mtool.pass_params()
	#print "%sSet Time: %s" % (mtool.preO, mtool.sendTime)
	#print "%sSet Mail: %s" % (mtool.preO, mtool.sendMail)
	
	# smtp approach
	#mtool.smtp_mail_setup()
	#mtool.smtp_send_mail()
	
	#selenium approach
	mtool.selenium_setup()
	mtool.selenium_mail_setup()
	mtool.selenium_send_mail()