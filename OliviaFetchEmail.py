import imaplib
import os
import subprocess
import time

class PortalEmail():
	"""
	Interface to downloading attachments from an email address, save them to a folder and execute.
	
	Credit for initial code: https://gist.github.com/johnpaulhayes/106c3e40dc04b6a6b516
	"""

	connection = None
	error = None

	def __init__(self, mail_server, username, password):
		self.connection = imaplib.IMAP4_SSL(mail_server)
		self.connection.login(username, password)
		self.connection.select(readonly=False) # so we can mark mails as read

	def close_connection(self):
		"""
		Close the connection to the IMAP server
		"""
		self.connection.close()

	def save_attachment(self, msg, download_folder="C:\Users\Olivia\Downloads"):
		"""
		Given a message, save its attachments to the specified download folder

		return: file path to attachment
		"""
		att_path = "No attachment found."
		for part in msg.walk():
			if part.get_content_maintype() == 'multipart':
				continue
			if part.get('Content-Disposition') is None:
				continue

			filename = part.get_filename()
			fixname = filename.replace("/","_")
			name, ext = os.path.splitext(fixname)
			att_path = os.path.join(download_folder, fixname)
			if os.path.isfile(att_path):
				att_path = os.path.join(download_folder, name + "1" + ext)
			fp = open(att_path, 'wb')
			fp.write(part.get_payload(decode=True))
			fp.close()
		return att_path

	def fetch_unread_messages(self):
		"""
		Retrieve unread messages
		"""
		self.connection.select()
		emails = []
		(result, messages) = self.connection.search(None, 'UnSeen')
		if result == "OK":
			for message in messages[0].split(' '):
				try: 
					ret, data = self.connection.fetch(message,'(RFC822)')
				except:
					print "No new emails to read."
					self.close_connection()
					exit()

				msg = email.message_from_string(data[0][1])
				if isinstance(msg, str) == False:
					emails.append(msg)
				response, data = self.connection.store(message, '+FLAGS','\\Seen')

			return emails

		self.error = "Failed to retreive emails."
		return emails

	def parse_email_address(self, email_address):
		"""
		Helper function to parse out the email address from the message

		return: tuple (name, address). Eg. ('John Doe', 'jdoe@example.com')
		"""
		return email.utils.parseaddr(email_address)

""" Usage """

portal_email = PortalEmail('10.20.160.10', 'olvia', 'test')

while True:
	emails = portal_email.fetch_unread_messages()

	for m in emails:
		try:
			attachment = portal_email.save_attachment(m)
			print(attachment)
			if not attachment == "No attachment found.":
				subprocess.Popen(["cmd.exe","/c","start",attachment])
		except Exception as e:
			print(e)
	
	print("Processed {0}".format(len(emails)))
	portal_email.close_connection()
	time.sleep(20)

	print "Processed {0}".format(len(emails))
	portal_email.close_connection()