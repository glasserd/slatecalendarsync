'''
	Slate - Google Calendar Sync
	This program syncs a users Slate calendar to Google Calendar
	
	David Glasser
	glasserd@union.edu
	Union College
	Schenectady, NY
	
	Version 1.4.1
	Released 8/1/19
	
	
	Copyright 2019 Union College NY

	Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

	The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
			
'''

import httplib2
import os
import argparse
import json
import sys
import time
import hashlib
import requests
import html
import configparser
import threading
import time
import urllib
import logging
from logging.handlers import TimedRotatingFileHandler
from urllib.parse import urlparse
from http.server import BaseHTTPRequestHandler, HTTPServer
from pprint import pprint
from datetime import date, datetime, timedelta

# Email libraries
import smtplib
from email.mime.text import MIMEText

# Timezone Library
from pytz import timezone
import pytz

# Google libraries
from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools
from oauth2client import file

# Read Configuration File

try:
	config = configparser.ConfigParser()
	config.read('config.ini')

	logFileName = config['Files']['LogFile']	
	CLIENT_SECRET_FILE = config['Files']['ClientSecretFile']
	
	logDays = int(config['Logging']['LogDaysArchive'])
	logLevel = config['Logging']['LogLevel']
	logModuleLevel = logging.getLevelName(logLevel)

	pastDays = int(config['CalendarSyncing']['NumberOfPriorDays'])
	futureDays = int(config['CalendarSyncing']['NumberOfFutureDays'])
	syncInterval = int(config['CalendarSyncing']['SyncInterval'])

	emailFrom = config['Emails']['EmailFromAddress']
	emailTo = config['Emails']['ErrorEmailAddress'].split(',')
	emailEventChanges = config['Emails'].getboolean('EmailEventChanges')
	
	syncServer = config['Servers']['SyncServer']
	syncServerPort = config['Servers']['syncServerPort']
	slateServer = config['Servers']['SlateServer']
	
	syncServerUrl = syncServer
	
	onCampusInterviewLocation = config['Settings']['OnCampusInterviewLocation']
	
except KeyError as err:
	print ("Unsuccessful read of configuration file config.ini")
	print (format(err))
	sys.exit()


# Start Logging
log_dir = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'logs')
if not os.path.exists(log_dir):
	os.makedirs(log_dir)
logPath = os.path.join(log_dir, logFileName)

logger = logging.getLogger('slate_sync')
logger.setLevel(logModuleLevel)

formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

handler = TimedRotatingFileHandler(logPath, when='midnight', interval=1, backupCount=logDays)
handler.setLevel(logModuleLevel)
handler.setFormatter(formatter)
logger.addHandler(handler)

try:
	import argparse
	parser = argparse.ArgumentParser(description='Union College Slate - Google Calendar Sync',epilog="Created by David Glasser at Union College", parents=[tools.argparser])
	group = parser.add_mutually_exclusive_group()
	group.add_argument("-a", "--add", type=str, metavar='email_address',
                    help="Add additional Google calendar")
	group.add_argument("-d", "--delete", type=str, metavar='email_address',
                    help="Delete Google calendar")
	group.add_argument("-c", "--clear", type=str, metavar='email_address',
                    help="Clear all Slate events from Google calendar")
	group.add_argument("-s", "--sync", type=str, metavar='email_address',
                    help="Sync a single existing calendar")	
	flags = parser.parse_args()
	
except ImportError:
	flags = None
	logger.warning('Failed to capture command line arguments.')

SCOPES = 'https://www.googleapis.com/auth/calendar https://www.googleapis.com/auth/userinfo.email'
APPLICATION_NAME = 'Union College Slate-Google Calendar Sync'

# Currently if an interview is cancelled the slot stays assigned to the person. To accomodate this we'll prefix empty slots with "Potential"
ONCAMPUS_INTERVIEW_TEXT_NOT_ASSIGNED = 'On Campus Interview'

# Get credentials
# Check to see if credentials directory exists
credential_dir = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'credentials')
if not os.path.exists(credential_dir):
	os.makedirs(credential_dir)

# Check to see if master list of calendars exists
calendar_list_file = 'calendar_list.json'
if not os.path.isfile(calendar_list_file):
	f = open(calendar_list_file, 'w')
	calendars = {}
	json.dump(calendars, f)
	f.close()

# Open calendar list
f = open(calendar_list_file, 'r')
calendars = json.load(f)
f.close()
logger.info('Found the following calendars: %s', calendars)


def main():
	logger.info('Start SlateSync')
	logger.debug('Log level set to: %s', logger.getEffectiveLevel())
	
	errors = []
	
	# Calculate sync windowBegin
	windowBegin = datetime.now(pytz.utc) - timedelta(days=pastDays)
	windowBegin = windowBegin.replace(hour=0, minute=0, second=0, microsecond=0)
	windowEnd = datetime.now(pytz.utc) + timedelta(days=futureDays)
	windowEnd = windowEnd.replace(hour=0, minute=0, second=0, microsecond=0)
	windowGrace = datetime.now(pytz.utc) - timedelta(days=pastDays-1)
	windowGrace = windowGrace.replace(hour=0, minute=0, second=0, microsecond=0)
	logger.info('Setting sync window. Window Begin: %s Window End: %s Window Grace: %s', windowBegin, windowEnd, windowGrace)
	

	# Loop through calendars
	for googleCalendar, calendarInfo in calendars.items():
		slateCalendar = calendarInfo['calendarUrl']
		try:
			eventColorOnCampus = calendarInfo['eventColorOnCampus']
		except:
			eventColorOnCampus = ''
		try:
			eventColorOther = calendarInfo['eventColorOther']
		except:
			eventColorOther = ''
	
		# Check to make sure that a single calendar sync wasn't requested
		if flags.sync is not None:
			sync_calendar = (flags.sync).strip()
			
			if googleCalendar != sync_calendar:
				continue #go to beginning of loop
	
		logger.info('Syncing events for calendar: %s', googleCalendar)
		print('Syncing events for calendar: ', googleCalendar)
	
		credential_file = googleCalendar + '.json'
		credential_path = os.path.join(credential_dir, credential_file)
		store = oauth2client.file.Storage(credential_path)
		credentials = store.get()
				
		if not credentials or credentials.invalid:
			# Try to refresh credentials
			try:
				logger.info ('Attempting to refresh credentials for calendar: %s ', googleCalendar)
				credentials = credentials.refresh(httplib2.Http())
			except Exception as e:
				logger.error ('Exception caught while refreshing credentials: %s', e)
				logger.exception(e)
			if not credentials or credentials.invalid:
				logger.error ('Google Calendar: %s could not synced. No valid OAuth Token. Have user reauthenticate.', googleCalendar)
				logger.info ('credential_file: %s credential_path: %s', credential_file, credential_path)
				logger.info ('credentials: %s ', credentials)
				errors.append ('Google Calendar: ' + googleCalendar + ' could not synced. No valid OAuth Token. Have user reauthenticate.')
		else:
			logger.info('Retrieved valid credentials for calendar: %s', googleCalendar)
		
			http = credentials.authorize(httplib2.Http(timeout=15))
			service = discovery.build('calendar', 'v3', http=http, cache_discovery=False)
			
			# Get users events
			googleEvents = readGoogleCalendar(service, googleCalendar, windowBegin, windowEnd)
			logger.info ('Google Calendar: %s Slate events in Google Calendar: %s', googleCalendar, googleEvents)
			googleEventKeys = list(googleEvents.keys())
			#print(googleEvents)
			
			# Get Slate events
			try:
				slateEvents = readSlateCalendar(googleCalendar, slateCalendar, windowBegin, windowEnd)
			except:
				print ('Unable to retrieve Slate Calendar ', slateCalendar)
			else:
				logger.info ('Google Calendar: %s Slate events in Slate: %s', googleCalendar, slateEvents)
				
				# Store changes
				calendarModifications = []
				
				# Compare differences and make updates
				for eventId, eventDetails in slateEvents.items(): # Iterate over Slate Events
				
					try:
					
						if (eventId in googleEvents): #Check if event exists in Google Calendar
							logger.debug('Event %s from calendar %s already exists in Google Calendar. Look for changes.', eventId, googleCalendar)

							googleEvent = googleEvents[eventId]
							googleEventKeys.remove(eventId)
							
							#print ('Slate Summary   ', eventDetails['summary'])
							#print ('Google Summary  ',  googleEvent['summary'])
							#print ('Slate location  ', eventDetails['location'])
							#print ('Google location ',  googleEvent['location'])
							#print ('Slate start     ', eventDetails['start'], type(eventDetails['start']))
							#print ('Google start    ',  googleEvent['start'], type(googleEvent['start']))
							#print ('Slate end       ', eventDetails['end'], type(eventDetails['end']))
							#print ('Google end      ',  googleEvent['end'], type(googleEvent['end']))
							#print ('Google color    ',  googleEvent['colorId'], type(googleEvent['colorId']))
							
							
							# Determine if event is on campus
							onCampusEvent = False
							if (eventDetails['location'].startswith(onCampusInterviewLocation)):
								onCampusEvent = True
								
							#Set eventColor
							if (onCampusEvent):
								eventColor = eventColorOnCampus
							else: 
								eventColor = eventColorOther
							
							# Check if event has changed
							summaryChange = False
							if (googleEvent['summary'] != eventDetails['summary']):
								summaryChange = True
							
							locationChange = False
							if (googleEvent['location'] != eventDetails['location']):
								locationChange = True
							
							startChange = False
							if (googleToDateTime(googleEvent['start']) != eventDetails['start']):
								startChange = True
							
							colorChange = False
							if (googleEvent['colorId'] != eventColor):
								colorChange = True
		
							## Check to see if the end of the event changed. 
							endChange = False
							
							# Slate does not return end date for all day events
							if ( type(eventDetails['start']) == date and eventDetails['end'] == '' ): 
								endChange = False
								
							# Check to see if the start date is of type datetime but the end date is of type date. If so, make sure event is one hour long
							elif ( type(eventDetails['start']) == datetime and type(eventDetails['end']) == date and googleToDateTime(googleEvent['end']) == (eventDetails['start'] + timedelta(hours=1))):
								endChange = False
							
							# No end time in Google, make sure end time is 1 hour after start time
							elif ( eventDetails['end'] == '' and googleToDateTime(googleEvent['end']) == (eventDetails['start'] + timedelta(hours=1)) ):
								endChange = False
								
							# Check to see if Google has end time by Slate does not
							elif ( eventDetails['end'] == '' and googleEvent['end'] != ''):
								endChange = True
								
							# Check to see if the event ends before it starts. If so, make sure end time is 1 hour after start time
							elif (eventDetails['end'] < eventDetails['start']  and googleToDateTime(googleEvent['end']) == (eventDetails['start'] + timedelta(hours=1)) ):
								endChange = False
							
							elif (googleToDateTime(googleEvent['end']) != eventDetails['end']):
								endChange = True
							
							# Check to see if event changed
							if (summaryChange or locationChange or startChange or endChange or colorChange):
								logger.debug ('Event has changed. summaryChange: %s locationChange: %s startChange: %s endChange: %s colorChange: %s', summaryChange, locationChange, startChange, endChange, colorChange)
								logger.debug(eventId, eventDetails)
								
								logger.debug ('Slate Summary   %s', eventDetails['summary'])
								logger.debug ('Google Summary  %s', googleEvent['summary'])
								logger.debug ('Slate location  %s', eventDetails['location'])
								logger.debug ('Google location %s', googleEvent['location'])
								logger.debug ('Slate start     %s %s', eventDetails['start'], type(eventDetails['start']))
								logger.debug ('Google start    %s %s', googleEvent['start'], type(googleEvent['start']))
								logger.debug ('Slate end       %s %s', eventDetails['end'], type(eventDetails['end']))
								logger.debug ('Google end      %s %s', googleEvent['end'], type(googleEvent['end']))
								logger.debug ('Window grace    %s', windowGrace)
								
								
								if (eventDetails['start'] < windowGrace):
									logger.debug('Event %s from calendar %s occurs during grace period. Make no changes to event', eventId, googleCalendar)
								else:
									#Event has changed. Delete old event and recreate.							
									deleteError = deleteEvent(service, googleCalendar, googleEvent['eventId'])
									if (deleteError != ''):
										errors.append(deleteError)
									
									addError = addEvent(service, googleCalendar, eventId, eventDetails['summary'], '', eventDetails['location'], eventDetails['start'], eventDetails['startTimeZone'], eventDetails['end'], eventDetails['endTimeZone'], eventColor)
									if (addError != ''):
										errors.append(addError)

									# Only send notification if summary or time change
									if (summaryChange or startChange or endChange):										
										calendarModifications.append('Deleting event: ' + googleToDateTime(googleEvent['start'], False).strftime("%B %d, %Y %I:%M %p")  + ' - ' +  googleEvents[eventId]['summary'])
										calendarModifications.append('Adding event: ' + eventDetails['start'].astimezone(timezone('America/New_York')).strftime("%B %d, %Y %I:%M %p") + ' - ' + eventDetails['summary'])
								
							
						else:
							if (eventDetails['start'] < windowGrace):
									logger.debug('Event %s from calendar %s occurs during grace period. Do not add event.', eventId, googleCalendar)
							else:
								logger.debug('Event %s from calendar %s does not exists in Google Calendar. Add event.', eventId, googleCalendar)
								
								# Determine if event is on campus
								onCampusEvent = False
								if (eventDetails['location'].startswith(onCampusInterviewLocation)):
									onCampusEvent = True
									
								#Set eventColor
								if (onCampusEvent):
									eventColor = eventColorOnCampus
								else: 
									eventColor = eventColorOther
								
								addError = addEvent(service, googleCalendar, eventId, eventDetails['summary'], '', eventDetails['location'], eventDetails['start'], eventDetails['startTimeZone'], eventDetails['end'], eventDetails['endTimeZone'], eventColor)
								calendarModifications.append('Adding event: ' + formatDate(eventDetails['start']) + ' - ' + eventDetails['summary'])
								if (addError != ''):
									errors.append(addError)
								
					except Exception as e:
							logger.error ('Error processing event. Event ID: : %s', eventId)
							logger.exception(e)
					
				#Remove Google Events that are no longer present in Slate Calendar
				try:
					for eventId in googleEventKeys:
						if (googleToDateTime(googleEvents[eventId]['start'], False) < windowGrace):
							logger.debug('Event %s from calendar %s occurs during grace period. Make no changes to event.', eventId, googleCalendar)

						else:
							logger.info('Deleting event %s from calendar %s. Event no longer in Slate calendar.', eventId, googleCalendar)
						
							deleteError = deleteEvent(service, googleCalendar, googleEvents[eventId]['eventId'])
							
							calendarModifications.append('Deleting event: ' + googleToDateTime(googleEvents[eventId]['start'], False).strftime("%B %d, %Y %I:%M %p")  + ' - ' +  googleEvents[eventId]['summary'])
							if (deleteError != ''):
								errors.append(deleteError)
				except Exception as e:
					logger.error ('Error deleting event. Event ID: : %s', eventId)
					logger.exception(e)
						
				if (len(calendarModifications) > 0 and emailEventChanges):
					
					msg = MIMEText('\n'.join(calendarModifications))
					msg['Subject'] = 'Slate Calendar Updates'
					msg['From'] = emailFrom
					msg['To'] = googleCalendar
					s = smtplib.SMTP('mail.union.edu')
					s.sendmail(emailFrom, googleCalendar, msg.as_string())
					s.quit()
					
					logger.info('Events have changed in calendar %s. Sending the following email to user: %s', googleCalendar, '***'.join(calendarModifications))
						
	
	#Send email if error occured
	if (len(errors) > 0):
		msg = MIMEText('\n'.join(errors))
		msg['Subject'] = 'Slate-Google Sync Errors'
		msg['From'] = emailFrom
		msg['To'] =  ', '.join(emailTo)
		s = smtplib.SMTP('mail.union.edu')
		s.sendmail(emailFrom, emailTo, msg.as_string())
		s.quit()
	
	#Finish
	logger.info('Finish SlateSync')

def formatDate(d):
	f = ''
	if (type(d) == date):
		f = d.strftime("%B %d, %Y")
	elif (type(d) == datetime):
		f = d.astimezone(timezone('America/New_York')).strftime("%B %d, %Y %I:%M %p")
	
	return f
	
	
def readSlateCalendar(calendar, slateCalendar, windowBegin, windowEnd):
	logger.info ('readSlateCalendar - Starting method for calendar: %s', calendar)
	logger.debug('test debug event')

	r = requests.get(slateCalendar)
	
	if r.status_code != 200:
		logger.error ('Unable to retrieve Slate Calendar %s. HTTP Status Code: %s', slateCalendar, r.status_code)
		raise Exception('No Slate Calendar')

	logger.debug('readSlateCalendar - raw output from Slate: %s', r.text)	
	
	events = {}
	tempEvent = {}
	
	i = 0
	inEvent = False
	for line in r.text.splitlines():
		i = i + 1
		l = ICalLine(line)
		
		#print (str(i).ljust(3), " In Event:", str(inEvent).ljust(5), "EventStart:", str(l.eventStart()).ljust(5), "EventEnd:", str(l.eventEnd()).ljust(5), " Line: ", line)
		#print (str(i).ljust(3), " Property:", str(l.property).ljust(20), "Attribute:", str(l.attribute).ljust(20), "Attribute Value:", str(l.attributeValue).ljust(20), "Value:", str(l.value).ljust(5), " Line: ", line)
		
		try: 
			if (inEvent):
				if (l.eventEnd()): # Event is over, write it to dictionary
					inEvent = False

					# Check to see if start date is between range we are looking for
					if (type(tempEvent['start']) == date):
						startDate = datetime.combine(tempEvent['start'], datetime.min.time(), pytz.utc)
					else:
						startDate = tempEvent['start']

					try:
						if (startDate >= windowBegin and startDate <= windowEnd and tempEvent['status'] == 'CONFIRMED'):
							# Calculate digest
							m = hashlib.md5()
							m.update(tempEvent['summary'].encode('utf-8'))
							m.update(tempEvent['location'].encode('utf-8'))
							m.update(tempEvent['start'].strftime("%A, %d. %B %Y %I:%M%p").encode('utf-8'))
							m.update(tempEvent['startTimeZone'].encode('utf-8'))
							if (tempEvent['end'] != ''):
								m.update(tempEvent['end'].strftime("%A, %d. %B %Y %I:%M%p").encode('utf-8'))
							m.update(tempEvent['endTimeZone'].encode('utf-8'))
							digest = m.hexdigest()

							# Store event
							if not (tempEvent['potentialInterview'] and tempEvent['start'].date() <= date.today()):
								# TODO event Id is currently stored as a hash. It looks like slate is now returning the GUID for the event. That may be a better option.
								events[tempEvent['slateGUID']] = tempEvent
						else:
							logger.debug('Event not in window. startDate: %s windowBegin: %s windowEnd: %s', startDate, windowBegin, windowEnd)
					except Exception as e:
						logger.error ('readSlateCalendar - Error parsing iCal Feed for calendar: : %s', calendar)
						logger.error ('startDate: %s windowBegin: %s windowEnd: %s', startDate, windowBegin, windowEnd)
						logger.exception(e)


				else: # We are in the middle of an event, check for a value we are tracking
					if (l.property == 'UID'):
						tempEvent['slateGUID'] = removeICalEscape(l.value)
					if (l.property == 'SUMMARY'):
						tempEvent['summary'] = html.unescape(removeICalEscape(l.value))

						if (tempEvent['summary'] == ONCAMPUS_INTERVIEW_TEXT_NOT_ASSIGNED):
							tempEvent['summary'] = 'Potential ' + ONCAMPUS_INTERVIEW_TEXT_NOT_ASSIGNED
							tempEvent['potentialInterview'] = True

					elif (l.property == 'STATUS'):
						tempEvent['status'] = l.value
					elif (l.property == 'LOCATION'):
						tempEvent['location'] = removeICalEscape(l.value)
					elif (l.property == 'DTSTART'):
						year   = int (l.value[0:4])
						month  = int (l.value[4:6])
						day    = int (l.value[6:8])
						if (l.attribute == 'VALUE' and l.attributeValue == 'DATE'): # Date object
							tempEvent['start'] = date(year, month, day)
						else:
							hour   = int (l.value[9:11])
							minute = int (l.value[11:13])
							second = int (l.value[13:15])
							tempEvent['start'] = datetime(year, month, day, hour, minute, second, tzinfo=pytz.utc)
						if (l.attribute == 'TZID'):
							tempEvent['startTimeZone'] = l.attributeValue

					elif (l.property == 'DTEND'):
						year   = int (l.value[0:4])
						month  = int (l.value[4:6])
						day    = int (l.value[6:8])
						if (l.attribute == 'VALUE' and l.attributeValue == 'DATE'): # Date object
							tempEvent['end'] = date(year, month, day)
						else:
							hour   = int (l.value[9:11])
							minute = int (l.value[11:13])
							second = int (l.value[13:15])
							tempEvent['end'] = datetime(year, month, day, hour, minute, second, tzinfo=pytz.utc)

						if (l.attribute == 'TZID'):
							tempEvent['endTimeZone'] = l.attributeValue
		
			else: #Not in an event
				if (l.eventStart()): # Check to see if this is the start of an event
					inEvent = True

					tempEvent = {
						'summary'			:'',
						'location'			:'',
						'status'			:'',
						'start'				:'',
						'startTimeZone'		:'',
						'end'				:'',
						'endTimeZone'		:'',
						'potentialInterview': False,
					}
		except Exception as e:
			logger.error ('readSlateCalendar - Error parsing iCal Feed for calendar: : %s', calendar)
			logger.exception(e)
	
	#print ('Events')
	#pprint (events)
	logger.info ('readSlateCalendar - Total Slate events for calendar %s: %s', calendar, len(events))
	print ('Total Slate events: ', len(events) )
	
	return events
	
def readGoogleCalendar(service, calendar, windowBegin, windowEnd):
	logger.info ('readGoogleCalendar - Starting method for calendar: %s', calendar)
	
	userEvents = {}
	
	# Calculate start and end dates for range search		
	startDateFrmt = windowBegin.isoformat()
	endDateFrmt = windowEnd.isoformat()
	
	try:
		eventsResult = service.events().list(
			calendarId='primary', timeMin=startDateFrmt, timeMax=endDateFrmt, maxResults=2500, singleEvents=True,
			orderBy='startTime').execute()
		events = eventsResult.get('items', [])	
	
		logger.info ('readGoogleCalendar - Retrieved event list for calendar: %s', calendar)	
		
		if not events:
			pass
		for event in events:
			slateID = ''
					
			# Find Slate ID
			if 'extendedProperties' in event:
				if 'private' in event['extendedProperties']:
					if 'SlateID' in event['extendedProperties']['private']:
						slateID = event['extendedProperties']['private']['SlateID']
						
						#print (event)
						
						try:
							if slateID in userEvents:
								logger.warning ('Google Calendar: %s Duplicate event found in Google Calendar. Deleting... SlateID =  %s', calendar, slateID)
								deleteEvent(service, calendar, event['id'])
						
							else:
								#Event is a Slate Event. Add it to the dictionary
								userEvents[slateID] = {
									'eventId'		: event['id'],
									'summary'		: event['summary'],
									'location'		: '',
									'start'			: '',
									'startTimeZone'	: '',
									'end'			: '',
									'endTimeZone'	: '',
									'colorId'		: '',
								}
								
								if 'location' in event:
									userEvents[slateID]['location'] = event['location']
								
								if 'date' in event['start']:
									userEvents[slateID]['start'] = event['start']['date']
								elif 'dateTime' in event['start']:
									userEvents[slateID]['start'] = event['start']['dateTime']
								if 'date' in event['end']:
									userEvents[slateID]['end'] = event['end']['date']
								elif 'dateTime' in event['end']:
									userEvents[slateID]['end'] = event['end']['dateTime']
									
								if 'timeZone' in event['start']:
									userEvents[slateID]['startTimeZone'] = event['start']['timeZone']
									
								if 'timeZone' in event['end']:
									userEvents[slateID]['endTimeZone'] = event['start']['timeZone']
									
								if 'colorId' in event:
									userEvents[slateID]['colorId'] = event['colorId']
							
						except Exception as e:
							logger.error ('Could not read Slate event from Google Calendar. Slate ID: : %s', slateID)
							logger.exception(e)
	except Exception as e:
		logger.error ('Could not retrieve events from Google calendar: %s', calendar)
		logger.exception(e)
						
	return userEvents
		

def addEvent(service, calendar, slateId, summary, description, location, start, startTimeZone, end, endTimeZone, eventColor):
	logger.debug('addEvent method. Calendar = [%s] slateId = [%s] summary = [%s] description = [%s] location = [%s] start = [%s] end = [%s]', calendar, slateId, summary, description, location, start, end)

	addError = ''

	
	# Check to see if this an all day event. If so set end date to start date
	if (type(start) == date and end == ''): 
		end = start
	# Check to see if the start date is of type datetime but the end date is of type date. If so, assume event is one hour long
	elif (type(start) == datetime and type(end) == date):
		end = start + timedelta(hours=1)
		logger.warning ('Google Calendar: %s Start date inclues date & time but end date has no time associated with it. Assuming end is 1 hour after start. %s %s', calendar, start, summary)
	# Check to see if there is a start time but no end time. If so, assume event is one hour long.
	elif (type(start) == datetime and end == ''):
		end = start + timedelta(hours=1)
		logger.warning ('Google Calendar: %s No end date provided. Assuming end is 1 hour after start. %s %s', calendar, start, summary)
	# Check to see if end is of type datetime and start is of type date
	elif (type(start) == date and type(end) == datetime):
		logger.warning ('Google Calendar: %s End time provided, but no start time. Letting code throw error. %s %s', calendar, start, summary)
	# Check to see if the event ends before it starts. If so, assume the event is 1 hour long
	elif (end < start):
		end = start + timedelta(hours=1)
		logger.warning ('Google Calendar: %s Event ends before it starts. Assuming end is 1 hour after start. %s %s', calendar, start, summary)
		
		
	startIso = start.isoformat()
	if (type(start) == datetime):
		startType = 'dateTime'
	else:
		startType = 'date'		
		
	if (type(end) == datetime):
		endType = 'dateTime'
		endIso = end.isoformat()
	elif (type(end) == date):
		endIso = end.isoformat()
		endType = 'date'
	else:
		endIso = ''
		endType = 'date'
	
	event = {
		'summary': summary,
		'description': description,
		'location': location,
		'start': {
			startType: startIso, #'2015-10-15T13:00:00'
			'timeZone': startTimeZone, #'America/New_York'
		},
		'end': {
			endType: endIso,
			'timeZone': endTimeZone,
		},
		"extendedProperties": {
			"private": {
				('SlateID'): slateId,
			},
		},
	}
	
	if eventColor != '':
		event['colorId'] = eventColor
	
	try:
		service.events().insert(calendarId='primary', body=event).execute()
		logger.info ('Google Calendar: %s Event created: %s', calendar, event)
	except Exception as e:
		logger.error ('Google Calendar: %s Could not create event: %s Exception: %s', calendar, event, e)
		addError = 'Google Calendar: ' + str(calendar) + ' Could not create event: '  + str(event) + 'Exception:' + str(e)
	
	return addError
	
	
def deleteEvent(service, calendar, eventId):
	deleteError = ''
	try:
		service.events().delete(calendarId='primary', eventId=eventId).execute()
		logger.info ('Google Calendar: %s Event deleted. Event Id: %s', calendar, eventId)
	except Exception as e:
		logger.error ('Google Calendar: %s Could not delete event: %s Exception: %s', calendar, eventId, e)
		deleteError = 'Google Calendar: ' + str(calendar) + ' Could not delete event: '  + str(eventId) + 'Exception:' + str(e)
		
	return deleteError
	
def getGoogleCredentials(email_address, credential_dir):
	"""Gets valid user credentials from storage.

	If nothing has been stored, or if the stored credentials are invalid,
	the OAuth2 flow is completed to obtain the new credentials.

	Returns:
		Credentials, the obtained credential.
	"""
	
	credential_file = email_address + '.json'
	credential_path = os.path.join(credential_dir, credential_file)
	store = oauth2client.file.Storage(credential_path)	
	credentials = store.get()	
	
	if not credentials or credentials.invalid:
		flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
		flow.user_agent = APPLICATION_NAME
		credentials = tools.run_flow(flow, store, flags)
		print ('Storing credentials to ', credential_path)
	return credentials

def googleToDateTime(date1, convertToUTC=True):	
	if (len(date1) == 10): # Check if date is YYYY-MM-DD format
		try:
			year = date1[0:4]
			month = date1[5:7]
			day = date1[8:10]			
			datef = date(int(year), int(month), int(day))
		except Exception as e:
			datef = date1
			logger.error("googleToDateTime: Unable to convert 10 character date %s. Exception: %s", date1, e)			
		return datef

	else:
		#Format:  2015-10-28T15:00:00-04:00
		tzf = date1[19:].replace(':','')		
		tdate = date1[0:19] + tzf		
		
		datef = datetime.strptime(tdate,"%Y-%m-%dT%H:%M:%S%z")		
		if (convertToUTC == True):
			datef = datef.astimezone(timezone('UCT'))
		return datef
	
def removeICalEscape(value):
	return value.replace('\\','')


class ICalLine:
	def __init__(self, line):
		self.line = line
		
		components = line.split(':')
		self.property = components[0]
		self.value = components[1]
		self.attribute = ''
		self.attributeValue = ''
		
		components = self.property.split(';')
		if (len(components) == 2):
			self.property = components[0]
			
			attrComp = components[1].split('=')
			self.attribute = attrComp[0]
			self.attributeValue = attrComp[1]
		
		
	def eventStart(self):
		if self.property == 'BEGIN' and self.value == 'VEVENT':
			return True
		else:
			return False
			
	def eventEnd(self):
		if self.property == 'END' and self.value == 'VEVENT':
			return True
		else:
			return False

			
# Manage Dictionary of Slate Calendars
def calendarExists(calendar):
	if calendar in calendars:
		return True
	else:
		return False
		
def createCalendarUrl(id):
	url = slateServer + '/manage/event/?user=' + id + '&output=ical'
	return url

def addCalendar(new_calendar):
	if new_calendar in calendars:
		message = ('Calendar ' + new_calendar + ' already exists.')
		logger.info (message)
		print (message)
		return message
	else:		
		slateUrl = input('Enter Slate Calendar URL: ')		
	
		credentials = getGoogleCredentials(new_calendar, credential_dir)
		
		with lock:
			calendars[new_calendar] = {'calendarUrl':slateUrl}
			f = open(calendar_list_file, 'w')
			json.dump(calendars, f)
			f.close()
		
		message = ('Calendar ' + new_calendar + ' added.')
		logger.info (message)
		print (message)
		return message

def deleteCalendar(delete_calendar):
	if delete_calendar in calendars:
		with lock:
			del calendars[delete_calendar]
			f = open(calendar_list_file, 'w')
			json.dump(calendars, f)
			f.close()
				
			credential_file = delete_calendar + '.json'
			credential_path = os.path.join(credential_dir, credential_file)
			os.remove(credential_path)				
				
			logger.info ('Calendar %s deleted.', delete_calendar)
			print ('Calendar ', delete_calendar, ' deleted.')
	else:
		logger.info ('Calendar %s does not exist.', delete_calendar)
		print ('Calendar ', delete_calendar, ' does not exist.')
			
			
			
# HTTP Server
class testHTTPServer_RequestHandler(BaseHTTPRequestHandler):

	# GET
	def do_GET(self):
		print(self.path)
		parsed = urlparse(self.path)
		parameters = (urllib.parse.parse_qs(parsed.query))
		message = ''
				
		if self.path.startswith('/?calendar='):
			# Initial page entered by user		
			new_calendar = parameters['calendar'][0]
			calendar_url = createCalendarUrl(parameters['id'][0])
			
			if not calendarExists(new_calendar):
				flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES, redirect_uri=syncServerUrl)
				flow.user_agent = APPLICATION_NAME
				flow.params['access_type'] = 'offline'
				flow.params['state'] = calendar_url
				auth_uri = flow.step1_get_authorize_url()
				
				self.send_response(302)
				self.send_header('Location', auth_uri)
				self.end_headers()
				return
					
			else:
				message = 'Calendar ' + new_calendar + ' already exists'
				print (message)
			
		elif self.path.startswith('/?error='):
			message = 'Error occured while requesting authorization from Google.'
			
		elif 'calendarlist' in self.path:
			message = ''
			for googleCalendar, calendarInfo in calendars.items():
				message += googleCalendar + '<br />'
			
		elif 'code' in self.path:
			# Page redirected back from auth server
			auth_code = parameters['code'][0]
			calendar_url = parameters['state'][0]
			flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES, redirect_uri=syncServerUrl)
			credentials = flow.step2_exchange(auth_code)
			
			http = credentials.authorize(httplib2.Http())
			user_info_service = discovery.build('oauth2', 'v2', http=http)
			
			user_info = None
			try:
				user_info = user_info_service.userinfo().get().execute()
				new_calendar = user_info.get('email')
				print('calendar', new_calendar)
				
				with lock:
					calendars[new_calendar] = {'calendarUrl':calendar_url}
					f = open(calendar_list_file, 'w')
					json.dump(calendars, f)
					f.close()
										
					credential_file = new_calendar + '.json'
					credential_path = os.path.join(credential_dir, credential_file)
					storage = oauth2client.file.Storage(credential_path)
					storage.put(credentials)
					print ('Storing credentials to ', credential_path)
				
				message = 'Succesfully added calendar ' + new_calendar
				
			except Exception as e:
				logger.error('An error occurred: %s', e)
				message = 'Error adding calendar.'
			
			
		else:
			message = 'Union College Slate-Google Calendar Sync. Please log in to Slate to set up the sync.'

		self.send_response(200)
		self.send_header('Content-type', 'text/html')
		self.send_header('Cache-Control', 'no-cache')
		self.end_headers()
		print(message)
		self.wfile.write(bytes('<html><head><title>Slate Calendar Sync</title></head>', "utf8"))
		self.wfile.write(bytes(message, "utf8"))  
  
		return
			

def web():
	print('starting server...')
 
	# Server settings
	server_address = ('localhost', int(syncServerPort))
	httpd = HTTPServer(server_address, testHTTPServer_RequestHandler)
	print('running server...')
	httpd.serve_forever()			

def sync():
	print ('Starting Calendar Sync')
	t_sync = threading.Timer(syncInterval, sync)
	t_sync.daemon = False
	t_sync.start()
	
	main()
			
if __name__ == '__main__':

	# Create lock object
	lock = threading.Lock()
	
	# Check to see if we need to add a new calendar
	if flags.add is not None:
		new_calendar = (flags.add).strip()
		logger.info ('Adding new calendar: %s', new_calendar)		
		addCalendar(new_calendar)
		sys.exit()
		
	# Check to see if we need to delete a calendar
	if flags.delete is not None:
		delete_calendar = (flags.delete).strip()
		logger.info ('Deleting calendar: %s ', delete_calendar)	
		deleteCalendar(delete_calendar)
		sys.exit()
		
	# Check to see if we need to clear out all Slate events on a calendar
	if flags.clear is not None:
		windowBegin = date.today() - timedelta(days=1000)
		windowEnd = date.today() + timedelta(days=1000)
	
		clear_calendar = (flags.clear).strip()
		logger.info ('Clearing all events from: %s', clear_calendar)
		
		if clear_calendar in calendars:
			
			credentials = getGoogleCredentials(clear_calendar, credential_dir)
			http = credentials.authorize(httplib2.Http())
			service = discovery.build('calendar', 'v3', http=http)
			
			googleEvents = readGoogleCalendar(service, clear_calendar, windowBegin, windowEnd)
		
			for event, eventDetails in googleEvents.items():
				deleteEvent(service, clear_calendar, eventDetails['eventId'])
			
			logger.info ('Calendar %s has been cleared.', clear_calendar)
			print ('Calendar ', clear_calendar, ' has been cleared.')
		else:
			logger.info ('Calendar %s does not exist.', clear_calendar)
			print ('Calendar ', clear_calendar, ' does not exist.')
		
		sys.exit()
		
	# Start Web Server	
	t_web = threading.Thread(target=web)
	t_web.daemon = True
	t_web.start()
	
	# Sync calendars
	sync()
