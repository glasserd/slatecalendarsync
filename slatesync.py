'''
	Slate - Google Calendar Sync
	This program syncs a users Slate calendar to Google Calendar
	
	David Glasser
	glasserd@union.edu
	Union College
	Schenectady, NY
	
	Version 1.6.4
	Released 9/21/21
	
	
	Copyright 2021 Union College NY

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
from os import curdir, sep

# Email libraries
import smtplib
from email.mime.text import MIMEText

# Timezone Library
from pytz import timezone
import pytz

# Google libraries
from apiclient import discovery
from googleapiclient.errors import HttpError
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
	mailServer = config['Emails']['MailServer']
	
	syncServer = config['Servers']['SyncServer']
	syncServerPort = config['Servers']['syncServerPort']
	slateServer = config['Servers']['SlateServer']
	slateEventWebService = config['Servers']['SlateEventWebService']
	slateEventWebServiceStops = config['Servers']['SlateEventWebServiceStops']
	slateEventWebServiceUsername = config['Servers']['SlateEventWebServiceUsername']
	slateEventWebServicePassword = config['Servers']['SlateEventWebServicePassword']
	
	syncServerUrl = syncServer

	
	openInterviewLabel = config['Settings']['OpenInterviewLabel']
	onCampusInterviewLocation = config['Settings']['OnCampusInterviewLocation']
	googleApiBackoff = config['Settings']['GoogleApiBackoff']
	
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
	
	currentThread = threading.current_thread()

	logger.info('Start SlateSync. Current thread: %s', currentThread.getName())
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
		slateCalendar = googleCalendar
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
			
			# Get Slate events
			try:
				slateEvents = readSlateCalendarWebService(googleCalendar, slateEventWebService, slateEventWebServiceStops, slateEventWebServiceUsername, slateEventWebServicePassword, windowBegin, windowEnd)
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

							# Check to see if the attendee count changed
							attendeeChange = False
							if (summaryChange and eventDetails['type'].lower() == 'event'):
								googleEventIndex = googleEvent['summary'].rfind('(')
								slateEventIndex = eventDetails['summary'].rfind('(')
								if (googleEvent['summary'][0:googleEventIndex] == eventDetails['summary'][0:slateEventIndex]):
									logger.debug ('Event attendance has changed. Event ID: %s', eventId)
									attendeeChange = True
									

							# Check for location change
							
							locationChange = False
							if (googleEvent['location'] != eventDetails['location']):
								locationChange = True
							
							startChange = False
							if (googleToDateTime(googleEvent['start']) != eventDetails['start']):
								startChange = True
							
							colorChange = False
							if (googleEvent['colorId'] != eventColor):
								colorChange = True

							descriptionChange = False
							if (googleEvent['description'] != eventDetails['description']):
								descriptionChange = True
		
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
							if (summaryChange or locationChange or startChange or endChange or colorChange or descriptionChange):
								logger.debug ('Event has changed. summaryChange: %s locationChange: %s descriptionChange: %s startChange: %s endChange: %s colorChange: %s', summaryChange, locationChange, descriptionChange, startChange, endChange, colorChange)
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
								
								

								#Event has changed. Delete old event and recreate.							
								deleteError = deleteEvent(service, googleApiBackoff, googleCalendar, googleEvent['eventId'])
								if (deleteError != ''):
									errors.append(deleteError)
								
								addError = addEvent(service, googleCalendar, eventId, eventDetails['summary'], eventDetails['location'], eventDetails['description'], eventDetails['start'], eventDetails['end'], eventColor)
								if (addError != ''):
									errors.append(addError)

								# Only send notification if summary or time change
								if ((summaryChange and eventDetails['type'].lower() == 'interview') or (summaryChange and not attendeeChange) or startChange):										
									calendarModifications.append('Deleting event: ' + googleToDateTime(googleEvent['start'], False).strftime("%B %d, %Y %I:%M %p")  + ' - ' +  googleEvents[eventId]['summary'])
									calendarModifications.append('Adding event: ' + formatDate(eventDetails['start']) + ' - ' + eventDetails['summary'])
							
							
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
							
							addError = addEvent(service, googleCalendar, eventId, eventDetails['summary'], eventDetails['location'], eventDetails['description'], eventDetails['start'], eventDetails['end'], eventColor)
							calendarModifications.append('Adding event: ' + formatDate(eventDetails['start']) + ' - ' + eventDetails['summary'])
							if (addError != ''):
								errors.append(addError)
								
					except Exception as e:
							logger.error ('Error processing event. Event ID: : %s', eventId)
							logger.exception(e)
					
				#Remove Google Events that are no longer present in Slate Calendar				
				for eventId in googleEventKeys:
					try:
						start = googleToDateTime(googleEvents[eventId]['start'], False)
						if (isinstance(start, datetime)):
							start = start.date()
							
						if (start < windowGrace.date()):
							logger.debug('Event %s from calendar %s occurs during grace period. Make no changes to event.', eventId, googleCalendar)

						else:
							logger.info('Deleting event %s from calendar %s. Event no longer in Slate calendar.', eventId, googleCalendar)
						
							deleteError = deleteEvent(service, googleApiBackoff, googleCalendar, googleEvents[eventId]['eventId'])
							
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
					s = smtplib.SMTP(mailServer)
					s.sendmail(emailFrom, googleCalendar, msg.as_string())
					s.quit()
					
					logger.info('Events have changed in calendar %s. Sending the following email to user: %s', googleCalendar, '***'.join(calendarModifications))
						
	
	#Send email if error occured
	if (len(errors) > 0):
		msg = MIMEText('\n'.join(errors))
		msg['Subject'] = 'Slate-Google Sync Errors'
		msg['From'] = emailFrom
		msg['To'] =  ', '.join(emailTo)
		s = smtplib.SMTP(mailServer)
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

def readSlateCalendarWebService (calendar, slateEventWebService, slateEventWebServiceStops, slateEventWebServiceUsername, slateEventWebServicePassword, windowBegin, windowEnd):
	logger.info ('readSlateCalendarWebService - Starting method for calendar: %s', calendar)

	events = {}

	webServices = [slateEventWebService]
	if (slateEventWebServiceStops != ''):
		webServices.append(slateEventWebServiceStops)

	for ws in webServices:
		r = requests.get(ws + calendar, auth=(slateEventWebServiceUsername, slateEventWebServicePassword))
		
		if r.status_code != 200:
			logger.error ('Unable to retrieve Slate Calendar %s. HTTP Status Code: %s', calendar, r.status_code)
			raise Exception('No Slate Calendar')

		
		for event in r.json()['row']:
			try:
				logger.debug('readSlateCalendarWebService - reading event for %s: %s', calendar, event)
				tempEvent = {
					'summary'			: '',
					'location'			: '',
					'start'				: '',
					'end'				: '',
					'description'		: '',
					'type'				: event['Type'],
				}

				if 'Title' in event:
					if event['Type'].lower() == 'interview':
						if 'Interviewee' in event:
							tempEvent['summary'] = event['Title'] + ' (' + event['Interviewee'] + ')'
						elif openInterviewLabel != '':
							tempEvent['summary'] = event['Title'] + ' (' + openInterviewLabel + ')'
						else:
							tempEvent['summary'] = event['Title']
					elif event['Type'] == 'Stop':
						tempEvent['summary'] = event['Title']
					else:
						tempEvent['summary'] = event['Title'] + ' (' + event['Attendees'] + ')'
				
				if 'Location' in event:
					tempEvent['location'] = event['Location']

				if 'Address' in event:
					tempEvent['location'] = tempEvent['location'] + event['Address']

				if 'Description' in event:
					tempEvent['description'] = event['Description'] 

				if 'TimezoneOffset' in event:
					offset = int(event['TimezoneOffset'])
				else:
					offset = 0
					logger.warning('readSlateCalendarWebService - no timezone for %s: %s', calendar, event['GUID'])

				if 'Start' not in event:
					# We can't create an event without a start time
					continue
				else:
					# Example format: 2019-08-28T12:00:00
					start = event['Start']

					year   = int (start[0:4])
					month  = int (start[5:7])
					day    = int (start[8:10])

					if 'T' in start:
						# Event has a date and a time
						hour   = int (start[11:13])
						minute = int (start[14:16])
						second = int (start[17:19])
						tempDateTime = datetime(year, month, day, hour, minute, second, tzinfo=pytz.utc)

						tempDateTime = tempDateTime - timedelta(minutes=offset)

						tempEvent['start'] = tempDateTime

					else:
						# Event is a date object
						tempEvent['start'] = date(year, month, day)

				if 'End' in event:
					# Example format: 2019-08-28T12:00:00
					end = event['End']

					year   = int (end[0:4])
					month  = int (end[5:7])
					day    = int (end[8:10])

					if 'T' in end:
						# Event has a date and a time
						hour   = int (end[11:13])
						minute = int (end[14:16])
						second = int (end[17:19])
						tempDateTime = datetime(year, month, day, hour, minute, second, tzinfo=pytz.utc)

						tempDateTime = tempDateTime - timedelta(minutes=offset)

						tempEvent['end'] = tempDateTime

					else:
						# Event is a date object
						tempEvent['end'] = date(year, month, day)

				# If event is an interview and occurs in the past delete it from the calendar
				if event['Type'].lower() == 'interview' and event['Attendees'] == '0' and tempEvent['start'].date() <= datetime.now().date():
					logger.debug('readSlateCalendarWebService - Removing unbooked expired interview %s for calendar %s', event['GUID'], calendar)
					continue

				# Check to see if event is in sync window
				if (type(tempEvent['start']) == date):
					startDate = datetime.combine(tempEvent['start'], datetime.min.time(), pytz.utc)
				else:
					startDate = tempEvent['start']

				try:
					if (startDate >= windowBegin and startDate <= windowEnd):
						# Store event					
						events[event['GUID']] = tempEvent
					else:
						logger.debug('Event %s not in window. startDate: %s windowBegin: %s windowEnd: %s', event['GUID'], startDate, windowBegin, windowEnd)
				except Exception as e:
					logger.error ('readSlateCalendar - Error parsing Slate event feed for calendar: : %s', calendar)
					logger.error ('startDate: %s windowBegin: %s windowEnd: %s', startDate, windowBegin, windowEnd)
					logger.exception(e)


				logger.debug('readSlateCalendarWebService - processed event for %s: %s', calendar, tempEvent)
				

			except Exception as e:
				logger.error ('Could not read Slate event from Slate Calendar Feed. Slate ID: : %s', event['GUID'])
				logger.exception(e)


	logger.info ('readSlateCalendarWebService - Total Slate events for calendar %s: %s', calendar, len(events))

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
								deleteEvent(service, googleApiBackoff, calendar, event['id'])
						
							else:
								description = ''
								if 'description' in event:
									description = event['description']

								#Event is a Slate Event. Add it to the dictionary
								userEvents[slateID] = {
									'eventId'		: event['id'],
									'summary'		: event['summary'],
									'location'		: '',
									'description'	: description,
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
		

def addEvent(service, calendar, slateId, summary, location, description, start, end, eventColor):
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
		},
		'end': {
			endType: endIso,
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
	
	
def deleteEvent(service, googleApiBackoff, calendar, eventId):
	deleteError = ''
	try:
		service.events().delete(calendarId='primary', eventId=eventId).execute()
		logger.info ('Google Calendar: %s Event deleted. Event Id: %s', calendar, eventId)
	except HttpError as e:
		if e.resp.status in [403]:
			logger.info('Google Calendar: %s Could not delete event: %s 403 error received. Backing off for %s seconds', calendar, eventId, googleApiBackoff)
			time.sleep(int(googleApiBackoff))

		logger.error ('Google Calendar: %s Could not delete event: %s Exception: %s', calendar, eventId, e)
		deleteError = 'Google Calendar: ' + str(calendar) + ' Could not delete event: '  + str(eventId) + 'Exception:' + str(e)

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
		# Check for time ending in 'Z'
		if tzf == 'Z':
			tzf = '+00:00'
		tdate = date1[0:19] + tzf		
		
		datef = datetime.strptime(tdate,"%Y-%m-%dT%H:%M:%S%z")		
		if (convertToUTC == True):
			datef = datef.astimezone(timezone('UCT'))
		return datef
			
# Manage Dictionary of Slate Calendars
def calendarExists(calendar):
	if calendar in calendars:
		return True
	else:
		return False
		
def createCalendarUrl(id, signature):
	url = slateServer + '/manage/event/ical?cmd=feed&identity=' + id + '&user=' + id + '&signature=' + signature
	return url

def deleteCalendar(delete_calendar):
	if delete_calendar in calendars:
		with lock:
			del calendars[delete_calendar]
			f = open(calendar_list_file, 'w')
			json.dump(calendars, f, indent=4, sort_keys=True)
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
				
		if self.path.startswith('/sync'):
			# Initial page entered by user		
			flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES, redirect_uri=syncServerUrl)
			flow.user_agent = APPLICATION_NAME
			flow.params['access_type'] = 'offline'
			auth_uri = flow.step1_get_authorize_url()
			
			self.send_response(302)
			self.send_header('Location', auth_uri)
			self.end_headers()
			return
					

		elif self.path.startswith('/?error='):
			message = 'Error occured while requesting authorization from Google.'
			
		elif 'calendarlist' in self.path:
			message = ''
			for googleCalendar in sorted(calendars):
				message += googleCalendar + '<br />'
			
		elif 'code' in self.path:
			# Page redirected back from auth server
			auth_code = parameters['code'][0]
			flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES, redirect_uri=syncServerUrl)
			credentials = flow.step2_exchange(auth_code)
			
			http = credentials.authorize(httplib2.Http())
			user_info_service = discovery.build('oauth2', 'v2', http=http)
			
			user_info = None
			try:
				user_info = user_info_service.userinfo().get().execute()
				new_calendar = user_info.get('email')
				
				if not calendarExists(new_calendar):
				
					with lock:
						calendars[new_calendar] = {'eventColorOnCampus':'','eventColorOther':''}
						f = open(calendar_list_file, 'w')
						json.dump(calendars, f, indent=4, sort_keys=True)
						f.close()
											
						credential_file = new_calendar + '.json'
						credential_path = os.path.join(credential_dir, credential_file)
						storage = oauth2client.file.Storage(credential_path)
						storage.put(credentials)
						print ('Storing credentials to ', credential_path)
					
					message = 'Succesfully added calendar ' + new_calendar
				
				else:
					message = 'Calendar already exists: ' + new_calendar
				
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

	t_sync = threading.Thread(target=main)
	t_sync.daemon = False
	t_sync.start()

	while True:
		time.sleep (syncInterval)

		if t_sync.is_alive():
			logger.warning ('Sync: Prior thread still running. Do not kick off another sync.')
		else:
			t_sync = threading.Thread(target=main)
			t_sync.daemon = False
			t_sync.start()
			
if __name__ == '__main__':

	# Create lock object
	lock = threading.Lock()
		
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
				deleteEvent(service, googleApiBackoff, clear_calendar, eventDetails['eventId'])
			
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
