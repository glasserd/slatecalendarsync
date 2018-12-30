# Slate-Google Calendar Sync

David Glasser  
glasserd@union.edu  
Union College  
Schenectady, NY

## Description

This program reads events from the Slate calendar and syncs them to a Google calendar.

## Prerequisites

* Google API Credentials (https://console.developers.google.com/apis/)
* Python 3.x
* pip (Python package manager)
* Python modules:
  * httplib2
  * pytz
  * urllib3
  * google-api-python-client
  * requests
 
## Setup

1. Download Google API Secret and save in same directory as `slatesync.py`. File name should be `client_secret.json`
2. Update config.ini as desired

## Usage

To run the program:
```
python slatesync.py
```
To set up a new calendar you need to call a URL with specific parameters:
	
Parameter: calendar  
Value: User's email address
	
Parameter: id  
Value: The guid of the user in slate
	
You may wish to create a query in Slate using custom SQL. This way users can self serve by running the query and clicking the link. An example query is below.

```
SELECT 'https://yourserveraddress/calendar=' + cast([email] as varchar(50)) + '&id=' + cast([id] as varchar(50)) as [link]
FROM [user]
WHERE id = @user
```
Most interactions will be through the web interface. However there are some command line options that are available:
	
To print a list of available commands:
```
python slatesync.py -h
```

To add a calendar to be synced:
```
python slatesync.py -a email_address
```

To delete a calendar that is currently being synced:
```
python slatesync.py -d email_address
```

## Notes
		
	Please send bugs and feature requests to glasserd@union.edu. There is no guarantee your need will be addressed, but I will consider all requests.