# Slate-Google Calendar Sync

David Glasser  
glasserd@union.edu  
Union College  
Schenectady, NY

## Description

This program reads events from the Slate calendar and syncs them to a Google calendar.

## Prerequisites

* Google API Credentials (https://console.developers.google.com/apis/)
* Python 3.6.x
* pip (Python package manager)
* Python modules:
    * httplib2
    * pytz
    * urllib3
    * google-api-python-client
    * oauth2client
    * requests
 
## Setup

1. Clone this repository. The most recent stable tag is 1.6.3. It is highly recommended that you clone using this tag. Development commits are stored in the master branch, using the latest commit from this branch is at your own risk. After you clone the repository you can get to a specific release with these commands:
    ```
    git fetch --all --tags --prune
    git checkout tags/1.6.3
    ```
2. Create a query in Slate that will feed your calendar events. You should use the configurable join form query base.

    Web Service Settings
    - Use the "Edit Web Service" function to turn the query into a web service
    - Service Type = JSON
    - Parameter:
        ```
        <param id="calendar" type="varchar" />
        ```
    - Set a username and complex password

    Exports **(names must match exactly)**:
    - GUID
    - Type
    - Title
    - Start (Default Format)
    - End (Default Format)
    - Timezone Offset (Subquery export with Form Street, Form City, Form Region)
    - Location
    - LocationKey
    - Address
    - Description (Notes field)
    - Attendees (Join Form Response - Registration Status in Attended or Registered)
    - Interviewee (Join Form Response and Person - Export First and Last Name. Filter on type = Interview and Registration Status in Attended or Registered)

    Filters:
    - Type: Event, Interview
    - Start Date: Today - 15  (This should roughly correspond to the NumberOfPriorDays variable in config.ini)
    - Status: Active/Confirmed
    - Use custom SQL to join user email to @calendar

    Joins:
    - Join Users to Forms

3. Make a query in Slate using the custom SQL query base. This query will pull your stop data. Sample SQL is provided below. As of October 2019 stop data is not available through a CJ query base.

    ```
    SELECT 
        ts.[id] AS 'GUID'
        ,'Stop' AS 'Type'
        ,ts.[name] AS [Title]
    ,ts.[dtstart] AS 'Start'
    ,ts.[dtend] AS 'End'
        ,ts.[timezone_offset] as 'Timezone Offset'     
    ,ts.[location] as 'Location'	
    ,char(10) + isnull(ts.[street] + char(10),'') + isnull(ts.[city]  + char(10),'') + isnull(ts.[region],'') AS 'Address'
        ,convert(nvarchar(max), ts.[notes].query('//text()')) as [Description]
        , '' AS 'Attendees'	
    FROM [trip.stop] ts
    LEFT OUTER JOIN [trip] t ON ts.[trip] = t.[id]
    LEFT OUTER JOIN [user] u ON t.[user] = u.[id]
    WHERE u.[email] = @calendar
        AND	
            (convert(date, ts.[dtstart]) >= dateadd(day, -15, convert(date, getdate())))
    ```

4. Make a copy of `config.ini.example` and name it `config.ini`. Populate this file with values for your enviornment.
5. Download Google API Secret and save in same directory as `slatesync.py`. File name should be `client_secret.json`


## Usage

To run the program:
```
python slatesync.py
```

Ideally this should be set up as a service on a server.

To set up a new calendar you need to browse to https://yourserveraddress/sync and follow the prompts.

A list of currently synced calendars can be found at https://yourserveraddress/calendarlist


Most interactions will be through the web interface. However there are some command line options that are available:
	
To print a list of available commands:
```
python slatesync.py -h
```

To delete a calendar that is currently being synced:
```
python slatesync.py -d email_address
```

## Known Issues
- None
## Notes
		
Please send bugs and feature requests to glasserd@union.edu. There is no guarantee your need will be addressed, but I will consider all requests.