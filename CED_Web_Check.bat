:: *****************************************************************************
:: CED_Web_Check.bat  9/8/2017 
:: Summary: CED Web Page Monitor
:: Author: Mike Murnane
::
:: Description: Monitor GADS web pages and sends email of results. 
::
:: Scheduled Task - Every day @ 7:35 am.
::
:: Path: \\Geobase-win\CED\GADS\R2017\R426\ScheduledTask\CED_Web_Check.bat
:: *****************************************************************************

:: Set log directory for process verification file
    SET LogDir=\\Geobase-win\CED\GADS\R2017\R426\ScheduledTask\log\

:: Set variable %theDate% to today's date (YYYYMMDD)
     for /f "tokens=2,3,4 delims=/ " %%a in ('date/t') do set theDate=%%c%%a%%b

:: Record starting time
Echo. > %LogDir%%theDate%.log
Echo ============================================================================  >> %LogDir%%theDate%.log
Echo GADS Web Page Report  >> %LogDir%%theDate%.log
Echo ============================================================================  >> %LogDir%%theDate%.log
Echo. >> %LogDir%%theDate%.log 2>&1
Echo Start time:  >> %LogDir%%theDate%.log
 time /T >> %LogDir%%theDate%.log

::  Echo. = blank line added
Echo. >> %LogDir%%theDate%.log 2>&1

:: Check pages and send Email with log file content
Echo Checking each GADS web page ...  >> %LogDir%%theDate%.log
cscript \\Geobase-win\CED\GADS\R2017\R426\ScheduledTask\CED_Web_Check.vbs %LogDir%%theDate%.log

:: Record stop time
Echo. >> %LogDir%%theDate%.log 2>&1
Echo Stop time:  >> %LogDir%%theDate%.log
 time /T >> %LogDir%%theDate%.log

::pause