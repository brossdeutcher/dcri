# email-handler


Assists with Microoft Outlook email client.

First version creates drafts based on data in a CSV file. Set the name, to, cc, subject, body, & attachments of each report preset; and pass an integer associated with the desired report to create a draft for that report.

For attachments with a date, insert %s into the string representing the file path under that Attachments column. Write the date variables associated with each %s of Attachments, in order, spaced by ', ' under Attachment Dates.
Below are the usable date variables for attachment date:
  today - current date, in yyyy-mm-dd format
  curMDY - concatinated current date, in mmddyyyy format
  curDay - number of day of month, with leading zero
  curMonInt -  number of month, with leading zero
  curMonLet - first three letters of current month, ex. 'Nov'
  curYr - number of current year
