# outlook-automail
VBA script that: automatically sends out an email to specified recipients, with a specified document attached, every time a task reminder with a specified title fires.

This script is notable in that it uses ONLY VBA through Outlook to send out emails on a set schedule. 

At the time I wrote this, I was working a non-tech job and didn't have admin access to install something like Python to do this, so I had to work with what was available to me - namely, Office. If I were doing this today, I would (and do) use a combination of Python scripting, .bat files and Task Scheduler to accomplish this goal.

I couldn't find anything in VBA/Outlook to schedule this to execute on a schedule (i.e. every M-F at a specified time), so instead I used Outlook's task reminders to handle the schedule part. I tried to get it working with calendar events, but they didn't work reliably; plus, we were using the Calendar for other things and weren't using Tasks for anything. 

The major downside of using Task reminders is that if the user accidentally clicks Dismiss (or Dismiss All), the task will remain uncompleted and will not reoccur - which means no more notifications, which means the script won't fire. Again, nowadays I would use Python rather than jump through these kinds of hoop, but if you're in a similar situation like I was where you don't have admin access, it's possible you could have the script first check if the task exists, and if not create a task with the specified name and reoccurrence schedule. (Maybe I'll add that one day if I'm bored!)
