@echo off
echo Create scheduled task CTIC\ModelNotify: Mon-Fri 15:00 and 17:00 (LINE model report)
schtasks /create /f /tn "CTIC\ModelNotify" /tr "\"C:\Users\User\CTIC_Dining_Prediction\task_model_notify.bat\"" /sc weekly /d MON,TUE,WED,THU,FRI /st 15:00 /du 02:05 /ri 120 /it
schtasks /query /tn "CTIC\ModelNotify"
