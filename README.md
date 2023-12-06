# GUI-Python

A large energy company provides a call centre to answer customers’ queries. The call centre management
is keen to understand the activity and performance levels that the call centre has experienced over the last
winter and what they might do to improve their experiences next year. They know that activity levels vary
considerably during the day and so data has been collected for the morning, afternoon and evening ‘peak’
hours. They are hoping to understand their workloads better, and also how they are related to the experience
of their customers when calling the call centre.


They are particularly interested in reducing the number of callers who abandon if this is possible. During the
last Winter they have been experimenting with virtual hold technology (VHT). When VHT is ‘on’ callers who
cannot be dealt with immediately are offered the chance to be rung back without losing their place in the
queue. Is there any evidence that this reduces the chance that callers abandon their call?


The Excel data file ‘EnergyCallCentre.xlsx’ contains the following data for 504 ‘peak’ hours of operation from
last Winter:

• Month – refers to three 2-month long periods of time;
• VHT – specifies whether Virtual Hold Technology was on or off;
• ToD – specifies which peak hour the data is from;
• Agents – an indication of the number of agents available to handle calls during the hour (not very
accurate);
• CallsOffered – number of callers calling during the hour;
• CallsAbandoned – number of calls arriving during the hour which rang off before speaking to an agent;
• CallsHandled – number of calls arriving during the hour which spoke to an agent;
• ASA – average speed of answer (minutes), i.e. average time between the caller first ringing the call
centre and speaking to an agents;
• Avehandletime – the average time that calls require from an agent, including ‘wrap-up’ time.


Draw a random sample of size 100 from this population of one hour periods and investigate your sample
using Python.
