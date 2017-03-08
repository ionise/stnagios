# stnagios
vbScript to poll ShoreTel health and return it to a Nagios instance over http/https

I made a thing to do some basic monitoring on a ShoreTel using the NCPA plugin for Nagios.
Its been running for over a year quite happily on our internal system which was upgraded from ShoreTel 14.2 to ShoreTel Connect. I have yet to write some documentation on how to set it up on a system from scratch but I'll get the under way in due course.
I am looking for contributors to help me to improve and maintain this since software development isn't my main area of expertise. I'm a full time employed at a ShoreTel reseller so I don't have much time to work on this or give support. 

If you know about Nagios then this script can be called by the NCPA agent installed directly on the ShoreTel HQ server and it basically hooks into the "shorewarestatus" monitoring database used by Diagnostics and Monitoring so you can get similar information available to there, but in Nagios. https://www.nagios.org/ncpa/help.php
I configure Nagios to accept PASSIVE checks from the NCPA agent since Nagios doesnt have access to poll the ShoreTel server directly so the agent polls the ShoreTel server locally and then relays the results back to the Nagios server over http/https - the idea being that I might pro-actively monitor systems belonging to my customers.

There's a function in the script to take an inventory of your ShoreGear switches and build an INI file that can be used by the NCPA agent to execute checks against your hardware using the script.

Anyway, here's the repository as is. Again, I'm afraid I cannot offer much technical support, but keep an eye out of some further documentation on getting it running.
