# Solarview

**Solarview** is a Python3-application that produces a per year heatmap of solar PV energy production
using the Growatt server data

Reads the production data from the Growatt server (growatt.server.com)
using the GrowattShinePhone api, stores them locally.

Runs on **Windows**, **Linux** (RaspBian, Ubuntu) and **Android** (using PyDroid3)

Needs a **solarview.ini** file with the following contents:  
[ini]  
username=YourUserName  
password=YourPassword  
pickle_dir="./"  
pickle_template="solarviewdata_????.pkl"  

Dependencies:
requests
         
Uses:
GrowattApi: https://github.com/Sjord/growatt_api_client,
included in this file
         

